"""Shared HTTP transport for Microsoft Graph v1.0."""

from __future__ import annotations

import math
import time
from collections.abc import Callable, Mapping
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import Any
from urllib.parse import unquote, urljoin, urlsplit

import requests

from ._version import __version__
from .auth import GRAPH_SCOPE, GraphAuthenticator, TokenCredential
from .exceptions import (
    GraphAuthenticationError,
    GraphBridgeError,
    GraphConflictError,
    GraphGoneError,
    GraphInvalidResponseError,
    GraphNetworkError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphPreconditionFailedError,
    GraphRequestError,
    GraphServerError,
    GraphThrottlingError,
)
from .models import GraphError

DEFAULT_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_TIMEOUT = (3.05, 30.0)
DEFAULT_USER_AGENT = f"GraphBridge/{__version__}"
DEFAULT_MAX_RETRY_DELAY = 120.0
MAX_RETRIES = 10
_IDEMPOTENT_METHODS = frozenset({"GET", "HEAD", "OPTIONS", "PUT", "DELETE"})


class GraphTransport:
    """Send authenticated requests through one reusable HTTP session.

    The transport centralizes authentication, timeouts, retry policy, safe URL
    resolution, response decoding, token redaction, and typed HTTP errors.

    Args:
        credential: Credential used to acquire Microsoft Graph tokens.
        session: Optional HTTP session reused for all requests.
        base_url: Microsoft Graph v1.0 base URL.
        timeout: Connection and response timeout.
        user_agent: Value sent in the User-Agent header.
        max_retries: Maximum retries for eligible requests.
        backoff_factor: Base factor used for exponential backoff.
        max_retry_delay: Maximum delay between retries.
        sleep: Function used to wait between retries.
        scope: OAuth scope requested from the credential.
    """

    def __init__(
        self,
        credential: TokenCredential,
        *,
        session: requests.Session | None = None,
        base_url: str = DEFAULT_BASE_URL,
        timeout: float | tuple[float, float] = DEFAULT_TIMEOUT,
        user_agent: str = DEFAULT_USER_AGENT,
        max_retries: int = 3,
        backoff_factor: float = 0.5,
        max_retry_delay: float = DEFAULT_MAX_RETRY_DELAY,
        sleep: Callable[[float], None] = time.sleep,
        scope: str = GRAPH_SCOPE,
    ) -> None:
        """Initialize and validate the HTTP transport.

        Configuration is rejected eagerly so every later request has a finite
        timeout, bounded retry behavior, and a safe Microsoft Graph v1.0 root.

        Args:
            credential: Credential used to acquire access tokens.
            session: Optional reusable HTTP session.
            base_url: Microsoft Graph v1.0 base URL.
            timeout: Connection and response timeout.
            user_agent: Value sent in the User-Agent header.
            max_retries: Maximum retries for eligible requests.
            backoff_factor: Base factor used for retry delays.
            max_retry_delay: Maximum delay between retries.
            sleep: Function used to wait between retries.
            scope: OAuth scope requested from the credential.

        Raises:
            TypeError: If a numeric or timeout option has an invalid type.
            ValueError: If a URL, timeout, or retry option is outside its allowed range.
        """
        if isinstance(max_retries, bool) or not isinstance(max_retries, int):
            raise TypeError("max_retries must be an integer")
        if max_retries < 0:
            raise ValueError("max_retries must be non-negative")
        if max_retries > MAX_RETRIES:
            raise ValueError(f"max_retries cannot exceed {MAX_RETRIES}")
        if isinstance(backoff_factor, bool) or not isinstance(backoff_factor, (int, float)):
            raise TypeError("backoff_factor must be a number")
        if not math.isfinite(backoff_factor) or backoff_factor < 0:
            raise ValueError("backoff_factor must be finite and non-negative")
        if isinstance(max_retry_delay, bool) or not isinstance(max_retry_delay, (int, float)):
            raise TypeError("max_retry_delay must be a number")
        if not math.isfinite(max_retry_delay) or max_retry_delay < 0:
            raise ValueError("max_retry_delay must be finite and non-negative")
        self.authenticator = GraphAuthenticator(credential, scope=scope)
        self.session = session if session is not None else requests.Session()
        self.base_url = self._validated_base_url(base_url)
        self.timeout = self._validated_timeout(timeout)
        self.user_agent = user_agent
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.max_retry_delay = float(max_retry_delay)
        self._sleep = sleep

    def request(
        self,
        method: str,
        url: str,
        *,
        params: Mapping[str, Any] | None = None,
        json: Any = None,
        headers: Mapping[str, str] | None = None,
        retry: bool | None = None,
    ) -> Any:
        """Send one authenticated Microsoft Graph request.

        A fresh token is requested for every attempt. Safe methods retry eligible
        network, throttling, and server failures within the configured budget;
        non-idempotent methods are not replayed unless explicitly requested.

        Args:
            method: HTTP method to send.
            url: Relative or permitted absolute Graph URL.
            params: Optional query parameters.
            json: Optional JSON request body.
            headers: Optional additional request headers.
            retry: Optional override for automatic retry eligibility.

        Returns:
            The decoded JSON response, or ``None`` for an empty success.

        Raises:
            GraphAuthenticationError: If an access token cannot be acquired.
            GraphInvalidResponseError: If the URL or successful response is invalid.
            GraphNetworkError: If the network request ultimately fails.
            GraphRequestError: If Microsoft Graph returns a non-success status.
        """

        normalized_method = method.upper()
        request_url = self._resolve_url(url)
        retry_allowed = normalized_method in _IDEMPOTENT_METHODS if retry is None else retry
        attempt = 0

        while True:
            token = self.authenticator.get_access_token()
            request_headers = {
                "Accept": "application/json",
                "User-Agent": self.user_agent,
            }
            if json is not None:
                request_headers["Content-Type"] = "application/json"
            if headers:
                request_headers.update(headers)
            request_headers["Authorization"] = f"Bearer {token}"

            try:
                response = self.session.request(
                    normalized_method,
                    request_url,
                    params=params,
                    json=json,
                    headers=request_headers,
                    timeout=self.timeout,
                )
            except requests.RequestException:
                if retry_allowed and attempt < self.max_retries:
                    self._sleep(self._limited_delay(self._backoff_delay(attempt)))
                    attempt += 1
                    continue
                raise GraphNetworkError("The Microsoft Graph request could not be completed") from None

            if 200 <= response.status_code < 300:
                return self._parse_success(response, token=token)

            retry_after = self._retry_after_seconds(response.headers.get("Retry-After"))
            retriable_status = response.status_code == 429 or response.status_code >= 500
            if retry_allowed and retriable_status and attempt < self.max_retries:
                delay = retry_after
                if delay is None:
                    delay = self._backoff_delay(attempt)
                self._sleep(self._limited_delay(delay))
                attempt += 1
                continue

            raise self._build_http_error(response, token=token, retry_after=retry_after)

    def get(self, url: str, **kwargs: Any) -> Any:
        """Send a GET request through the shared request pipeline.

        GET is retry-eligible by default because it is treated as idempotent.

        Args:
            url: Relative or permitted absolute Graph URL.
            **kwargs: Options forwarded to :meth:`request`.
        """
        return self.request("GET", url, **kwargs)

    def post(self, url: str, **kwargs: Any) -> Any:
        """Send a POST request through the shared request pipeline.

        POST is not retried automatically by default to avoid duplicating writes.

        Args:
            url: Relative or permitted absolute Graph URL.
            **kwargs: Options forwarded to :meth:`request`.
        """
        return self.request("POST", url, **kwargs)

    def patch(self, url: str, **kwargs: Any) -> Any:
        """Send a PATCH request through the shared request pipeline.

        PATCH is not retried automatically by default because replay safety
        depends on the operation and its concurrency conditions.

        Args:
            url: Relative or permitted absolute Graph URL.
            **kwargs: Options forwarded to :meth:`request`.
        """
        return self.request("PATCH", url, **kwargs)

    def put(self, url: str, **kwargs: Any) -> Any:
        """Send a PUT request through the shared request pipeline.

        PUT is considered idempotent and is retry-eligible by default.

        Args:
            url: Relative or permitted absolute Graph URL.
            **kwargs: Options forwarded to :meth:`request`.
        """
        return self.request("PUT", url, **kwargs)

    def delete(self, url: str, **kwargs: Any) -> Any:
        """Send a DELETE request through the shared request pipeline.

        DELETE is retry-eligible by default and can include caller-provided
        conditional headers such as ``If-Match``.

        Args:
            url: Relative or permitted absolute Graph URL.
            **kwargs: Options forwarded to :meth:`request`.
        """
        return self.request("DELETE", url, **kwargs)

    def close(self) -> None:
        """Close the underlying HTTP session.

        Closing releases connection-pool resources owned by the injected or
        internally created session.
        """
        self.session.close()

    def __enter__(self) -> GraphTransport:
        """Return the transport when entering a context manager."""
        return self

    def __exit__(self, *_args: object) -> None:
        """Close the transport when leaving a context manager.

        Args:
            *_args: Context-manager exception details, if any.
        """
        self.close()

    def __repr__(self) -> str:
        """Return a representation containing only safe configuration."""
        return f"GraphTransport(base_url={self.base_url!r}, timeout={self.timeout!r}, max_retries={self.max_retries!r})"

    def _resolve_url(self, url: str) -> str:
        """Resolve and validate a request URL.

        Absolute continuation links are accepted only when they remain on the
        configured origin and below its v1.0 path, preventing token forwarding.

        Args:
            url: Relative or absolute URL to validate.

        Raises:
            GraphInvalidResponseError: If the URL escapes the configured v1.0 root.
        """
        target = urlsplit(url)
        resolved = url if target.scheme or target.netloc else urljoin(f"{self.base_url}/", url.lstrip("/"))
        configured = urlsplit(self.base_url)
        target = urlsplit(resolved)
        configured_path = configured.path.rstrip("/")
        same_origin = (
            target.scheme.casefold(),
            target.netloc.casefold(),
        ) == (
            configured.scheme.casefold(),
            configured.netloc.casefold(),
        )
        target_path = unquote(target.path)
        within_base_path = target_path == configured_path or target_path.startswith(f"{configured_path}/")
        contains_dot_segment = any(part in {".", ".."} for part in target_path.split("/"))
        if not same_origin or not within_base_path or contains_dot_segment:
            raise GraphInvalidResponseError("Microsoft Graph URLs must remain under the configured base URL (v1.0)")
        return resolved

    def _parse_success(self, response: requests.Response, *, token: str) -> Any:
        """Decode and redact a successful response.

        Empty HTTP successes remain distinct from an empty JSON object. Any
        occurrence of the current access token is removed recursively.

        Args:
            response: Successful HTTP response.
            token: Access token that must be redacted.

        Raises:
            GraphInvalidResponseError: If a non-empty response is not valid JSON.
        """
        if response.status_code in {204, 205} or not response.content:
            return None
        try:
            return self._redact_value(response.json(), token)
        except ValueError:
            raise GraphInvalidResponseError("Microsoft Graph returned invalid JSON for a successful response") from None

    def _build_http_error(
        self,
        response: requests.Response,
        *,
        token: str,
        retry_after: float | None,
    ) -> GraphBridgeError:
        """Convert a failed HTTP response into a typed exception.

        Graph error metadata and response headers are preserved after redaction,
        and common status codes map to specialized exception subclasses.

        Args:
            response: Failed HTTP response.
            token: Access token that must be redacted.
            retry_after: Parsed retry delay, when available.
        """
        raw_text = self._redact(response.text, token)
        payload: Any = None
        try:
            payload = response.json()
        except ValueError:
            pass

        error_payload: Mapping[str, Any] = {}
        if isinstance(payload, Mapping):
            candidate = payload.get("error")
            if isinstance(candidate, Mapping):
                error_payload = candidate
        raw_inner_error = error_payload.get("innerError", error_payload.get("innererror", {}))
        if not isinstance(raw_inner_error, Mapping):
            raw_inner_error = {}
        inner_error = self._redact_value(dict(raw_inner_error), token)
        if not isinstance(inner_error, Mapping):  # pragma: no cover - mapping input stays a mapping
            inner_error = {}
        message = error_payload.get("message")
        safe_message = (
            self._redact(str(message), token) if message else f"Microsoft Graph returned HTTP {response.status_code}"
        )
        code = error_payload.get("code")
        graph_error = GraphError(
            code=self._redact(str(code), token) if code is not None else "httpError",
            message=safe_message,
            status_code=response.status_code,
            request_id=self._string_or_none(inner_error.get("request-id", inner_error.get("requestId"))),
            date=self._string_or_none(inner_error.get("date")),
            inner_error=dict(inner_error),
        )

        if response.status_code == 401:
            return GraphAuthenticationError(
                graph_error.message,
                status_code=response.status_code,
                response_text=raw_text,
                error=graph_error,
            )
        error_type: type[GraphRequestError]
        if response.status_code == 403:
            error_type = GraphPermissionError
        elif response.status_code == 404:
            error_type = GraphNotFoundError
        elif response.status_code == 409:
            error_type = GraphConflictError
        elif response.status_code == 410:
            error_type = GraphGoneError
        elif response.status_code == 412:
            error_type = GraphPreconditionFailedError
        elif response.status_code == 429:
            error_type = GraphThrottlingError
        elif response.status_code >= 500:
            error_type = GraphServerError
        else:
            error_type = GraphRequestError
        return error_type(
            graph_error,
            response_text=raw_text,
            retry_after=retry_after,
            response_headers=self._redact_value(dict(response.headers), token),
        )

    @staticmethod
    def _validated_base_url(value: str) -> str:
        """Validate and normalize a Graph v1.0 base URL.

        The URL must be HTTPS, contain no credentials or query data, and terminate
        at a v1.0 root used to confine later absolute links.

        Args:
            value: Base URL to validate.

        Raises:
            ValueError: If the URL is not a safe HTTPS v1.0 root.
        """
        if not isinstance(value, str) or not value:
            raise ValueError("base_url cannot be empty")
        parsed = urlsplit(value)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("base_url must be an absolute HTTPS URL")
        if parsed.username is not None or parsed.password is not None:
            raise ValueError("base_url must not contain credentials")
        if parsed.query or parsed.fragment:
            raise ValueError("base_url must not contain a query or fragment")
        normalized = value.rstrip("/")
        if not urlsplit(normalized).path.rstrip("/").endswith("/v1.0"):
            raise ValueError("base_url must target a Microsoft Graph v1.0 root")
        return normalized

    @staticmethod
    def _validated_timeout(
        value: float | tuple[float, float],
    ) -> float | tuple[float, float]:
        """Validate an HTTP timeout value.

        Both scalar timeouts and separate connect/read pairs are supported, but
        every component must be finite and strictly positive.

        Args:
            value: Positive scalar or connect/read timeout pair.

        Raises:
            TypeError: If the timeout shape or types are invalid.
            ValueError: If any timeout is non-finite or not positive.
        """
        values: tuple[object, ...] = value if isinstance(value, tuple) else (value,)
        if len(values) not in {1, 2}:
            raise TypeError("timeout must be a positive number or a two-number tuple")
        for item in values:
            if isinstance(item, bool) or not isinstance(item, (int, float)):
                raise TypeError("timeout must be a positive number or a two-number tuple")
            if not math.isfinite(item) or item <= 0:
                raise ValueError("timeout values must be finite and greater than zero")
        return value

    def _limited_delay(self, value: float) -> float:
        """Cap a retry delay at the configured maximum.

        Args:
            value: Proposed delay in seconds.
        """
        return min(value, self.max_retry_delay)

    def _backoff_delay(self, attempt: int) -> float:
        """Calculate exponential backoff for an attempt.

        Args:
            attempt: Zero-based retry attempt number.
        """
        return self.backoff_factor * (2**attempt)

    @staticmethod
    def _retry_after_seconds(value: str | None) -> float | None:
        """Parse a Retry-After header into seconds.

        Args:
            value: Header value as seconds or an HTTP date.
        """
        if value is None:
            return None
        try:
            return max(0.0, float(value))
        except ValueError:
            try:
                retry_at = parsedate_to_datetime(value)
            except (TypeError, ValueError, OverflowError):
                return None
            if retry_at.tzinfo is None:
                retry_at = retry_at.replace(tzinfo=timezone.utc)
            return max(0.0, (retry_at - datetime.now(timezone.utc)).total_seconds())

    @staticmethod
    def _redact(value: str, token: str) -> str:
        """Replace an access token in a string.

        Args:
            value: Text that may contain the token.
            token: Sensitive token to replace.
        """
        return value.replace(token, "<redacted>") if token else value

    @classmethod
    def _redact_value(cls, value: Any, token: str) -> Any:
        """Recursively redact an access token from JSON-like data.

        Args:
            value: Value to sanitize.
            token: Sensitive token to replace.
        """
        if isinstance(value, str):
            return cls._redact(value, token)
        if isinstance(value, Mapping):
            return {cls._redact(str(key), token): cls._redact_value(item, token) for key, item in value.items()}
        if isinstance(value, list):
            return [cls._redact_value(item, token) for item in value]
        return value

    @staticmethod
    def _string_or_none(value: Any) -> str | None:
        """Convert a value to text while preserving ``None``.

        Args:
            value: Value to convert.
        """
        return str(value) if value is not None else None

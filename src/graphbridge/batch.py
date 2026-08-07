"""Generic Microsoft Graph v1.0 JSON batch execution and retry primitives."""

from __future__ import annotations

import math
import time
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import TYPE_CHECKING, Any, TypeVar

from .exceptions import GraphInvalidResponseError
from .models import GraphError

if TYPE_CHECKING:
    from .transport import GraphTransport

T = TypeVar("T")
MAX_BATCH_REQUESTS = 20
MAX_BATCH_ATTEMPTS = 11
DEFAULT_MAX_RETRY_DELAY = 120.0


@dataclass(frozen=True, slots=True)
class BatchRequest:
    """Represent one request in a Microsoft Graph JSON batch.

    Subrequests use paths relative to the configured v1.0 root and carry their
    own correlation ID, optional headers, body, and original input position.

    Args:
        id: Identifier used to correlate the subrequest.
        method: HTTP method for the subrequest.
        url: URL relative to the configured v1.0 root.
        headers: Optional subrequest headers.
        body: Optional JSON-compatible request body.
        input_index: Optional position in the caller's original input.
    """

    id: str
    method: str
    url: str
    headers: dict[str, str] = field(default_factory=dict, repr=False)
    body: Any = field(default=None, repr=False)
    input_index: int | None = None

    def to_payload(self) -> dict[str, Any]:
        """Serialize the subrequest for the Graph batch endpoint.

        Validation prevents beta or absolute URLs and rejects embedded bearer
        authorization because authentication belongs to the outer batch call.

        Returns:
            A JSON-compatible subrequest mapping.

        Raises:
            ValueError: If the identifier, method, URL, or headers are unsafe.
        """
        if not self.id:
            raise ValueError("batch request id cannot be empty")
        if not self.url.startswith("/") or self.url.startswith("//") or self.url.casefold().startswith("/beta/"):
            raise ValueError("batch request URLs must be relative to the configured v1.0 root")
        method = self.method.upper()
        if method not in {"GET", "POST", "PATCH", "PUT", "DELETE"}:
            raise ValueError("unsupported Microsoft Graph batch request method")
        payload: dict[str, Any] = {
            "id": self.id,
            "method": method,
            "url": self.url,
        }
        headers = dict(self.headers)
        if any(name.casefold() == "authorization" for name in headers):
            raise ValueError("batch subrequests must not contain Authorization headers")
        if self.body is not None and not any(name.lower() == "content-type" for name in headers):
            headers["Content-Type"] = "application/json"
        if headers:
            payload["headers"] = headers
        if self.body is not None:
            payload["body"] = self.body
        return payload


@dataclass(frozen=True, slots=True)
class BatchResponse:
    """Represent one batch response and its original request.

    Correlation is retained even when Microsoft Graph returns subresponses in a
    different order from the submitted requests.

    Args:
        request: Original correlated batch request.
        status_code: HTTP status returned for the subrequest.
        headers: Response headers for the subrequest.
        body: Decoded response body.
        attempts: Number of attempts used for the subrequest.
    """

    request: BatchRequest
    status_code: int
    headers: Mapping[str, Any] = field(default_factory=dict, repr=False)
    body: Any = field(default=None, repr=False)
    attempts: int = 1

    @property
    def succeeded(self) -> bool:
        """Return whether the subrequest completed successfully.

        Any HTTP status in the inclusive 200-299 range is treated as success.
        """
        return 200 <= self.status_code < 300

    def to_error(self) -> GraphError:
        """Convert a failed subresponse into a structured Graph error.

        The generated error includes correlation and attempt metadata so partial
        batch failures can be traced back to their original inputs.
        """
        error_payload: Mapping[str, Any] = {}
        if isinstance(self.body, Mapping) and isinstance(self.body.get("error"), Mapping):
            error_payload = self.body["error"]
        code = error_payload.get("code", "batchError")
        message = error_payload.get("message", f"Microsoft Graph batch request returned HTTP {self.status_code}")
        inner_error: dict[str, Any] = {
            "batch_request_id": self.request.id,
            "attempts": self.attempts,
        }
        if self.request.input_index is not None:
            inner_error["input_index"] = self.request.input_index
        raw_inner = error_payload.get("innerError", error_payload.get("innererror"))
        if isinstance(raw_inner, Mapping):
            inner_error["graph_inner_error"] = dict(raw_inner)
        return GraphError(
            code=str(code),
            message=str(message),
            status_code=self.status_code,
            inner_error=inner_error,
        )


def chunks(values: Sequence[T], size: int = MAX_BATCH_REQUESTS) -> Iterator[Sequence[T]]:
    """Split values into Graph-compatible batch chunks.

    Microsoft Graph accepts at most twenty requests per JSON batch, so larger
    sequences are divided without changing their order.

    Args:
        values: Ordered values to split.
        size: Maximum number of values per chunk.

    Raises:
        ValueError: If the chunk size is outside the Graph batch limit.
    """
    if not 1 <= size <= MAX_BATCH_REQUESTS:
        raise ValueError(f"batch size must be between 1 and {MAX_BATCH_REQUESTS}")
    for index in range(0, len(values), size):
        yield values[index : index + size]


def batch_payload(requests: Sequence[BatchRequest]) -> dict[str, list[dict[str, Any]]]:
    """Build the payload for one Graph batch call.

    Request IDs are validated case-insensitively because response correlation
    treats IDs the same way.

    Args:
        requests: Subrequests to include in the batch.

    Raises:
        ValueError: If the batch is too large or request IDs are not unique.
    """
    if len(requests) > MAX_BATCH_REQUESTS:
        raise ValueError(f"a Microsoft Graph batch accepts at most {MAX_BATCH_REQUESTS} requests")
    _validate_unique_ids(requests)
    return {"requests": [request.to_payload() for request in requests]}


def execute_batch(
    transport: GraphTransport,
    requests: Sequence[BatchRequest],
    *,
    max_attempts: int = 3,
    backoff_factor: float = 0.5,
    max_retry_delay: float = DEFAULT_MAX_RETRY_DELAY,
    sleep: Callable[[float], None] = time.sleep,
) -> list[BatchResponse]:
    """Execute batches and retry only transient subrequests.

    Inputs are chunked at twenty requests, correlated by ID, and returned in
    original order. Only HTTP 408, 429, and 5xx subrequests are replayed; already
    successful operations are never submitted again.

    Args:
        transport: Transport used to call the Graph batch endpoint.
        requests: Ordered subrequests to execute.
        max_attempts: Maximum attempts per transient subrequest.
        backoff_factor: Base factor used for retry delays.
        max_retry_delay: Maximum delay between retry waves.
        sleep: Function used to wait between retry waves.

    Returns:
        Responses correlated in original input order.

    Raises:
        TypeError: If an attempt or delay option has an invalid type.
        ValueError: If an option is outside its allowed range.
        GraphInvalidResponseError: If Graph returns an invalid batch response.
    """

    if isinstance(max_attempts, bool) or not isinstance(max_attempts, int):
        raise TypeError("max_attempts must be an integer")
    if max_attempts < 1:
        raise ValueError("max_attempts must be at least one")
    if max_attempts > MAX_BATCH_ATTEMPTS:
        raise ValueError(f"max_attempts cannot exceed {MAX_BATCH_ATTEMPTS}")
    if isinstance(backoff_factor, bool) or not isinstance(backoff_factor, (int, float)):
        raise TypeError("backoff_factor must be a number")
    if not math.isfinite(backoff_factor) or backoff_factor < 0:
        raise ValueError("backoff_factor must be finite and non-negative")
    if isinstance(max_retry_delay, bool) or not isinstance(max_retry_delay, (int, float)):
        raise TypeError("max_retry_delay must be a number")
    if not math.isfinite(max_retry_delay) or max_retry_delay < 0:
        raise ValueError("max_retry_delay must be finite and non-negative")
    _validate_unique_ids(requests)
    completed: dict[str, BatchResponse] = {}

    for request_chunk in chunks(requests):
        pending = {_normalized_id(request.id): request for request in request_chunk}
        attempts = {request_id: 0 for request_id in pending}

        while pending:
            active = list(pending.values())
            for request in active:
                attempts[_normalized_id(request.id)] += 1
            payload = transport.post(
                "/$batch",
                json=batch_payload(active),
                retry=False,
            )
            received = _parse_batch_responses(payload, pending)
            retrying: dict[str, BatchRequest] = {}
            delays: list[float] = []

            for request_id, request in pending.items():
                raw = received[request_id]
                status = raw["status"]
                headers = raw["headers"]
                response = BatchResponse(
                    request=request,
                    status_code=status,
                    headers=headers,
                    body=raw.get("body"),
                    attempts=attempts[request_id],
                )
                if _is_transient(status) and attempts[request_id] < max_attempts:
                    retrying[request_id] = request
                    retry_after = _retry_after_seconds(_header(headers, "Retry-After"))
                    delays.append(
                        retry_after if retry_after is not None else backoff_factor * (2 ** (attempts[request_id] - 1))
                    )
                else:
                    completed[request_id] = response

            if retrying:
                sleep(min(max(delays, default=0.0), max_retry_delay))
            pending = retrying

    return [completed[_normalized_id(request.id)] for request in requests]


def _parse_batch_responses(payload: Any, expected: Mapping[str, BatchRequest]) -> dict[str, dict[str, Any]]:
    """Validate and index raw batch responses by request ID.

    Args:
        payload: Decoded Graph batch response.
        expected: Expected subrequests indexed by normalized ID.

    Raises:
        GraphInvalidResponseError: If response correlation is invalid.
    """
    if not isinstance(payload, Mapping) or not isinstance(payload.get("responses"), list):
        raise GraphInvalidResponseError("Microsoft Graph batch response must contain responses")
    parsed: dict[str, dict[str, Any]] = {}
    for response in payload["responses"]:
        if not isinstance(response, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph batch responses must be JSON objects")
        raw_id = response.get("id")
        status = response.get("status")
        if raw_id is None or not isinstance(status, int) or isinstance(status, bool):
            raise GraphInvalidResponseError("Microsoft Graph returned an invalid batch result")
        request_id = _normalized_id(str(raw_id))
        if request_id not in expected:
            raise GraphInvalidResponseError("Microsoft Graph returned an unknown batch request id")
        if request_id in parsed:
            raise GraphInvalidResponseError("Microsoft Graph returned a duplicate batch request id")
        raw_headers = response.get("headers", {})
        if not isinstance(raw_headers, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph returned invalid batch response headers")
        parsed[request_id] = {
            "status": status,
            "headers": dict(raw_headers),
            "body": response.get("body"),
        }
    missing = set(expected) - set(parsed)
    if missing:
        raise GraphInvalidResponseError("Microsoft Graph omitted one or more batch responses")
    return parsed


def _validate_unique_ids(requests: Sequence[BatchRequest]) -> None:
    """Ensure request IDs are unique case-insensitively.

    Args:
        requests: Batch requests to validate.

    Raises:
        ValueError: If an identifier is duplicated.
    """
    normalized = [_normalized_id(request.id) for request in requests]
    if len(normalized) != len(set(normalized)):
        raise ValueError("batch request ids must be unique (case-insensitive)")


def _normalized_id(value: str) -> str:
    """Normalize a request identifier.

    Args:
        value: Identifier to normalize.
    """
    return value.casefold()


def _is_transient(status: int) -> bool:
    """Return whether a status is eligible for retry.

    Args:
        status: HTTP status code.
    """
    return status in {408, 429} or 500 <= status <= 599


def _header(headers: Mapping[str, Any], name: str) -> str | None:
    """Read a header case-insensitively.

    Args:
        headers: Response header mapping.
        name: Header name to retrieve.
    """
    expected = name.casefold()
    for key, value in headers.items():
        if str(key).casefold() == expected:
            return str(value)
    return None


def _retry_after_seconds(value: str | None) -> float | None:
    """Parse a Retry-After value into seconds.

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

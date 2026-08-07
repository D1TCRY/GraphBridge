"""Typed exceptions raised by the composed GraphBridge API."""

from __future__ import annotations

from typing import Any, Mapping

from .models import GraphError


class GraphBridgeError(Exception):
    """Base class for GraphBridge errors."""


class GraphAuthenticationError(GraphBridgeError):
    """Report a token acquisition failure or HTTP 401 response.

    Args:
        message: Safe error message.
        status_code: Optional HTTP status code.
        response_text: Redacted response text.
        error: Optional structured Graph error.
    """

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None = None,
        response_text: str = "",
        error: GraphError | None = None,
    ) -> None:
        """Initialize an authentication error.

        Args:
            message: Safe error message.
            status_code: Optional HTTP status code.
            response_text: Redacted response text.
            error: Optional structured Graph error.
        """
        self.status_code = status_code
        self.response_text = response_text
        self.error = error
        super().__init__(message)


class GraphNetworkError(GraphBridgeError):
    """The HTTP request could not be completed."""


class GraphInvalidResponseError(GraphBridgeError):
    """Microsoft Graph returned a response with an unexpected shape."""


class GraphAmbiguousMatchError(GraphBridgeError):
    """Report a name that matched multiple Graph resources.

    Args:
        resource: Type of resource being resolved.
        name: Name that matched multiple resources.
        candidate_ids: Identifiers of all matching resources.
    """

    def __init__(self, resource: str, name: str, candidate_ids: list[str]) -> None:
        """Initialize an ambiguous-name error.

        Args:
            resource: Type of resource being resolved.
            name: Name that matched multiple resources.
            candidate_ids: Identifiers of all matching resources.
        """
        self.resource = resource
        self.name = name
        self.candidate_ids = list(candidate_ids)
        candidates = ", ".join(candidate_ids)
        super().__init__(f"{resource} name {name!r} is ambiguous; matching ids: {candidates}")


class GraphUnsupportedOperationError(GraphBridgeError):
    """The requested operation is not available on Microsoft Graph v1.0."""


class GraphRequestError(GraphBridgeError):
    """Represent a structured non-success Graph response.

    Args:
        error: Structured Microsoft Graph error.
        response_text: Redacted raw response text.
        retry_after: Optional retry delay in seconds.
        response_headers: Optional redacted response headers.
    """

    def __init__(
        self,
        error: GraphError,
        *,
        response_text: str = "",
        retry_after: float | None = None,
        response_headers: Mapping[str, Any] | None = None,
    ) -> None:
        """Initialize a structured Graph request error.

        Args:
            error: Structured Microsoft Graph error.
            response_text: Redacted raw response text.
            retry_after: Optional retry delay in seconds.
            response_headers: Optional redacted response headers.
        """
        self.error = error
        self.status_code = error.status_code
        self.response_text = response_text
        self.retry_after = retry_after
        self.response_headers = dict(response_headers or {})
        super().__init__(error.message)


class GraphPermissionError(GraphRequestError):
    """Microsoft Graph returned HTTP 403."""


class GraphNotFoundError(GraphRequestError):
    """Microsoft Graph returned HTTP 404."""


class GraphConflictError(GraphRequestError):
    """Microsoft Graph returned HTTP 409."""


class GraphGoneError(GraphRequestError):
    """Microsoft Graph returned HTTP 410."""


class GraphPreconditionFailedError(GraphRequestError):
    """An eTag or another precondition failed with HTTP 412."""


class GraphThrottlingError(GraphRequestError):
    """Microsoft Graph throttled the request with HTTP 429."""


class GraphServerError(GraphRequestError):
    """Microsoft Graph returned an HTTP 5xx response."""


class DeltaResetRequiredError(GraphBridgeError):
    """Report an expired delta cursor that requires resynchronization.

    Args:
        error: Structured HTTP 410 Graph error.
        restart_link: Optional link supplied for restarting synchronization.
        strategy: Graph resynchronization strategy code.
    """

    def __init__(
        self,
        error: GraphError,
        *,
        restart_link: str | None,
        strategy: str,
    ) -> None:
        """Initialize an expired-delta-cursor error.

        Args:
            error: Structured HTTP 410 Graph error.
            restart_link: Optional link supplied for restarting synchronization.
            strategy: Graph resynchronization strategy code.
        """
        self.error = error
        self.status_code = 410
        self.restart_link = restart_link
        self.strategy = strategy
        super().__init__(
            f"Microsoft Graph cannot continue this delta feed ({strategy}); "
            "review and perform an explicit full-state reconciliation"
        )


class SyncValidationError(GraphBridgeError, ValueError):
    """A synchronization source or remote key set is unsafe to plan."""


class SyncMissingKeyError(SyncValidationError):
    """Report rows that do not contain the synchronization key.

    Args:
        key_field: Field used as the synchronization key.
        locations: Source or remote locations missing the key.
    """

    def __init__(self, key_field: str, locations: list[str]) -> None:
        """Initialize a missing synchronization key error.

        Args:
            key_field: Field used as the synchronization key.
            locations: Source or remote locations missing the key.
        """
        self.key_field = key_field
        self.locations = list(locations)
        super().__init__(
            f"sync key {key_field!r} is missing or empty at: {', '.join(locations)}"
        )


class SyncDuplicateKeyError(SyncValidationError):
    """Report synchronization keys that are not unique.

    Args:
        key_field: Field used as the synchronization key.
        duplicates: Duplicate values mapped to their locations.
    """

    def __init__(self, key_field: str, duplicates: Mapping[Any, list[str]]) -> None:
        """Initialize a duplicate synchronization key error.

        Args:
            key_field: Field used as the synchronization key.
            duplicates: Duplicate values mapped to their locations.
        """
        self.key_field = key_field
        self.duplicates = {key: list(value) for key, value in duplicates.items()}
        values = ", ".join(repr(value) for value in duplicates)
        super().__init__(f"sync key {key_field!r} contains duplicate values: {values}")

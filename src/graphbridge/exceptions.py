"""Typed exceptions raised by the composed GraphBridge API."""

from __future__ import annotations

from typing import Any, Mapping

from .models import GraphError


class GraphBridgeError(Exception):
    """Base class for errors raised by the composed GraphBridge API.

    Catching this type handles library-specific failures without also catching
    unrelated application exceptions.
    """


class GraphAuthenticationError(GraphBridgeError):
    """Report a token acquisition failure or HTTP 401 response.

    The exception keeps only redacted response details and may carry the
    structured Graph error when the failure originated from an HTTP response.

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

        Response details are stored for compatibility with the legacy adapter,
        while the main exception message remains safe for logs.

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
    """Report a request that could not be completed at the network layer.

    This is raised only after the configured retry budget, when applicable, has
    been exhausted.
    """


class GraphInvalidResponseError(GraphBridgeError):
    """Report a Microsoft Graph response with an unexpected shape.

    GraphBridge raises this instead of guessing when required JSON properties,
    collection entries, or continuation links are malformed.
    """


class GraphAmbiguousMatchError(GraphBridgeError):
    """Report a name that matched multiple Graph resources.

    Candidate identifiers are retained so the caller can choose an immutable ID
    rather than allowing GraphBridge to select an arbitrary match.

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
    """Report an operation unavailable on stable Microsoft Graph v1.0.

    The library does not silently substitute beta endpoints or emulate a
    capability whose stable semantics are unavailable.
    """


class GraphRequestError(GraphBridgeError):
    """Represent a structured non-success Microsoft Graph response.

    The original status, safe response text, headers, and retry information are
    preserved when available for diagnostics and caller policy decisions.

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
    """Report HTTP 403 when Graph rejects the caller's authorization.

    This generally indicates missing application permissions or a missing
    resource assignment for a Selected permission.
    """


class GraphNotFoundError(GraphRequestError):
    """Report HTTP 404 when a requested Graph resource does not exist.

    It is also used by explicit name-resolution helpers when enumeration finds
    no exact match.
    """


class GraphConflictError(GraphRequestError):
    """Report HTTP 409 when a Graph operation conflicts with current state.

    Callers should inspect the structured error before deciding whether to retry
    or revise the requested mutation.
    """


class GraphGoneError(GraphRequestError):
    """Report a generic HTTP 410 response from Microsoft Graph.

    Recognized expired delta cursors are converted to the more specific
    :class:`DeltaResetRequiredError` by the delta resource.
    """


class GraphPreconditionFailedError(GraphRequestError):
    """Report an eTag or other precondition failure returned as HTTP 412.

    The exception keeps concurrency conflicts visible so callers can re-read or
    re-plan instead of silently overwriting newer remote state.
    """


class GraphThrottlingError(GraphRequestError):
    """Report that Microsoft Graph throttled a request with HTTP 429.

    When available, the parsed ``Retry-After`` value is stored on the inherited
    request-error attributes.
    """


class GraphServerError(GraphRequestError):
    """Report an HTTP 5xx failure returned by Microsoft Graph.

    Safe methods may already have been retried by the transport before this
    exception reaches the caller.
    """


class DeltaResetRequiredError(GraphBridgeError):
    """Report an expired delta cursor that requires resynchronization.

    GraphBridge exposes the server strategy and restart link but deliberately
    leaves full-state reconciliation to the application.

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
    """Report synchronization input that is unsafe to plan or apply.

    This validation family prevents ambiguous matching before any mutation is
    attempted.
    """


class SyncMissingKeyError(SyncValidationError):
    """Report rows that do not contain the synchronization key.

    Every source and remote row needs a non-empty key so planning can establish a
    deterministic one-to-one match.

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

    Duplicate values are grouped with their source or remote locations to make
    the unsafe records straightforward to identify.

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

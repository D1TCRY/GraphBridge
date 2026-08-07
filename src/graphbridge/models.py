"""Small dependency-free models used by the composed GraphBridge API."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any, Generic, Literal, Mapping, TypeVar

T = TypeVar("T")


@dataclass(frozen=True, slots=True)
class GraphError:
    """Store a structured Microsoft Graph error.

    Args:
        code: Graph error code.
        message: Safe human-readable message.
        status_code: Optional HTTP status code.
        request_id: Optional Graph request identifier.
        date: Optional date reported by Graph.
        inner_error: Additional structured error details.
    """
    code: str
    message: str
    status_code: int | None = None
    request_id: str | None = None
    date: str | None = None
    inner_error: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class SiteInfo:
    """Store common SharePoint site metadata.

    Args:
        id: Graph site identifier.
        display_name: Optional display name.
        name: Optional internal name.
        web_url: Optional browser URL.
        raw: Complete original payload.
    """
    id: str
    display_name: str | None = None
    name: str | None = None
    web_url: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> SiteInfo:
        """Build site metadata from a Graph payload.

        Args:
            payload: Site response mapping.
        """
        return cls(
            id=str(payload.get("id", "")),
            display_name=_optional_string(payload.get("displayName")),
            name=_optional_string(payload.get("name")),
            web_url=_optional_string(payload.get("webUrl")),
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class ListInfo:
    """Store common SharePoint list metadata.

    Args:
        id: Graph list identifier.
        display_name: Optional display name.
        name: Optional internal name.
        web_url: Optional browser URL.
        description: Optional list description.
        etag: Optional concurrency token.
        created_date_time: Optional creation timestamp.
        last_modified_date_time: Optional modification timestamp.
        template: Optional SharePoint list template.
        raw: Complete original payload.
    """
    id: str
    display_name: str | None = None
    name: str | None = None
    web_url: str | None = None
    description: str | None = None
    etag: str | None = None
    created_date_time: str | None = None
    last_modified_date_time: str | None = None
    template: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> ListInfo:
        """Build list metadata from a Graph payload.

        Args:
            payload: List response mapping.
        """
        list_facet = payload.get("list", {})
        if not isinstance(list_facet, Mapping):
            list_facet = {}
        return cls(
            id=str(payload.get("id", "")),
            display_name=_optional_string(payload.get("displayName")),
            name=_optional_string(payload.get("name")),
            web_url=_optional_string(payload.get("webUrl")),
            description=_optional_string(payload.get("description")),
            etag=_optional_string(payload.get("eTag") or payload.get("@odata.etag")),
            created_date_time=_optional_string(payload.get("createdDateTime")),
            last_modified_date_time=_optional_string(payload.get("lastModifiedDateTime")),
            template=_optional_string(list_facet.get("template")),
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class ListItem:
    """Store a SharePoint list item and its fields.

    Args:
        id: Graph item identifier.
        fields: SharePoint field values.
        etag: Optional concurrency token.
        created_date_time: Optional creation timestamp.
        last_modified_date_time: Optional modification timestamp.
        web_url: Optional browser URL.
        response_empty: Whether Graph returned an empty success response.
        raw: Complete original payload.
    """
    id: str
    fields: Mapping[str, Any] = field(default_factory=dict)
    etag: str | None = None
    created_date_time: str | None = None
    last_modified_date_time: str | None = None
    web_url: str | None = None
    response_empty: bool = False
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any], *, fallback_id: str = "") -> ListItem:
        """Build a list item from a Graph payload.

        Args:
            payload: List-item response mapping.
            fallback_id: Identifier used when the payload omits one.
        """
        fields_payload = payload.get("fields", {})
        raw_fields = dict(fields_payload) if isinstance(fields_payload, Mapping) else {}
        field_etag = raw_fields.get("@odata.etag")
        fields = {
            str(name): value
            for name, value in raw_fields.items()
            if not str(name).startswith("@odata.")
        }
        return cls(
            id=str(payload.get("id", fallback_id)),
            fields=fields,
            etag=_optional_string(
                payload.get("eTag") or payload.get("@odata.etag") or field_etag
            ),
            created_date_time=_optional_string(payload.get("createdDateTime")),
            last_modified_date_time=_optional_string(payload.get("lastModifiedDateTime")),
            web_url=_optional_string(payload.get("webUrl")),
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class ColumnInfo:
    """Store a SharePoint column definition.

    Args:
        id: Graph column identifier.
        name: Optional internal field name.
        display_name: Optional display name.
        description: Optional column description.
        column_type: Detected Graph column facet.
        type_properties: Properties of the detected type facet.
        raw: Complete original payload.
    """
    id: str
    name: str | None = None
    display_name: str | None = None
    description: str | None = None
    column_type: str | None = None
    type_properties: Any = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> ColumnInfo:
        """Build column metadata from a Graph payload.

        Args:
            payload: Column response mapping.
        """
        column_type = next((name for name in _COLUMN_TYPE_PROPERTIES if name in payload), None)
        return cls(
            id=str(payload.get("id", "")),
            name=_optional_string(payload.get("name")),
            display_name=_optional_string(payload.get("displayName")),
            description=_optional_string(payload.get("description")),
            column_type=column_type,
            type_properties=payload.get(column_type) if column_type is not None else None,
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class ListItemVersion:
    """Store one retained SharePoint list item version.

    Args:
        id: Version identifier.
        fields: Field values stored in the version.
        last_modified_by: Identity that last modified the version.
        last_modified_date_time: Optional modification timestamp.
        published: Optional publication metadata.
        raw: Complete original payload.
    """

    id: str
    fields: Mapping[str, Any] = field(default_factory=dict)
    last_modified_by: Mapping[str, Any] = field(default_factory=dict)
    last_modified_date_time: str | None = None
    published: Mapping[str, Any] = field(default_factory=dict)
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> ListItemVersion:
        """Build a version from a Graph payload.

        Args:
            payload: Version response mapping.
        """
        fields_payload = payload.get("fields", {})
        modified_by = payload.get("lastModifiedBy", {})
        published = payload.get("published", {})
        return cls(
            id=str(payload.get("id", "")),
            fields=dict(fields_payload) if isinstance(fields_payload, Mapping) else {},
            last_modified_by=(
                dict(modified_by) if isinstance(modified_by, Mapping) else {}
            ),
            last_modified_date_time=_optional_string(
                payload.get("lastModifiedDateTime")
            ),
            published=dict(published) if isinstance(published, Mapping) else {},
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class DeletedListItem:
    """Store a list-item tombstone returned by a delta feed.

    Args:
        id: Deleted item identifier.
        state: Optional deletion state.
        raw: Complete original payload.
    """

    id: str
    state: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> DeletedListItem:
        """Build a deletion tombstone from a Graph payload.

        Args:
            payload: Delta entry mapping.
        """
        deleted = payload.get("deleted", {})
        if not isinstance(deleted, Mapping):
            deleted = {}
        return cls(
            id=str(payload.get("id", "")),
            state=_optional_string(deleted.get("state")),
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class DeltaResult:
    """Store one fully traversed list-item delta round.

    Args:
        created: Items classified as newly created.
        modified: Items classified as modified.
        deleted: Deleted-item tombstones.
        unclassified: Changes that could not be classified without known IDs.
        delta_link: Opaque link for the next delta round.
        pages: Number of pages traversed.
    """

    created: list[ListItem] = field(default_factory=list)
    modified: list[ListItem] = field(default_factory=list)
    deleted: list[DeletedListItem] = field(default_factory=list)
    unclassified: list[ListItem] = field(default_factory=list)
    delta_link: str | None = None
    pages: int = 0

    @property
    def cursor(self) -> str | None:
        """Return the opaque delta link as a cursor alias."""
        return self.delta_link


@dataclass(frozen=True, slots=True)
class Page(Generic[T]):
    """Store one page from a Graph collection.

    Args:
        items: Parsed items in the page.
        next_link: Optional link to the next page.
        delta_link: Optional final delta cursor.
        raw: Complete original payload.
    """
    items: list[T] = field(default_factory=list)
    next_link: str | None = None
    delta_link: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class BatchResult(Generic[T]):
    """Store aggregate and per-input batch outcomes.

    Args:
        successes: Successfully parsed values.
        failures: Structured failures.
        results: Ordered outcome for every input.
    """
    successes: list[T] = field(default_factory=list)
    failures: list[GraphError] = field(default_factory=list)
    results: list[BatchItemResult[T]] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class BatchItemResult(Generic[T]):
    """Store one batch outcome in original input order.

    Args:
        input_index: Position in the original input.
        request_id: Correlated batch request identifier.
        status_code: Subrequest HTTP status.
        value: Optional successful value.
        error: Optional structured failure.
        attempts: Number of attempts used.
        response_headers: Raw subresponse headers.
        response_body: Raw decoded subresponse body.
    """

    input_index: int
    request_id: str
    status_code: int
    value: T | None = None
    error: GraphError | None = None
    attempts: int = 1
    response_headers: Mapping[str, Any] = field(default_factory=dict)
    response_body: Any = None


@dataclass(frozen=True, slots=True)
class SyncFieldDifference:
    """Store one field-level synchronization difference.

    Args:
        field: Field name.
        local_value: Desired source value.
        remote_value: Current SharePoint value.
    """

    field: str
    local_value: Any
    remote_value: Any

    def to_dict(self) -> dict[str, Any]:
        """Serialize the field difference to a dictionary."""
        return {
            "field": self.field,
            "local_value": self.local_value,
            "remote_value": self.remote_value,
        }


SyncOperationKind = Literal["create", "update", "delete", "unchanged"]


@dataclass(frozen=True, slots=True)
class SyncOperation:
    """Store one inspectable synchronization operation.

    Args:
        operation: Operation kind.
        key: Business key identifying the row.
        fields: Field values involved in the operation.
        item_id: Optional remote item identifier.
        etag: Optional remote concurrency token.
        differences: Field-level differences for an update.
        reason: Human-readable reason for the operation.
        source_index: Optional source-row position.
        remote_index: Optional remote-item position.
    """

    operation: SyncOperationKind
    key: Any
    fields: Mapping[str, Any] = field(default_factory=dict)
    item_id: str | None = None
    etag: str | None = None
    differences: list[SyncFieldDifference] = field(default_factory=list)
    reason: str = ""
    source_index: int | None = None
    remote_index: int | None = None

    def to_dict(self) -> dict[str, Any]:
        """Serialize the synchronization operation to a dictionary."""
        return {
            "operation": self.operation,
            "key": self.key,
            "fields": dict(self.fields),
            "item_id": self.item_id,
            "etag": self.etag,
            "differences": [difference.to_dict() for difference in self.differences],
            "reason": self.reason,
            "source_index": self.source_index,
            "remote_index": self.remote_index,
        }


@dataclass(frozen=True, slots=True)
class SyncPlan:
    """Store a side-effect-free synchronization plan.

    Args:
        creates: Planned create operations.
        updates: Planned update operations.
        deletes: Planned delete operations.
        unchanged: Rows that require no mutation.
        key_field: Field used as the business key.
        prune: Whether remote-only rows should be deleted.
        dry_run: Whether applying the plan should avoid writes.
        field_names: Field-name convention used by the plan.
        local_count: Number of source rows inspected.
        remote_count: Number of remote items inspected.
    """

    # The first three defaults retain source compatibility with the phase-2
    # placeholder model while new plans contain SyncOperation values.
    creates: Sequence[SyncOperation | Mapping[str, Any]] = field(default_factory=list)
    updates: Sequence[SyncOperation | Mapping[str, Any]] = field(default_factory=list)
    deletes: Sequence[SyncOperation | str] = field(default_factory=list)
    unchanged: Sequence[SyncOperation] = field(default_factory=list)
    key_field: str = ""
    prune: bool = False
    dry_run: bool = False
    field_names: Literal["internal", "display"] = "internal"
    local_count: int = 0
    remote_count: int = 0

    @property
    def operation_count(self) -> int:
        """Return the number of mutating operations in the plan."""
        return len(self.creates) + len(self.updates) + len(self.deletes)

    def to_dict(self) -> dict[str, Any]:
        """Serialize the synchronization plan to a dictionary."""
        def serialize(value: object) -> Any:
            """Serialize one plan entry.

            Args:
                value: Plan entry to serialize.
            """
            if isinstance(value, SyncOperation):
                return value.to_dict()
            if isinstance(value, Mapping):
                return dict(value)
            return value

        return {
            "key_field": self.key_field,
            "prune": self.prune,
            "dry_run": self.dry_run,
            "field_names": self.field_names,
            "local_count": self.local_count,
            "remote_count": self.remote_count,
            "creates": [serialize(value) for value in self.creates],
            "updates": [serialize(value) for value in self.updates],
            "deletes": [serialize(value) for value in self.deletes],
            "unchanged": [value.to_dict() for value in self.unchanged],
        }


@dataclass(frozen=True, slots=True)
class SyncOperationResult:
    """Store the outcome of one synchronization operation.

    Args:
        operation: Operation that was evaluated.
        value: Optional successful item or identifier.
        error: Optional structured failure.
        status_code: Optional HTTP status code.
        attempts: Number of attempts used.
        deferred: Whether the operation was deliberately postponed.
    """

    operation: SyncOperation
    value: ListItem | str | None = None
    error: GraphError | None = None
    status_code: int | None = None
    attempts: int = 1
    deferred: bool = False

    @property
    def succeeded(self) -> bool:
        """Return whether the operation succeeded without deferral."""
        return self.error is None and not self.deferred

    def to_dict(self) -> dict[str, Any]:
        """Serialize the operation result to a dictionary."""
        value: Any = self.value
        if isinstance(value, ListItem):
            value = {
                "id": value.id,
                "fields": dict(value.fields),
                "etag": value.etag,
                "response_empty": value.response_empty,
            }
        error = None
        if self.error is not None:
            error = {
                "code": self.error.code,
                "message": self.error.message,
                "status_code": self.error.status_code,
                "inner_error": dict(self.error.inner_error),
            }
        return {
            "operation": self.operation.to_dict(),
            "value": value,
            "error": error,
            "status_code": self.status_code,
            "attempts": self.attempts,
            "deferred": self.deferred,
        }


@dataclass(frozen=True, slots=True)
class SyncResult:
    """Store the aggregate result of applying a synchronization plan.

    Args:
        created: Items created successfully.
        updated: Items updated successfully.
        deleted: Item identifiers deleted successfully.
        unchanged: Operations that required no mutation.
        failures: Structured operation failures.
        results: Correlated outcomes for all operations.
        plan: Original applied plan, when available.
        applied: Whether mutations were attempted.
        dry_run: Whether the result came from a dry run.
    """
    created: list[ListItem] = field(default_factory=list)
    updated: list[ListItem] = field(default_factory=list)
    deleted: list[str] = field(default_factory=list)
    unchanged: list[SyncOperation] = field(default_factory=list)
    failures: list[GraphError] = field(default_factory=list)
    results: list[SyncOperationResult] = field(default_factory=list)
    plan: SyncPlan | None = None
    applied: bool = True
    dry_run: bool = False

    def retry_plan(self) -> SyncPlan:
        """Build a plan containing only failed or deferred operations."""

        failed = [result.operation for result in self.results if not result.succeeded]
        source = self.plan
        return SyncPlan(
            creates=[item for item in failed if item.operation == "create"],
            updates=[item for item in failed if item.operation == "update"],
            deletes=[item for item in failed if item.operation == "delete"],
            key_field=source.key_field if source is not None else "",
            prune=source.prune if source is not None else any(
                item.operation == "delete" for item in failed
            ),
            field_names=source.field_names if source is not None else "internal",
            local_count=source.local_count if source is not None else 0,
            remote_count=source.remote_count if source is not None else 0,
        )

    def to_dict(self) -> dict[str, Any]:
        """Serialize the synchronization result to a dictionary."""
        def item(value: ListItem) -> dict[str, Any]:
            """Serialize one list item.

            Args:
                value: List item to serialize.
            """
            return {
                "id": value.id,
                "fields": dict(value.fields),
                "etag": value.etag,
                "response_empty": value.response_empty,
            }

        return {
            "created": [item(value) for value in self.created],
            "updated": [item(value) for value in self.updated],
            "deleted": list(self.deleted),
            "unchanged": [value.to_dict() for value in self.unchanged],
            "failures": [
                {
                    "code": error.code,
                    "message": error.message,
                    "status_code": error.status_code,
                    "inner_error": dict(error.inner_error),
                }
                for error in self.failures
            ],
            "results": [value.to_dict() for value in self.results],
            "applied": self.applied,
            "dry_run": self.dry_run,
        }


def _optional_string(value: object) -> str | None:
    """Convert a value to text while preserving ``None``.

    Args:
        value: Value to convert.
    """
    return str(value) if value is not None else None


_COLUMN_TYPE_PROPERTIES = (
    "boolean",
    "calculated",
    "choice",
    "contentApprovalStatus",
    "currency",
    "dateTime",
    "geolocation",
    "hyperlinkOrPicture",
    "lookup",
    "number",
    "personOrGroup",
    "term",
    "text",
    "thumbnail",
)

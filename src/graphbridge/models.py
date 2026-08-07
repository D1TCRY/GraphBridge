"""Small dependency-free models used by the composed GraphBridge API."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any, Generic, Literal, Mapping, TypeVar

T = TypeVar("T")


@dataclass(frozen=True, slots=True)
class GraphError:
    code: str
    message: str
    status_code: int | None = None
    request_id: str | None = None
    date: str | None = None
    inner_error: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class SiteInfo:
    id: str
    display_name: str | None = None
    name: str | None = None
    web_url: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> SiteInfo:
        return cls(
            id=str(payload.get("id", "")),
            display_name=_optional_string(payload.get("displayName")),
            name=_optional_string(payload.get("name")),
            web_url=_optional_string(payload.get("webUrl")),
            raw=dict(payload),
        )


@dataclass(frozen=True, slots=True)
class ListInfo:
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
    id: str
    name: str | None = None
    display_name: str | None = None
    description: str | None = None
    column_type: str | None = None
    type_properties: Any = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> ColumnInfo:
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
    """One retained SharePoint list item version."""

    id: str
    fields: Mapping[str, Any] = field(default_factory=dict)
    last_modified_by: Mapping[str, Any] = field(default_factory=dict)
    last_modified_date_time: str | None = None
    published: Mapping[str, Any] = field(default_factory=dict)
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> ListItemVersion:
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
    """A list item tombstone returned by a delta feed."""

    id: str
    state: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)

    @classmethod
    def from_payload(cls, payload: Mapping[str, Any]) -> DeletedListItem:
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
    """One fully traversed list-item delta round and its next opaque cursor.

    ``unclassified`` contains non-deleted changes when the caller did not
    provide its known item IDs. Microsoft Graph reports the latest item state,
    but does not label a non-deleted entry as a create or an update.
    """

    created: list[ListItem] = field(default_factory=list)
    modified: list[ListItem] = field(default_factory=list)
    deleted: list[DeletedListItem] = field(default_factory=list)
    unclassified: list[ListItem] = field(default_factory=list)
    delta_link: str | None = None
    pages: int = 0

    @property
    def cursor(self) -> str | None:
        return self.delta_link


@dataclass(frozen=True, slots=True)
class Page(Generic[T]):
    items: list[T] = field(default_factory=list)
    next_link: str | None = None
    delta_link: str | None = None
    raw: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class BatchResult(Generic[T]):
    successes: list[T] = field(default_factory=list)
    failures: list[GraphError] = field(default_factory=list)
    results: list[BatchItemResult[T]] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class BatchItemResult(Generic[T]):
    """One correlated batch outcome, retained in original input order."""

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
    """A reviewable field-level difference between source and SharePoint."""

    field: str
    local_value: Any
    remote_value: Any

    def to_dict(self) -> dict[str, Any]:
        return {
            "field": self.field,
            "local_value": self.local_value,
            "remote_value": self.remote_value,
        }


SyncOperationKind = Literal["create", "update", "delete", "unchanged"]


@dataclass(frozen=True, slots=True)
class SyncOperation:
    """One deterministic, inspectable operation in a synchronization plan."""

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
    """A side-effect-free description of a proposed list synchronization."""

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
        return len(self.creates) + len(self.updates) + len(self.deletes)

    def to_dict(self) -> dict[str, Any]:
        def serialize(value: object) -> Any:
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
    """One correlated outcome from applying a synchronization operation."""

    operation: SyncOperation
    value: ListItem | str | None = None
    error: GraphError | None = None
    status_code: int | None = None
    attempts: int = 1
    deferred: bool = False

    @property
    def succeeded(self) -> bool:
        return self.error is None and not self.deferred

    def to_dict(self) -> dict[str, Any]:
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
        """Build a plan containing only failed or safely deferred operations."""

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
        def item(value: ListItem) -> dict[str, Any]:
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

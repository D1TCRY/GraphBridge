"""SharePoint list column schema resources for Microsoft Graph v1.0."""

from __future__ import annotations

import builtins
from collections.abc import Iterator, Mapping, Sequence
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

from ..exceptions import GraphAmbiguousMatchError, GraphInvalidResponseError, GraphNotFoundError
from ..models import ColumnInfo, GraphError, Page
from ..pagination import iter_pages
from ..query import ODataQuery

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .lists import SharePointListResource


class ColumnsResource:
    """Read and mutate list ``columnDefinition`` resources."""

    def __init__(self, client: GraphBridgeClient, sharepoint_list: SharePointListResource) -> None:
        self.client = client
        self.transport = client.transport
        self.sharepoint_list = sharepoint_list
        self._schema_cache: builtins.list[ColumnInfo] | None = None

    def list(
        self,
        *,
        select: Sequence[str] | None = None,
        top: int | None = None,
        orderby: Sequence[str] | None = None,
    ) -> Page[ColumnInfo]:
        return next(self.iter_pages(select=select, top=top, orderby=orderby))

    def iter_pages(
        self,
        *,
        select: Sequence[str] | None = None,
        top: int | None = None,
        orderby: Sequence[str] | None = None,
    ) -> Iterator[Page[ColumnInfo]]:
        query = ODataQuery(
            select=tuple(select or ()),
            top=top,
            orderby=tuple(orderby or ()),
        )
        return iter_pages(
            self.transport,
            self._base_path,
            params=query.to_params() or None,
            parser=ColumnInfo.from_payload,
        )

    def iter_all(
        self,
        *,
        select: Sequence[str] | None = None,
        top: int | None = None,
        orderby: Sequence[str] | None = None,
    ) -> Iterator[ColumnInfo]:
        for page in self.iter_pages(select=select, top=top, orderby=orderby):
            yield from page.items

    def get(self, column_id: str) -> ColumnInfo:
        """Compatibility alias for :meth:`get_by_id`."""

        return self.get_by_id(column_id)

    def get_by_id(self, column_id: str) -> ColumnInfo:
        if not isinstance(column_id, str) or not column_id:
            raise ValueError("column_id cannot be empty")
        payload = self.transport.get(f"{self._base_path}/{quote(column_id, safe='')}")
        return self._column_from_payload(payload)

    def get_by_name(self, name: str, *, include_display_name: bool = True) -> ColumnInfo:
        """Resolve a column by internal name, optionally falling back to displayName."""

        if not isinstance(name, str) or not name:
            raise ValueError("column name cannot be empty")
        schema = self._load_schema()
        internal_matches = [column for column in schema if column.name == name]
        matches = internal_matches
        if not matches and include_display_name:
            matches = [column for column in schema if column.display_name == name]
        if not matches:
            raise GraphNotFoundError(
                GraphError(
                    code="itemNotFound",
                    message=f"No SharePoint column named {name!r} was found",
                    status_code=404,
                )
            )
        if len(matches) > 1:
            raise GraphAmbiguousMatchError("SharePoint column", name, [item.id for item in matches])
        return matches[0]

    def create(self, definition: Mapping[str, Any]) -> ColumnInfo:
        """Create one stable ``columnDefinition`` without discarding unknown facets."""

        if not isinstance(definition, Mapping):
            raise TypeError("column definition must be a mapping")
        payload = self.transport.post(self._base_path, json=dict(definition))
        column = self._column_from_payload(payload)
        self.invalidate_schema()
        return column

    def update(
        self,
        column_id: str,
        changes: Mapping[str, Any],
    ) -> ColumnInfo:
        """Update any mutable v1.0 column property except ``id``."""

        if not isinstance(column_id, str) or not column_id:
            raise ValueError("column_id cannot be empty")
        if "id" in changes:
            raise ValueError("column id is read-only")
        payload = self.transport.patch(
            f"{self._base_path}/{quote(column_id, safe='')}",
            json=dict(changes),
        )
        column = self._column_from_payload(payload)
        self.invalidate_schema()
        return column

    def delete(self, column_id: str) -> None:
        if not isinstance(column_id, str) or not column_id:
            raise ValueError("column_id cannot be empty")
        payload = self.transport.delete(f"{self._base_path}/{quote(column_id, safe='')}")
        if payload is not None:
            raise GraphInvalidResponseError("Microsoft Graph column delete response must be empty")
        self.invalidate_schema()

    def display_name_map(self, *, refresh: bool = False) -> dict[str, str]:
        """Return the authoritative ``displayName -> name`` schema mapping."""

        mapping: dict[str, str] = {}
        ids_by_display: dict[str, builtins.list[str]] = {}
        for column in self._load_schema(refresh=refresh):
            if column.display_name is None or column.name is None:
                continue
            ids_by_display.setdefault(column.display_name, []).append(column.id)
            mapping[column.display_name] = column.name
        duplicate = next(
            ((name, ids) for name, ids in ids_by_display.items() if len(ids) > 1), None
        )
        if duplicate is not None:
            name, ids = duplicate
            raise GraphAmbiguousMatchError("SharePoint column displayName", name, ids)
        return mapping

    @property
    def name_map(self) -> dict[str, str]:
        """Cached authoritative alias for :meth:`display_name_map`."""

        return self.display_name_map()

    def to_internal_fields(
        self, fields: Mapping[str, Any], *, strict: bool = True
    ) -> dict[str, Any]:
        """Translate display-name keys to internal SharePoint field names."""

        mapping = self.display_name_map()
        internal_names = set(mapping.values())
        translated: dict[str, Any] = {}
        for name, value in fields.items():
            if name in internal_names:
                translated[name] = value
            elif name in mapping:
                translated[mapping[name]] = value
            elif strict:
                raise KeyError(f"unknown SharePoint column displayName: {name!r}")
            else:
                translated[name] = value
        return translated

    def to_display_fields(self, fields: Mapping[str, Any]) -> dict[str, Any]:
        """Translate known internal field keys to their display names."""

        reverse = {internal: display for display, internal in self.display_name_map().items()}
        return {reverse.get(name, name): value for name, value in fields.items()}

    def invalidate_schema(self) -> None:
        self._schema_cache = None

    @property
    def _base_path(self) -> str:
        site_id = quote(self.sharepoint_list.site.id, safe=",")
        list_id = quote(self.sharepoint_list.id, safe="")
        return f"/sites/{site_id}/lists/{list_id}/columns"

    def _load_schema(self, *, refresh: bool = False) -> builtins.list[ColumnInfo]:
        if refresh or self._schema_cache is None:
            self._schema_cache = builtins.list(self.iter_all())
        return builtins.list(self._schema_cache)

    @staticmethod
    def _column_from_payload(payload: Any) -> ColumnInfo:
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph column response must be a JSON object")
        column = ColumnInfo.from_payload(payload)
        if not column.id:
            raise GraphInvalidResponseError("Microsoft Graph column response does not contain an id")
        return column

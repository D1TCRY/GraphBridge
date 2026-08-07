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
    """Read and mutate SharePoint column definitions.

    Besides column CRUD, this resource owns the authoritative mapping between
    user-facing display names and SharePoint internal field names. It is available
    as ``sharepoint_list.columns`` and shares the parent client's transport.

    Args:
        client: Shared GraphBridge client.
        sharepoint_list: Parent SharePoint list.
    """

    def __init__(self, client: GraphBridgeClient, sharepoint_list: SharePointListResource) -> None:
        """Initialize the columns resource.

        Construction creates an empty schema cache and performs no HTTP request;
        definitions are loaded only when an operation needs them.

        Args:
            client: Shared GraphBridge client.
            sharepoint_list: Parent SharePoint list.
        """
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
        """Return the first page of columns.

        This convenience method is suitable when only the initial Graph page is
        needed; complete schema discovery should use the lazy iterators. The page
        retains both typed columns and the original Graph response.

        Args:
            select: Optional column properties to return.
            top: Optional maximum page size.
            orderby: Optional ordering expressions.
        """
        return next(self.iter_pages(select=select, top=top, orderby=orderby))

    def iter_pages(
        self,
        *,
        select: Sequence[str] | None = None,
        top: int | None = None,
        orderby: Sequence[str] | None = None,
    ) -> Iterator[Page[ColumnInfo]]:
        """Lazily iterate through column pages.

        Selection, ordering, and page-size options are validated through the
        shared OData query builder before the first request.

        Args:
            select: Optional column properties to return.
            top: Optional maximum page size.
            orderby: Optional ordering expressions.
        """
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
        """Lazily iterate through all columns.

        Parsed definitions are yielded one at a time without eagerly materializing
        the entire schema.

        Args:
            select: Optional column properties to return.
            top: Optional maximum page size.
            orderby: Optional ordering expressions.
        """
        for page in self.iter_pages(select=select, top=top, orderby=orderby):
            yield from page.items

    def get(self, column_id: str) -> ColumnInfo:
        """Retrieve a column by ID.

        This compatibility alias delegates directly to :meth:`get_by_id`; the
        explicit form is clearer in new code.

        Args:
            column_id: Graph column identifier.
        """

        return self.get_by_id(column_id)

    def get_by_id(self, column_id: str) -> ColumnInfo:
        """Retrieve a column by its Graph ID.

        The response must be a JSON object containing an ID; malformed successful
        responses fail explicitly. Direct lookup does not populate the schema
        enumeration cache.

        Args:
            column_id: Graph column identifier.

        Raises:
            ValueError: If the identifier is empty.
        """
        if not isinstance(column_id, str) or not column_id:
            raise ValueError("column_id cannot be empty")
        payload = self.transport.get(f"{self._base_path}/{quote(column_id, safe='')}")
        return self._column_from_payload(payload)

    def get_by_name(self, name: str, *, include_display_name: bool = True) -> ColumnInfo:
        """Resolve a column by internal or display name.

        Internal-name matches take precedence. Display names are used only as an
        optional fallback and ambiguous matches are never selected arbitrarily.
        Resolving by name may load and cache the complete schema.

        Args:
            name: Internal or display column name.
            include_display_name: Whether display names may be matched.

        Raises:
            GraphNotFoundError: If no matching column exists.
            GraphAmbiguousMatchError: If multiple columns match.
        """

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
        """Create one SharePoint column definition.

        The mapping is forwarded without stripping unknown stable properties, and
        a successful mutation invalidates the cached schema. This administrative
        operation changes the contract seen by every list consumer.

        Args:
            definition: Graph column definition payload.

        Raises:
            TypeError: If the definition is not a mapping.
        """

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
        """Update mutable properties of a column.

        The immutable ID cannot appear in the changes. After Graph accepts the
        update, the display-name mapping cache is cleared. Callers should provide
        only properties documented as mutable by Graph.

        Args:
            column_id: Graph column identifier.
            changes: Mutable column properties to update.

        Raises:
            ValueError: If the ID is empty or changes include the read-only ID.
        """

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
        """Delete a SharePoint column.

        Graph must return an empty success response; the local schema cache is
        invalidated only after that contract is satisfied. This destructive
        schema action should be part of a reviewed administrative workflow.

        Args:
            column_id: Graph column identifier.

        Raises:
            ValueError: If the identifier is empty.
            GraphInvalidResponseError: If Graph returns a non-empty success body.
        """
        if not isinstance(column_id, str) or not column_id:
            raise ValueError("column_id cannot be empty")
        payload = self.transport.delete(f"{self._base_path}/{quote(column_id, safe='')}")
        if payload is not None:
            raise GraphInvalidResponseError("Microsoft Graph column delete response must be empty")
        self.invalidate_schema()

    def display_name_map(self, *, refresh: bool = False) -> dict[str, str]:
        """Return the display-to-internal-name mapping.

        The complete schema is cached to avoid repeated enumeration, while
        duplicate display names are rejected because translation would be unsafe.
        The returned dictionary cannot mutate the internal cache.

        Args:
            refresh: Whether to reload the schema before mapping.

        Raises:
            GraphAmbiguousMatchError: If display names are duplicated.
        """

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
        """Return the cached display-to-internal-name mapping.

        Accessing the property loads the schema only when the cache is empty. Call
        :meth:`display_name_map` directly when a refresh is required.
        """

        return self.display_name_map()

    def to_internal_fields(
        self, fields: Mapping[str, Any], *, strict: bool = True
    ) -> dict[str, Any]:
        """Translate display names to internal field names.

        Keys that are already valid internal names pass through unchanged. Strict
        mode rejects unknown names instead of guessing their encoded form. This
        is the authoritative replacement for the legacy character codec.

        Args:
            fields: Field values keyed by display or internal name.
            strict: Whether unknown display names should fail.

        Raises:
            KeyError: If an unknown name is found in strict mode.
        """

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
        """Translate known internal names to display names.

        Unknown internal fields are preserved, which allows system or future Graph
        fields to remain visible to callers. Known translations come from the
        cached authoritative list schema.

        Args:
            fields: Field values keyed by internal name.
        """

        reverse = {internal: display for display, internal in self.display_name_map().items()}
        return {reverse.get(name, name): value for name, value in fields.items()}

    def invalidate_schema(self) -> None:
        """Clear the cached column schema.

        The next mapping or schema lookup will enumerate columns again from Graph.
        Column mutations call this automatically after a valid success.
        """
        self._schema_cache = None

    @property
    def _base_path(self) -> str:
        """Return the stable Graph path for columns in the bound list.

        Site and list IDs are quoted as separate path components.
        """
        site_id = quote(self.sharepoint_list.site.id, safe=",")
        list_id = quote(self.sharepoint_list.id, safe="")
        return f"/sites/{site_id}/lists/{list_id}/columns"

    def _load_schema(self, *, refresh: bool = False) -> builtins.list[ColumnInfo]:
        """Load and cache the complete column schema.

        Args:
            refresh: Whether to bypass the existing cache.
        """
        if refresh or self._schema_cache is None:
            self._schema_cache = builtins.list(self.iter_all())
        return builtins.list(self._schema_cache)

    @staticmethod
    def _column_from_payload(payload: Any) -> ColumnInfo:
        """Validate and parse a column payload.

        Args:
            payload: Decoded Graph response.

        Raises:
            GraphInvalidResponseError: If the payload is invalid or lacks an ID.
        """
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError("Microsoft Graph column response must be a JSON object")
        column = ColumnInfo.from_payload(payload)
        if not column.id:
            raise GraphInvalidResponseError("Microsoft Graph column response does not contain an id")
        return column

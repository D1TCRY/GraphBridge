"""SharePoint list resources backed only by Microsoft Graph v1.0 endpoints."""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

from ..exceptions import GraphAmbiguousMatchError, GraphInvalidResponseError, GraphNotFoundError
from ..models import GraphError, ListInfo, Page
from ..pagination import iter_pages
from ..query import ODataQuery

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .columns import ColumnsResource
    from .items import ListItemsResource
    from .sites import SiteResource
    from .sync import SyncService
    from .versions import VersionsResource

_LIST_RELATIONSHIPS = frozenset(
    {"columns", "contentTypes", "drive", "items", "operations", "subscriptions"}
)


class ListsResource:
    """Enumerate, resolve, and create SharePoint lists.

    Args:
        client: Shared GraphBridge client.
        site: Parent SharePoint site.
    """

    def __init__(self, client: GraphBridgeClient, site: SiteResource) -> None:
        """Initialize the lists resource.

        Args:
            client: Shared GraphBridge client.
            site: Parent SharePoint site.
        """
        self.client = client
        self.transport = client.transport
        self.site = site

    def get(
        self,
        identifier: str,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
    ) -> SharePointListResource:
        """Resolve a list using a Graph ID-or-title path segment.

        Args:
            identifier: List ID or title accepted by Graph.
            select: Optional metadata properties to return.
            expand: Optional stable relationships to expand.
        """

        return self._get_direct(identifier, select=select, expand=expand)

    def get_by_id(
        self,
        list_id: str,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
    ) -> SharePointListResource:
        """Retrieve one list by its immutable Graph ID.

        Args:
            list_id: Graph list identifier.
            select: Optional metadata properties to return.
            expand: Optional stable relationships to expand.
        """

        return self._get_direct(list_id, select=select, expand=expand)

    def get_by_name(self, name: str) -> SharePointListResource:
        """Resolve an exact list name and reject duplicates.

        Args:
            name: Exact display or internal list name.

        Raises:
            GraphNotFoundError: If no matching list exists.
            GraphAmbiguousMatchError: If multiple lists have the same name.
        """

        if not isinstance(name, str) or not name:
            raise ValueError("list name cannot be empty")
        matches = [
            info
            for info in self.iter_all()
            if info.display_name == name or info.name == name
        ]
        if not matches:
            raise GraphNotFoundError(
                GraphError(
                    code="itemNotFound",
                    message=f"No SharePoint list named {name!r} was found",
                    status_code=404,
                )
            )
        if len(matches) > 1:
            raise GraphAmbiguousMatchError("SharePoint list", name, [item.id for item in matches])
        return SharePointListResource(self.client, self.site, matches[0])

    def list(
        self,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
        top: int | None = None,
    ) -> Page[ListInfo]:
        """Return the first page of lists.

        Args:
            select: Optional metadata properties to return.
            expand: Optional relationships to expand.
            top: Optional maximum page size.
        """
        return next(self.iter_pages(select=select, expand=expand, top=top))

    def iter_pages(
        self,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
        top: int | None = None,
    ) -> Iterator[Page[ListInfo]]:
        """Lazily iterate through list pages.

        Args:
            select: Optional metadata properties to return.
            expand: Optional relationships to expand.
            top: Optional maximum page size.
        """
        query = ODataQuery(
            select=tuple(select or ()),
            expand=tuple(expand or ()),
            top=top,
        )
        return iter_pages(
            self.transport,
            self._base_path,
            params=query.to_params() or None,
            parser=ListInfo.from_payload,
        )

    def iter_all(
        self,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
        top: int | None = None,
    ) -> Iterator[ListInfo]:
        """Lazily iterate through all lists.

        Args:
            select: Optional metadata properties to return.
            expand: Optional relationships to expand.
            top: Optional maximum page size.
        """
        for page in self.iter_pages(select=select, expand=expand, top=top):
            yield from page.items

    def create(
        self,
        display_name: str,
        *,
        template: str = "genericList",
        columns: Sequence[Mapping[str, Any]] | None = None,
        description: str | None = None,
    ) -> SharePointListResource:
        """Create a list with optional initial columns.

        Args:
            display_name: Display name of the new list.
            template: SharePoint list template.
            columns: Optional initial column definitions.
            description: Optional list description.

        Raises:
            ValueError: If the display name or template is empty.
        """

        if not isinstance(display_name, str) or not display_name:
            raise ValueError("display_name cannot be empty")
        if not isinstance(template, str) or not template:
            raise ValueError("template cannot be empty")
        payload: dict[str, Any] = {
            "displayName": display_name,
            "list": {"template": template},
        }
        if description is not None:
            payload["description"] = description
        if columns is not None:
            payload["columns"] = [dict(column) for column in columns]
        response = self.transport.post(self._base_path, json=payload)
        return SharePointListResource(self.client, self.site, _list_info(response))

    def bind(self, sharepoint_list: ListInfo | Mapping[str, Any]) -> SharePointListResource:
        """Bind known list metadata without an HTTP request.

        Args:
            sharepoint_list: Existing list model or payload.

        Raises:
            ValueError: If the metadata does not contain an ID.
        """

        info = (
            sharepoint_list
            if isinstance(sharepoint_list, ListInfo)
            else ListInfo.from_payload(sharepoint_list)
        )
        if not info.id:
            raise ValueError("list metadata must contain an id")
        return SharePointListResource(self.client, self.site, info)

    @property
    def _base_path(self) -> str:
        """Return the Graph path for lists in the parent site."""
        return f"/sites/{quote(self.site.id, safe=',')}/lists"

    def _get_direct(
        self,
        identifier: str,
        *,
        select: Sequence[str] | None,
        expand: Sequence[str] | None,
    ) -> SharePointListResource:
        """Retrieve a list from a direct Graph path.

        Args:
            identifier: List ID or title path segment.
            select: Optional metadata properties to return.
            expand: Optional stable relationships to expand.

        Raises:
            ValueError: If the identifier is empty.
        """
        if not isinstance(identifier, str) or not identifier:
            raise ValueError("list identifier cannot be empty")
        endpoint = f"{self._base_path}/{quote(identifier, safe='')}"
        query = ODataQuery(select=tuple(select or ()), expand=tuple(expand or ()))
        payload = self.transport.get(endpoint, params=query.to_params() or None)
        return SharePointListResource(self.client, self.site, _list_info(payload))


class SharePointListResource:
    """Represent a list and its subordinate resources.

    Args:
        client: Shared GraphBridge client.
        site: Parent SharePoint site.
        info: Parsed list metadata.
    """

    def __init__(self, client: GraphBridgeClient, site: SiteResource, info: ListInfo) -> None:
        """Initialize a bound SharePoint list.

        Args:
            client: Shared GraphBridge client.
            site: Parent SharePoint site.
            info: Parsed list metadata.
        """
        from .columns import ColumnsResource
        from .items import ListItemsResource
        from .sync import SyncService
        from .versions import VersionsResource

        self.client = client
        self.transport = client.transport
        self.site = site
        self.info = info
        self.items: ListItemsResource = ListItemsResource(client, self)
        self.columns: ColumnsResource = ColumnsResource(client, self)
        self.sync: SyncService = SyncService(client, self)
        self.versions: VersionsResource = VersionsResource(client, self)

    @property
    def id(self) -> str:
        """Return the Graph list identifier."""
        return self.info.id

    @property
    def display_name(self) -> str | None:
        """Return the list display name."""
        return self.info.display_name

    @property
    def name(self) -> str | None:
        """Return the list internal name."""
        return self.info.name

    @property
    def metadata(self) -> Mapping[str, Any]:
        """Return the complete list metadata payload."""

        return self.info.raw

    def relationship(self, name: str) -> Any:
        """Read a relationship retrieved with ``$expand``.

        Args:
            name: Stable list relationship name.

        Raises:
            ValueError: If the relationship is unknown or was not expanded.
        """

        if name not in _LIST_RELATIONSHIPS:
            raise ValueError(f"unknown stable list relationship: {name!r}")
        if name not in self.info.raw:
            raise ValueError(f"relationship {name!r} was not expanded in this list response")
        return self.info.raw[name]

    def refresh(
        self,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
    ) -> SharePointListResource:
        """Refresh list metadata in place.

        Args:
            select: Optional metadata properties to return.
            expand: Optional relationships to expand.
        """

        refreshed = self.site.lists.get_by_id(self.id, select=select, expand=expand)
        self.info = refreshed.info
        return self

    def __repr__(self) -> str:
        """Return a concise list representation."""
        return f"SharePointListResource(id={self.id!r}, display_name={self.info.display_name!r})"


def _list_info(payload: Any) -> ListInfo:
    """Validate and parse a list payload.

    Args:
        payload: Decoded Graph response.

    Raises:
        GraphInvalidResponseError: If the payload is invalid or lacks an ID.
    """
    if not isinstance(payload, Mapping):
        raise GraphInvalidResponseError("Microsoft Graph list response must be a JSON object")
    info = ListInfo.from_payload(payload)
    if not info.id:
        raise GraphInvalidResponseError("Microsoft Graph list response does not contain an id")
    return info

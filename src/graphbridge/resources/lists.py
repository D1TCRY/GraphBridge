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
    """Enumerate, resolve and create lists for one SharePoint site."""

    def __init__(self, client: GraphBridgeClient, site: SiteResource) -> None:
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
        """Compatibility lookup accepting the Graph list ID-or-title path segment.

        New code should call :meth:`get_by_id` or :meth:`get_by_name` so the
        identifier's meaning and duplicate-name behavior are explicit.
        """

        return self._get_direct(identifier, select=select, expand=expand)

    def get_by_id(
        self,
        list_id: str,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
    ) -> SharePointListResource:
        """Get one list by its immutable Microsoft Graph ID."""

        return self._get_direct(list_id, select=select, expand=expand)

    def get_by_name(self, name: str) -> SharePointListResource:
        """Resolve an exact list title/name and reject ambiguous duplicates.

        Graph v1.0 permits a title in the direct list path, but that response
        cannot expose duplicate titles. A single paginated enumeration is used
        here so duplicate handling is deterministic and explicit.
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
        return next(self.iter_pages(select=select, expand=expand, top=top))

    def iter_pages(
        self,
        *,
        select: Sequence[str] | None = None,
        expand: Sequence[str] | None = None,
        top: int | None = None,
    ) -> Iterator[Page[ListInfo]]:
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
        """Create a list, optionally with initial v1.0 column definitions."""

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
        """Bind known list metadata without issuing an HTTP request."""

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
        return f"/sites/{quote(self.site.id, safe=',')}/lists"

    def _get_direct(
        self,
        identifier: str,
        *,
        select: Sequence[str] | None,
        expand: Sequence[str] | None,
    ) -> SharePointListResource:
        if not isinstance(identifier, str) or not identifier:
            raise ValueError("list identifier cannot be empty")
        endpoint = f"{self._base_path}/{quote(identifier, safe='')}"
        query = ODataQuery(select=tuple(select or ()), expand=tuple(expand or ()))
        payload = self.transport.get(endpoint, params=query.to_params() or None)
        return SharePointListResource(self.client, self.site, _list_info(payload))


class SharePointListResource:
    """A list plus its stable item, column and version composition anchors."""

    def __init__(self, client: GraphBridgeClient, site: SiteResource, info: ListInfo) -> None:
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
        return self.info.id

    @property
    def display_name(self) -> str | None:
        return self.info.display_name

    @property
    def name(self) -> str | None:
        return self.info.name

    @property
    def metadata(self) -> Mapping[str, Any]:
        """Return all metadata, including unknown v1.0 properties, unchanged."""

        return self.info.raw

    def relationship(self, name: str) -> Any:
        """Read a relationship already retrieved with ``$expand``."""

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
        """Refresh metadata and optional v1.0 relationships in place."""

        refreshed = self.site.lists.get_by_id(self.id, select=select, expand=expand)
        self.info = refreshed.info
        return self

    def __repr__(self) -> str:
        return f"SharePointListResource(id={self.id!r}, display_name={self.info.display_name!r})"


def _list_info(payload: Any) -> ListInfo:
    if not isinstance(payload, Mapping):
        raise GraphInvalidResponseError("Microsoft Graph list response must be a JSON object")
    info = ListInfo.from_payload(payload)
    if not info.id:
        raise GraphInvalidResponseError("Microsoft Graph list response does not contain an id")
    return info

"""Stable Microsoft Graph v1.0 list item version operations."""

from __future__ import annotations

import builtins
from collections.abc import Iterator, Mapping
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

from ..exceptions import GraphInvalidResponseError
from ..models import ListItemVersion, Page
from ..pagination import iter_pages

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .lists import SharePointListResource


class VersionsResource:
    """Read retained list item versions and restore one as the current state.

    Version availability and retention depend on SharePoint list configuration
    and tenant policy. Restoring creates a new current version and does not
    remove the existing version history.
    """

    def __init__(
        self, client: GraphBridgeClient, sharepoint_list: SharePointListResource
    ) -> None:
        self.client = client
        self.transport = client.transport
        self.sharepoint_list = sharepoint_list

    def list(self, item_id: str) -> Page[ListItemVersion]:
        """Return the first page of versions retained for ``item_id``."""

        return next(self.iter_pages(item_id))

    def iter_pages(self, item_id: str) -> Iterator[Page[ListItemVersion]]:
        """Lazily traverse version pages, preserving Graph nextLink values."""

        return iter_pages(
            self.transport,
            self._versions_path(item_id),
            parser=self._version_from_payload,
        )

    def iter_all(self, item_id: str) -> Iterator[ListItemVersion]:
        for page in self.iter_pages(item_id):
            yield from page.items

    def versions(self, item_id: str) -> builtins.list[ListItemVersion]:
        """Return all retained versions for one list item."""

        return list(self.iter_all(item_id))

    def get(self, item_id: str, version_id: str) -> ListItemVersion:
        """Return one retained list item version."""

        self._validate_version_id(version_id)
        payload = self.transport.get(
            f"{self._versions_path(item_id)}/{quote(version_id, safe='')}"
        )
        return self._version_from_payload(payload)

    def restore_version(self, item_id: str, version_id: str) -> None:
        """Restore a version with the stable ``restoreVersion`` action."""

        self._validate_version_id(version_id)
        payload = self.transport.post(
            f"{self._versions_path(item_id)}/{quote(version_id, safe='')}/restoreVersion"
        )
        if payload is not None:
            raise GraphInvalidResponseError(
                "Microsoft Graph version restore response must be empty"
            )

    def _versions_path(self, item_id: str) -> str:
        if not isinstance(item_id, str) or not item_id:
            raise ValueError("item_id cannot be empty")
        site_id = quote(self.sharepoint_list.site.id, safe=",")
        list_id = quote(self.sharepoint_list.id, safe="")
        return (
            f"/sites/{site_id}/lists/{list_id}/items/"
            f"{quote(item_id, safe='')}/versions"
        )

    @staticmethod
    def _validate_version_id(version_id: str) -> None:
        if not isinstance(version_id, str) or not version_id:
            raise ValueError("version_id cannot be empty")

    @staticmethod
    def _version_from_payload(payload: Mapping[str, Any]) -> ListItemVersion:
        if not isinstance(payload, Mapping):
            raise GraphInvalidResponseError(
                "Microsoft Graph list item version must be a JSON object"
            )
        version = ListItemVersion.from_payload(payload)
        if not version.id:
            raise GraphInvalidResponseError(
                "Microsoft Graph list item version does not contain an id"
            )
        return version

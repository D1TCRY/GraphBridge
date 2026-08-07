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
    """Read and restore retained list-item versions.

    Version availability depends on SharePoint list configuration and retention
    policy. Restore uses the stable action and does not erase existing history.

    Args:
        client: Shared GraphBridge client.
        sharepoint_list: Parent SharePoint list.
    """

    def __init__(
        self, client: GraphBridgeClient, sharepoint_list: SharePointListResource
    ) -> None:
        """Initialize the versions resource.

        Args:
            client: Shared GraphBridge client.
            sharepoint_list: Parent SharePoint list.
        """
        self.client = client
        self.transport = client.transport
        self.sharepoint_list = sharepoint_list

    def list(self, item_id: str) -> Page[ListItemVersion]:
        """Return the first page of retained versions.

        Use the lazy iterators when version history may span multiple Graph pages.

        Args:
            item_id: Graph list-item identifier.
        """

        return next(self.iter_pages(item_id))

    def iter_pages(self, item_id: str) -> Iterator[Page[ListItemVersion]]:
        """Lazily iterate through version pages.

        Continuation links are handled by the shared pagination helper and are
        forwarded unchanged through the safe transport.

        Args:
            item_id: Graph list-item identifier.
        """

        return iter_pages(
            self.transport,
            self._versions_path(item_id),
            parser=self._version_from_payload,
        )

    def iter_all(self, item_id: str) -> Iterator[ListItemVersion]:
        """Lazily iterate through all retained versions.

        Individual parsed versions are yielded without loading the complete
        history into memory first.

        Args:
            item_id: Graph list-item identifier.
        """
        for page in self.iter_pages(item_id):
            yield from page.items

    def versions(self, item_id: str) -> builtins.list[ListItemVersion]:
        """Return all retained versions for one item.

        This convenience method materializes the lazy version iterator into a
        list for callers that need the complete history at once.

        Args:
            item_id: Graph list-item identifier.
        """

        return list(self.iter_all(item_id))

    def get(self, item_id: str, version_id: str) -> ListItemVersion:
        """Retrieve one retained item version.

        Both item and version identifiers are quoted as independent Graph path
        segments before the response is validated and parsed.

        Args:
            item_id: Graph list-item identifier.
            version_id: Version identifier.

        Raises:
            ValueError: If either identifier is empty.
        """

        self._validate_version_id(version_id)
        payload = self.transport.get(
            f"{self._versions_path(item_id)}/{quote(version_id, safe='')}"
        )
        return self._version_from_payload(payload)

    def restore_version(self, item_id: str, version_id: str) -> None:
        """Restore a retained version as the current item state.

        The stable Graph action creates a new current version while retaining the
        prior history, and a successful response must be empty.

        Args:
            item_id: Graph list-item identifier.
            version_id: Version identifier to restore.

        Raises:
            ValueError: If either identifier is empty.
            GraphInvalidResponseError: If Graph returns a non-empty success body.
        """

        self._validate_version_id(version_id)
        payload = self.transport.post(
            f"{self._versions_path(item_id)}/{quote(version_id, safe='')}/restoreVersion"
        )
        if payload is not None:
            raise GraphInvalidResponseError(
                "Microsoft Graph version restore response must be empty"
            )

    def _versions_path(self, item_id: str) -> str:
        """Build the Graph path for an item's versions.

        Args:
            item_id: Graph list-item identifier.

        Raises:
            ValueError: If the identifier is empty.
        """
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
        """Validate a version identifier.

        Args:
            version_id: Version identifier to validate.

        Raises:
            ValueError: If the identifier is empty.
        """
        if not isinstance(version_id, str) or not version_id:
            raise ValueError("version_id cannot be empty")

    @staticmethod
    def _version_from_payload(payload: Mapping[str, Any]) -> ListItemVersion:
        """Validate and parse a version payload.

        Args:
            payload: Decoded Graph response.

        Raises:
            GraphInvalidResponseError: If the payload is invalid or lacks an ID.
        """
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

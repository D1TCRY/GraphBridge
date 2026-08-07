"""SharePoint site resources."""

from __future__ import annotations

from collections.abc import Mapping
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

from ..exceptions import GraphInvalidResponseError
from ..models import SiteInfo

if TYPE_CHECKING:
    from ..client import GraphBridgeClient
    from .lists import ListsResource


class SitesResource:
    """Resolve and bind SharePoint sites.

    Args:
        client: Shared GraphBridge client.
    """

    def __init__(self, client: GraphBridgeClient) -> None:
        """Initialize the site collection resource.

        Args:
            client: Shared GraphBridge client.
        """
        self.client = client
        self.transport = client.transport

    def get_by_path(self, *, hostname: str, path: str) -> SiteResource:
        """Resolve a SharePoint site by hostname and path.

        Args:
            hostname: SharePoint tenant hostname.
            path: Server-relative site path.

        Raises:
            ValueError: If the hostname or path is empty.
        """
        if not hostname:
            raise ValueError("hostname cannot be empty")
        if not path.startswith("/"):
            path = f"/{path}"
        endpoint = f"/sites/{quote(hostname, safe='.-')}:{quote(path, safe='/')}"
        payload = self.transport.get(endpoint)
        return SiteResource(self.client, _site_info(payload), hostname=hostname, path=path)

    def get(self, site_id: str) -> SiteResource:
        """Retrieve a SharePoint site by Graph ID.

        Args:
            site_id: Graph site identifier.

        Raises:
            ValueError: If the identifier is empty.
        """
        if not site_id:
            raise ValueError("site_id cannot be empty")
        payload = self.transport.get(f"/sites/{quote(site_id, safe=',')}")
        return SiteResource(self.client, _site_info(payload))

    def bind(self, site: SiteInfo | Mapping[str, Any]) -> SiteResource:
        """Bind known site metadata without an HTTP request.

        Args:
            site: Existing site model or payload.

        Raises:
            ValueError: If the metadata does not contain an ID.
        """

        info = site if isinstance(site, SiteInfo) else SiteInfo.from_payload(site)
        if not info.id:
            raise ValueError("site metadata must contain an id")
        return SiteResource(self.client, info)


class SiteResource:
    """Represent one SharePoint site and its lists.

    Args:
        client: Shared GraphBridge client.
        info: Parsed site metadata.
        hostname: Optional hostname used to resolve the site.
        path: Optional path used to resolve the site.
    """

    def __init__(
        self,
        client: GraphBridgeClient,
        info: SiteInfo,
        *,
        hostname: str | None = None,
        path: str | None = None,
    ) -> None:
        """Initialize a bound site resource.

        Args:
            client: Shared GraphBridge client.
            info: Parsed site metadata.
            hostname: Optional hostname used for resolution.
            path: Optional path used for resolution.
        """
        from .lists import ListsResource

        self.client = client
        self.transport = client.transport
        self.info = info
        self.hostname = hostname
        self.path = path
        self.lists: ListsResource = ListsResource(client, self)

    @property
    def id(self) -> str:
        """Return the Graph site identifier."""
        return self.info.id

    def __repr__(self) -> str:
        """Return a concise site representation."""
        return f"SiteResource(id={self.id!r}, display_name={self.info.display_name!r})"


def _site_info(payload: Any) -> SiteInfo:
    """Validate and parse a site payload.

    Args:
        payload: Decoded Graph response.

    Raises:
        GraphInvalidResponseError: If the payload is invalid or lacks an ID.
    """
    if not isinstance(payload, Mapping):
        raise GraphInvalidResponseError("Microsoft Graph site response must be a JSON object")
    info = SiteInfo.from_payload(payload)
    if not info.id:
        raise GraphInvalidResponseError("Microsoft Graph site response does not contain an id")
    return info

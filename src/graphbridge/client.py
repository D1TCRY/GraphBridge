"""Public composed client for Microsoft Graph SharePoint resources."""

from __future__ import annotations

import time
from collections.abc import Callable

import requests

from .auth import GRAPH_SCOPE, TokenCredential
from .resources.sites import SitesResource
from .transport import (
    DEFAULT_BASE_URL,
    DEFAULT_MAX_RETRY_DELAY,
    DEFAULT_TIMEOUT,
    DEFAULT_USER_AGENT,
    GraphTransport,
)


class GraphBridgeClient:
    """Provide the composed entry point to GraphBridge resources.

    A client owns one transport and exposes site navigation through ``sites``.
    Every bound resource created from it reuses the same session and credential.
    New users should normally create one client per configuration, resolve a
    site through ``client.sites``, and keep that client alive for the complete
    unit of work so HTTP connections can be reused.

    Args:
        credential: Credential used when constructing a transport.
        transport: Optional preconfigured transport.
        session: Optional HTTP session reused by the new transport.
        base_url: Microsoft Graph v1.0 base URL.
        timeout: Connection and response timeout.
        user_agent: Value sent in the User-Agent header.
        max_retries: Maximum retries for eligible requests.
        backoff_factor: Base factor used for exponential backoff.
        max_retry_delay: Maximum delay between retries.
        sleep: Function used to wait between retries.
        scope: OAuth scope requested from the credential.
    """

    def __init__(
        self,
        *,
        credential: TokenCredential | None = None,
        transport: GraphTransport | None = None,
        session: requests.Session | None = None,
        base_url: str = DEFAULT_BASE_URL,
        timeout: float | tuple[float, float] = DEFAULT_TIMEOUT,
        user_agent: str = DEFAULT_USER_AGENT,
        max_retries: int = 3,
        backoff_factor: float = 0.5,
        max_retry_delay: float = DEFAULT_MAX_RETRY_DELAY,
        sleep: Callable[[float], None] = time.sleep,
        scope: str = GRAPH_SCOPE,
    ) -> None:
        """Initialize the client and its shared transport.

        Callers may inject a complete transport or let the client construct one
        from a credential, but those two construction modes cannot be mixed.
        Credential-based construction is the normal application path; transport
        injection is useful for tests or advanced shared-network configuration.

        Args:
            credential: Credential used when no transport is supplied.
            transport: Optional preconfigured transport.
            session: Optional HTTP session for a newly created transport.
            base_url: Microsoft Graph v1.0 base URL.
            timeout: Connection and response timeout.
            user_agent: Value sent in the User-Agent header.
            max_retries: Maximum retries for eligible requests.
            backoff_factor: Base factor used for retry delays.
            max_retry_delay: Maximum delay between retries.
            sleep: Function used to wait between retries.
            scope: OAuth scope requested from the credential.

        Raises:
            TypeError: If neither a credential nor a transport is supplied.
            ValueError: If incompatible construction options are combined.
        """
        if transport is not None:
            if credential is not None or session is not None:
                raise ValueError("credential and session cannot be supplied with transport")
            self.transport = transport
        else:
            if credential is None:
                raise TypeError("credential is required when transport is not supplied")
            self.transport = GraphTransport(
                credential,
                session=session,
                base_url=base_url,
                timeout=timeout,
                user_agent=user_agent,
                max_retries=max_retries,
                backoff_factor=backoff_factor,
                max_retry_delay=max_retry_delay,
                sleep=sleep,
                scope=scope,
            )
        self.sites = SitesResource(self)

    def close(self) -> None:
        """Close the shared HTTP transport.

        This releases the reusable HTTP session and should be called when the
        client is no longer needed. Closing does not revoke credentials or alter
        any SharePoint resource; it only releases local networking resources.
        """
        self.transport.close()

    def __enter__(self) -> GraphBridgeClient:
        """Return the client when entering a context manager.

        Context-manager use provides deterministic cleanup of the HTTP session.
        The returned object is the same client, so resources are accessed in the
        usual way inside the ``with`` block.
        """
        return self

    def __exit__(self, *_args: object) -> None:
        """Close the client when leaving a context manager.

        Cleanup occurs whether the managed block succeeds or raises an exception.
        Exceptions from the managed block are not suppressed.

        Args:
            *_args: Context-manager exception details, if any.
        """
        self.close()

    def __repr__(self) -> str:
        """Return a safe representation of the client.

        Authentication material is omitted; only the transport's safe summary is
        included. The result can be logged when diagnosing configuration without
        printing the injected credential or bearer token.
        """
        return f"GraphBridgeClient(transport={self.transport!r})"

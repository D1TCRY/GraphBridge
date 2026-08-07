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
        """Close the shared HTTP transport."""
        self.transport.close()

    def __enter__(self) -> GraphBridgeClient:
        """Return the client when entering a context manager."""
        return self

    def __exit__(self, *_args: object) -> None:
        """Close the client when leaving a context manager.

        Args:
            *_args: Context-manager exception details, if any.
        """
        self.close()

    def __repr__(self) -> str:
        """Return a safe representation of the client."""
        return f"GraphBridgeClient(transport={self.transport!r})"

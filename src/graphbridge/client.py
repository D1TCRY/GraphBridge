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
    """Entry point for the composed GraphBridge API."""

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
        self.transport.close()

    def __enter__(self) -> GraphBridgeClient:
        return self

    def __exit__(self, *_args: object) -> None:
        self.close()

    def __repr__(self) -> str:
        return f"GraphBridgeClient(transport={self.transport!r})"

"""Credential abstraction for Microsoft Graph authentication."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, Protocol, runtime_checkable

from .exceptions import GraphAuthenticationError

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


@runtime_checkable
class AccessToken(Protocol):
    token: str
    expires_on: int


@runtime_checkable
class TokenCredential(Protocol):
    """Structural subset implemented by Azure Identity credentials."""

    @property
    def get_token(self) -> Callable[..., Any]: ...


class GraphAuthenticator:
    """Acquire a fresh access token for each HTTP attempt without caching it."""

    def __init__(self, credential: TokenCredential, *, scope: str = GRAPH_SCOPE) -> None:
        if not hasattr(credential, "get_token"):
            raise TypeError("credential must provide get_token()")
        self._credential = credential
        self._scope = scope

    @property
    def scope(self) -> str:
        return self._scope

    def get_access_token(self) -> str:
        try:
            access_token = self._credential.get_token(self._scope)
            token = access_token.token
        except Exception:
            raise GraphAuthenticationError("Unable to acquire a Microsoft Graph access token") from None
        if not isinstance(token, str) or not token:
            raise GraphAuthenticationError("The credential returned an invalid Microsoft Graph access token")
        return token

    def __repr__(self) -> str:
        return "GraphAuthenticator(scope=<configured>, credential=<redacted>)"

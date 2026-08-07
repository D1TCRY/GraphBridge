"""Credential abstraction for Microsoft Graph authentication."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, Protocol, runtime_checkable

from .exceptions import GraphAuthenticationError

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


@runtime_checkable
class AccessToken(Protocol):
    """Describe the token value returned by a credential.

    Implementations expose both the bearer-token text and its expiration time,
    matching the small portion of Azure Identity used by GraphBridge.
    """

    token: str
    expires_on: int


@runtime_checkable
class TokenCredential(Protocol):
    """Describe the credential interface accepted by GraphBridge.

    Any credential with a compatible ``get_token`` callable can be injected, so
    callers are not restricted to one concrete Azure Identity credential type.
    """

    @property
    def get_token(self) -> Callable[..., Any]:
        """Return the callable that acquires an access token.

        The callable is expected to accept an OAuth scope and return an object
        compatible with :class:`AccessToken`.
        """

        ...


class GraphAuthenticator:
    """Acquire Microsoft Graph tokens without caching them locally.

    Token lifetime and renewal remain the credential's responsibility. Asking
    for a token on every HTTP attempt also allows retries to use renewed tokens.

    Args:
        credential: Credential that provides a ``get_token`` method.
        scope: OAuth scope requested for each access token.
    """

    def __init__(self, credential: TokenCredential, *, scope: str = GRAPH_SCOPE) -> None:
        """Initialize the authenticator.

        Only the structural credential contract is required; no network request
        is performed until :meth:`get_access_token` is called.

        Args:
            credential: Credential used to acquire access tokens.
            scope: OAuth scope requested from the credential.

        Raises:
            TypeError: If the credential does not provide ``get_token``.
        """
        if not hasattr(credential, "get_token"):
            raise TypeError("credential must provide get_token()")
        self._credential = credential
        self._scope = scope

    @property
    def scope(self) -> str:
        """Return the configured OAuth scope.

        This is normally Microsoft Graph's application ``.default`` scope.
        """
        return self._scope

    def get_access_token(self) -> str:
        """Acquire and validate a Microsoft Graph access token.

        Credential-specific failures are deliberately hidden so secrets or
        provider details cannot leak through the public exception message.

        Returns:
            The non-empty access token string.

        Raises:
            GraphAuthenticationError: If acquisition fails or returns an invalid token.
        """
        try:
            access_token = self._credential.get_token(self._scope)
            token = access_token.token
        except Exception:
            raise GraphAuthenticationError("Unable to acquire a Microsoft Graph access token") from None
        if not isinstance(token, str) or not token:
            raise GraphAuthenticationError("The credential returned an invalid Microsoft Graph access token")
        return token

    def __repr__(self) -> str:
        """Return a representation that does not expose credentials.

        The output indicates that authentication is configured while redacting
        the credential object itself.
        """
        return "GraphAuthenticator(scope=<configured>, credential=<redacted>)"

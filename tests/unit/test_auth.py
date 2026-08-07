from __future__ import annotations

from types import SimpleNamespace

import pytest
from conftest import ACCESS_TOKEN, CLIENT_ID, CLIENT_SECRET, TENANT_ID, FakeCredential

import graphbridge
import graphbridge.graph_bridge as graph_bridge_module
from graphbridge import GbAuth
from graphbridge.graph_bridge import deduplicate_dicts


def test_public_package_exports() -> None:
    assert {
        "__version__",
        "GraphBridgeClient",
        "GraphTransport",
        "TokenCredential",
        "ListItem",
        "SyncPlan",
        "GraphRequestError",
        "GbAuth",
        "GbSite",
        "GbList",
    } <= set(graphbridge.__all__)
    assert len(graphbridge.__all__) == len(set(graphbridge.__all__))
    assert all(hasattr(graphbridge, name) for name in graphbridge.__all__)
    assert graphbridge.GbAuth is GbAuth


def test_deduplicate_dicts_preserves_first_occurrence() -> None:
    first = {"b": 2, "a": 1}
    duplicate = {"a": 1, "b": 2}
    third = {"a": 3}

    assert deduplicate_dicts([first, duplicate, third]) == [first, third]


@pytest.mark.parametrize("field", ["tenant_id", "client_id", "client_secret"])
def test_auth_rejects_empty_values(field: str) -> None:
    values = {
        "tenant_id": TENANT_ID,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }
    values[field] = ""

    with pytest.raises(ValueError, match="cannot be empty"):
        GbAuth(**values)


@pytest.mark.parametrize("field", ["tenant_id", "client_id", "client_secret"])
def test_auth_rejects_non_string_values(field: str) -> None:
    values = {
        "tenant_id": TENANT_ID,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }
    values[field] = 123

    with pytest.raises(TypeError, match="must be string"):
        GbAuth(**values)


def test_credential_and_token_are_lazy_and_cached(auth: GbAuth) -> None:
    assert FakeCredential.instances == []

    credential = auth.credential
    assert auth.credential is credential
    assert credential.tenant_id == TENANT_ID
    assert credential.client_id == CLIENT_ID
    assert credential.client_secret == CLIENT_SECRET

    assert auth.token == ACCESS_TOKEN
    assert auth.token == ACCESS_TOKEN
    assert credential.get_token_calls == ["https://graph.microsoft.com/.default"]
    assert auth.headers == {"Authorization": f"Bearer {ACCESS_TOKEN}"}


@pytest.mark.parametrize(
    ("field", "new_value"),
    [
        ("tenant_id", "new-tenant"),
        ("client_id", "new-client"),
        ("client_secret", "new-secret"),
    ],
)
def test_changing_auth_input_invalidates_cached_credential_and_token(
    auth: GbAuth,
    field: str,
    new_value: str,
) -> None:
    old_credential = auth.credential
    assert auth.token == ACCESS_TOKEN

    setattr(auth, field, new_value)

    assert auth.credential is not old_credential
    assert auth.token == ACCESS_TOKEN
    assert len(FakeCredential.instances) == 2


def test_token_acquisition_error_is_wrapped(monkeypatch: pytest.MonkeyPatch) -> None:
    class BrokenCredential:
        def __init__(self, *_args: object) -> None:
            pass

        def get_token(self, _scope: str) -> SimpleNamespace:
            raise ValueError("simulated credential failure")

    monkeypatch.setattr(graph_bridge_module, "ClientSecretCredential", BrokenCredential)
    auth = GbAuth(TENANT_ID, CLIENT_ID, CLIENT_SECRET)

    with pytest.raises(RuntimeError, match="Failed to get token") as exc_info:
        _ = auth.token

    assert isinstance(exc_info.value.__cause__, ValueError)


def test_token_acquisition_error_redacts_credential_details(monkeypatch: pytest.MonkeyPatch) -> None:
    class BrokenCredential:
        def __init__(self, *_args: object) -> None:
            pass

        def get_token(self, _scope: str) -> SimpleNamespace:
            raise ValueError(f"failure contains {CLIENT_SECRET} and {ACCESS_TOKEN}")

    monkeypatch.setattr(graph_bridge_module, "ClientSecretCredential", BrokenCredential)
    auth = GbAuth(TENANT_ID, CLIENT_ID, CLIENT_SECRET)

    with pytest.raises(RuntimeError) as caught:
        _ = auth.token

    rendered = str(caught.value) + str(caught.value.__cause__)
    assert CLIENT_SECRET not in rendered
    assert ACCESS_TOKEN not in rendered


def test_auth_string_representations_redact_secret_and_token(auth: GbAuth) -> None:
    _ = auth.token

    rendered = f"{auth!s}\n{auth!r}"

    assert CLIENT_SECRET not in rendered
    assert ACCESS_TOKEN not in rendered
    assert "<redacted>" in rendered

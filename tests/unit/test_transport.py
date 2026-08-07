from __future__ import annotations

import json
from collections.abc import Iterator
from types import SimpleNamespace
from typing import Any

import pytest
import requests

from graphbridge.auth import GRAPH_SCOPE, GraphAuthenticator
from graphbridge.exceptions import (
    GraphAuthenticationError,
    GraphConflictError,
    GraphInvalidResponseError,
    GraphNetworkError,
    GraphNotFoundError,
    GraphPermissionError,
    GraphPreconditionFailedError,
    GraphRequestError,
    GraphServerError,
    GraphThrottlingError,
)
from graphbridge.transport import GraphTransport

TOKEN_ONE = "transport-token-one"
TOKEN_TWO = "transport-token-two"
BASE_URL = "https://graph.example.invalid/v1.0"


class SequenceCredential:
    def __init__(self, *tokens: str) -> None:
        self._tokens: Iterator[str] = iter(tokens)
        self.calls: list[tuple[tuple[str, ...], dict[str, Any]]] = []

    def get_token(self, *scopes: str, **kwargs: Any) -> SimpleNamespace:
        self.calls.append((scopes, kwargs))
        return SimpleNamespace(token=next(self._tokens), expires_on=9999999999)

    def __repr__(self) -> str:
        return "SequenceCredential(secret=credential-secret)"


class RecordingSession:
    def __init__(self, *actions: requests.Response | requests.RequestException) -> None:
        self.actions = list(actions)
        self.calls: list[tuple[str, str, dict[str, Any]]] = []
        self.closed = False

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        self.calls.append((method, url, kwargs))
        action = self.actions.pop(0)
        if isinstance(action, requests.RequestException):
            raise action
        return action

    def close(self) -> None:
        self.closed = True


def response(status: int, payload: object | None = None, *, text: str | None = None) -> requests.Response:
    result = requests.Response()
    result.status_code = status
    result.url = f"{BASE_URL}/test"
    result.headers = {}
    if text is not None:
        result._content = text.encode("utf-8")
    elif payload is None:
        result._content = b""
    else:
        result._content = json.dumps(payload).encode("utf-8")
        result.headers["Content-Type"] = "application/json"
    return result


def test_token_is_requested_for_every_attempt_and_never_cached() -> None:
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO)
    session = RecordingSession(response(200, {"page": 1}), response(200, {"page": 2}))
    transport = GraphTransport(credential, session=session, base_url=BASE_URL)

    assert transport.get("/first") == {"page": 1}
    assert transport.get("/second") == {"page": 2}

    assert credential.calls == [((GRAPH_SCOPE,), {}), ((GRAPH_SCOPE,), {})]
    assert session.calls[0][2]["headers"]["Authorization"] == f"Bearer {TOKEN_ONE}"
    assert session.calls[1][2]["headers"]["Authorization"] == f"Bearer {TOKEN_TWO}"
    assert TOKEN_ONE not in repr(transport)
    assert "credential-secret" not in repr(transport.authenticator)


def test_timeout_user_agent_json_and_absolute_url_are_forwarded_unchanged() -> None:
    credential = SequenceCredential(TOKEN_ONE)
    session = RecordingSession(response(200, {"ok": True}))
    transport = GraphTransport(
        credential,
        session=session,
        base_url=BASE_URL,
        timeout=(1.0, 9.0),
        user_agent="GraphBridge-tests/1",
    )
    absolute_url = f"{BASE_URL}/page?$skiptoken=opaque%2Bvalue"

    transport.post(
        absolute_url,
        params={"test": "one"},
        json={"Title": "Hello"},
        headers={"If-Match": "etag"},
    )

    method, url, kwargs = session.calls[0]
    assert (method, url) == ("POST", absolute_url)
    assert kwargs["timeout"] == (1.0, 9.0)
    assert kwargs["params"] == {"test": "one"}
    assert kwargs["json"] == {"Title": "Hello"}
    assert kwargs["headers"] == {
        "Accept": "application/json",
        "User-Agent": "GraphBridge-tests/1",
        "Content-Type": "application/json",
        "If-Match": "etag",
        "Authorization": f"Bearer {TOKEN_ONE}",
    }


def test_success_payload_and_repr_redact_a_reflected_access_token() -> None:
    credential = SequenceCredential(TOKEN_ONE)
    session = RecordingSession(response(200, {"echo": TOKEN_ONE, "nested": [TOKEN_ONE]}))
    transport = GraphTransport(
        credential,
        session=session,
        base_url=BASE_URL,
        user_agent=TOKEN_ONE,
    )

    payload = transport.get("/reflected")

    assert payload == {"echo": "<redacted>", "nested": ["<redacted>"]}
    assert TOKEN_ONE not in repr(transport)


@pytest.mark.parametrize(
    "base_url",
    [
        "graph.example.invalid/v1.0",
        "http://graph.example.invalid/v1.0",
        "https://user:password@graph.example.invalid/v1.0",
        "https://graph.example.invalid/beta",
        "https://graph.example.invalid/v1.0?secret=value",
    ],
)
def test_transport_rejects_unsafe_or_non_v1_base_urls(base_url: str) -> None:
    with pytest.raises(ValueError):
        GraphTransport(SequenceCredential(TOKEN_ONE), base_url=base_url)


@pytest.mark.parametrize("timeout", [0, -1, float("inf"), float("nan"), (1.0, 0)])
def test_transport_rejects_unbounded_or_invalid_timeouts(timeout: object) -> None:
    with pytest.raises((TypeError, ValueError)):
        GraphTransport(SequenceCredential(TOKEN_ONE), timeout=timeout)  # type: ignore[arg-type]


def test_transport_rejects_unbounded_retry_configuration() -> None:
    with pytest.raises(TypeError):
        GraphTransport(SequenceCredential(TOKEN_ONE), max_retries=True)  # type: ignore[arg-type]
    with pytest.raises(ValueError):
        GraphTransport(SequenceCredential(TOKEN_ONE), backoff_factor=float("inf"))
    with pytest.raises(ValueError):
        GraphTransport(SequenceCredential(TOKEN_ONE), max_retries=11)
    with pytest.raises(ValueError):
        GraphTransport(SequenceCredential(TOKEN_ONE), max_retry_delay=float("inf"))


def test_absolute_urls_cannot_escape_the_configured_v1_root() -> None:
    transport = GraphTransport(
        SequenceCredential(TOKEN_ONE),
        session=RecordingSession(),
        base_url=BASE_URL,
    )

    with pytest.raises(GraphInvalidResponseError, match="configured base URL"):
        transport.get("https://attacker.example.invalid/v1.0/items")
    with pytest.raises(GraphInvalidResponseError, match="configured base URL"):
        transport.get("https://graph.example.invalid/beta/items")
    with pytest.raises(GraphInvalidResponseError, match="configured base URL"):
        transport.get("HTTPS://attacker.example.invalid/v1.0/items")
    with pytest.raises(GraphInvalidResponseError, match="configured base URL"):
        transport.get("../beta/items")
    with pytest.raises(GraphInvalidResponseError, match="configured base URL"):
        transport.get("https://graph.example.invalid/v1.0/%2e%2e/beta/items")

    assert transport.session.calls == []


def test_network_errors_retry_safe_requests_with_injected_sleep() -> None:
    sleeps: list[float] = []
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO)
    session = RecordingSession(requests.ConnectionError("sensitive details"), response(200, {"ok": True}))
    transport = GraphTransport(
        credential,
        session=session,
        max_retries=1,
        backoff_factor=0.25,
        sleep=sleeps.append,
    )

    assert transport.get("/retry") == {"ok": True}
    assert sleeps == [0.25]
    assert len(session.calls) == 2


def test_network_error_is_sanitized_after_retry_budget_is_exhausted() -> None:
    credential = SequenceCredential(TOKEN_ONE)
    session = RecordingSession(requests.ConnectionError("secret network detail"))
    transport = GraphTransport(credential, session=session, max_retries=0)

    with pytest.raises(GraphNetworkError) as caught:
        transport.get("/failure")

    assert "secret network detail" not in str(caught.value)


def test_retry_after_seconds_is_respected_and_token_is_renewed() -> None:
    throttled = response(429, {"error": {"code": "tooManyRequests", "message": "slow down"}})
    throttled.headers["Retry-After"] = "2"
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO)
    session = RecordingSession(throttled, response(200, {"ok": True}))
    sleeps: list[float] = []
    transport = GraphTransport(credential, session=session, max_retries=1, sleep=sleeps.append)

    assert transport.get("/throttled") == {"ok": True}
    assert sleeps == [2.0]
    assert len(credential.calls) == 2


def test_retry_after_is_capped_to_prevent_unbounded_sleep() -> None:
    throttled = response(429, {"error": {"message": "slow down"}})
    throttled.headers["Retry-After"] = "999999"
    sleeps: list[float] = []
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO)
    transport = GraphTransport(
        credential,
        session=RecordingSession(throttled, response(200, {"ok": True})),
        max_retries=1,
        max_retry_delay=5,
        sleep=sleeps.append,
    )

    assert transport.get("/limited-retry") == {"ok": True}
    assert sleeps == [5.0]
    assert len(credential.calls) == 2


def test_server_errors_retry_safe_requests() -> None:
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO)
    session = RecordingSession(
        response(503, {"error": {"message": "temporarily unavailable"}}),
        response(200, {"ok": True}),
    )
    sleeps: list[float] = []
    transport = GraphTransport(
        credential,
        session=session,
        max_retries=1,
        backoff_factor=0.1,
        sleep=sleeps.append,
    )

    assert transport.get("/temporary-failure") == {"ok": True}
    assert sleeps == [0.1]
    assert len(session.calls) == 2


def test_non_idempotent_request_is_not_replayed_by_default() -> None:
    credential = SequenceCredential(TOKEN_ONE)
    session = RecordingSession(response(503, {"error": {"message": "try later"}}))
    sleeps: list[float] = []
    transport = GraphTransport(credential, session=session, max_retries=5, sleep=sleeps.append)

    with pytest.raises(GraphServerError):
        transport.post("/items", json={"fields": {"Title": "Once"}})

    assert len(session.calls) == 1
    assert sleeps == []


@pytest.mark.parametrize(
    ("status", "expected"),
    [
        (401, GraphAuthenticationError),
        (403, GraphPermissionError),
        (404, GraphNotFoundError),
        (409, GraphConflictError),
        (412, GraphPreconditionFailedError),
        (429, GraphThrottlingError),
        (500, GraphServerError),
        (400, GraphRequestError),
    ],
)
def test_http_statuses_map_to_typed_errors_and_redact_tokens(
    status: int,
    expected: type[Exception],
) -> None:
    payload = {
        "error": {
            "code": f"code-{TOKEN_ONE}",
            "message": f"server reflected {TOKEN_ONE}",
            "innerError": {"request-id": TOKEN_ONE, "nested": [TOKEN_ONE]},
        }
    }
    credential = SequenceCredential(TOKEN_ONE)
    reflected = response(status, payload)
    reflected.headers["X-Reflected-Token"] = TOKEN_ONE
    transport = GraphTransport(
        credential,
        session=RecordingSession(reflected),
        max_retries=0,
    )

    with pytest.raises(expected) as caught:
        transport.get("/typed-error")

    rendered = repr(caught.value.__dict__) + str(caught.value)
    assert TOKEN_ONE not in rendered
    assert getattr(caught.value, "status_code") == status
    assert getattr(caught.value, "error").status_code == status


def test_successful_empty_and_invalid_json_responses_are_distinguished() -> None:
    credential = SequenceCredential(TOKEN_ONE, TOKEN_TWO, "transport-token-three")
    session = RecordingSession(response(200, {"replaced": True}), response(204), response(200, text="not-json"))
    transport = GraphTransport(credential, session=session)

    assert transport.put("/items/1", json={"Title": "Replacement"}) == {"replaced": True}
    assert transport.delete("/items/1") is None
    with pytest.raises(GraphInvalidResponseError):
        transport.get("/invalid-json")


def test_invalid_credential_and_acquisition_failure_are_safe() -> None:
    with pytest.raises(TypeError):
        GraphAuthenticator(object())  # type: ignore[arg-type]

    class BrokenCredential:
        def get_token(self, *_scopes: str) -> None:
            raise RuntimeError("real-client-secret")

    authenticator = GraphAuthenticator(BrokenCredential())  # type: ignore[arg-type]
    with pytest.raises(GraphAuthenticationError) as caught:
        authenticator.get_access_token()
    assert "real-client-secret" not in str(caught.value)


def test_context_manager_closes_injected_session() -> None:
    session = RecordingSession()
    with GraphTransport(SequenceCredential(TOKEN_ONE), session=session) as transport:
        assert transport.session is session
    assert session.closed is True

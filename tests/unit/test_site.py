from __future__ import annotations

import pytest
import responses
from conftest import ACCESS_TOKEN, CLIENT_ID, CLIENT_SECRET, HOSTNAME, SITE_ID, SITE_PATH, TENANT_ID

from graphbridge import GbSite


def test_site_can_be_built_from_auth_keywords() -> None:
    site = GbSite(
        hostname=HOSTNAME,
        site_path=SITE_PATH,
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
    )

    assert site.hostname == HOSTNAME
    assert site.site_path == SITE_PATH
    assert site.tenant_id == TENANT_ID


def test_site_rejects_invalid_auth_object() -> None:
    with pytest.raises(TypeError, match="sp_auth"):
        GbSite(hostname=HOSTNAME, site_path=SITE_PATH, gb_auth=object())


@pytest.mark.parametrize(("field", "value", "error"), [("hostname", "", ValueError), ("site_path", "", ValueError), ("hostname", 1, TypeError), ("site_path", 1, TypeError)])
def test_site_validates_location(field: str, value: object, error: type[Exception]) -> None:
    values = {"hostname": HOSTNAME, "site_path": SITE_PATH}
    values[field] = value

    with pytest.raises(error):
        GbSite(
            **values,
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
        )


def test_site_url_shape(site: GbSite) -> None:
    assert site.site_url == f"https://graph.microsoft.com/v1.0/sites/{HOSTNAME}:{SITE_PATH}"
    assert CLIENT_SECRET not in repr(site)
    assert "<redacted>" in repr(site)


@responses.activate
def test_site_data_is_requested_once_and_cached(site: GbSite, fixture_json: object) -> None:
    site_payload = fixture_json("site.json")
    responses.get(site.site_url, json=site_payload, status=200)

    assert site.site_data == site_payload
    assert site.site_data is site.site_data
    assert site.site_id == SITE_ID
    assert len(responses.calls) == 1

    request = responses.calls[0].request
    assert request.method == "GET"
    assert request.url == site.site_url
    assert request.headers["Authorization"] == f"Bearer {ACCESS_TOKEN}"
    assert str(site) == f"< GbSite | Hostname: {HOSTNAME}, Site Path: {SITE_PATH}, Site ID: {SITE_ID} >"


@responses.activate
def test_site_data_error_raises_runtime_error(site: GbSite) -> None:
    responses.get(site.site_url, body="simulated forbidden", status=403)

    with pytest.raises(RuntimeError, match="403 simulated forbidden"):
        _ = site.site_data


@responses.activate
def test_missing_site_id_returns_legacy_warning(site: GbSite) -> None:
    responses.get(site.site_url, json={}, status=200)

    assert site.site_id == "<WARNING SPM | Site ID not found>"


@responses.activate
def test_site_cache_is_not_invalidated_when_location_changes(site: GbSite) -> None:
    responses.get(site.site_url, json={"id": SITE_ID}, status=200)
    assert site.site_id == SITE_ID

    site.hostname = "other.example.invalid"

    # Legacy behavior documented for the future refactor: setters do not clear site_data.
    assert site.site_id == SITE_ID
    assert len(responses.calls) == 1

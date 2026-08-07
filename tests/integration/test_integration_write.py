"""Explicitly enabled write check that cleans up only its own created item."""

from uuid import uuid4

import pytest

from graphbridge.resources import SharePointListResource

pytestmark = [pytest.mark.integration, pytest.mark.integration_write]


def test_create_update_and_remove_only_the_created_item(
    dedicated_list: SharePointListResource,
    integration_writes_enabled: None,
) -> None:
    marker = f"GraphBridge integration {uuid4()}"
    created = dedicated_list.items.create({"Title": marker})
    if not created.id:
        pytest.fail("Graph did not return the ID needed for safe cleanup", pytrace=False)

    try:
        current = dedicated_list.items.get(created.id, fields=("Title",))
        if current.etag is None:
            pytest.fail("Graph did not return an eTag for concurrency testing", pytrace=False)
        updated = dedicated_list.items.update(
            created.id,
            {"Title": f"{marker} updated"},
            etag=current.etag,
        )
        assert updated.id == created.id
    finally:
        # Cleanup is intentionally scoped to the ID created in this test.
        dedicated_list.items.delete(created.id, etag="*")

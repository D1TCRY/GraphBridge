"""Read-only smoke checks for a dedicated Microsoft 365 test environment."""

import pytest

from graphbridge.resources import SharePointListResource

pytestmark = pytest.mark.integration


def test_dedicated_list_can_read_items_and_schema(
    dedicated_list: SharePointListResource,
) -> None:
    dedicated_list.items.list(top=1)
    dedicated_list.columns.list(top=1)

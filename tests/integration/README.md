# Opt-in integration tests

These tests make real Microsoft Graph calls. They are skipped unless
`--run-integration` is supplied and all dedicated-environment variables are
configured. Never point them at a production site or list.

Required for read-only tests:

```text
AZURE_TENANT_ID
AZURE_CLIENT_ID
AZURE_CLIENT_SECRET
GRAPHBRIDGE_INTEGRATION_SITE_ID
GRAPHBRIDGE_INTEGRATION_SITE_NAME=GraphBridge Integration - <test site>
GRAPHBRIDGE_INTEGRATION_LIST_ID
GRAPHBRIDGE_INTEGRATION_LIST_NAME=GraphBridge Integration - <test list>
GRAPHBRIDGE_INTEGRATION_CONFIRM=GRAPHBRIDGE_DEDICATED_TEST_ENVIRONMENT
```

Run only after reviewing the target:

```bash
pytest --run-integration -m "integration and not integration_write" tests/integration
```

The write test additionally requires
`GRAPHBRIDGE_INTEGRATION_ALLOW_WRITES=CREATE_UPDATE_DELETE_OWN_ITEMS`. It creates
one uniquely named item, updates it, and deletes only that same returned item ID
in cleanup. It never creates or removes a site, list, or column. A process crash
can leave that one marked item behind, so inspect the dedicated list after a
failed run.

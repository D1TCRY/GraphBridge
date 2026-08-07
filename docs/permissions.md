# Microsoft Graph permissions

GraphBridge uses app-only authentication and the Microsoft Graph `.default`
scope. The application permissions below require Microsoft Entra administrator
consent. Grant only the permission and resource role needed by the workload.

This document describes authorization; GraphBridge does not create Entra
consents or SharePoint permission assignments.

## `Sites.Read.All`

Tenant-wide read access to documents and list items in all site collections.
Use it for site/list discovery, item and column reads, OData queries, pagination,
delta reads, and version-history reads. It cannot create/update/delete list items
or change list schema.

One edge case from the Graph documentation: reading non-approved items from a
list with content approval enabled can require `Sites.Manage.All`. Treat that as
an exceptional workload requirement, not a default.

## `Sites.ReadWrite.All`

Tenant-wide create/read/update/delete access to documents and list items. It is
the broad application permission for item CRUD, batch item operations,
`sync.apply()`, item deletion, and version restore. It is unnecessary for a
read-only process.

This permission covers all site collections, so prefer a Selected permission
when the application works with one controlled site or list.

## `Sites.Manage.All`

Tenant-wide management of lists and document libraries. The stable Graph
endpoints identify it as the least-privileged application permission for
creating a list and for creating, updating, or deleting a column definition.
Schema changes alter the contract for every item and consumer, so Graph separates
them from ordinary item writes.

Use it for `site.lists.create()` and list column create/update/delete only when a
resource-scoped Selected design cannot satisfy the deployment. Reading columns
does not by itself require this permission.

## `Sites.Selected`

Resource-scoped access at a selected site collection. Entra consent alone grants
no site access. An authorized administrator must also assign the application a
role on each intended site. Available resource roles are `read`, `write`,
`owner`, and `fullcontrol`.

- `read` fits item/schema reads and delta/version reads.
- `write` fits normal content/item CRUD when supported by the target operation.
- schema administration and list creation need the appropriate management-level
  resource role, normally `fullcontrol`, and must be validated in the target
  tenant against the v1.0 endpoint.

Site-level assignment is the appropriate Selected boundary when the app must
create a list, because a list-level assignment cannot exist before the list.

## `Lists.SelectedOperations.Selected`

Resource-scoped access at an individual list. Entra consent initially grants no
access. The application must be assigned a role on the specific list through
`POST /sites/{site-id}/lists/{list-id}/permissions`, and its token must contain
the selected permission. The same `read`, `write`, `owner`, and `fullcontrol`
roles apply.

This is the narrowest natural scope for a worker that reads or synchronizes one
existing list. It cannot authorize creating that list. Assigning permissions at
list/list-item/file level breaks permission inheritance and consumes unique
permission capacity, so account for SharePoint service limits.

## Entra consent versus object assignment

Selected authorization has two independent control planes plus the token:

1. an administrator consents to `Sites.Selected` or
   `Lists.SelectedOperations.Selected` in Microsoft Entra ID;
2. an authorized owner/administrator grants the application a role on the exact
   site or list object;
3. the application obtains a token containing the consented Selected scope.

Missing any step means no access. Revoking Entra consent disables access to all
previously assigned objects; deleting one object assignment removes access only
to that object. Do not put real application, tenant, site, or list identifiers in
source control. Use placeholders such as `{application-id}` in runbooks.

## Recommended configurations

| Workload | Preferred least-privilege design | Broad fallback |
|---|---|---|
| Read one existing list | `Lists.SelectedOperations.Selected` + `read` on that list | `Sites.Read.All` |
| Read several lists in one site | `Sites.Selected` + `read` on that site | `Sites.Read.All` |
| Item CRUD or safe synchronization on one list | `Lists.SelectedOperations.Selected` + `write` on that list | `Sites.ReadWrite.All` |
| Item CRUD across one controlled site | `Sites.Selected` + `write` on that site | `Sites.ReadWrite.All` |
| Create lists or administer schema in one controlled site | `Sites.Selected` + management-capable role (normally `fullcontrol`), validated against the operation | `Sites.Manage.All` |

Never present `Sites.Manage.All`, `Sites.ReadWrite.All`, or full control as the
default for read-only use. Keep the integration-test application and its
dedicated site/list separate from production.

## Primary Microsoft documentation

- [Permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Selected permissions overview](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)
- [Create a list](https://learn.microsoft.com/en-us/graph/api/list-create?view=graph-rest-1.0)
- [Create a list column](https://learn.microsoft.com/en-us/graph/api/list-post-columns?view=graph-rest-1.0)
- [Update a column](https://learn.microsoft.com/en-us/graph/api/columndefinition-update?view=graph-rest-1.0)
- [Create a list item](https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0)
- [Update a list item](https://learn.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0)

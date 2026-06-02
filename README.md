# SharePoint EDC Connector for Simba Intelligence

A native Enterprise Data Connector (EDC) that connects Microsoft SharePoint
to [Simba Intelligence](https://insightsoftware.com/simba-intelligence/) and
Logi Composer via the Microsoft Graph API.

Built from the same Zoomdata EDC template as
[edc-graphql](https://github.com/isw-da/edc-graphql), upgraded to the
production stack (edc-api 25.4.0, Thrift 0.21.0, Spring Boot 3.2.5, Java 17).

## What it does

- Connects to **SharePoint Online** via Microsoft Graph (app-only, client
  credentials flow). No user-interactive sign-in, no MFA prompts.
- Auto-discovers every visible SharePoint **List** in the configured site
  and exposes it as an EDC collection.
- **OData $filter pushdown** for Lists: EQ / GT / GE / LT / LE /
  CONTAINS / STARTS_WITH / ENDS_WITH / IS_NULL are pushed to Graph. The QE
  handles anything more complex locally.
- Registers with Composer via Consul (`subStorageType=SHAREPOINT`).
- Queries through SI Playground using natural language.

## Tested against

| Surface | Auth | Status |
|-----|---------|--------|
| SharePoint Online Lists | Service principal | v1, primary target |
| Excel-in-SharePoint (tables, sheets, named ranges) | Service principal | v1.1, validated E2E |
| SharePoint on-prem | n/a | Out of scope |

### Excel-in-SharePoint (v1.1)

Set `INCLUDE_EXCEL` to a comma-separated list of `.xlsx` paths in the site's
default document library (e.g. `/sophos-test.xlsx, /reports/agents.xlsx`).
Each file contributes up to three kinds of collection:

| Collection name | Source | Header row |
|---|---|---|
| `excel_table__<file>__<tableName>` | every named Excel table | yes (table header) |
| `excel_sheet__<file>__<sheetName>` | every visible worksheet's used-range | yes (first row) |
| `excel_range__<file>__<rangeName>` | every visible named range | no (synthetic column names) |

Column types are inferred from the data (integers, doubles, strings). Cell
values are read via the Graph Workbook API as raw 2D arrays (the typed SDK
cannot represent them). No filter pushdown for Excel — the QE filters
in-memory, so keep workbook files to a reasonable size (tens of thousands of
rows). Dates may appear as Excel serial numbers depending on cell formatting.

## Quick start

### Prerequisites

- Java 17
- Maven 3.x
- Docker (for containerised deployment)
- A running Simba Intelligence instance ([setup guide](https://github.com/isw-da/simba-intelligence-skill))
- An Azure app registration with Graph application permissions (see below)

### Azure app registration

In the target tenant's Azure portal:

1. **Microsoft Entra ID > App registrations > New registration.** Give it a
   name like `simba-intelligence-sharepoint-edc`. Leave redirect URI blank
   (app-only flow).
2. **Certificates & secrets > New client secret.** Save the value
   immediately; you cannot see it again.
3. **API permissions > Add a permission > Microsoft Graph > Application
   permissions.** Add at minimum:
   - `Sites.Read.All` (read SharePoint sites and Lists)
   - `Files.Read.All` (required if you will use INCLUDE_EXCEL when v1.1 ships)
4. **Grant admin consent** for the tenant. Application permissions need
   admin consent before they can be used.
5. Note the values you'll paste into Composer:
   - Tenant ID (Overview page)
   - Application (client) ID (Overview page)
   - Client secret (from step 2)

**Tighter security alternative:** use `Sites.Selected` instead of
`Sites.Read.All`, then grant the app access to specific sites via
[Graph site permissions](https://learn.microsoft.com/graph/api/site-post-permissions).
Tenant-wide read becomes per-site read, at the cost of one provisioning
step per site.

### Build

```bash
export JAVA_HOME=/path/to/jdk-17
mvn clean package -Dlicense.skip=true -DskipTests
```

### Run locally

```bash
java -Duser.timezone=UTC -jar target/connector-server-sharepoint-1.0.0-exec.jar
```

Server starts on port 7339 at `/connector/`.

### Deploy to Kubernetes

```bash
# Build and load Docker image
docker build -t edc-sharepoint:latest .
# For kind:
kind load docker-image edc-sharepoint:latest --name simba-intel-lab
# For cloud: push to your container registry

# Deploy pod and service
kubectl apply -n simba-intel -f - <<EOF
apiVersion: v1
kind: Pod
metadata:
  name: edc-sharepoint
  labels:
    app: edc-sharepoint
spec:
  securityContext:
    runAsNonRoot: true
    runAsUser: 1000
    runAsGroup: 1000
    fsGroup: 1000
  containers:
  - name: edc-sharepoint
    image: edc-sharepoint:latest
    imagePullPolicy: Never
    ports:
    - containerPort: 7339
    resources:
      requests:
        cpu: "100m"
        memory: "256Mi"
      limits:
        cpu: "1000m"
        memory: "1Gi"
    readinessProbe:
      httpGet:
        path: /actuator/health
        port: 7339
      initialDelaySeconds: 5
      periodSeconds: 5
      timeoutSeconds: 2
      failureThreshold: 3
    livenessProbe:
      httpGet:
        path: /actuator/health
        port: 7339
      initialDelaySeconds: 30
      periodSeconds: 15
      timeoutSeconds: 3
      failureThreshold: 3
    securityContext:
      allowPrivilegeEscalation: false
      readOnlyRootFilesystem: false  # Jetty needs to write to /tmp/jetty-docbase.*
      capabilities:
        drop:
        - ALL
---
apiVersion: v1
kind: Service
metadata:
  name: edc-sharepoint
spec:
  selector:
    app: edc-sharepoint
  ports:
  - port: 7339
    targetPort: 7339
EOF
```

### Register in Composer

```bash
# 1. Register in Consul
kubectl -n simba-intel exec si-consul-server-0 -c consul -- \
  consul services register \
    -name=edc-sharepoint \
    -address=edc-sharepoint.simba-intel.svc.cluster.local \
    -port=7339

# 2. Register in Composer
curl -s -X POST "http://localhost:8080/discovery/api/connectors" \
  -u "admin:<password>" \
  -H "Content-Type: application/vnd.composer.v3+json" \
  -d '{
    "name": "SharePoint",
    "type": "DISCOVERY",
    "params": {
      "SERVICE_NAME": "edc-sharepoint",
      "BEHIND_GATEWAY": "false"
    }
  }'
```

### Create a connection

In the SI UI: **Connections > Create > SharePoint**

| Parameter | Required | Description |
|-----------|----------|-------------|
| TENANT_ID | Yes | Azure AD tenant ID (GUID or `contoso.onmicrosoft.com`) |
| CLIENT_ID | Yes | App registration client (application) ID |
| CLIENT_SECRET | Yes | App registration client secret |
| SITE_URL | Yes | SharePoint site URL, e.g. `https://contoso.sharepoint.com/sites/sales`. One site per connection. |
| INCLUDE_LISTS | No | Comma-separated allowlist of List displayNames. Empty = all visible Lists. |
| INCLUDE_EXCEL | No | Comma-separated paths to .xlsx files in the site drive (e.g. `/sophos-test.xlsx`). Each file's tables, sheets, and named ranges become collections. See "Excel-in-SharePoint" above. |
| AUTHORITY | No | Override `https://login.microsoftonline.com` for sovereign clouds (US Gov, China). |

## Architecture

```
Composer QE  -->  Thrift RPC      -->  SharePoint EDC  -->  Microsoft Graph  -->  SharePoint Online
                  /connector/          (this repo)          https://graph.microsoft.com
                                                            (client-credentials, app-only)
```

The connector implements `ConnectorService.Iface` (Thrift) and translates
EDC requests into Graph SDK calls. List columns are discovered via
`/sites/{id}/lists/{id}/columns`; rows are fetched via
`/sites/{id}/lists/{id}/items?$expand=fields(...)` with OData
`$select`/`$filter`/`$top` pushed down where supported.

## Known limitations

- **Excel-in-SharePoint:** implemented in v1.1 (tables, worksheet
  used-ranges, named ranges) via the Graph Workbook API. No filter or
  aggregation pushdown — the whole range is read and the QE filters
  in-memory; keep files to tens of thousands of rows. Dates may surface as
  Excel serial numbers depending on cell formatting.
- **One site per connection.** SITE_URL is single-valued. Multi-site is a
  v1.1 candidate; for now create one Composer connection per SharePoint
  site.
- **Lookup columns:** returned as the lookup value's string representation.
  No automatic FK relations are emitted in `describeSchemas`.
- **List >5000 items + non-indexed filter:** Graph honours these with the
  `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly` header, which the
  connector sets by default. Some queries may still fail with
  `tooManyItems`; index the column in SharePoint to fix.
- **Date pushdown:** datetime columns are typed as DATE for downstream QE
  use but filter pushdown for date comparisons hasn't been hardened on
  ISO 8601 strings yet. If date filters misbehave, treat the column as
  STRING via a custom view.
- **No pushdown for aggregations.** All aggregation is QE-side; this
  matches the GraphQL EDC and is constrained by the EDC framework.

## Related

Other Logi Symphony / Simba Intelligence developer toolkit components in
the same org:

- **[edc-graphql](https://github.com/isw-da/edc-graphql)** — GraphQL EDC
  built from the same Zoomdata template.
- **[simba-intelligence-skill](https://github.com/isw-da/simba-intelligence-skill)** —
  Install, configure, and troubleshoot Simba Intelligence. Includes a
  comprehensive guide for building custom EDC connectors at
  `references/custom-edc-build.md`.
- **[composer-mcp](https://github.com/isw-da/composer-mcp)** — MCP server
  wrapping the Composer REST API as 30+ tools.

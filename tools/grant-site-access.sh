#!/usr/bin/env bash
#
# grant-site-access.sh — grant the SharePoint EDC app read/write access to a
# SINGLE site, the tight "Sites.Selected" model.
#
# Why: with the app granted Graph application permission **Sites.Selected**
# (instead of the tenant-wide Sites.Read.All), it has NO access to any site
# until a SharePoint admin explicitly grants it per site. This script performs
# that per-site grant. Security teams prefer this: the connector can only read
# the exact sites you list, nothing else.
#
# Prerequisites:
#   - Azure CLI (az), logged in as a user with the SharePoint Administrator
#     (or Global Administrator) role in the target tenant:
#         az login --tenant <tenant>
#   - The EDC app registration already has the Graph application permission
#     "Sites.Selected" (admin-consented). This script grants it a site role;
#     it does NOT change the app's API permissions.
#
# Usage:
#   ./grant-site-access.sh \
#       --site-url https://contoso.sharepoint.com/sites/sales \
#       --app-id   <edc-app-client-id> \
#       [--role read|write|fullcontrol]   (default: read)
#
# Notes:
#   - "read" is sufficient for the connector (it only reads).
#   - The running identity needs Sites.FullControl.All (delegated) or the
#     SharePoint admin role; a plain user token will get HTTP 403 and the
#     script will say so.
#
set -euo pipefail

ROLE="read"
SITE_URL=""
APP_ID=""
APP_NAME="SharePoint EDC"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --site-url) SITE_URL="$2"; shift 2 ;;
    --app-id)   APP_ID="$2"; shift 2 ;;
    --role)     ROLE="$2"; shift 2 ;;
    --app-name) APP_NAME="$2"; shift 2 ;;
    -h|--help)  grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "Unknown arg: $1" >&2; exit 2 ;;
  esac
done

[[ -z "$SITE_URL" ]] && { echo "ERROR: --site-url is required" >&2; exit 2; }
[[ -z "$APP_ID"  ]] && { echo "ERROR: --app-id is required (the EDC app client id)" >&2; exit 2; }
case "$ROLE" in read|write|fullcontrol) ;; *) echo "ERROR: --role must be read|write|fullcontrol" >&2; exit 2 ;; esac

command -v az >/dev/null || { echo "ERROR: az (Azure CLI) not found" >&2; exit 3; }

# Host + server-relative path from the site URL.
HOST="$(printf '%s' "$SITE_URL" | sed -E 's#^https?://([^/]+).*#\1#')"
SPATH="$(printf '%s' "$SITE_URL" | sed -E 's#^https?://[^/]+##; s#/$##')"
if [[ -z "$SPATH" || "$SPATH" == "/" ]]; then SITE_ADDR="$HOST"; else SITE_ADDR="${HOST}:${SPATH}:"; fi

echo "Acquiring Graph token via az (run as a SharePoint/Global admin)..."
TOKEN="$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)"

echo "Resolving site: $SITE_URL"
SITE_ID="$(curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/sites/${SITE_ADDR}?\$select=id,webUrl" \
  | python3 -c 'import sys,json; print(json.load(sys.stdin).get("id",""))')"
[[ -z "$SITE_ID" ]] && { echo "ERROR: could not resolve site '$SITE_URL' (check the URL and your admin rights)" >&2; exit 4; }
echo "  siteId: $SITE_ID"

echo "Granting role '$ROLE' to app $APP_ID on the site..."
BODY="$(python3 - "$ROLE" "$APP_ID" "$APP_NAME" <<'PY'
import sys, json
role, app_id, app_name = sys.argv[1], sys.argv[2], sys.argv[3]
print(json.dumps({"roles":[role],
                  "grantedToIdentities":[{"application":{"id":app_id,"displayName":app_name}}]}))
PY
)"
HTTP="$(curl -s -o /tmp/grant-resp.json -w '%{http_code}' \
  -X POST "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions" \
  -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d "$BODY")"

if [[ "$HTTP" == "200" || "$HTTP" == "201" ]]; then
  GID="$(python3 -c 'import json;print(json.load(open("/tmp/grant-resp.json")).get("id",""))')"
  echo "SUCCESS: granted '$ROLE' to app $APP_ID on $SITE_URL (permission id: $GID)"
elif [[ "$HTTP" == "403" ]]; then
  echo "ERROR: HTTP 403 — the identity running this lacks rights to grant site permissions." >&2
  echo "       Run as a SharePoint Administrator / Global Administrator, or use an app with" >&2
  echo "       Graph application permission Sites.FullControl.All." >&2
  exit 5
else
  echo "ERROR: grant failed (HTTP $HTTP):" >&2
  head -c 400 /tmp/grant-resp.json >&2; echo >&2
  exit 6
fi

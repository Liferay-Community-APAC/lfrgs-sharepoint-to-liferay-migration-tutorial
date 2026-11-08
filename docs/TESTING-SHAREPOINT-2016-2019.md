# Testing SharePoint 2016 and 2019 migrations without a farm

This guide is for anyone, person or tool, who needs to exercise the SharePoint 2016 and 2019 code paths
of this repo. There's no SharePoint Server you can run in a container (it needs Windows Server, SQL
Server and Active Directory), so the repo ships its own **mock SharePoint Server** and runs it in Docker.

> **What this proves, and what it doesn't.** The mock runs the real `sp2lr` code against HTTP endpoints
> that behave the way SharePoint 2016 and 2019 are documented to behave, including real NTLM sign-in. It
> can't prove a real farm answers the same way: the responses were written from Microsoft's
> documentation, not captured from SharePoint. Items S-SE1 to S-SE10 in
> [VALIDATION-CHECKLIST.md](VALIDATION-CHECKLIST.md) still need a real farm.

## 1. What the tooling is

| Piece | What it is | Where |
|---|---|---|
| Mock SharePoint | Python web app (Starlette + uvicorn) that answers SharePoint REST and client object model calls. NTLM is checked for real with `pyspnego` | `mock/sharepoint_mock/` |
| Test site | Lists, a document library, a term set, wiki/publishing/modern pages, users, AD group, permissions | `mock/sharepoint_mock/scenario.py` |
| Docker image | `python:3.12-slim` + the mock | `docker/sharepoint-mock/Dockerfile` |
| Containers | One per platform; 2019 on **8093**, 2016 on **8094** (also Online 8091, SE 8092) | `docker/docker-compose.sharepoint-mock.yml` |
| Probe script | `curl` checks of every call the migrator relies on | `docker/sharepoint-mock/check.sh` |
| Migrator | `sp2lr` CLI: `extract → plan → load → verify` | `src/sp2lr/` |
| Automated tests | pytest, starts the mock in-process in all four modes | `tests/test_platforms.py` |

## 2. How 2019 and 2016 differ (and what the mock reproduces)

| Behaviour | SharePoint 2019 mock (8093) | SharePoint 2016 mock (8094) |
|---|---|---|
| Sign-in | NTLM, `CONTOSO\spadmin` / `Passw0rd!`; wrong password → `401` | same |
| Microsoft Graph | none | none |
| Modern pages | yes: Site Pages items carry `CanvasContent1` | **none**: Site Pages has no `CanvasContent1` field |
| `$select` of a column the list lacks | `400 Column '…' does not exist` | same, so `$select=CanvasContent1` → `400` |
| Wiki pages | `WikiField` (one classic page) | `WikiField` (**every** page) |
| Publishing pages | `Pages` library (template 850), `PublishingPageContent` | same |
| News posts | `PromotedState = 2` | none (no news in 2016) |
| Term store REST `/_api/v2.1/termStore` | `404` | `404` |
| Term store through the client object model | `POST /_api/contextinfo` for a form digest, then `POST /_vti_bin/client.svc/ProcessQuery`; without the digest: *"The security validation for this page is invalid."* | same |
| Users / groups | `i:0#.w\|contoso\jane`, AD group `c:0+.w\|<SID>`, `PeopleManager` profiles | same |
| Paging | one item per page, with `odata.nextLink` | same |

## 3. Start the mocks

Prerequisites: Docker, `curl` with NTLM support, Python 3.11+.

```bash
git clone https://github.com/LGSAPAC/lfrgs-sharepoint-to-liferay-migration-tutorial.git
cd lfrgs-sharepoint-to-liferay-migration-tutorial
docker compose -f docker/docker-compose.sharepoint-mock.yml up -d --build \
  sharepoint-online sharepoint-server2019 sharepoint-server2016
curl -s localhost:8093/_health    # {"mode":"server2019","site":"http://localhost:8093/sites/marketing"}
curl -s localhost:8094/_health    # {"mode":"server2016","site":"http://localhost:8094/sites/marketing"}
```

`sharepoint-online` is listed because the other services reuse the image it builds.

## 4. Probe the endpoints

```bash
docker/sharepoint-mock/check.sh 8093    # 2019
docker/sharepoint-mock/check.sh 8094    # 2016
```

Expected output. Lines 3 to 5 are the ones that differ between versions:

| # | Check | 2019 (8093) | 2016 (8094) |
|---|---|---|---|
| 1 | sign-in: none / wrong password / right password | `401 / 401 / 200` | `401 / 401 / 200` |
| 2 | lists and templates | `Product Catalog 100, Orders 100, Policies 101, Site Pages 119, Pages 850` | same |
| 3 | Site Pages has a `CanvasContent1` field | `True` | `False` |
| 4 | `$select=CanvasContent1` on Site Pages | `HTTP 200` | `HTTP 400` |
| 5 | first page item | `Home.aspx modern`, `nextLink: True` | `Home.aspx wiki`, `nextLink: True` |
| 6 | publishing page | `Press-Release.aspx <h2>Widget 2.0 launches</h2>…` | same |
| 7 | term store REST v2.1 | `HTTP 404` | `HTTP 404` |
| 8 | client object model without digest | `The security validation for this page is invalid.` | same |
| 9 | client object model with digest | `['Sales', 'Sales;EMEA']` | same |
| 10 | site users | `i:0#.w\|contoso\jane`, `…bob`, `…old`, `c:0+.w\|s-1-5-21-…-1105` | same |
| 11 | `PeopleManager` profile | `Jane Doe Marketing Lead` | same |
| 12 | file download | `read me` | same |

## 5. Run a migration from them

```bash
python3 -m venv .venv && . .venv/bin/activate && pip install -e '.[dev]'

export SP_PLATFORM=server SP_USERNAME='CONTOSO\spadmin' SP_PASSWORD='Passw0rd!'
export SP_SITE_URL=http://localhost:8093/sites/marketing    # 8094 for 2016
export SP2LR_OUT=./out-2019                                  # ./out-2016 for 2016
sp2lr extract && sp2lr plan --mapping mapping.example.yaml
```

Expected `extract` summary, identical for both versions:

```json
{"users": 3, "groups": 3, "lists": 2, "items": 4, "libraries": 1, "files": 3,
 "termSets": 1, "pages": 5, "uniquePermissionScopes": 2}
```

What to look for in `out-*/`:

| File | 2019 | 2016 |
|---|---|---|
| `plan.md` pages | Home, About-Us, Old-Wiki, Press-Release (+1 news post) | Home, About-Us, Q3-Results, Old-Wiki, Press-Release (no news) |
| `extract/termsets.json` | Sales → EMEA (hierarchy from the client object model) | same |
| `plan.md` warnings | AD group *members can't be read*, Calculated column skipped, 300 MB video skipped, current version only | same |

To load into Liferay as well, start it (`docker compose -f docker/docker-compose.yml up -d`, wait for
*Server startup in*), then:

```bash
export LIFERAY_URL=http://localhost:8080 LIFERAY_USER=test@liferay.com LIFERAY_PASSWORD=learn
sp2lr load --dry-run && sp2lr load && sp2lr verify    # verify: 11/11 on 2019, 10/10 on 2016
sp2lr load                                            # re-run: 0 errors, nothing new
```

Both mocks serve the same site, so give Liferay a fresh start (or delete the *Marketing* site, its
Space and the two object definitions) before loading the other version.

## 6. Automated tests

```bash
pytest -q tests/test_platforms.py
```

These start the mock in-process (no Docker needed) in all four modes and assert, among other things:

* every version extracts the same counts
* 2019 and SE produce the same Liferay page definitions as SharePoint Online
* 2016 yields wiki pages and no news
* the term hierarchy survives on 2019 and 2016 (client object model)
* `$select=CanvasContent1` gets `400` on 2016
* a wrong NTLM password gets `401`

## 7. Pointing the same tests at a real farm

The commands in sections 4 and 5 work against a real SharePoint 2016 or 2019 site:

* For `check.sh`, edit `SITE` and `AUTH` at the top. Real list GUIDs differ, so look them up first
  with `/_api/web/lists?$select=Id,Title,BaseTemplate`.
* For `sp2lr`, set `SP_SITE_URL` to the real site and `SP_USERNAME` / `SP_PASSWORD` to a site collection
  administrator.

Any difference from the expected output is a finding for the S-SE items in the
[validation checklist](VALIDATION-CHECKLIST.md).

## 8. Troubleshooting

| Symptom | Fix |
|---|---|
| `Cannot connect to the Docker daemon` | Start Docker (in a bare Linux container: `setsid nohup dockerd >/tmp/dockerd.log 2>&1 &`) |
| Image build fails with `CERTIFICATE_VERIFY_FAILED` | A TLS-intercepting proxy: build from a copy of the Dockerfile that adds the proxy's CA and sets `PIP_CERT`. Don't disable TLS checks |
| `toomanyrequests` from Docker Hub | Anonymous pull limit; wait and retry, or `docker login` |
| `curl: option --ntlm: … not supported` | Use a curl built with NTLM (most distributions' packages are) |
| Liferay answers `500` right after start | It's still initialising; wait a minute after *Server startup in* |

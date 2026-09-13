# AssignLetters

AssignLetters is a Microsoft Outlook add-in for **Outlook on the web** plus a **Python (FastAPI)** backend.

It opens in **Read mode** (an opened email, not compose). The task pane reads a custom Internet message header with Office.js. If that header is present, the form is shown; otherwise the pane shows a message from configuration. Saving POSTs from the **browser** to the API and displays **Success** or **Failed**.

## Defaults

| Item | Default |
| --- | --- |
| Add-in name | AssignLetters |
| Custom header | `X-AssignLetters-Id` |
| Missing-header message | `This message does not include the required assignment header, so AssignLetters cannot open the form.` |
| API bind | `0.0.0.0:8000` |
| Staff list | `data/staff.json` |
| Log file | `data/logs/assignletters.log` |

Copy `.env.example` to `.env` and change values as needed.

## Environment variables

| Variable | Purpose | Example |
| --- | --- | --- |
| `ASSIGNLETTERS_HEADER_NAME` | Internet header the add-in looks for | `X-AssignLetters-Id` |
| `ASSIGNLETTERS_MISSING_HEADER_MESSAGE` | Shown when the header is absent | (see `.env.example`) |
| `ASSIGNLETTERS_PUBLIC_BASE_URL` | HTTPS origin substituted into `/manifest.xml` on each request | `https://abc.ngrok-free.app` |
| `ASSIGNLETTERS_API_URL` | Destination that receives each saved assignment (`POST` JSON). Origin only → `{origin}/api/save`; a full path is used as-is | `https://letters.example/v1/assignments` |
| `ASSIGNLETTERS_API_TIMEOUT_SECONDS` | Timeout for the save push | `15` |
| `ASSIGNLETTERS_LOG_PATH` | Text log file | `data/logs/assignletters.log` |
| `ASSIGNLETTERS_STAFF_PATH` | JSON staff list | `data/staff.json` |
| `ASSIGNLETTERS_CORS_ALLOW_ORIGINS` | CORS origins (`*` so Outlook Web can call the API) | `*` |
| `HOST` / `PORT` | Bind address | `0.0.0.0` / `8000` |

## Run the backend

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python -m backend
```

Or:

```bash
uvicorn backend.main:app --host 0.0.0.0 --port 8000 --reload
```

Open http://127.0.0.1:8000

| Path | Role |
| --- | --- |
| `/` | Landing page |
| `/health` | Health check |
| `/api/config` | Header name, missing-header message, API URL |
| `/api/staff` | Staff JSON (`?email=` filters by `Department`) |
| `POST /api/save` | Save assignment (called from the task pane in the browser) |
| `/api/logs` | Log file as text (`?format=json`, `?tail=200`) |
| `/logs` | HTML log viewer |
| `/manifest.xml` | Outlook manifest generated from `ASSIGNLETTERS_PUBLIC_BASE_URL` |
| `/taskpane.html` | Add-in task pane |
| `/taskpane.html?mock=1` | Browser preview with a sample header |
| `/taskpane.html?mock=1&header=` | Browser preview of the missing-header message |

```bash
curl -s http://127.0.0.1:8000/health
curl -s http://127.0.0.1:8000/api/staff
curl -s http://127.0.0.1:8000/api/staff?email=ada@contoso.com
curl -s http://127.0.0.1:8000/api/logs
curl -s -X POST http://127.0.0.1:8000/api/save \
  -H 'Content-Type: application/json' \
  -H 'Origin: https://outlook.office.com' \
  -d '{"staffName":"Ada Lovelace","department":"ada@contoso.com","deadlineDate":"2026-09-30","internetHeaderName":"X-AssignLetters-Id","internetHeaderValue":"AL-1001"}'
```

## Staff JSON

`data/staff.json` is a list of objects with `StaffName` and `Department`. **`Department` stores the email address of the mailbox user** so the dropdown can be filtered to the person using the add-in (`Office.context.mailbox.userProfile.emailAddress`).

```json
{
  "staff": [
    { "StaffName": "Ada Lovelace", "Department": "ada@contoso.com" },
    { "StaffName": "Alan Turing", "Department": "alan@contoso.com" }
  ]
}
```

`GET /api/staff?email=ada@contoso.com` returns only rows whose `Department` equals or contains that email.

## Sideload in Outlook on the web

Outlook loads add-ins only over **HTTPS**. Serve the Python app, tunnel it, then sideload the generated manifest.

1. Start the API locally (port **8000**).
2. Expose it with HTTPS, for example `ngrok http 8000`.
3. Put the tunnel origin in `.env` as `ASSIGNLETTERS_PUBLIC_BASE_URL` (and restart). Then download `https://YOUR-TUNNEL/manifest.xml`. The XML is generated on each request from that environment variable. To write a static copy:

```bash
python scripts/set_manifest_url.py
# or: python scripts/set_manifest_url.py https://YOUR-SUBDOMAIN.ngrok-free.app
```

That writes `addin/manifest.xml`.

4. In [Outlook on the web](https://outlook.office.com), sign in with a Microsoft 365 work or school account.
5. Open a **received message** (Read mode — do not start a new compose).
6. **Apps** / **Get Add-ins** → **My add-ins** → **Custom add-ins** → **Add a custom add-in** → **Add from file**.
7. Upload `addin/manifest.xml` (or the file you downloaded from `/manifest.xml`).
8. On the open message, launch **AssignLetters** from the ribbon. The task pane appears on the side.

To exercise the header gate, send yourself a message that includes:

```
X-AssignLetters-Id: AL-1001
```

If that header is missing, the pane shows `ASSIGNLETTERS_MISSING_HEADER_MESSAGE` instead of the form.

### Notes

- Do not upload a manifest that still points at `localhost`. Outlook cannot fetch those URLs.
- Every URL in the manifest, including `<AppDomain>`, must be an absolute `https://…` origin. A hostname without a scheme (`localhost` or `example.com`) makes Outlook Web report **Error in reading the manifest, Failed to construct URL**.
- Keep the API and the tunnel running while the add-in is in use.
- The task pane calls `/api/staff` and `/api/save` from the browser. CORS is enabled (`*` by default).
- `getAllInternetHeadersAsync` needs Mailbox **1.8** (declared in the manifest). Exchange may strip some custom headers; if the form never appears, inspect the raw MIME headers on the message.

## Save payload

The task pane POSTs JSON to `/api/save`:

```json
{
  "staffName": "Ada Lovelace",
  "department": "ada@contoso.com",
  "deadlineDate": "2026-09-30",
  "internetHeaderName": "X-AssignLetters-Id",
  "internetHeaderValue": "AL-1001",
  "userEmail": "ada@contoso.com",
  "subject": "Example",
  "itemId": "..."
}
```

The API appends a structured line to the log file. If `ASSIGNLETTERS_API_URL` is set, it then `POST`s the same JSON to that URL. The pane shows **Success** when the remote API returns HTTP 2xx, or **Failed** otherwise. With no `ASSIGNLETTERS_API_URL`, the save is logged locally only.

## Deploy (HTTPS host)

Set `ASSIGNLETTERS_PUBLIC_BASE_URL` to the public HTTPS origin, deploy with the included `Dockerfile`, then sideload `https://YOUR-HOST/manifest.xml`. Health check: `/health`.

## Tests

```bash
pip install -r requirements.txt
pytest -q
```

## Layout

| Path | Role |
| --- | --- |
| `backend/main.py` | FastAPI app: health, config, staff, save, logs, manifest, static add-in |
| `backend/push.py` | POST saved assignments to `ASSIGNLETTERS_API_URL` |
| `addin/manifest.template.xml` | Outlook Web Read-mode template (`{{PUBLIC_BASE_URL}}` filled at runtime) |
| `addin/taskpane.html` / `taskpane.js` | Office.js task pane |
| `data/staff.json` | Staff dropdown source |
| `.env.example` | All supported environment variables |

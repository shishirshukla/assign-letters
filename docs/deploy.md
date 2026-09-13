# Deploy AssignLetters

Use any HTTPS host (Railway, a VM, or a tunnel). Outlook requires HTTPS.

1. Set `ASSIGNLETTERS_PUBLIC_BASE_URL` to `https://YOUR-HOST`.
2. Set `ASSIGNLETTERS_API_URL` to the HTTPS endpoint that should receive saved assignments.
3. Deploy with the repo `Dockerfile` (health check `/health`).
4. Sideload `https://YOUR-HOST/manifest.xml`. That file is generated on each request from `ASSIGNLETTERS_PUBLIC_BASE_URL`.

Local development with ngrok is documented in the root `README.md`.

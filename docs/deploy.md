# Deploy AssignLetters

Use any HTTPS host (Railway, a VM, or a tunnel). Outlook requires HTTPS.

1. Set `ASSIGNLETTERS_PUBLIC_BASE_URL` to `https://YOUR-HOST`.
2. Deploy with the repo `Dockerfile` (health check `/health`).
3. Sideload `https://YOUR-HOST/manifest.xml`. That file is generated on each request from the environment variable.

Local development with ngrok is documented in the root `README.md`.

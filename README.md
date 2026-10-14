# PMS Workflow Web App - Vercel runtime fix

## What was fixed

- Runtime data and uploads now use `os.tmpdir()` instead of `/var/task/data` and `/var/task/uploads`.
- The Express app is exported with `module.exports = app` for Vercel.
- `app.listen()` runs only during local execution.
- `vercel.json` routes requests to the Vercel Node function.
- A health endpoint is available at `/api/health`.

## Replace the GitHub repository

Replace the existing repository files with the contents of this ZIP, commit, push, and redeploy in Vercel.

## Vercel environment variables

Set these variables for Production, Preview, and Development as needed:

- `BASE_URL` = your production Vercel URL, without a trailing slash
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `MAIL_SENDER`

The Entra application requires Microsoft Graph `Mail.Send` application permission with admin consent.

## Important storage limitation

This package fixes the reported startup error. On Vercel, `/tmp` is temporary. Therefore SQLite records, uploaded attachments, and generated PDFs are not guaranteed to persist between serverless invocations or deployments. For production use, replace SQLite with a hosted PostgreSQL database and store attachments/PDFs in persistent object storage.

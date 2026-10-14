# PMS Vercel Run Fix

This version removes the native `better-sqlite3` dependency and removes the deprecated `builds` configuration that appeared in the Vercel build warnings.

## Replace the repository

1. Extract this ZIP.
2. Replace all existing GitHub repository files with the extracted files.
3. Commit and push.
4. In Vercel, redeploy without using the old build cache.
5. Keep Framework Preset as **Other**. Do not specify a custom Build Command or Output Directory.

## Environment variables

Set `BASE_URL`, `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, and `MAIL_SENDER` in Vercel.

## Test

Open `/api/health`. The response should show `ok: true` and `nativeSqlite: false`.

## Storage warning

This package is designed to build and start on Vercel without native SQLite. It uses Vercel's temporary writable directory. Temporary records are not guaranteed to persist between function instances. For durable production approvals, connect the application to a hosted database and object storage.

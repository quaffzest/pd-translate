# Tencent Cloud deployment notes

This project can use Tencent Cloud COS as the durable workbook store. The app still keeps a local cache in `DATA_DIR`, but when COS is configured the source of durability is the COS bucket.

## Required services

- Lighthouse or CVM: runs the Node.js app and WebSocket collaboration server.
- COS: stores workbooks and `_meta.json` under one prefix.
- Database: not required by this migration step yet; use it in the next phase for users, permissions, file index, and version history.

## Environment variables

Set these on the Tencent Cloud server:

```bash
NODE_ENV=production
PORT=3000
SESSION_SECRET=change-to-a-long-random-string
WORKSPACE_ALLOWED_EMAILS=your@email.com,collaborator@email.com

TENCENT_SECRET_ID=AKIDxxxxxxxxxxxxxxxx
TENCENT_SECRET_KEY=xxxxxxxxxxxxxxxx
TENCENT_COS_BUCKET=pd-translate-1250000000
TENCENT_COS_REGION=ap-guangzhou
TENCENT_COS_PREFIX=pd-translate
```

The simplest way is to put the values into a `.env` file in the project root and let the app load it automatically.

`TENCENT_COS_PREFIX` keeps this app's files isolated in the bucket. The app writes:

- `pd-translate/workbooks/_meta.json`
- `pd-translate/workbooks/<folder>/<workbook>.xlsx`

## Behavior

- On startup, if COS already contains workbook objects, the app treats COS as the durable source and rebuilds the local cache from COS.
- If COS is empty, the app treats the current local workbooks as first-run migration data and uploads them to COS after startup.
- Upload, save, move, rename, and delete actions are mirrored to COS.
- If COS variables are not configured, the app falls back to the existing local file behavior.

## Current transition status

Google Drive routes still exist for compatibility, but the production target is Tencent Cloud. The next phase should remove Drive from the main UI and replace it with the COS-backed cloud workbook list.

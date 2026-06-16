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

## One-command bootstrap

After you clone the repo on the Tencent Cloud server, run:

```bash
sudo bash deploy/tencent-cloud-setup.sh
```

You do not need to mark the script executable first, because `bash` runs it directly.

The script will install dependencies, ask you for the required keys, write `.env`, and register a `systemd` service so the app starts automatically after reboot.

### What the script will ask you for

1. Your public base URL, such as `http://YOUR_SERVER_IP:3000` or `https://your.domain.com`
2. `GOOGLE_CLIENT_ID`
3. `GOOGLE_CLIENT_SECRET`
4. The Google email address that is allowed to use the workbench
5. `TENCENT_SECRET_ID`
6. `TENCENT_SECRET_KEY`
7. `TENCENT_COS_BUCKET`
8. `TENCENT_COS_REGION`
9. `TENCENT_COS_PREFIX`

If you only want a quick local-style setup, you can keep the default values for the region and prefix prompts.

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

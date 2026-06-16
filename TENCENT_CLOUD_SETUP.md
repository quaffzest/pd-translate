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
SESSION_COOKIE_SECURE=false
WORKSPACE_AUTH_MODE=local
WORKSPACE_PASSWORD=change-this-to-a-strong-password
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
2. Whether to enable Google login on this server
3. If Google login is enabled: `GOOGLE_CLIENT_ID`
4. If Google login is enabled: `GOOGLE_CLIENT_SECRET`
5. If Google login is enabled: the Google email address allowed to use the workbench
6. If Google login is disabled: a local workspace password
7. `TENCENT_SECRET_ID`
8. `TENCENT_SECRET_KEY`
9. `TENCENT_COS_BUCKET`
10. `TENCENT_COS_REGION`
11. `TENCENT_COS_PREFIX`

If you only want the Tencent Cloud workspace to run first, choose the local password mode and skip the Google prompts for now.

`TENCENT_COS_PREFIX` keeps this app's files isolated in the bucket. The app writes:

- `pd-translate/workbooks/_meta.json`
- `pd-translate/workbooks/<folder>/<workbook>.xlsx`

## Behavior

- On startup, if COS already contains workbook objects, the app treats COS as the durable source and rebuilds the local cache from COS.
- If COS is empty, the app treats the current local workbooks as first-run migration data and uploads them to COS after startup.
- Upload, save, move, rename, and delete actions are mirrored to COS.
- If COS variables are not configured, the app falls back to the existing local file behavior.
- If `WORKSPACE_AUTH_MODE=local`, the main workbench uses the local password login and does not require Google OAuth.
- If `WORKSPACE_AUTH_MODE=google`, the workbench login still uses Google OAuth and the allowed email list.

## Current transition status

Google Drive routes still exist for compatibility, but the production target is Tencent Cloud. The next phase should remove Drive from the main UI and replace it with the COS-backed cloud workbook list.

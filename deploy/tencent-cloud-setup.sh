#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SERVICE_NAME="pd-translate"
ENV_FILE="$APP_DIR/.env"
SERVICE_FILE="/etc/systemd/system/${SERVICE_NAME}.service"
SERVICE_USER="root"

if [[ "${EUID:-$(id -u)}" -ne 0 ]]; then
  exec sudo -E bash "$0" "$@"
fi

need_cmd() {
  command -v "$1" >/dev/null 2>&1
}

install_packages() {
  if need_cmd dnf; then
    dnf install -y git nodejs npm openssl
  elif need_cmd yum; then
    yum install -y git nodejs npm openssl
  elif need_cmd apt-get; then
    apt-get update
    apt-get install -y git nodejs npm openssl
  else
    echo "No supported package manager found. Please install git, nodejs, npm, and openssl manually." >&2
    exit 1
  fi
}

prompt() {
  local label="$1"
  local default="${2:-}"
  local value=""
  if [[ -n "$default" ]]; then
    read -r -p "$label [$default]: " value
    echo "${value:-$default}"
  else
    read -r -p "$label: " value
    echo "$value"
  fi
}

prompt_secret() {
  local label="$1"
  local value=""
  read -r -s -p "$label: " value
  echo
  echo "$value"
}

require_value() {
  local name="$1"
  local value="$2"
  if [[ -z "$value" ]]; then
    echo "$name is required." >&2
    exit 1
  fi
}

write_env_file() {
  local public_base_url="$1"
  local auth_mode="$2"
  local workspace_password="$3"
  local google_client_id="$4"
  local google_client_secret="$5"
  local workspace_emails="$6"
  local tencent_secret_id="$7"
  local tencent_secret_key="$8"
  local cos_bucket="$9"
  local cos_region="${10}"
  local cos_prefix="${11}"
  local session_secret
  local session_cookie_secure
  local escape_env

  session_secret="$(openssl rand -hex 32)"
  if [[ "$public_base_url" =~ ^https:// ]]; then
    session_cookie_secure="true"
  else
    session_cookie_secure="false"
  fi

  escape_env() {
    printf '%s' "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'
  }

  umask 077
  cat > "$ENV_FILE" <<EOF
NODE_ENV="production"
PORT="3000"
SESSION_SECRET="$(escape_env "$session_secret")"
SESSION_COOKIE_SECURE="$(escape_env "$session_cookie_secure")"
WORKSPACE_AUTH_MODE="$(escape_env "$auth_mode")"
WORKSPACE_PASSWORD="$(escape_env "$workspace_password")"
WORKSPACE_ALLOWED_EMAILS="$(escape_env "$workspace_emails")"

GOOGLE_CLIENT_ID="$(escape_env "$google_client_id")"
GOOGLE_CLIENT_SECRET="$(escape_env "$google_client_secret")"
GOOGLE_CALLBACK_URL="$(escape_env "${public_base_url%/}/auth/google/callback")"

TENCENT_SECRET_ID="$(escape_env "$tencent_secret_id")"
TENCENT_SECRET_KEY="$(escape_env "$tencent_secret_key")"
TENCENT_COS_BUCKET="$(escape_env "$cos_bucket")"
TENCENT_COS_REGION="$(escape_env "$cos_region")"
TENCENT_COS_PREFIX="$(escape_env "$cos_prefix")"
EOF

  chown "$SERVICE_USER:$SERVICE_USER" "$ENV_FILE" || true
  chmod 600 "$ENV_FILE"
}

write_service_file() {
  local npm_path
  npm_path="$(command -v npm)"
  cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=pd-translate collaborative workspace
After=network.target

[Service]
Type=simple
User=$SERVICE_USER
WorkingDirectory=$APP_DIR
EnvironmentFile=$ENV_FILE
ExecStart=$npm_path start
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF
}

open_firewall_port() {
  if need_cmd firewall-cmd && systemctl is-active --quiet firewalld; then
    firewall-cmd --permanent --add-port=3000/tcp >/dev/null
    firewall-cmd --reload >/dev/null
  elif need_cmd ufw; then
    ufw allow 3000/tcp >/dev/null || true
  fi
}

echo "Installing system packages..."
install_packages

echo "Installing Node dependencies..."
cd "$APP_DIR"
npm install

echo "Writing deployment env file..."
PUBLIC_BASE_URL="$(prompt "Enter your public base URL (for example http://YOUR_SERVER_IP:3000 or https://your.domain.com)")"
USE_GOOGLE_LOGIN="$(prompt "Enable Google login for this server? (y/N)" "N")"
USE_GOOGLE_LOGIN="${USE_GOOGLE_LOGIN,,}"
WORKSPACE_AUTH_MODE="local"
WORKSPACE_PASSWORD=""
GOOGLE_CLIENT_ID=""
GOOGLE_CLIENT_SECRET=""
WORKSPACE_ALLOWED_EMAILS=""
if [[ "$USE_GOOGLE_LOGIN" == "y" || "$USE_GOOGLE_LOGIN" == "yes" ]]; then
  WORKSPACE_AUTH_MODE="google"
  GOOGLE_CLIENT_ID="$(prompt "Enter GOOGLE_CLIENT_ID")"
  GOOGLE_CLIENT_SECRET="$(prompt_secret "Enter GOOGLE_CLIENT_SECRET")"
  WORKSPACE_ALLOWED_EMAILS="$(prompt "Enter allowed Google email(s), comma-separated" "")"
else
  WORKSPACE_PASSWORD="$(prompt_secret "Set a workspace password")"
fi
TENCENT_SECRET_ID="$(prompt "Enter TENCENT_SECRET_ID")"
TENCENT_SECRET_KEY="$(prompt_secret "Enter TENCENT_SECRET_KEY")"
TENCENT_COS_BUCKET="$(prompt "Enter TENCENT_COS_BUCKET")"
TENCENT_COS_REGION="$(prompt "Enter TENCENT_COS_REGION" "ap-guangzhou")"
TENCENT_COS_PREFIX="$(prompt "Enter TENCENT_COS_PREFIX" "pd-translate")"

if [[ "$WORKSPACE_AUTH_MODE" == "google" ]]; then
  require_value "GOOGLE_CLIENT_ID" "$GOOGLE_CLIENT_ID"
  require_value "GOOGLE_CLIENT_SECRET" "$GOOGLE_CLIENT_SECRET"
else
  require_value "WORKSPACE_PASSWORD" "$WORKSPACE_PASSWORD"
fi
require_value "TENCENT_SECRET_ID" "$TENCENT_SECRET_ID"
require_value "TENCENT_SECRET_KEY" "$TENCENT_SECRET_KEY"
require_value "TENCENT_COS_BUCKET" "$TENCENT_COS_BUCKET"

write_env_file \
  "$PUBLIC_BASE_URL" \
  "$WORKSPACE_AUTH_MODE" \
  "$WORKSPACE_PASSWORD" \
  "$GOOGLE_CLIENT_ID" \
  "$GOOGLE_CLIENT_SECRET" \
  "$WORKSPACE_ALLOWED_EMAILS" \
  "$TENCENT_SECRET_ID" \
  "$TENCENT_SECRET_KEY" \
  "$TENCENT_COS_BUCKET" \
  "$TENCENT_COS_REGION" \
  "$TENCENT_COS_PREFIX"

echo "Installing systemd service..."
write_service_file
systemctl daemon-reload
systemctl enable "$SERVICE_NAME"
systemctl restart "$SERVICE_NAME"
open_firewall_port

echo
echo "Deployment finished."
echo "Service: systemctl status $SERVICE_NAME"
echo "Logs: journalctl -u $SERVICE_NAME -f"
echo "Open: $PUBLIC_BASE_URL"
echo "Remember to allow TCP 3000 in the Tencent Cloud security group if needed."

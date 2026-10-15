#!/usr/bin/env bash
set -euo pipefail

# --- mise: trust, install, and activate in every new zsh terminal ---
mise trust /workspaces/setup
mise install
echo 'eval "$(mise activate zsh)"' >> ~/.zshrc

# --- AWS config ---
mkdir -p ~/.aws
cat > ~/.aws/config << 'EOF'
[sso-session default]
sso_start_url = https://exxeta.awsapps.com/start/#
sso_region = eu-central-1
sso_registration_scopes = sso:account:access

EOF
chmod 600 ~/.aws/config

# --- Granted config ---
mkdir -p ~/.granted
cat > ~/.granted/config << 'EOF'
DefaultBrowser = "STDOUT"
Ordering = "Frecency"
ExportCredentialSuffix = ""
EOF


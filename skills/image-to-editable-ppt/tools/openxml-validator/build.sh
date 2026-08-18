#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
dotnet restore --locked-mode "$SCRIPT_DIR/openxml-validator.csproj"
dotnet build --configuration Release --no-restore "$SCRIPT_DIR/openxml-validator.csproj"

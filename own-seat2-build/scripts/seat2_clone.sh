#!/usr/bin/env bash
# Rebuild the desktop second-seat app bundles from the current first-seat apps.
#
# Why this exists: a second seat is a full copy of the vendor bundle whose
# CFBundleIdentifier was changed (local.launchers.*) so LaunchServices treats it
# as a separate app. Both vendors update in place by locating, inside the
# downloaded archive, a bundle whose identifier matches the running app's:
#
#   [updater] Auto-update error: Could not locate update bundle for
#   local.launchers.claude-seat2 within .../local.launchers.claude-seat2.ShipIt/
#
# The archive only ever contains com.anthropic.claudefordesktop (Squirrel) or
# com.openai.codex (Sparkle), so the in-app updater of a second seat can never
# succeed — and if it did, it would replace the whole bundle and take the seat
# shim with it. The seat is therefore refreshed by re-cloning it from the
# first-seat app, which does update normally.
#
# Run this after the first-seat app updates. --check reports drift, --install
# rebuilds. Adding a third seat is one line in SEATS plus its profile in
# shim_env_for.
#
# The clone differs from its source in exactly three places:
#   1. Contents/MacOS/<exe> is a shim that sets the seat's profile directory and
#      execs the renamed vendor binary (<exe>Seat2Payload).
#   2. Info.plist: display name, bundle identifier, URL schemes.
#   3. The bundle is re-signed ad hoc with --deep, because 1 and 2 invalidate
#      the vendor signature. The rename alone invalidates the payload's own
#      signature too: a payload left with the vendor signature is SIGKILLed on
#      launch, so every nested binary is re-signed with it.

set -euo pipefail

install_dir="${LAUNCHER_INSTALL_DIR:-/Applications}"

# seat : source app : installed name : display name : bundle identifier
SEATS=(
  "claude:Claude:Claude-Seat2:Claude Seat2:local.launchers.claude-seat2"
  "codex:ChatGPT:Codex-Seat2:ChatGPT Seat2:local.launchers.codex-seat2"
)

usage() {
  echo "Usage: seat2_clone.sh [--check|--install] [claude|codex|all]" >&2
}

version_of() {
  /usr/libexec/PlistBuddy -c "Print :CFBundleShortVersionString" \
    "$1/Contents/Info.plist" 2>/dev/null || true
}

# The shim body, per seat. $script_dir and $payload are expanded at run time by
# the shim itself, so everything here is single-quoted.
shim_env_for() {
  case "$1" in
    claude)
      echo 'seat2_profile="${HOME:?}/.claude-seat2"'
      ;;
    codex)
      echo 'seat2_profile="${HOME:?}/.codex-seat2/electron"'
      echo 'export CODEX_HOME="${HOME:?}/.codex-seat2"'
      echo 'export CODEX_ELECTRON_USER_DATA_PATH="$seat2_profile"'
      ;;
  esac
}

write_shim() {
  local seat="$1" path="$2" payload="$3"
  {
    echo '#!/bin/sh'
    echo 'set -eu'
    echo 'script_dir=$(CDPATH= cd "$(dirname "$0")" && pwd -P)'
    shim_env_for "$seat"
    cat <<'BODY'
has_user_data=0
for arg in "$@"; do
    case "$arg" in
        --user-data-dir=*) has_user_data=1 ;;
    esac
done
if [ "$has_user_data" -eq 0 ]; then
    set -- "--user-data-dir=$seat2_profile" "$@"
fi
BODY
    printf 'exec "$script_dir/%s" "$@"\n' "$payload"
  } > "$path"
  chmod 755 "$path"
}

# Identity edits. Schemes are suffixed rather than listed, so a vendor that adds
# a scheme does not silently give the seat a duplicate registration; http and
# https are dropped so the seat never competes to be the default browser.
patch_plist() {
  local plist="$1" display="$2" ident="$3"
  DISPLAY_NAME="$display" BUNDLE_ID="$ident" /usr/bin/python3 - "$plist" <<'PY'
import os, plistlib, sys

path = sys.argv[1]
with open(path, "rb") as fh:
    info = plistlib.load(fh)

info["CFBundleDisplayName"] = os.environ["DISPLAY_NAME"]
info["CFBundleIdentifier"] = os.environ["BUNDLE_ID"]

url_types = []
for entry in info.get("CFBundleURLTypes", []):
    schemes = []
    for scheme in entry.get("CFBundleURLSchemes", []):
        if scheme in ("http", "https"):
            continue
        schemes.append(scheme + (".seat2" if scheme.startswith("msauth.") else "-seat2"))
    if not schemes:
        continue
    entry["CFBundleURLSchemes"] = schemes
    url_types.append(entry)
if url_types:
    info["CFBundleURLTypes"] = url_types
elif "CFBundleURLTypes" in info:
    del info["CFBundleURLTypes"]

with open(path, "wb") as fh:
    plistlib.dump(info, fh)
PY
}

# ps, not pgrep: pgrep did not report the app hosting the agent session that
# wrote this script, while ps did (measured 2026-09-23). Any process under the
# bundle counts — replacing a bundle with helpers still live is what leaves a
# half-updated seat behind.
running_pids() {
  /bin/ps -Ao pid=,command= |
    sed -n "s|^ *\([0-9][0-9]*\) *$1/Contents/.*|\1|p"
}

clone_seat() {
  local seat="$1" src="$2" app="$3" display="$4" ident="$5"

  [[ -d "$src" ]] || { echo "ERROR: source $src is not installed" >&2; return 1; }

  local pids
  pids=$(running_pids "$app")
  if [[ -n "$pids" ]]; then
    echo "ERROR: $app is running (pid $(echo $pids | tr '\n' ' ')); quit it first" >&2
    return 1
  fi

  local exe payload staging
  exe=$(/usr/libexec/PlistBuddy -c "Print :CFBundleExecutable" "$src/Contents/Info.plist")
  payload="${exe}Seat2Payload"
  staging="$install_dir/.$(basename "$app" .app).staging.app"

  rm -rf "$staging"
  ditto "$src" "$staging"

  mv "$staging/Contents/MacOS/$exe" "$staging/Contents/MacOS/$payload"
  write_shim "$seat" "$staging/Contents/MacOS/$exe" "$payload"
  patch_plist "$staging/Contents/Info.plist" "$display" "$ident"

  rm -rf "$staging/Contents/_CodeSignature"
  codesign --force --deep --sign - "$staging"

  rm -rf "$app.previous"
  [[ -d "$app" ]] && mv "$app" "$app.previous"
  mv "$staging" "$app"
  rm -rf "$app.previous"

  echo "rebuilt $(basename "$app") at $(version_of "$app") from $(basename "$src")"
}

selected() {
  local want="$1" seat="$2"
  [[ "$want" == "all" || "$want" == "$seat" ]]
}

do_check() {
  local want="$1" status=0
  for entry in "${SEATS[@]}"; do
    IFS=: read -r seat srcname name display ident <<<"$entry"
    selected "$want" "$seat" || continue
    local src="$install_dir/$srcname.app" app="$install_dir/$name.app"
    if [[ ! -d "$app" ]]; then
      echo "SKIP: $name.app is not installed"
      continue
    fi
    local have want_v
    have=$(version_of "$app"); want_v=$(version_of "$src")
    if [[ "$have" != "$want_v" ]]; then
      echo "DRIFT: $name.app is $have, $srcname.app is $want_v" >&2
      status=1
    else
      echo "OK: $name.app $have matches $srcname.app"
    fi
  done
  return $status
}

do_install() {
  local want="$1"
  for entry in "${SEATS[@]}"; do
    IFS=: read -r seat srcname name display ident <<<"$entry"
    selected "$want" "$seat" || continue
    clone_seat "$seat" "$install_dir/$srcname.app" "$install_dir/$name.app" \
      "$display" "$ident"
  done
  echo
  echo "A rebuilt seat is a new bundle: macOS re-asks for its permissions"
  echo "(Automation, Accessibility, screen recording) on first use."
}

mode="${1:---check}"
target="${2:-all}"
case "$target" in
  claude|codex|all) ;;
  *) usage; exit 64 ;;
esac
case "$mode" in
  --check) do_check "$target" ;;
  --install) do_install "$target" ;;
  *) usage; exit 64 ;;
esac

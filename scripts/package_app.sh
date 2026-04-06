#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "$0")" && pwd)"
ROOT_DIR="$(cd -- "$SCRIPT_DIR/.." && pwd)"
BUILD_DIR="$ROOT_DIR/.build/arm64-apple-macosx/debug"
EXECUTABLE="$BUILD_DIR/TourAutoLayout"
APP_NAME="TourAutoLayout.app"
APP_DIR="$ROOT_DIR/$APP_NAME"
CONTENTS_DIR="$APP_DIR/Contents"
MACOS_DIR="$CONTENTS_DIR/MacOS"
RESOURCES_DIR="$CONTENTS_DIR/Resources"
INFO_PLIST="$CONTENTS_DIR/Info.plist"
PKGINFO="$CONTENTS_DIR/PkgInfo"
DOCUMENTS_ALIAS_NAME="TourAutoLayout.app"
DOCUMENTS_DIR="$HOME/Documents"
DISPLAY_NAME="旅游行程自动排版"
ICON_SOURCE="$ROOT_DIR/.build/arm64-apple-macosx/debug/ZIPFoundation_ZIPFoundation.bundle/PrivacyInfo.xcprivacy"

cd "$ROOT_DIR"
if ! swift build >/dev/null 2>&1; then
  rm -rf "$ROOT_DIR/.build"
  swift build >/dev/null
fi

if [[ ! -x "$EXECUTABLE" ]]; then
  echo "Missing executable: $EXECUTABLE" >&2
  exit 1
fi

rm -rf "$APP_DIR"
mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

cp "$EXECUTABLE" "$MACOS_DIR/TourAutoLayout"
chmod 755 "$MACOS_DIR/TourAutoLayout"

cat > "$INFO_PLIST" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDevelopmentRegion</key>
  <string>zh_CN</string>
  <key>CFBundleExecutable</key>
  <string>TourAutoLayout</string>
  <key>CFBundleIdentifier</key>
  <string>com.codex.tourautolayout</string>
  <key>CFBundleInfoDictionaryVersion</key>
  <string>6.0</string>
  <key>CFBundleName</key>
  <string>TourAutoLayout</string>
  <key>CFBundleDisplayName</key>
  <string>旅游行程自动排版</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>1.0</string>
  <key>CFBundleVersion</key>
  <string>1</string>
  <key>LSMinimumSystemVersion</key>
  <string>14.0</string>
  <key>NSHighResolutionCapable</key>
  <true/>
</dict>
</plist>
PLIST

printf 'APPL????' > "$PKGINFO"

if [[ -f "$ICON_SOURCE" ]]; then
  cp "$ICON_SOURCE" "$RESOURCES_DIR/.keep"
fi

/usr/bin/touch "$APP_DIR"

rm -f "$DOCUMENTS_DIR/$DOCUMENTS_ALIAS_NAME"

osascript <<APPLESCRIPT
tell application "Finder"
  set targetApp to POSIX file "$APP_DIR" as alias
  set documentsFolder to POSIX file "$DOCUMENTS_DIR" as alias
  set createdAlias to make new alias file at documentsFolder to targetApp
  set name of createdAlias to "$DOCUMENTS_ALIAS_NAME"
end tell
APPLESCRIPT

open "$APP_DIR"

echo "App bundle created at: $APP_DIR"
echo "Documents alias created at: $DOCUMENTS_DIR/$DOCUMENTS_ALIAS_NAME"
echo "Display name: $DISPLAY_NAME"

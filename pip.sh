#!/bin/bash
set -euo pipefail

SITE_PACKAGES="$1"
OUTPUT_DIR="$2"

if [[ -z "$SITE_PACKAGES" || -z "$OUTPUT_DIR" ]]; then
    echo "Usage: $0 <site-packages-dir> <output-dir>"
    exit 1
fi

mkdir -p "$OUTPUT_DIR"

find "$SITE_PACKAGES" -maxdepth 1 -name "*.dist-info" -type d | while read -r dist_info; do
    RECORD="$dist_info/RECORD"
    [[ ! -f "$RECORD" ]] && echo "WARN: no RECORD in $dist_info, skipping" && continue

    PACKAGE_NAME="$(basename "$dist_info" .dist-info)"
    STAGING_DIR="$(mktemp -d)/wheel_staging"
    mkdir -p "$STAGING_DIR"

    echo "Packing: $PACKAGE_NAME"

    # העתק את כל הקבצים מה-RECORD לתיקיית staging שטוחה
    while IFS=',' read -r filepath _ _; do
        [[ -z "$filepath" ]] && continue

        SRC="$(realpath -m "$SITE_PACKAGES/$filepath")"
        [[ ! -f "$SRC" ]] && echo "  WARN: missing $SRC" && continue

        # שומר נתיב יחסי ל-site-packages בלבד (wheel pack לא מבין /usr/bin וכו')
        REL="$(realpath -m --relative-to="$SITE_PACKAGES" "$SRC")"
        DEST="$STAGING_DIR/$REL"
        mkdir -p "$(dirname "$DEST")"
        cp --preserve=mode "$SRC" "$DEST"

    done < "$RECORD"

    python -m wheel pack "$STAGING_DIR" --dest-dir "$OUTPUT_DIR" \
        && echo "  ✓ $PACKAGE_NAME" \
        || echo "  ✗ FAILED: $PACKAGE_NAME"

    rm -rf "$(dirname "$STAGING_DIR")"
done

echo "Done → $OUTPUT_DIR"

#!/usr/bin/env bash
# Netlify build step: assemble the published site/ from the source files and
# stamp a fresh service-worker cache so every deploy reaches installed apps.
set -euo pipefail
mkdir -p site
cp -f index.html manifest.webmanifest icon-192.png icon-512.png site/
ver="${COMMIT_REF:-$(date +%Y%m%d-%H%M%S)}"
sed "s|__CACHE__|${ver}|g" sw.js > site/sw.js
echo "build.sh: synced site/ with cache ${ver}"

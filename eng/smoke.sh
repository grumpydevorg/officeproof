#!/usr/bin/env bash
# Runs a built officeproof end to end on the example template: help, generate, validate, a tracked
# redline that verify passes, and an untracked replace that verify must fail. The release workflow
# runs it on every published binary; it also runs under Git Bash on Windows.
#
#   eng/smoke.sh <officeproof executable>
set -euo pipefail

exe=$1
template=$(cd "$(dirname "$0")/.." && pwd)/examples/minimal-template
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT

run() {
    echo "+ officeproof $*"
    "$exe" "$@"
}

run --version
run --help > /dev/null
run generate --template "$template" -o "$work/gen"
sample=$work/gen/Example-sample.docx
run validate "$sample"

# Redirected output is UTF-8 everywhere: the sample's list bullets survive (Windows' OEM code
# page has no "•").
outline=$(run outline "$sample")
echo "$outline"
if [[ $outline != *"• Cover fields"* ]]; then
    echo "smoke: outline lost the list bullet; console output is not UTF-8" >&2
    exit 1
fi

cat > "$work/edits.json" <<'EOF'
[{ "find": "Cover fields come from", "replace": "report.yaml", "with": "the report's metadata file",
   "comment": "smoke test" }]
EOF
run redline "$sample" "$work/redline.docx" --edits "$work/edits.json" --author "Smoke Test"
run verify "$sample" "$work/redline.docx"

# An untracked change is exactly what verify exists to catch: exit 1, an error finding.
run replace "$sample" "$work/replaced.docx" --find "come from the bindings" --with "follow the bindings" --expect 1
status=0
run verify "$sample" "$work/replaced.docx" || status=$?
if [[ $status != 1 ]]; then
    echo "smoke: verify of an untracked replace exited $status, expected 1" >&2
    exit 1
fi
echo "smoke: ok"

#!/usr/bin/env bash
# Builds the Clippit package officeproof pins (Directory.Packages.props) into the local feed that
# nuget.config names. Clippit 3.9.2-rmorgans.2 is upstream 3.9.1 plus eight RevisionProcessor
# fixes (sergey-tihon/Clippit#548-#555), merged on the fork https://github.com/rmorgans/Clippit
# at the commit below.
#
# Built from `git archive` of that exact commit, so the state of the Clippit checkout does not
# matter. Upstream's build.fsx takes its version from CHANGELOG.md (3.9.1) and writes the
# gitignored Clippit/Properties/AssemblyInfo.g.cs; this script writes that file itself with the
# patched version and passes the package version to `dotnet pack`.
#
# Delete this script, nuget.config and the feed once a nuget.org Clippit contains the fixes.
set -euo pipefail

readonly COMMIT=4fc1f95c7a509e941b6cb4a528951d9530f5a7aa
readonly FORK=https://github.com/rmorgans/Clippit
readonly PACKAGE_VERSION=3.9.2-rmorgans.2
readonly ASSEMBLY_VERSION=3.9.2
CLIPPIT_DIR=${CLIPPIT_DIR:-$(cd "$(dirname "$0")/.." && pwd)/.local/Clippit}
FEED=${FEED:-$HOME/.local/share/nuget/local-feed}

if [[ ! -d $CLIPPIT_DIR/.git ]]; then
    git clone --quiet "$FORK" "$CLIPPIT_DIR"
fi
if ! git -C "$CLIPPIT_DIR" cat-file -e "$COMMIT^{commit}" 2>/dev/null; then
    git -C "$CLIPPIT_DIR" fetch --quiet "$FORK" "$COMMIT"
fi

work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT
git -C "$CLIPPIT_DIR" archive "$COMMIT" | tar -x -C "$work"

cat > "$work/Clippit/Properties/AssemblyInfo.g.cs" <<EOF
// Written by officeproof/eng/pack-clippit.sh (upstream generates this file in build.fsx).
using System.Reflection;

[assembly: AssemblyTitle("Clippit")]
[assembly: AssemblyProduct("Clippit")]
[assembly: AssemblyDescription("Fresh PowerTools for OpenXml")]
[assembly: AssemblyVersion("$ASSEMBLY_VERSION")]
[assembly: AssemblyFileVersion("$ASSEMBLY_VERSION")]
[assembly: AssemblyInformationalVersion("$PACKAGE_VERSION+${COMMIT:0:7}")]
EOF

dotnet pack "$work/Clippit/Clippit.csproj" -c Release --nologo -v quiet \
    -o "$work/out" -p:PackageVersion="$PACKAGE_VERSION"

mkdir -p "$FEED"
cp "$work/out/Clippit.$PACKAGE_VERSION.nupkg" "$work/out/Clippit.$PACKAGE_VERSION.snupkg" "$FEED/"
# NuGet caches a package by id and version: a rebuilt package with the same version is only
# picked up after the cached copy is removed.
rm -rf "${NUGET_PACKAGES:-$HOME/.nuget/packages}/clippit/$PACKAGE_VERSION"
echo "$FEED/Clippit.$PACKAGE_VERSION.nupkg"

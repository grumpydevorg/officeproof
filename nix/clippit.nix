# Clippit 3.9.2-rmorgans.2: upstream 3.9.1 plus the RevisionProcessor fixes proposed as
# sergey-tihon/Clippit#548-#555, from the fork commit eng/pack-clippit.sh also builds. Packed as a
# NuGet package so the officeproof build can restore it (projectReferences). Delete this file
# once a nuget.org Clippit contains the fixes.
{
  lib,
  buildDotnetModule,
  dotnetCorePackages,
  fetchFromGitHub,
}:

buildDotnetModule rec {
  pname = "clippit";
  version = "3.9.2-rmorgans.2";

  src = fetchFromGitHub {
    owner = "rmorgans";
    repo = "Clippit";
    rev = "4fc1f95c7a509e941b6cb4a528951d9530f5a7aa";
    hash = "sha256-dHI7xX+ADsGV0bILqGYJ8sPDm0Ht2pR6UhINdjfOSvY=";
  };

  # Upstream's build.fsx generates this (gitignored) file from CHANGELOG.md; write it with the
  # patched version instead, as eng/pack-clippit.sh does. The repo's dotnet tools (csharpier,
  # husky, dotnet-outdated) are development-only and cannot be restored offline.
  postPatch = ''
    rm .config/dotnet-tools.json
    cat > Clippit/Properties/AssemblyInfo.g.cs <<EOF
    using System.Reflection;
    [assembly: AssemblyTitle("Clippit")]
    [assembly: AssemblyProduct("Clippit")]
    [assembly: AssemblyDescription("Fresh PowerTools for OpenXml")]
    [assembly: AssemblyVersion("3.9.2")]
    [assembly: AssemblyFileVersion("3.9.2")]
    EOF
  '';

  projectFile = "Clippit/Clippit.csproj";
  nugetDeps = ./clippit-deps.json;
  dotnet-sdk = dotnetCorePackages.sdk_10_0;
  dotnet-runtime = null;
  # The compiler server's server GC cannot start in the Linux build sandbox, which hides
  # /sys/devices/system/cpu ("process is affinitize to 0 CPUs", measured on a NixOS x86_64 VM,
  # 2026-10-15); workstation GC builds. Build-time only.
  env.DOTNET_gcServer = "0";

  # buildDotnetModule derives -p:Version from the numeric parts only (3.9.2.2); the package
  # version officeproof pins is the prerelease string.
  packNupkg = true;
  dontPublish = true;
  dotnetPackFlags = [ "-p:PackageVersion=${version}" ];

  meta = {
    description = "Open XML PowerTools for .NET, with officeproof's RevisionProcessor fixes";
    homepage = "https://github.com/rmorgans/Clippit";
    license = lib.licenses.mit;
  };
}

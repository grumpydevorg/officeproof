{
  lib,
  buildDotnetModule,
  dotnetCorePackages,
  callPackage,
  poppler-utils,
}:

let
  clippit = callPackage ./clippit.nix { };
in
buildDotnetModule {
  pname = "officeproof";
  version = "0.1.0";

  src = lib.cleanSource ../.;

  # nuget.config points Clippit at a per-user local feed for source builds; here it comes from
  # the clippit derivation through projectReferences, and the sandbox has its own sources.
  postPatch = ''
    rm nuget.config
  '';

  projectFile = "src/OfficeProof.Cli/OfficeProof.Cli.csproj";
  nugetDeps = ./deps.json;
  projectReferences = [ clippit ];
  dotnet-sdk = dotnetCorePackages.sdk_10_0;
  dotnet-runtime = dotnetCorePackages.runtime_10_0;
  executables = [ "officeproof" ];

  # verify --word compares Word's PDFs page by page through poppler's pdftoppm.
  makeWrapperArgs = [
    "--prefix"
    "PATH"
    ":"
    (lib.makeBinPath [ poppler-utils ])
  ];

  # The test suite runs from the source tree (dotnet test); the package proves the installed
  # binary works end to end, Clippit included: generate a sample from the example template,
  # validate it, and verify it against itself.
  doInstallCheck = true;
  installCheckPhase = ''
    runHook preInstallCheck
    export HOME=$TMPDIR
    $out/bin/officeproof generate --template ${../examples/minimal-template} -o $TMPDIR/gen
    sample=$(echo $TMPDIR/gen/*-sample.docx)
    $out/bin/officeproof validate "$sample"
    $out/bin/officeproof verify "$sample" "$sample"
    runHook postInstallCheck
  '';

  passthru = { inherit clippit; };

  meta = {
    description = "Edit Word and Excel documents as an agent and prove the result";
    homepage = "https://github.com/grumpydevorg/officeproof";
    license = lib.licenses.mit;
    mainProgram = "officeproof";
    platforms = lib.platforms.unix;
  };
}

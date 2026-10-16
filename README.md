# officeproof

Edit Word and Excel files as an agent, and prove the result.

officeproof is a command-line tool (.NET 10) for changing `.docx`, `.dotx` and `.xlsx`
packages without breaking them, and for checking that a change is exactly what it claims to be:

- **Tracked redlines you can trust.** `redline` writes tracked changes and comments into
  someone else's document; `verify` proves that Reject All restores the original in *every*
  story (body, headers, footers, footnotes, endnotes, text boxes, comments), compared
  canonically: text, style, list numbers, paragraph and run formatting, tables, sections and
  what each relationship resolves to. It checks in-process with
  [Clippit](https://github.com/sergey-tihon/Clippit) and, with `--word`, also has Microsoft
  Word reject all changes and compares Word's result story by story and page by page.
- **Lint** for structure that breaks: typed heading and caption numbers, duplicate numbers,
  typed cross-references and missing targets, TOC direct formatting, dangling styles,
  off-palette colours, unresolved revisions, revisions hidden inside style definitions.
- **Untracked replace** for your own documents: find/replace across run boundaries in every
  story, with match-count guards and a schema check.
- **Templates as data**: generate a `.dotx` and sample from a directory of parts, or restyle an
  existing report with it.
- **Workbooks**: inspect structure and risks, and set cells while every other part stays
  byte-identical.

Every writing verb takes an input and a separate output; nothing is edited in place.

## Install

Download the archive for your platform from
[Releases](https://github.com/grumpydevorg/officeproof/releases): one self-contained executable
(no .NET needed) with this README and the licence; `SHA256SUMS` lists the archives' checksums.

| Platform | Archive |
|---|---|
| Windows x64, Arm64 | `officeproof-<version>-win-x64.zip`, `officeproof-<version>-win-arm64.zip` |
| macOS Apple silicon, Intel | `officeproof-<version>-osx-arm64.tar.gz`, `officeproof-<version>-osx-x64.tar.gz` |
| Linux x64, Arm64 (glibc) | `officeproof-<version>-linux-x64.tar.gz`, `officeproof-<version>-linux-arm64.tar.gz` |

```powershell
# Windows (PowerShell); then add the folder to PATH
Expand-Archive officeproof-<version>-win-x64.zip -DestinationPath $env:LOCALAPPDATA\Programs
& "$env:LOCALAPPDATA\Programs\officeproof-<version>-win-x64\officeproof.exe" --help
```

```sh
# macOS and Linux; then copy officeproof onto PATH
tar -xzf officeproof-<version>-osx-arm64.tar.gz
./officeproof-<version>-osx-arm64/officeproof --help
```

The executables are not signed by a developer certificate. macOS refuses to open one a browser
downloaded until its quarantine flag is removed
(`xattr -d com.apple.quarantine officeproof`; `curl` downloads carry no flag). The Linux builds
need ICU (`libicu`), which desktop distributions include. Page comparison needs poppler's
`pdftoppm` on PATH (`pdftoppm.exe` on Windows).

With Nix (macOS or Linux; builds the patched Clippit too, and puts `pdftoppm` on the path):

```sh
nix run github:grumpydevorg/officeproof -- --help
nix profile install github:grumpydevorg/officeproof
```

The flake also exports `overlays.default` (adds `officeproof` to a package set).

## Build from source

Requires the .NET 10 SDK and Git. While the patched Clippit is unreleased, build it into the
repository-local NuGet feed `.local/nuget-feed` once per checkout. The script is bash: on
Windows run it in Git Bash (installed with Git for Windows); from PowerShell call Git's bash by
path (`& "$env:ProgramFiles\Git\bin\bash.exe" eng/pack-clippit.sh`), since `bash` there may be WSL's.

```sh
bash eng/pack-clippit.sh      # clones the pinned Clippit fork commit into .local/Clippit (override with CLIPPIT_DIR)
dotnet build officeproof.slnx
dotnet test --solution officeproof.slnx --no-build
```

A release-style executable (any of the runtime identifiers above, e.g. `win-x64`), and the
end-to-end check the release workflow runs on it:

```sh
dotnet publish src/OfficeProof.Cli -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o publish
bash eng/smoke.sh publish/officeproof.exe
```

Run the CLI with `dotnet src/OfficeProof.Cli/bin/Debug/net10.0/officeproof.dll <verb>` (or
`dotnet run --project src/OfficeProof.Cli -- <verb>`). `--word` and the Word tests need
Microsoft Word (for Mac, or for Windows in a desktop session); page comparison needs poppler's
`pdftoppm`.

`nix/` packages the same build: `nix/clippit.nix` packs the Clippit fork commit that
`eng/pack-clippit.sh` uses, and `nix/deps.json` / `nix/clippit-deps.json` lock the NuGet
dependencies. After changing a package reference, regenerate them with
`nix build .#officeproof.fetch-deps -o fd && ./fd nix/deps.json` (and `.clippit.fetch-deps`
for Clippit).

## Examples

```sh
# Redline someone else's report, then prove Reject All gives their document back
officeproof redline report.docx report-redline.docx --edits edits.json --author "Reviewer"
officeproof verify report.docx report-redline.docx --word

# What numbers does Word actually show, before and after the changes?
officeproof outline report-redline.docx --view reject

# Lint for typed numbers, broken references and formatting debt
officeproof lint report.docx --template examples/minimal-template

# Untracked replace in your own document; refuse unless exactly 3 matches
officeproof replace draft.docx final.docx --find "Acme Pty" --with "Acme Pty Ltd" --expect 3

# Templates
officeproof generate --template examples/minimal-template -o out/
officeproof retemplate old.docx restyled.docx --template examples/minimal-template --metadata report.yaml

# Workbooks
officeproof xlsx inspect budget.xlsx --findings
officeproof xlsx set budget.xlsx budget-2.xlsx 'Summary!B4=1250'
```

`officeproof --help` and `officeproof <verb> --help` list every option. Exit codes: 0 clean,
1 an error finding, 2 bad input or a refused operation. The template directory format is
documented in [examples/minimal-template/README.md](examples/minimal-template/README.md).

## How it compares

Redline validators that diff the text of `document.xml` after Reject All miss most of what
goes wrong in practice: a style or numbering definition changed outside tracked changes, a
formatting change with no revision mark, a retargeted hyperlink or replaced image, edits in
headers, footers or notes. On a real 42-page report, such a validator passed all four
untracked defects planted in a redline; `verify` fails each of them. It also catches a trap
Word itself sets: tracked changes inside style definitions are *not* undone by Word's Reject
All, so `verify` refuses them and `redline` numbers headings with new styles instead.

## Limits

- Comment anchor positions, bookmark ids, rsids and paragraph ids are treated as bookkeeping
  and not compared; inherited headers are not resolved per section; charts are compared only
  as bytes; custom XML parts are warnings only; the glossary document is not compared.
- Revision kinds where Clippit and Word, or Word for Mac and Word for Windows, disagree, or
  that have no planted test case (`w:cellIns`, `w:cellMerge`, `w:cellDel`, `w:tblPrExChange`,
  revisions in `w:dir`/`w:bdo`, run-level `w:customXml`) fail closed as unverifiable.
- `--word` drives Microsoft Word for Mac (AppleScript) or Word for Windows (COM, in a desktop
  session, not over SSH or as a service); elsewhere only the in-process checks run.
- Depends on a patched Clippit build (`eng/pack-clippit.sh`) until the fixes reach nuget.org.

## Contributing

See [AGENTS.md](AGENTS.md) for the architecture, methodology and conventions.

## Licence

MIT; see [LICENSE](LICENSE).

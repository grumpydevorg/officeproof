# officeproof: contributor and agent guide

officeproof is a .NET solution for doing real work with Word `.docx`/`.dotx` and Excel
`.xlsx` packages and proving the result: generate a house template, restyle a report with it,
lint a document for structure that will break, find/replace in your own documents, redline
someone else's document with tracked changes, verify the redline the way Word will see it,
and inspect or edit a workbook without disturbing it. User-facing documentation is in
[README.md](README.md); this file is for people and agents changing the code.

## Toolchain

- **.NET 10** (`global.json` pins SDK 10.0.400, roll-forward on feature band). Solution file
  `officeproof.slnx`; shared build settings in `Directory.Build.props` (nullable, warnings as
  errors, latest-recommended analyzers, code style enforced in build); every package version
  in `Directory.Packages.props` (central package management).
- **Open XML SDK** (`DocumentFormat.OpenXml` 3.5): packaging, typed model, schema validator.
- **Clippit** (Verify, Edit and the Redline tests): revision accept/reject
  (`RevisionProcessor`), list-number retrieval (`ListItemRetriever`), run-splitting
  find/replace (`OpenXmlRegex`). Pinned to 3.9.2-rmorgans.2: 3.9.1 plus eight
  RevisionProcessor fixes, open upstream as sergey-tihon/Clippit#548-#555 and merged on the
  fork https://github.com/rmorgans/Clippit. `eng/pack-clippit.sh` builds it from that exact
  commit (cloning into `CLIPPIT_DIR`, default `.local/Clippit`) into the repository-local feed `.local/nuget-feed`
  that `nuget.config` names; run it once per checkout before the first restore (on Windows in
  Git Bash: from PowerShell, `bash` may be WSL's). `nix/clippit.nix`
  packs the same commit for the Nix package. When a nuget.org release contains the fixes, pin
  it and delete the script, `nuget.config`, `nix/clippit.nix` (and the `projectReferences`
  entry in `nix/package.nix`), the Clippit feed steps in `.github/workflows/` and `ClippitWorkarounds`.
- **Nix**: `flake.nix` exports `packages.<system>.officeproof` and `overlays.default`;
  `nix/package.nix` builds the CLI with `buildDotnetModule` and its install check generates,
  validates and verifies the example template's sample. `nix/deps.json` and
  `nix/clippit-deps.json` lock NuGet packages; regenerate after any package-reference change:
  `nix build .#officeproof.fetch-deps -o fd && ./fd nix/deps.json` (`.clippit.fetch-deps` for
  Clippit). Tests run from source (`dotnet test`), not in the Nix build.
- **poppler** `pdftoppm` on PATH for `PageDiff` (Word's PDFs to page images; the Nix wrapper
  adds it). On Windows `PageDiff.FindOnPath` resolves `pdftoppm.exe` through `PATHEXT`.
- **CI** (`.github/workflows/ci.yml`, push to `main` and pull requests): build with warnings as
  errors and test on Ubuntu, Windows and macOS (Clippit feed cached on the script's hash; Word
  tests skip), and `nix build .#officeproof`. Windows checkouts are CRLF (`core.autocrlf`) and
  the tests pass on them; `.gitattributes` keeps `*.sh` LF for shells other than Git Bash.
- **Releases** (`.github/workflows/release.yml`): a `v*` tag publishes a self-contained
  single-file executable (`-p:Version` from the tag) for win-x64, win-arm64, linux-x64,
  linux-arm64, osx-arm64 and osx-x64, each on a native runner, runs `eng/smoke.sh` on it
  (generate, validate, a redline that verifies, an untracked replace that does not) and attaches
  the archives and `SHA256SUMS` to the tag's release; run by hand it builds the archives as
  artifacts only. No trimming or AOT (the Windows Word oracle uses built-in COM interop). The
  archive holds the executable alone: no verb loads Clippit's SkiaSharp native library (Clippit
  uses it only for text metrics in `FormattingAssembler.AssembleFormatting` and the HTML
  converters); a verb that reaches those must ship it (`-p:IncludeNativeLibrariesForSelfExtract=true`).
- **System.CommandLine** 2.0 for the host; **xUnit v3** on Microsoft.Testing.Platform.
- Homebrew .NET: the `officeproof` apphost cannot find the runtime unless
  `DOTNET_ROOT=/opt/homebrew/opt/dotnet/libexec`; run `dotnet <path>/officeproof.dll` instead.

```
bash eng/pack-clippit.sh                                    # once per checkout: the pinned Clippit
dotnet build officeproof.slnx
dotnet test --solution officeproof.slnx --no-build          # the 68 Word tests skip unless OFFICEPROOF_WORD_TESTS=1
dotnet run --project src/OfficeProof.Cli -- <verb> ...      # or: dotnet src/OfficeProof.Cli/bin/Debug/net10.0/officeproof.dll
```

Word tests drive the real Microsoft Word (see [The Word oracle](#the-word-oracle)); set
`OFFICEPROOF_WORD_TESTS=1` only where Word is installed: on a Mac, or on Windows in a desktop
session. A smoke test runs `verify` on a real pair of documents from outside the tree:
`OFFICEPROOF_SMOKE_ORIGINAL`, `OFFICEPROOF_SMOKE_REDLINE`, optionally `OFFICEPROOF_SMOKE_OUT`.

## Components

Each component is a class library that references `OfficeProof.Core` (Verify and Edit also
Clippit), returns results or `Finding`s and never prints. Written packages get normal zip
entry modes (`Docx.NormalizeEntryModes`; the SDK writes mode 0, which `unzip` extracts
unreadable). `OfficeProof.Cli` composes them into verbs.

| Project | Owns | Key types |
|---|---|---|
| `OfficeProof.Core` | Namespaces, run-model text, findings, package IO, schema validation | `Ns`, `Wml`, `Finding`, `FindingReport`, `Docx` |
| `OfficeProof.Templating` | A template directory as code: generate, retemplate, doc props, data store, font embedding, header/footer, settings, identity | `TemplateSource`, `TemplateManifest`, `ReportMetadata`, `Generator`, `Retemplater`, `FontEmbedder`, `IdentityReader` |
| `OfficeProof.Lint` | Read-only rules over a document | `ILintRule`, `Linter`, `LintContext`, `BrandPalette` |
| `OfficeProof.Edit` | Untracked find/replace in every story of our own documents | `Replacer`, `ReplacePair`, `ReplaceOptions` |
| `OfficeProof.Redline` | Tracked-change authoring on someone else's document | `TrackedEditor`, `EditScript`, `CaptionFields`, `HeadingNumbering`, `TocStyles` |
| `OfficeProof.Verify` | Accept/Reject-All views of every story, relationship-resolved comparison, the revision allowlist, Microsoft Word as oracle (outline, Reject-All save, PDF), page diff | `Outline`, `RevisionCheck`, `StoryView`, `StoryComparison`, `RevisionAllowlist`, `ClippitWorkarounds`, `WordOracle`, `WordCheck`, `PageDiff` |
| `OfficeProof.Spreadsheet` | Workbook inspection and part-preserving cell edits | `WorkbookInspector`, `CellEditor` |
| `tests/OfficeProof.TestKit` | Code-built fixtures (`DocxBuilder.TypedNumberReport`: a report with typed heading/caption numbers and custom heading styles `Head1`/`Head2`/`Head3`) | `DocxBuilder` |

## Verbs

| Verb | Does |
|---|---|
| `generate [--template] [--metadata] [--fonts] [-o]` | Emit `<name>.dotx` + `<name>-sample.docx` from the template directory; validate both |
| `retemplate <in> <out> [--template] [--metadata] [--force]` | Restyle a report (untracked rewrite): remap the template's `headingStyles` to its headings, drop typed numbers, strip direct fonts (keep cover sizing), swap canonical parts, bind the cover, logo footer, fonts, `updateFields`; then validate + `dangling-style` + `brand-colour`. Refuses a document that does not carry the template's `identityKeys` unless `--metadata` is given |
| `validate <file>` | Open XML SDK schema validation (Office 2019), `.docx` or `.xlsx` |
| `lint <in> [--template] [--rule id]... [--list-rules]` | Typed heading/caption numbers, numbers shown twice, duplicate numbers, typed cross-references (plural lists included), missing reference targets, TOC direct formatting, dangling styles, off-palette colours (needs a template palette), unresolved revisions, revisions inside style definitions, tables without `w:tblLook` (Word styles them at random per open), `w:updateFields` (Word asks to update fields on open). Chapter numbers ("Figure 2.1") are not single-counter captions and are left alone |
| `replace <in> <out> --find --with [--regex] [--story] [--expect] [--force] [--json]` | Untracked find/replace in our own documents, across runs, in every story; refuses documents with tracked changes, a pair that matches nothing, a total other than `--expect`, and matches covering fields or drawings; reports new schema errors |
| `redline <in> <out> [--edits] [--captions] [--fix] [--headings] [--level] [--toc-font] [--force]` | Tracked changes + comments, then the `verify` checks |
| `verify <original> <redline> [--allow-style ids] [--word [--keep-pages dir]]` | Reject All restores the original in every story (see below), no revisions in styles/numbering, revisions only where the allowlist has evidence, references resolve after Accept All; an existing style, list, theme or layout setting changed outside tracked changes is an error unless the style is declared with `--allow-style` (`redline` declares the TOC styles `--toc-font` changes); `--word` has Word reject all changes in both documents and compares every story and every printed page, compares heading numbers in both views and updates fields |
| `outline <in> [--view as-is\|accept\|reject] [--word]` | Headings and numbered paragraphs with the numbers Word shows |
| `xlsx inspect <file> [--findings]`, `xlsx set <in> <out> Sheet!Cell=value...` | Workbook structure and risks; cell edits that keep every other part byte-identical, drop calcChain and set `fullCalcOnLoad` |

Exit codes: 0 clean, 1 an error finding, 2 bad input or refused operation. Writing verbs refuse
an existing output unless `--force`, and a failed run removes only the output it wrote. Errors
print as `officeproof: <message>`.

Templates are data: `--template <dir>` (or `OFFICEPROOF_TEMPLATE`) names a directory whose
schema is documented in [examples/minimal-template/README.md](examples/minimal-template/README.md).
There is no built-in template.

## The Word oracle

`WordOracle` (outline, Accept/Reject-All save, PDF, field update) asks the real Microsoft Word
through one backend per platform, with one API and the same answers: `MacWord` drives Word for
Mac through AppleScript (osascript; Word is sandboxed, so copies live in its container), and
`WindowsWord` drives Word for Windows through in-process COM (late-bound `dynamic`; each call
starts its own invisible Word instance on an STA thread and quits it, so a Word the user has
open is never touched). Both pass the Word tests: Word for Mac 16.113 (macOS 26.6) and Word for
Windows 16.0.20430 (Microsoft 365, Windows 11 on arm64), 2026-10-15/16.

- Every call opens a temporary copy, one officeproof document in Word at a time across
  processes (a lock file), alerts off, closed without saving even on error. At the timeout the
  Mac backend kills osascript and closes the copy; the Windows backend kills the Word process
  it started (never one it cannot identify as its own).
- The copy (`WordOracle.CopyForWord`): run-level `w:customXml` is refused (Word for Mac
  answers it with a modal dialog that blocks automation); `w:updateFields`, which `generate`
  and `retemplate` set, is dropped (Word for Mac asks about it in a modal dialog that may open
  off screen; `verify --word` on generate's sample hung until its timeout); a table without
  `w:tblLook` gets the documented 0x04A0 (both Words otherwise read it as 0000 or 04A0 at
  random, per open). Word's verdict then covers the adjusted copies, not the files as the
  recipient opens them, so `verify --word` says so (`word-adjusted`, Warning) whenever either
  adjustment applied, and `lint` reports both conditions in the file (`table-look-missing`,
  `update-fields-on-open`).
- Windows: Word cannot open or save documents from session 0 (an SSH logon, a service), so the
  backend refuses there; run it in a desktop session (over SSH: a scheduled task with an
  interactive logon whose settings allow battery power, or it never starts on a laptop host).
  Word's PDF after Accept/Reject All is exported from a fresh open of its save of the view: the
  live view laid an edited line out 0.12 pt off. COM needs built-in interop, so the CLI is
  never trimmed or AOT-compiled.
- One recorded difference: Word for Windows' Reject All of a row-level `w:tblPrExChange`
  rebuilds the table grid, moving the row's cell edges and changing the printed page; Word for
  Mac and Clippit restore the row. A redline the two Words reject differently cannot be shown
  to reject back to the original, so the allowlist refuses it.

### Revision evidence and re-measuring Word

`src/OfficeProof.Verify/revision-evidence.json` (embedded, read by `RevisionEvidence`) is the
one record of what is known: each engine (Word for Mac, Word for Windows, Clippit,
officeproof's view) and the Word builds measured, each probe's Reject/Accept outcome per engine,
and each revision site's evidence (a planted case, the probes it cites). `RevisionAllowlist`
admits a site only when a planted case exercises it and every cited probe agrees in every
engine; an `unverifiable-revision` finding names the engines that disagree. The tests hold the
file to the code: probe fixtures (`ContainerProbeTests`, `RevisionKindProbeTests`) must match
its Clippit and view columns on every run and its Word column for the installed Word with
`OFFICEPROOF_WORD_TESTS=1`; each `"planted": true` must be exercised by a `PlantedDefectTests`
case.

`verify --word` reports the installed Word's version (`word-version`): Info on a build of the
recorded version, Warning otherwise. Microsoft ships Word monthly, so before a release run the
Word tests on a Mac and on Windows (in a desktop session); after a Word update
`Installed_word_is_a_build_the_revision_evidence_was_measured_on` fails until the other Word
tests pass and the new version and date are recorded under `engines`. A changed outcome is
recorded in the probe's column, which changes what the allowlist admits.

## The methodology

The goal is **separation of content and presentation**: content references named styles,
formatting lives once in the styles/theme, direct formatting is near zero, and every number a
reader sees (headings, captions, cross-references, TOC) is generated by Word, not typed.

- **Canonical bundle** = `styles.xml` + `theme1.xml` + `docDefaults` (the theme resolves
  `minorHAnsi` to the concrete font, so all three are the spec, not styles alone).
- **Our own documents** are restyled directly (`retemplate`) or edited untracked
  (`replace`), then linted.
- **Someone else's document** is changed only as tracked changes they accept or reject
  (`redline`), and every redline is verified before it is sent (`verify`, `--word` when
  numbering, styles or fields changed).

### What a 42-page client report (2026-10-15) established

- Typed numbers everywhere (headings, captions, cross-references) are the common failure;
  only the TOC was a field. `lint` now finds each kind.
- **Word does not undo tracked changes inside style definitions on Reject All.** Linking the
  existing heading styles to a list left "2.0 1.0 Summary" after Reject All. Heading numbering
  that survives Reject All uses new `<style> numbered` styles applied per paragraph as a
  tracked `w:pPrChange` (`HeadingNumbering`). `verify` flags `revision-in-styles` and lists
  every untracked part change.
- A text-only Reject-All diff (pandoc) cannot see numbering, and Anthropic's validator (text of
  `document.xml`) passed all four untracked defects planted in that report's redline. `verify`
  compares every story after Reject All: main, headers and footers paired by section and
  type, referenced footnotes and endnotes, text boxes, and comments (pre-existing ones must be
  unchanged). Paragraphs are compared canonically (text, style, list number, paragraph and run
  properties, tables, sections) with every relationship reference replaced by what it resolves
  to (external URI, or a digest of the target part), so a retargeted hyperlink or replaced
  image fails even with `document.xml` byte-identical. `PlantedDefectTests` holds one
  untracked (must fail) and one tracked (must pass) case per construct.
- Clippit 3.9.1 drops deleted text inside `w:hyperlink`, `w:fldSimple`, `w:dir` and `w:bdo` on
  Reject All, empties header references on a rejected `w:sectPrChange`, leaves an empty
  paragraph for a moved paragraph mark, skips comment text and throws accepting a `w:cellDel`
  after a cell without `w:tcPr`; the pinned 3.9.2-rmorgans.2 fixes all of these.
  `ClippitWorkarounds` (marker pairs for hyperlinks and simple fields, the other two fixed
  up front) stays until a nuget.org release has the fixes. Revisions anywhere else fail closed
  (`unverifiable-revision`): `RevisionAllowlist` admits only the sites
  [the revision evidence](#revision-evidence-and-re-measuring-word) supports. Still off it:
  table, row, cell and section property changes (`w:tblPrChange`, `w:trPrChange`,
  `w:tcPrChange`, `w:sectPrChange`: Word for Mac's Reject All of the planted cases rebuilt the
  table grid and wrote its own header and footer distances, 2026-10-16), `w:cellIns` (the pinned
  Clippit leaves the row a cell short), `w:cellMerge` (it keeps a rejected merge) and
  `w:tblPrExChange` (Word for Windows rejects it differently from Word for Mac), where the
  engines disagree; `w:cellDel` and revisions in `w:dir`/`w:bdo`, which agree but have no
  planted cases yet (the SDK validator also rejects `w:dir`/`w:bdo`); and run-level
  `w:customXml` (Word answers it with a modal dialog). `redline` writes none of these; for a
  reviewer's redline only `--word` can check them.
- Word for Mac (16.113) and Word for Windows (16.0.20430) accept and reject tracked changes in
  headers, footers, notes, comments, text boxes, hyperlinks, content controls, smart tags,
  simple fields, table cells, moves and numbering changes the same way the pinned Clippit
  does. Word leaves an empty `w:hyperlink` behind when rejecting an inserted link and renumbers
  drawing ids on every save; the comparison ignores both. `--word` also compares Word's own
  Reject-All save of each document story by story, and their PDFs page by page at zero pixel
  tolerance (two exports of one document are pixel-identical).
- Untracked changes to definitions (`definition-changed`, Error): an existing style changed or
  removed unless declared, `docDefaults`, an existing list (`abstractNum`/`num`), the theme,
  and settings that affect layout. New styles and lists, review-state settings (rsids, zoom,
  proof state) and fonts, web settings and custom XML parts are warnings. `--word` gives the
  original only the redline's new and declared definitions before Word rejects both, so an
  undeclared change shows in Word's comparison too.
- Bounds that remain: comment anchor positions, bookmark ids and `_Ref`/`_GoBack` bookmarks,
  rsids and paragraph ids are ignored as bookkeeping; inherited headers are not resolved per
  section; text inside charts is compared only as the chart part's bytes; custom XML parts
  (which bound content controls display) are warnings only; the glossary document is never
  compared.
- Update Field rebuilds TOC lines from the TOC1..n styles and drops direct run formatting
  (Arial became Times New Roman). Put the font in the styles (`--toc-font`; lint
  `toc-direct-formatting`).

## Principles

- **Change the template, not the content.** A formatting fix is a style/theme edit, regenerate,
  never a per-document edit. Reviewers argue about the template, not the words.
- **Metadata is per report.** `--metadata <report.yaml>`; the template's `report.yaml` is only
  the default for `generate`. A shared metadata file once stamped one project's identity on
  another's report (2026-08-26).
- **Never edit in place.** Every writing verb takes an input and a different output.
- **Measure, don't ship the flatten.** Flatten only to read/match effective formatting; ship
  style-referencing documents.
- **Deterministic verdict.** Pass/fail from rules and structural comparison; reserve a model
  for the semantic residue only, never for labelling.
- **Preserve the cover's sizing** when restyling a real report (keep `sz` before the first
  heading); normalise everything after.
- **Real documents never enter the repository.** Tests use `DocxBuilder` fixtures and
  `examples/minimal-template`; real files are smoke-tested from outside the tree (keep local
  ones in the ignored `.local/`).

## Not built yet

- `normalise`: fingerprint effective formatting (Clippit `FormattingAssembler`), match to the
  canonical bundle, promote direct formatting to styles, emit the change as a redline.
- Lint for effective body font/size and direct-formatting debt, heading nesting, leftover
  highlights/placeholders, non-Word round-trip (rsid == 0); list levels that print no number;
  built-in Word style colours (Hyperlink, Heading1) still trip `brand-colour`.
- "Appendix N" references (typed eleven times in the 2026-10-15 report) are neither linted
  nor linked.

## Layout

```
officeproof.slnx, global.json, Directory.Build.props, Directory.Packages.props, nuget.config, .gitattributes
eng/pack-clippit.sh, eng/smoke.sh
.github/workflows/{ci.yml,release.yml}
flake.nix, flake.lock, nix/{package.nix,clippit.nix,deps.json,clippit-deps.json}
src/OfficeProof.{Core,Templating,Lint,Edit,Redline,Verify,Spreadsheet,Cli}/
tests/OfficeProof.{Templating,Lint,Edit,Redline,Verify,Spreadsheet}.Tests/, tests/OfficeProof.TestKit/
examples/minimal-template/   a complete, house-neutral template directory and its schema
```

`generate` validates both outputs with the Open XML SDK validator; 0 errors before shipping.

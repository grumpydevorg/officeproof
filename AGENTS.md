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
  commit (cloning into `CLIPPIT_DIR`, default `.local/Clippit`) into the local feed
  `nuget.config` names; run it once per machine before the first restore. When a nuget.org
  release contains the fixes, pin it and delete the script, `nuget.config` and
  `ClippitWorkarounds`.
- **poppler** `pdftoppm` on PATH for `PageDiff` (Word's PDFs to page images).
- **System.CommandLine** 2.0 for the host; **xUnit v3** on Microsoft.Testing.Platform.
- Homebrew .NET: the `officeproof` apphost cannot find the runtime unless
  `DOTNET_ROOT=/opt/homebrew/opt/dotnet/libexec`; run `dotnet <path>/officeproof.dll` instead.

```
eng/pack-clippit.sh                                         # once per machine: the pinned Clippit
dotnet build officeproof.slnx
dotnet test --solution officeproof.slnx --no-build          # Word tests (about 40) skip unless OFFICEPROOF_WORD_TESTS=1
dotnet run --project src/OfficeProof.Cli -- <verb> ...      # or: dotnet src/OfficeProof.Cli/bin/Debug/net10.0/officeproof.dll
```

Word tests drive the real Microsoft Word for Mac through AppleScript; set
`OFFICEPROOF_WORD_TESTS=1` only on a Mac with Word installed. A smoke test runs `verify` on a
real pair of documents from outside the tree: `OFFICEPROOF_SMOKE_ORIGINAL`,
`OFFICEPROOF_SMOKE_REDLINE`, optionally `OFFICEPROOF_SMOKE_OUT`.

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
| `lint <in> [--template] [--rule id]... [--list-rules]` | Typed heading/caption numbers, numbers shown twice, duplicate numbers, typed cross-references (plural lists included), missing reference targets, TOC direct formatting, dangling styles, off-palette colours (needs a template palette), unresolved revisions, revisions inside style definitions. Chapter numbers ("Figure 2.1") are not single-counter captions and are left alone |
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
  (`unverifiable-revision`): `RevisionAllowlist` lists each allowed site with its evidence.
  Still off it: `w:cellIns` (the pinned Clippit leaves the row a cell short) and `w:cellMerge`
  (it keeps a rejected merge), where Clippit and Word disagree (`RevisionKindProbeTests`);
  `w:cellDel` and revisions in `w:dir`/`w:bdo`, which now agree but have no planted cases yet
  (the SDK validator also rejects `w:dir`/`w:bdo`); and run-level `w:customXml` (Word answers
  it with a modal dialog).
- Word (16.113) accepts and rejects tracked changes in headers, footers, notes, comments, text
  boxes, hyperlinks, content controls, smart tags, simple fields, table cells, moves, numbering
  changes and table-exception changes the same way the pinned Clippit does
  (`ContainerProbeTests`, `RevisionKindProbeTests`). `--word` therefore
  also compares Word's own Reject-All save of each document story by story, and their PDFs
  page by page at zero pixel tolerance (two exports of one document are pixel-identical).
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
officeproof.slnx, global.json, Directory.Build.props, Directory.Packages.props, nuget.config
eng/pack-clippit.sh
src/OfficeProof.{Core,Templating,Lint,Edit,Redline,Verify,Spreadsheet,Cli}/
tests/OfficeProof.{Templating,Lint,Edit,Redline,Verify,Spreadsheet}.Tests/, tests/OfficeProof.TestKit/
examples/minimal-template/   a complete, house-neutral template directory and its schema
```

`generate` validates both outputs with the Open XML SDK validator; 0 errors before shipping.

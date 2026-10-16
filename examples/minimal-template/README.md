# Minimal template

A plain, house-neutral template directory: theme fonts (Calibri), a blue-grey palette, no logo,
no embedded fonts. The tests use it, and it is the starting point for a real house template:
copy the directory, edit the parts and `template.yaml`, and keep it in its own repository.

```
officeproof generate --template examples/minimal-template -o out/      # Example.dotx, Example-sample.docx
officeproof retemplate in.docx out.docx --template <dir> --metadata report.yaml
officeproof lint in.docx --template <dir>                             # brand-colour uses the palette
```

`--template` may be omitted when `OFFICEPROOF_TEMPLATE` names the directory. officeproof has no
built-in template: `generate` and `retemplate` need one, and `lint` without one skips
`brand-colour` with an Info finding.

## Files

| File | Required | Role |
|---|---|---|
| `template.yaml` | yes | The manifest (below) |
| `styles.xml`, `theme1.xml`, `numbering.xml` | yes | The canonical bundle, copied into every output (`retemplate` replaces the document's own) |
| `settings.xml` | yes | `generate`'s settings part; `retemplate` keeps the document's and adds `embedTrueTypeFonts` and `updateFields` |
| `cover-document.xml` | yes | `generate`'s `document.xml`: cover and sample body |
| `header.xml` | yes | Running header on every page but the first |
| `header-first.xml` | yes | First-page header; holds the optional full-bleed cover image |
| `footer.xml` | yes | Footer on every page |
| `cover-block.xml` | with `coverAnchor` | Cover block `retemplate` inserts in place of a hand-typed one |
| `report.yaml` | no | Default metadata for `generate` (a top-level `report:` map); per-report files go to `--metadata` |
| logo image | when `logo` is set | Footer image |

Part templates are filled before they are written. Structural tokens are replaced verbatim:
`{{STOREID}}` (the data store's item id, for `w:dataBinding/@w:storeItemID`),
`{{HEADERID}}`, `{{FIRSTHEADERID}}` and `{{FOOTERID}}` (relationship ids, in
`cover-document.xml`), `{{COVERIMAGE}}` (`header-first.xml`) and `{{LOGOID}}`
(`footer.xml`, only with `logo`). Every metadata key fills `{{key}}`, XML-escaped.

## template.yaml

Unknown keys are errors. Only `name` and `dataStore` are required.

```yaml
name: Example                     # generate writes Example.dotx and Example-sample.docx

dataStore:                        # the custom-XML store bound controls read and write
  namespace: https://example.org/report   # root <prefix:report>, one child per metadata key;
  prefix: rpt                             # the parts' w:dataBinding prefixMappings name the same URI

logo: logo.png                    # optional: .png/.jpg/.jpeg for footer.xml's {{LOGOID}}

fonts:                            # optional: a family embedded in every output (obfuscated .odttf);
  family: Example Sans            # faces are looked up in --fonts, else the default font
  regular: ExampleSans-Regular.otf          # directories (below); a missing face is reported, not fatal
  bold: ExampleSans-Bold.otf
  italic: ExampleSans-Italic.otf
  boldItalic: ExampleSans-BoldItalic.otf

coreProperties:                   # optional: docProps/core.xml
  title: "{{title}}"              # "{{key}}" takes the metadata value (blank when absent);
  subject: "{{preparedFor}}"      # any other value is fixed text. Allowed: title, subject,
  creator: "{{author}}"           # creator, description, language, keywords, category,
  category: Report                # contentStatus. generate also sets lastModifiedBy to the
                                  # creator and revision 1; retemplate leaves both alone

customProperties:                 # optional: docProps/custom.xml (File > Properties > Custom);
  ProjectCode: "{{projectCode}}"  # a bound property with no metadata value is not written

identityKeys: [title, projectCode]   # optional: what a document must carry itself (in its data
                                     # store or bound properties) before retemplate runs without
                                     # --metadata; empty or absent means --metadata is always needed

headingStyles:                    # optional: retemplate maps these legacy paragraph styles to
  OldHeading1: Heading1           # the template's headings and drops their typed numbers
  OldHeading2: Heading2

coverAnchor: "Registered office"  # optional: retemplate replaces every body paragraph before the
                                  # first one containing this text with cover-block.xml

palette:                          # optional: lint brand-colour flags explicit colours outside it
  colours:
    1f4e79: dark blue             # 6-digit hex (with or without '#') -> name
  nearMisses:
    1f4e78: 1f4e79 (dark blue) - one digit out   # hex -> hint shown with the finding
```

Font faces are found by file name, at the top of a directory or in its subdirectories. With
`--fonts <dir>` only that directory is searched; otherwise the user's font directory, then the
machine's (`generate --help` prints them):

| OS | Default font directories, in order |
|---|---|
| macOS | `~/Library/Fonts`, `/Library/Fonts` |
| Windows | `%LOCALAPPDATA%\Microsoft\Windows\Fonts` (installed for the current user), `%WINDIR%\Fonts` |
| Linux | `$XDG_DATA_HOME/fonts` (default `~/.local/share/fonts`), `/usr/local/share/fonts`, `/usr/share/fonts` |

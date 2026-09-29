# onenote-md

Exports your local OneNote notebooks to a Markdown folder tree.

It talks to the **OneNote desktop app you already have installed** through its
local COM automation interface. There is no Microsoft Graph, no Azure AD app
registration, no OAuth, and no API permission to grant — nothing leaves the
machine.

```
onenote-md                      # export everything to ./onenote-export
onenote-md D:\notes --list      # show the hierarchy, write nothing
onenote-md D:\notes --notebook <name>
onenote-md D:\notes --dry-run
```

## Build

```
build.cmd
```

Produces `bin\onenote-md.exe`. It needs nothing but the .NET Framework that
ships with Windows; the only external reference is
`Microsoft.Office.Interop.OneNote`, which is part of the OneNote/Office
install and is located automatically.

## Requirements

* **OneNote for Windows desktop** (Office 2016/2019/365) — the classic
  `ONENOTE.EXE`, not the Microsoft Store / UWP app. The Store app has no COM
  automation surface at all.
* You must be **signed in** to OneNote, because the tool reads the same
  hierarchy OneNote itself shows. Cloud notebooks come down through your
  existing sync; it does not fetch anything extra.
* Run it as your normal desktop user. An *elevated* shell can fail to attach to
  the already-running OneNote instance.

## Output layout

The hierarchy maps straight onto folders:

```
onenote-export/
  README.md                       index of everything that was written
  <Notebook>/
    <Section>/
      <Page title>.md
      assets/
        p01-img001.png             image extracted from that page
        report.pdf                 attachment, original filename preserved
    <Section group>/
      <Section>/
        <Page title>.md
```

Each page gets YAML front matter:

```markdown
---
title: "Page title"
notebook: "Notebook"
section: "Section"
group: "Group"
onenote_id: "{...}{1}{...}"
---
```

## Options

| Option | Meaning |
| --- | --- |
| `-o, --out <dir>` | Output directory (default `onenote-export`) |
| `--notebook <name>` | Only notebooks whose name contains `<name>` |
| `--section <name>` | Only sections whose name contains `<name>` |
| `--no-images` | Skip image and attachment extraction (faster, much smaller) |
| `--list` | Print notebooks / sections / pages and exit |
| `--overwrite` | Overwrite existing files (default behaviour) |
| `--skip-existing` | Leave existing files untouched, for incremental runs |
| `-n, --dry-run` | Report what would be written, write nothing |
| `-h, --help` | Usage |

Filters are case-insensitive substrings, so `--notebook Reports` also matches
`Team Reports 2026`.

## What gets converted

| OneNote | Markdown |
| --- | --- |
| `quickStyleIndex` → `h1`–`h6` | `#`–`######` |
| `<one:List>` bullet | `-` list item, nested by outline depth |
| `<one:r>` runs with `font-weight:bold` | `**bold**` |
| `<one:r>` runs with `font-style:italic` | `*italic*` |
| `<span style="font-weight:bold">` | `**bold**` |
| `<one:Image>` with inline `<one:Data>` | `![alt](assets/…)` |
| `<one:InsertedFile>` / `<one:MediaFile>` | `[report.pdf (117 KB)](assets/report.pdf)` |
| `tableHTML` | GitHub-flavoured Markdown table |
| `<a href>` | `text (href)` |
| HTML entities in CDATA (`&amp;`, `&nbsp;`, `&#xA;`) | decoded characters |

Section groups become an extra directory level.

Page text arrives as a mix of three different encodings depending on how it was
typed, and all three are handled: plain CDATA, structured `<one:r>` runs, and
literal HTML fragments that OneNote stores directly inside `<one:T>`. The HTML
path is a tolerant tag scanner rather than an XML parse, because those
fragments are routinely unclosed, so no content is ever dropped.

## Notes and limits

* **Images and attachments land in one shared `assets/` folder per section.**
  Images get generated names (`p07-img002.png`) because a page can hold many and
  their originals are not uniquely named. Attachments keep their **original
  filename**, since that is what you will look for; if a section contains two
  attachments with the same name, the second becomes `name (2).ext` rather than
  overwriting the first.
* **Attachment bytes are copied from OneNote's own cache.**
  `one:InsertedFile` and `one:MediaFile` carry a `pathCache` attribute pointing
  at the cached file, which is copied verbatim. The copy is byte-exact: the
  result keeps whatever magic number the original had (`%PDF` for PDFs, `PK` for
  Office and zip files, `SQLite format 3` for `.db`/`.sqlite`).
* **If an attachment has not synced to this machine**, `pathCache` is empty and
  the bytes do not exist locally. The Markdown then shows
  `[report.pdf] (not available locally)` instead of a link, so the gap is
  visible rather than silent.
* **Print artefacts are skipped.** Pages pasted in from a print-out contain
  `isPrintOut="true"` images whose `<one:Data>` element exists but is empty
  (0 bytes). There is no payload to recover, so they are not written and no
  dangling link is produced.
* **OLE embeddings are not exported.** Word/Excel objects embedded *as* OLE
  (rather than attached as files) are not written out.
* **Image-heavy pages are expensive.** Fetching a page with images uses
  `PageInfo.piAll`, which inlines every image as base64. Some pages are tens of
  megabytes, so a notebook with many screenshots runs at a reasonable speed but
  will use noticeable memory. `--no-images` avoids it, and also skips
  attachment extraction.
* **Notebooks are read as OneNote sees them**, so only what has synced to the
  desktop appears. The OneNote window will visibly move around while the export
  runs; that is normal.
* **Existing files are overwritten** by default. Use `--skip-existing` for
  incremental re-runs.
* Page titles are used for filenames, so invalid Windows characters are
  replaced, very long titles are truncated to 120 characters, and duplicate
  titles within one section get a ` (2)` suffix so no page is ever lost.

## How it works

1. `CoCreateInstance` on `OneNote.Application`
   (`{DC67E480-C3CB-49F8-8232-60B0C2056C8E}`), apartment-threaded.
2. `GetHierarchy(hsPages)` once, to get every notebook, section and page id.
3. `GetPageContent(id, PageInfo.piAll)` per page.
4. The page XML is walked and rendered to Markdown. Images are base64-decoded
   from their inline `<one:Data>`; attachments are copied from the local file
   that `pathCache` points at.

The walk follows `Outline/OutlineChildren/OE`, but OneNote also places pictures
and attachments **directly under `<one:Page>`** and in other positions. Rather
than hard-code every possible location, whatever the walk does not emit is
recovered by a second pass over the document, so no object is silently dropped.

Because step 1 and step 2 go through the interop assembly rather than late
binding, the tool does not depend on the OneNote type library being correctly
registered. On machines where that registration is broken, PowerShell's
`New-Object -ComObject OneNote.Application` fails with
`TYPE_E_LIBNOTREGISTERED` while this tool works normally.

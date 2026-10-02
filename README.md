# worddocs2pdf

Convert a folder of Word documents to PDF and merge them into one file, in a single
command.

```sh
worddocs2pdf ~/Documents/reports -o Q3.pdf
```
```
Merge order (4 document(s)) -> /Users/you/Documents/reports/Q3.pdf
   1. 01_intro.docx
   2. 02_methods.docx
   3. 03_results.docx
   4. 10_appendix.doc
  [1/4] 01_intro.docx         -> 01_intro.pdf
  [2/4] 02_methods.docx       -> 02_methods.pdf
  [3/4] 03_results.docx       -> 03_results.pdf
  [4/4] 10_appendix.doc       -> 10_appendix.pdf
  merging 4 file(s), 31 pages -> Q3.pdf
  done: /Users/you/Documents/reports/Q3.pdf
```

The conversion is done by Microsoft Word itself, so the output is exactly what you would
get from **File → Save As → PDF**: fonts, styles, headers, footers, tables, numbering and
page breaks all as Word renders them. No third-party rendering engine, no approximation.

## Requirements

This is a **macOS-only** tool and it **requires Microsoft Word to be installed**. It
drives your local Word via AppleScript; there is no bundled renderer and no fallback.

- macOS (tested on macOS 26 and 27, Apple silicon)
- Microsoft Word for Mac, installed at `/Applications/Microsoft Word.app`
  (tested on 16.113; any reasonably recent version should work)
- `pdfunite`, from poppler — `brew install poppler`

It will not run on Windows or Linux, and there are no plans for that. If you want
Word-free conversion on another platform, LibreOffice's `soffice --convert-to pdf` is the
usual starting point — it does not match Word's rendering, which is the whole point of
this tool.

## Install

```sh
git clone git@github.com:ronsm/worddocs2pdf.git
cd worddocs2pdf
chmod +x worddocs2pdf
ln -s "$PWD/worddocs2pdf" ~/.local/bin/worddocs2pdf   # or anywhere on your $PATH
```

## Usage

```
worddocs2pdf [options] <folder>
worddocs2pdf [options] <file.docx> [file2.docx ...]
```

Documents are merged in **natural filename order**, so `Doc2` comes before `Doc10`. Name
files `01_`, `02_`, … to control the order. Word lock files (`~$...`) and non-Word files
are skipped automatically.

| Option | Meaning |
|---|---|
| `-o FILE` | Output path. Default: `merged.pdf` beside the input documents. |
| `-f` | Overwrite `FILE` if it already exists. |
| `--keep` | Also save the per-document PDFs beside their source files. |
| `--clear-quarantine` | Strip `com.apple.quarantine` from inputs first (see below). |
| `--dry-run` | Print the merge order and exit without touching Word. |
| `-h`, `--help` | Usage. |

Accepts `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.rtf` and `.odt`.

Run `--dry-run` first to confirm the order before converting a large batch.

## First run

macOS will ask *"&lt;your terminal&gt; wants to control Microsoft Word"*. Click **OK**.

If you decline, conversions fail with error `-1743`. Re-enable at **System Settings →
Privacy & Security → Automation → &lt;your terminal&gt; → Microsoft Word**.

## How it works, and why

Word on macOS is sandboxed: it puts up a permission dialog every time it reads or writes a
file outside its own container. Converting twenty documents the naive way means twenty
dialogs.

So this script does all of Word's work *inside* Word's own container
(`~/Library/Containers/com.microsoft.Word/Data/tmp/`). The shell — which is not sandboxed —
copies each document in, Word converts it in place with no prompt, and the shell merges the
PDFs back out to wherever you asked. The working directory is removed afterwards.

Consequences worth knowing:

- Your source documents are never modified, and nothing is written next to them unless you
  pass `--keep`.
- An existing `<name>.pdf` is never overwritten; `--keep` writes `<name>.converted.pdf`
  instead.
- If Word was already running it is left running, with its open documents untouched. If the
  script started Word, it quits it afterwards.
- A document that fails to convert is reported and skipped; the rest are still merged.

One quirk, in case you are reading the source and wondering: Word's `open` command returns
no value despite what its scripting dictionary advertises, so the script picks the opened
document up as `active document` instead.

## Troubleshooting

**A document fails, and it came from the internet or an email attachment.** Downloaded
files carry the `com.apple.quarantine` flag and Word may open them read-only in Protected
View, which blocks PDF export. Re-run with `--clear-quarantine`. That removes a macOS
security marker, so only do it for files you trust — which is why it is opt-in rather than
automatic.

**Word shows an error about a `~$` file.** That is Word's lock file for a document that is
currently open, or a stale one left by a crash. Close the document in Word, or delete the
stale `~$` file.

**A password-protected document.** The script cannot supply a password. Those are reported
as failures and skipped.

**`pdfunite: command not found`.** `brew install poppler`.

## Tests

`test/` holds three documents of 1, 2 and 3 pages, named to exercise natural sort (`01_`,
`02_`, `10_`), plus an orphan `~$` lock file and a `notes.txt` that should both be skipped.

```sh
worddocs2pdf --dry-run test          # expect 3 documents, 10_ sorting last
worddocs2pdf test -o test/out.pdf    # expect "merging 3 file(s), 6 pages"
pdfinfo test/out.pdf                 # expect Pages: 6
```

The fixtures are generated, and can be rebuilt with `python3 test/make-fixtures.py` from
the repository root.

## Licence

MIT — see [LICENSE](LICENSE).

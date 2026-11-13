# Aspose.Note FOSS for Rust

A free, open-source Rust library for reading Microsoft OneNote `.one` files. This is a
Rust port of the [`aspose-note-foss`](../aspose-note-for-python) Python package: it keeps the
**same module structure, the same public API surface, and the same test set** — translated to
idiomatic Rust (`snake_case` methods, `Result`-based errors, an `Rc<RefCell>` DOM).

It parses the MS-ONESTORE / MS-ONE binary container with **no external dependencies** and exposes
the notebook as a DOM of pages, outlines, rich text, images, attachments, tables, tags, and
numbered lists — with no Microsoft OneNote installation or COM automation involved.

> **Status — Stage 1.** This build **reads** OneNote files from the test folder and exposes the
> full read-only DOM and traversal/query API. PDF export and full page-history reconstruction are
> planned for later stages (see [Roadmap](#roadmap)).

## Quick start

```rust
use aspose_note::Document;

let doc = Document::open("testfiles/SimpleTable.one")?;
println!("{:?}", doc.display_name());
for page in doc.child_pages() {
    if let Some(text) = page.title().and_then(|t| t.title_text()) {
        println!("{}", text.text());
    }
}
```

Load from any byte stream:

```rust
use aspose_note::Document;
use std::io::Cursor;

let bytes = std::fs::read("testfiles/SimpleTable.one")?;
let doc = Document::from_reader(&mut Cursor::new(bytes))?;
```

Traverse the DOM by node type:

```rust
use aspose_note::{Document, Image, RichText, Table, TableRow, TableCell};

let doc = Document::open("testfiles/SimpleTable.one")?;
for table in doc.get_child_nodes::<Table>() {
    for row in table.get_child_nodes::<TableRow>() {
        let values: Vec<String> = row
            .get_child_nodes::<TableCell>()
            .iter()
            .map(|cell| {
                cell.get_child_nodes::<RichText>()
                    .iter()
                    .map(|rt| rt.text())
                    .collect::<Vec<_>>()
                    .join(" ")
            })
            .collect();
        println!("{values:?}");
    }
}
let _images = doc.get_child_nodes::<Image>();
```

## Crate layout

The module structure mirrors the Python package 1:1.

| Rust | Python | Role |
|---|---|---|
| [`src/lib.rs`](src/lib.rs) | `aspose/note/__init__.py` | Public facade / re-exports |
| [`src/enums.rs`](src/enums.rs) | `enums.py` | `FileFormat`, `SaveFormat`, `HorizontalAlignment`, `TagStatus`, `NodeType` |
| [`src/error.rs`](src/error.rs) | `exceptions.py` | `AsposeNoteError` (base + 5 variants) |
| [`src/model.rs`](src/model.rs) | `model.py` | DOM nodes, `Document`, visitor, `Save` dispatch |
| [`src/saving/`](src/saving) | `saving/` | `SaveOptions`, `PdfSaveOptions` (PDF writer: later stage) |
| [`src/internal/onestore/parser.rs`](src/internal/onestore/parser.rs) | `_internal/onestore/parser.py` | MS-ONESTORE container parser |
| [`src/internal/ms_one/loader.rs`](src/internal/ms_one/loader.rs) | `_internal/ms_one/loader.py` | Object graph → public DOM |
| [`src/internal/ms_one/entities.rs`](src/internal/ms_one/entities.rs) | `_internal/ms_one/entities/` | Intermediate logical tree |

Everything under `aspose_note::internal` is an implementation detail and may change.

## API naming

The public surface matches the Python/.NET shape with idiomatic Rust naming:

| Python / .NET | Rust |
|---|---|
| `Document("f.one")` | `Document::open("f.one")? ` |
| `doc.DisplayName` | `doc.display_name()` |
| `doc.GetChildNodes(Page)` | `doc.get_child_nodes::<Page>()` |
| `node.ParentNode`, `node.Accept(v)` | `node.parent_node()`, `node.accept(&mut v)` |
| `page.Title`, `page.Level` | `page.title()`, `page.level()` |
| `rich_text.Text`, `.TextRuns`, `.Append(...)` | `rich_text.text()`, `.text_runs()`, `.append(...)` |
| `image.Bytes`, `.FileName`, `.Alignment` | `image.bytes()`, `.file_name()`, `.alignment()` |
| `NoteTag.CreateYellowStar(label)` | `NoteTag::create_yellow_star(Some(label))` |
| `doc.Save(target, SaveFormat.Pdf)` | `doc.save_to_path(target, Some(SaveFormat::Pdf))` |
| exceptions (`FileCorruptedException`, …) | `AsposeNoteError` enum variants |

## Development and testing

The crate has **no runtime dependencies**. Build and test with Cargo:

```bash
cargo build
cargo test
cargo clippy --all-targets
```

The test suite mirrors the Python `tests/` directory (open/DOM/content/tags/rich-text/decoding/
truncation/exception coverage) and runs against the same `.one` fixtures in [`testfiles/`](testfiles).

## Roadmap

- **Stage 1 (current)** — open `.one` files, build the read-only DOM, and expose the
  traversal/query API, tags, and rich-text runs. ✅
- **Stage 2** — full page-history reconstruction (`LoadOptions { load_history: true }`).
- **Stage 3** — exhaustive rich-text / tag fidelity hardening.
- **Stage 4** — PDF export (`Document::save`, `PdfSaveOptions`) and golden-output tests.
- **Stage 5** — runnable examples, published docs, and CI.

## Scope and limitations

- Reading only — writing `.one` back out is not implemented.
- Password-protected documents are not supported (`LoadOptions { document_password }` raises
  `AsposeNoteError::IncorrectPassword`).
- `Document::save` currently reports `AsposeNoteError::UnsupportedSaveFormat`; PDF export lands in
  Stage 4.

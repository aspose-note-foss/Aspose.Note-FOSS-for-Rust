# Aspose.Note for Rust — Implementation Plan & Roadmap

This document is the living implementation plan for porting
[`aspose-note-for-python`](../aspose-note-for-python) (`aspose-note-foss`, v26.9.0) to Rust as the
crate **`aspose-note-foss`** (import name `aspose_note`). It records the **completed Stage 1** and
the **planned future stages**.

> Goal: same module structure, same public API surface, and the same test set as the Python
> package — expressed idiomatically in Rust. Delivered incrementally.

---

## 1. Architecture (locked decisions)

| Decision | Choice |
|---|---|
| API naming | Idiomatic Rust `snake_case`, 1:1 with the Python/.NET surface (`GetChildNodes` → `get_child_nodes::<T>()`, `Page.Title` → `page.title()`, …). |
| DOM representation | Reference-counted tree: `Rc<RefCell<NodeInner>>` with a `NodeData` enum and thin typed handles (`Page`, `RichText`, …). Preserves parent links, node identity (`Node::ptr_eq`), and in-place mutation. |
| Errors | `Result<_, AsposeNoteError>`; the enum plays the role of the Python base exception, with one variant per Python subclass. |
| Dependencies | **Zero** external crates — GUID formatting, UTF-16LE, and Windows-1252 decoding are hand-rolled on `std`. |
| Crate / lib name | package `aspose-note-foss`, `[lib] name = "aspose_note"`. |

### Module map (mirrors Python 1:1)

```
src/lib.rs                         ↔ aspose/note/__init__.py      (public facade / re-exports)
src/enums.rs                       ↔ enums.py
src/error.rs                       ↔ exceptions.py
src/model.rs                       ↔ model.py                     (DOM, Document, visitor, save dispatch)
src/saving/{mod,options}.rs        ↔ saving/{__init__,options}.py
src/saving/pdf_writer.rs           ↔ saving/pdf_writer.py         (STAGE 4 — not yet created)
src/internal/onestore/parser.rs    ↔ _internal/onestore/parser.py
src/internal/ms_one/loader.rs      ↔ _internal/ms_one/loader.py
src/internal/ms_one/entities.rs    ↔ _internal/ms_one/entities/
testfiles/                         ↔ testfiles/                   (12 .one fixtures, copied)
tests/                             ↔ tests/                       (ported integration tests)
```

---

## 2. Stage 1 — Open files & read-only DOM — ✅ COMPLETE

Everything required to open the OneNote files in `testfiles/` and expose the full read-only DOM and
traversal/query API.

### Delivered

- **OneStore parser** (`internal/onestore/parser.rs`): header validation (≥1024 bytes, section
  file-type GUID), root FileNodeList, object-space manifests, revision manifests, object groups,
  global-ID tables, the file-data store, property-set decoding for all value types (`0x2`–`0x11`),
  revision-chain resolution, and "latest root" materialization. Includes the best-effort truncation
  guards (zero-padded GUIDs, clamped property counts, stop-on-short-read).
- **MS-ONE loader** (`internal/ms_one/loader.rs`): section/content-space selection, current-revision
  snapshotting, logical-tree build, and builders for Page / Title / Outline / OutlineElement /
  RichText / Image / AttachedFile / Table / Row / Cell / NumberList / NoteTag. Text decoding
  (UTF-16LE + UTF-8/Windows-1252 fallback), per-run `TextRunFormatting` + `TextRunIndex` boundaries
  (handles both inline sets **and** object-reference arrays), timestamp conversions, page ordering,
  and synthetic-title fallback.
- **DOM & public API** (`model.rs`): `Rc<RefCell>` node tree; typed handles with getters/builders;
  `Document::open` / `from_reader` (+ `*_with_options`); `get_child_nodes::<T>()`; `DocumentVisitor`;
  `Page::deep_clone`; `PageHistory` (current-only); `LoadOptions` (password → `IncorrectPassword`);
  `License` / `Metered` no-ops; `DateTime`; `save_*` returning `UnsupportedSaveFormat` for now.
- **Save options** (`saving/options.rs`): `SaveOptions` base + `PdfSaveOptions`.
- **Fixtures**: all 12 `.one` files copied to `testfiles/`.

### Test coverage — 65 tests passing, `cargo clippy` clean

| Python test module | Rust location | Status |
|---|---|---|
| `test_imports_smoke` | `tests/imports_smoke.rs` | ✅ ported |
| `test_aspose_note_dom_document` | `tests/dom_document.rs` | ✅ ported |
| `test_aspose_note_dom_content` | `tests/dom_content.rs` | ✅ ported |
| `test_aspose_note_exceptions_and_stubs` | `tests/exceptions_stubs.rs` | ✅ ported |
| `test_aspose_note_richtext_formatting` | `tests/richtext_formatting.rs` | ✅ ported |
| `test_aspose_note_tags` | `tests/tags.rs` | ✅ ported |
| `test_aspose_note_parser_truncation` | unit tests in `parser.rs` | ✅ ported |
| `test_aspose_note_text_decoding` | unit tests in `loader.rs` | ✅ ported |
| `test_aspose_note_history` | `tests/history.rs` | ◑ parser-level roots ported; full `LoadHistory` ordering → Stage 2 |
| `test_aspose_note_compat_smoke` | `dom_document.rs` + `imports_smoke.rs` | ◑ non-PDF parts ported; `Save(..Pdf)` → Stage 4 |
| `test_aspose_note_save_options` | — | ◻ Stage 4 (PDF) |
| `test_aspose_note_pdf_goldens` | — | ◻ Stage 4 (PDF) |

Validated against Python reference values, e.g. the SimpleTable grid
`[["1","2","3"],["6","5","4"],["7","8","9"],["b","a","0"]]`, subpage titles/levels
`["Page1","Subpage2"]` / `[1,2]`, 3-image alignment `[Right,Center,Left]`, and FormattedRichText
run styles (`This`→highlight 65535, `text`→bold, `is `→highlight 16776960, `not`→italic,
`hyperlink`→underline).

### Build/test notes

- Cargo/rustc 1.92 live in `%USERPROFILE%\.cargo\bin` (not on `PATH`).
- Linking the test binaries needs the MSVC environment; run `cargo test` after
  `call "...\VC\Auxiliary\Build\vcvars64.bat"` (VS 2022). `cargo build` (rlib) works without it.

---

## 3. Stage 2 — Page history (`LoadOptions.load_history`)

Reconstruct full multi-revision page history, matching the Python `_Builder` history logic.

- Port `_published_page_revision_indices`, `_build_page_histories`, `_page_text_signature`,
  `_page_identity_keys_from_objects`, `_build_page_pairs_for_revision`, and the `_run_in_space`
  snapshot-restore pattern (already expressible via the explicit `objects`/`space`/`revision`
  threading used in the Stage 1 loader).
- Wire `include_page_history` through `Builder::build` so `page_histories` carries ordered prior
  versions instead of the current-only placeholder.
- Populate `Document::from_loaded` so `get_page_history(page)` returns `Current` plus ordered
  versions.
- **Port the deferred `history` tests**: `get_page_history_returns_all_revisions_in_order`
  (`SimpleHistory` → `["", "First text", "Second text"]`, current "Third text"),
  `incremental_history_reconstructs_simple_table_versions`,
  `simple_table_published_history_matches_ui_structure`,
  `incremental_history_reconstructs_image_alignment_versions`.

---

## 4. Stage 3 — Tag & rich-text fidelity hardening

- Exhaustive tag extraction parity (labels, icons, status, colors, created/completed times) across
  shared definitions and `NoteTagStates`, with the `_unique_tags` dedup semantics.
- Rich-text edge cases: alignment, paragraph styles, hyperlink extraction, and run-boundary
  robustness on all fixtures.
- Add targeted cases only where Stage 2/3 behavior is newly introduced (avoid duplicating Stage 1
  coverage).

---

## 5. Stage 4 — PDF export

- Create `saving/pdf_writer.rs` (`write_pdf`) and make `Document::save*` produce PDF for
  `SaveFormat::Pdf` instead of returning `UnsupportedSaveFormat`.
- Port `PdfSaveOptions` behavior (image compression, JPEG quality, page settings/range).
- **Golden tests**: the Python goldens assert an exact ReportLab operator sequence
  (`tests/goldens/pdf/*.manifest.json` — text + `layout_ops` + links digest). Rust has no ReportLab,
  so either (a) reproduce the operator stream with a chosen PDF backend, or (b) regenerate goldens
  against the Rust writer while preserving the manifest schema. **Open decision** — recommend (b).
- Port `test_aspose_note_save_options` and `test_aspose_note_pdf_goldens`.

---

## 6. Stage 5 — Examples, CI, docs

- `examples/`: ports of the Python examples (`extract_text`, `save_images`, `export_pdf` once Stage 4
  lands) as Cargo examples.
- CI: GitHub Actions running `cargo build` / `cargo test` / `cargo clippy` / `cargo fmt --check`
  (with the MSVC toolchain on Windows runners; Linux runner for the dependency-free path).
- Docs: finalize `README.md`, crate-level rustdoc, and publishing metadata.

---

## 7. Risks / open questions

- **PDF goldens** are tightly coupled to ReportLab output (Stage 4) — likely require regenerating
  goldens for the Rust backend.
- **DOM identity**: Python `is`/`==` semantics map to `Node::ptr_eq` / handle `PartialEq` in Rust;
  keep this in mind when porting identity-based history assertions.
- **Snapshot cost**: object-map snapshots are cloned and cached; fine for the small fixtures, revisit
  only if large notebooks become a target.

//! Smoke test that the public API surface is importable and constructible,
//! mirroring `tests/test_imports_smoke.py`.

use aspose_note::saving::{PdfSaveOptions, SaveOptions};
use aspose_note::{
    AttachedFile, Document, DocumentVisitor, FileFormat, HorizontalAlignment, Image, License,
    LoadOptions, Metered, Node, NoteTag, NumberList, Outline, OutlineElement, Page, PageHistory,
    ParagraphStyle, RichText, SaveFormat, TagStatus, Table, TableCell, TableColumn, TableRow,
    TextRun, TextStyle, Title, TypedNode,
};

#[test]
fn public_types_are_constructible() {
    let _ = Document::new();
    let _ = Page::new();
    let _ = Outline::new();
    let _ = OutlineElement::new();
    let _ = RichText::new();
    let _ = Title::new();
    let _ = Table::new();
    let _ = TableRow::new();
    let _ = TableCell::new();
    let _ = Image::new();
    let _ = AttachedFile::new();
    let _ = ParagraphStyle::default();
    let _ = TextStyle::default();
    let _ = TextRun::new("x");
    let _ = TableColumn::new(None);
    let _ = NumberList::default();
    let _ = NoteTag::create_yellow_star(None);
    let _ = LoadOptions::new();
    let _ = License::new();
    let _ = Metered::new();
    let _ = PdfSaveOptions::new();
}

#[test]
fn enums_are_available() {
    let _ = FileFormat::OneNote2010;
    let _ = SaveFormat::Pdf;
    let _ = HorizontalAlignment::Center;
    let _ = TagStatus::Open;
    assert_eq!(PdfSaveOptions::new().save_format(), SaveFormat::Pdf);
}

// Keep traits/types referenced so the imports are exercised.
fn _assert_traits<V: DocumentVisitor, T: TypedNode>() {}

#[test]
fn trait_bounds_resolve() {
    fn use_node(n: &Node) -> Option<Page> {
        Page::from_node(n)
    }
    let doc = Document::new();
    let _ = use_node(&doc.as_node());
    let _ = PageHistory::new(Page::new());
    let _: fn() = _assert_traits::<NoopVisitor, Page>;
}

#[derive(Default)]
struct NoopVisitor;
impl DocumentVisitor for NoopVisitor {}

// `SaveOptions` is only reachable through the `saving` module, never the crate
// root — mirror that by referencing it here via the module path.
fn _save_options_lives_in_saving(options: &SaveOptions) -> SaveFormat {
    options.save_format()
}

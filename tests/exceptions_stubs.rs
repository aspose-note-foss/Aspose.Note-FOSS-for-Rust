use std::path::PathBuf;

use aspose_note::{
    AsposeNoteError, DateTime, Document, FileFormat, License, LoadOptions, Metered, Node, NoteTag,
    Page, PageHistory, SaveFormat, TagStatus,
};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

#[test]
fn license_and_metered_are_callable() {
    License::new().set_license("dummy.lic");
    Metered::new().set_metered_key("pub", "priv");
}

#[test]
fn encrypted_load_raises_incorrect_password() {
    let options = LoadOptions {
        document_password: Some("pass".to_string()),
        load_history: false,
    };
    let result = Document::open_with_options(fixture("FormattedRichText.one"), &options);
    assert!(matches!(result, Err(AsposeNoteError::IncorrectPassword(_))));
}

#[test]
fn parent_node_reports_container() {
    let doc = Document::new();
    let page = Page::new();
    doc.append_child_last(page.as_node());
    assert!(Node::ptr_eq(&page.parent_node().unwrap(), &doc.as_node()));
    assert_eq!(doc.child_nodes().len(), 1);
}

#[test]
fn note_tag_read_only_members_match_shape() {
    let tag = NoteTag::create_question_mark(Some("Question"));
    assert_eq!(tag.status(), TagStatus::Open);
    assert!(tag.completed_time().is_none());
}

#[test]
fn note_tag_creation_time_uses_datetime_values() {
    let created = DateTime::from_ymd_hms_micro(2024, 5, 6, 7, 8, 0, 0);
    let mut tag = NoteTag::create_yellow_star(Some("Created"));
    tag.set_creation_time(Some(created));
    assert_eq!(tag.creation_time(), Some(created));
    assert_eq!(tag.status(), TagStatus::Open);

    let updated = DateTime::from_ymd_hms_micro(2024, 5, 7, 8, 9, 0, 0);
    tag.set_creation_time(Some(updated));
    assert_eq!(tag.creation_time(), Some(updated));
}

#[test]
fn unsupported_file_format_exposes_file_format() {
    let error = AsposeNoteError::unsupported_file_format("bad", Some("bogus-guid"));
    assert_eq!(error.file_format(), Some(FileFormat::Unknown));

    let known = AsposeNoteError::unsupported_file_format(
        "bad",
        Some("7b5c52e4-d88c-4da7-aeb1-5378d02996d3"),
    );
    assert_eq!(known.file_format(), Some(FileFormat::OneNote2010));
}

#[test]
fn page_history_tracks_current_page() {
    let current = Page::new();
    let history = PageHistory::new(current.clone());
    assert!(Node::ptr_eq(&history.current().as_node(), &current.as_node()));
    assert_eq!(history.count(), 0);
}

#[test]
fn save_non_pdf_path_raises() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let out = std::env::temp_dir().join("aspose_note_rust_out.txt");
    let result = doc.save_to_path(&out, None);
    assert!(matches!(result, Err(AsposeNoteError::UnsupportedSaveFormat(_))));
}

#[test]
fn save_pdf_is_not_yet_implemented() {
    // Stage 1 opens files only; PDF export arrives in a later stage.
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let out = std::env::temp_dir().join("aspose_note_rust_out.pdf");
    let result = doc.save_to_path(&out, Some(SaveFormat::Pdf));
    assert!(matches!(result, Err(AsposeNoteError::UnsupportedSaveFormat(_))));
}

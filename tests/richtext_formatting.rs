use std::collections::HashMap;
use std::path::PathBuf;

use aspose_note::{Document, HorizontalAlignment, NoteTag, ParagraphStyle, RichText, TextRun, TextStyle};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

#[test]
fn richtext_runs_preserved() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let rich_texts = doc.get_child_nodes::<RichText>();
    assert!(!rich_texts.is_empty());
    assert!(rich_texts.iter().any(|rt| !rt.text_runs().is_empty()));
}

#[test]
fn richtext_has_some_non_default_style() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let rich_texts = doc.get_child_nodes::<RichText>();
    let styled = rich_texts.iter().any(|rt| {
        rt.text_runs().iter().any(|run| {
            let s = &run.style;
            s.is_bold
                || s.is_italic
                || s.is_underline
                || s.is_strikethrough
                || s.font_name.is_some()
                || s.font_size.is_some()
                || s.font_color.is_some()
                || s.highlight.is_some()
                || s.hyperlink_address.is_some()
        })
    });
    assert!(styled);
}

#[test]
fn textstyle_default_and_font_style() {
    let mut style = TextStyle::default_style();
    assert_eq!(style.language, Some(1033));
    style.is_bold = true;
    style.is_italic = true;
    style.is_underline = true;
    style.is_strikethrough = true;
    style.highlight = Some(123);
    style.language = Some(1031);
    assert_eq!(style.highlight, Some(123));
    assert_eq!(style.language, Some(1031));
    assert_eq!(style.font_style(), 15);
}

#[test]
fn richtext_stores_alignment() {
    let rich_text = RichText::build(
        "Aligned".to_string(),
        None,
        ParagraphStyle::default(),
        Some(HorizontalAlignment::Center),
        Vec::new(),
    );
    assert_eq!(rich_text.alignment(), Some(HorizontalAlignment::Center));
}

#[test]
fn paragraphstyle_default_text_properties() {
    let mut style = ParagraphStyle::default_style();
    style.font_name = Some("Courier New".to_string());
    style.font_size = Some(20.0);
    style.font_color = Some(0x112233);
    style.highlight = Some(0xAABBCC);
    style.is_bold = true;
    style.is_italic = true;
    style.is_underline = true;
    style.is_strikethrough = true;
    assert_eq!(style.font_name.as_deref(), Some("Courier New"));
    assert_eq!(style.font_size, Some(20.0));
    assert_eq!(style.font_style(), 15);
}

#[test]
fn richtext_uses_paragraphstyle_as_default_text_style() {
    let paragraph_style = ParagraphStyle {
        font_name: Some("Courier New".to_string()),
        font_size: Some(20.0),
        is_bold: true,
        highlight: Some(0xFFFF00),
        ..ParagraphStyle::default()
    };
    let rich_text = RichText::build(
        "Paragraph defaults".to_string(),
        None,
        paragraph_style,
        None,
        Vec::new(),
    );
    let runs = rich_text.text_runs();
    assert_eq!(runs.len(), 1);
    let style = &runs[0].style;
    assert_eq!(style.font_name.as_deref(), Some("Courier New"));
    assert_eq!(style.font_size, Some(20.0));
    assert!(style.is_bold);
    assert_eq!(style.highlight, Some(0xFFFF00));
}

#[test]
fn richtext_append_uses_paragraphstyle_defaults() {
    let rich_text = RichText::build(
        String::new(),
        None,
        ParagraphStyle {
            font_name: Some("Courier New".to_string()),
            font_size: Some(18.0),
            is_italic: true,
            ..ParagraphStyle::default()
        },
        None,
        Vec::new(),
    );
    rich_text.append("First", None);
    rich_text.append("Second", None);
    let runs = rich_text.text_runs();
    assert_eq!(
        runs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>(),
        ["First", "Second"]
    );
    for run in &runs {
        assert_eq!(run.style.font_name.as_deref(), Some("Courier New"));
        assert_eq!(run.style.font_size, Some(18.0));
        assert!(run.style.is_italic);
    }
}

#[test]
fn richtext_collection_properties_are_mutable_in_place() {
    let rich_text = RichText::build(
        String::new(),
        Some(vec![TextRun::new("segment")]),
        ParagraphStyle::default(),
        None,
        vec![NoteTag::create_yellow_star(Some("Важно"))],
    );
    rich_text.append(" tail", None);
    rich_text
        .tags_mut()
        .push(NoteTag::create_question_mark(Some("Вопрос")));
    assert_eq!(rich_text.text(), "segment tail");
    assert_eq!(
        rich_text.tags().iter().map(|t| t.label()).collect::<Vec<_>>(),
        vec![Some("Важно".to_string()), Some("Вопрос".to_string())]
    );
}

#[test]
fn formatted_richtext_run_boundaries_align_with_visible_text() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let target = doc
        .get_child_nodes::<RichText>()
        .into_iter()
        .find(|rt| rt.text().contains("hyperlink. This text is not a hyperlink."))
        .unwrap();

    let visible: Vec<(String, TextStyle)> = target
        .text_runs()
        .into_iter()
        .filter(|run| !run.text.is_empty() && !run.text.contains("HYPERLINK"))
        .map(|run| (run.text.clone(), run.style.clone()))
        .collect();
    let texts: Vec<String> = visible.iter().map(|(t, _)| t.clone()).collect();

    for expected in ["This", "text", "is ", "not", "hyperlink"] {
        assert!(texts.iter().any(|t| t == expected), "missing run {expected:?}");
    }

    let style_by_text: HashMap<String, TextStyle> = visible.into_iter().collect();
    assert_eq!(style_by_text["This"].highlight, Some(65535));
    assert!(style_by_text["text"].is_bold);
    assert_eq!(style_by_text["is "].highlight, Some(16776960));
    assert!(style_by_text["not"].is_italic);
    assert!(style_by_text["hyperlink"].is_underline);
}

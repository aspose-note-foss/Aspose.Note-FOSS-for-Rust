use std::collections::HashSet;
use std::path::PathBuf;

use aspose_note::{
    AttachedFile, Document, Image, Outline, OutlineElement, RichText, Table, TypedNode,
};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

fn tag_is_meaningful(label: &Option<String>, icon: Option<i64>) -> bool {
    label.as_deref().is_some_and(|l| !l.is_empty()) || icon.is_some()
}

fn rt_text_and_labels(rich_texts: &[RichText]) -> (String, HashSet<String>) {
    match rich_texts.first() {
        None => (String::new(), HashSet::new()),
        Some(rt) => {
            let labels: HashSet<String> = rt.tags().iter().filter_map(|t| t.label()).collect();
            (rt.text(), labels)
        }
    }
}

fn collect_all_tags(doc: &Document) -> Vec<(Option<String>, Option<i64>)> {
    let mut tags = Vec::new();
    for rt in doc.get_child_nodes::<RichText>() {
        tags.extend(rt.tags().iter().map(|t| (t.label(), t.icon())));
    }
    for img in doc.get_child_nodes::<Image>() {
        tags.extend(img.tags().iter().map(|t| (t.label(), t.icon())));
    }
    for table in doc.get_child_nodes::<Table>() {
        tags.extend(table.tags().iter().map(|t| (t.label(), t.icon())));
    }
    for af in doc.get_child_nodes::<AttachedFile>() {
        tags.extend(af.tags().iter().map(|t| (t.label(), t.icon())));
    }
    tags
}

#[test]
fn image_with_tag_exposes_tags() {
    let doc = Document::open(fixture("ImageWithTag.one")).unwrap();
    let images = doc.get_child_nodes::<Image>();
    assert!(!images.is_empty());
    let tagged: Vec<_> = images.iter().filter(|img| !img.tags().is_empty()).collect();
    assert!(!tagged.is_empty());
    assert!(tagged
        .iter()
        .any(|img| img.tags().iter().any(|t| tag_is_meaningful(&t.label(), t.icon()))));
}

#[test]
fn table_with_tag_exposes_tags() {
    let doc = Document::open(fixture("TableWithTag.one")).unwrap();
    let tables = doc.get_child_nodes::<Table>();
    assert!(!tables.is_empty());
    let tagged: Vec<_> = tables.iter().filter(|t| !t.tags().is_empty()).collect();
    assert!(!tagged.is_empty());
    assert!(tagged
        .iter()
        .any(|t| t.tags().iter().any(|tag| tag_is_meaningful(&tag.label(), tag.icon()))));
}

#[test]
fn richtext_tags_exposed() {
    let doc = Document::open(fixture("TagSizes.one")).unwrap();
    let rich_texts = doc.get_child_nodes::<RichText>();
    assert!(!rich_texts.is_empty());
    assert!(rich_texts.iter().any(|rt| !rt.tags().is_empty()));

    let outline_elements = doc.get_child_nodes::<OutlineElement>();
    let oe_texts: Vec<String> = outline_elements
        .iter()
        .map(|oe| rt_text_and_labels(&oe.get_child_nodes::<RichText>()).0)
        .filter(|t| !t.is_empty())
        .collect();
    assert!(oe_texts.len() >= 4);
    assert!(oe_texts.iter().any(|t| t == "66(6-9)"));
    assert!(oe_texts.iter().any(|t| t == "10(10-17)"));
    assert!(oe_texts.iter().any(|t| t == "18(18-23)"));
    assert!(oe_texts.iter().any(|t| t == "24(24-…)" || t == "24(242-…)"));

    for expected in ["66(6-9)", "10(10-17)", "18(18-23)"] {
        let oe = outline_elements
            .iter()
            .find(|oe| rt_text_and_labels(&oe.get_child_nodes::<RichText>()).0 == expected)
            .unwrap();
        let (_, labels) = rt_text_and_labels(&oe.get_child_nodes::<RichText>());
        assert!(labels.contains("Важно"), "{expected}");
    }

    let last = outline_elements
        .iter()
        .find(|oe| {
            let text = rt_text_and_labels(&oe.get_child_nodes::<RichText>()).0;
            text == "24(24-…)" || text == "24(242-…)"
        })
        .unwrap();
    assert!(rt_text_and_labels(&last.get_child_nodes::<RichText>())
        .1
        .contains("Важно"));
}

#[test]
fn outline_element_tags_and_list_metadata() {
    let doc = Document::open(fixture("NumberedListWithTags.one")).unwrap();

    let outlines = doc.get_child_nodes::<Outline>();
    assert_eq!(outlines.len(), 2);

    let tag_sizes_outline = outlines
        .iter()
        .find(|o| o.get_child_nodes::<RichText>().iter().any(|rt| rt.text() == "66(6-9)"))
        .expect("tag-sizes outline");
    let tag_sizes_oes = tag_sizes_outline.get_child_nodes::<OutlineElement>();
    let tag_sizes_map: std::collections::HashMap<String, HashSet<String>> = tag_sizes_oes
        .iter()
        .map(|oe| rt_text_and_labels(&oe.get_child_nodes::<RichText>()))
        .filter(|(text, _)| !text.is_empty())
        .collect();
    for expected in ["66(6-9)", "10(10-17)", "18(18-23)"] {
        assert!(tag_sizes_map[expected].contains("Важно"), "{expected}");
    }

    let list_outline = outlines
        .iter()
        .find(|o| o.get_child_nodes::<RichText>().iter().any(|rt| rt.text() == "First"))
        .expect("numbered list outline");

    let top_level: Vec<OutlineElement> = list_outline
        .child_nodes()
        .iter()
        .filter_map(OutlineElement::from_node)
        .collect();
    assert_eq!(top_level.len(), 2);

    let (text0, labels0) = rt_text_and_labels(&top_level[0].get_child_nodes::<RichText>());
    let (text1, labels1) = rt_text_and_labels(&top_level[1].get_child_nodes::<RichText>());
    assert_eq!(text0, "First");
    assert_eq!(text1, "Second");
    assert!(labels0.contains("Важно"));
    assert!(labels0.contains("Вопрос"));
    assert!(labels1.contains("Запланировать собрание"));

    let nested: Vec<OutlineElement> = top_level[0]
        .child_nodes()
        .iter()
        .filter_map(OutlineElement::from_node)
        .collect();
    assert!(nested.len() >= 3);
    let nested_map: std::collections::HashMap<String, HashSet<String>> = nested
        .iter()
        .map(|oe| rt_text_and_labels(&oe.get_child_nodes::<RichText>()))
        .collect();
    assert!(nested_map["First-first"].contains("Важно"));
    assert!(nested_map["First-first"].contains("Вопрос"));
    assert!(nested_map["First-second"].contains("Важно"));
    assert!(nested_map["First-second"].contains("Вопрос"));
    assert!(nested_map["First-third"].contains("Контакт"));
    assert!(nested_map["First-third"].contains("Послушать музыку"));
    assert!(nested_map["First-third"].contains("Запланировать собрание"));

    assert!(top_level.iter().all(|oe| oe.number_list().is_some()));

    let all_oes = list_outline.get_child_nodes::<OutlineElement>();
    let formats: HashSet<String> = all_oes
        .iter()
        .filter_map(|oe| oe.number_list().and_then(|nl| nl.format))
        .collect();
    assert!(formats.len() >= 2);

    let tags = collect_all_tags(&doc);
    assert!(!tags.is_empty());
    assert!(tags.iter().any(|(label, icon)| tag_is_meaningful(label, *icon)));
}

#[test]
fn attachment_with_tag_fixture_has_tags() {
    let doc = Document::open(fixture("AttachedFileWithTag.one")).unwrap();
    let attachments = doc.get_child_nodes::<AttachedFile>();
    assert!(!attachments.is_empty());
    let tags = collect_all_tags(&doc);
    assert!(!tags.is_empty());
    assert!(tags.iter().any(|(label, icon)| tag_is_meaningful(label, *icon)));
}

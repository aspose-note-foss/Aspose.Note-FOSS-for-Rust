use std::io::Cursor;
use std::path::PathBuf;

use aspose_note::{
    Document, DocumentVisitor, FileFormat, Outline, OutlineElement, Page, RichText, Title,
};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

#[test]
fn construct_from_path() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    assert!(!doc.child_nodes().is_empty());
    assert!(doc.first_child().is_some());
    assert!(doc.last_child().is_some());
}

#[test]
fn construct_from_stream() {
    let data = std::fs::read(fixture("FormattedRichText.one")).unwrap();
    let doc = Document::from_reader(&mut Cursor::new(data)).unwrap();
    assert!(!doc.child_nodes().is_empty());
}

#[test]
fn document_file_format_enum() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    assert_eq!(doc.file_format(), FileFormat::OneNote2010);
}

#[test]
fn get_child_nodes_page_and_title() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let pages = doc.get_child_nodes::<Page>();
    assert!(!pages.is_empty());

    let titles = doc.get_child_nodes::<Title>();
    assert!(titles.len() >= pages.len());

    let page0 = &pages[0];
    let title = page0.title();
    assert!(title.is_some());
    assert!(aspose_note::Node::ptr_eq(
        &page0.first_child().unwrap(),
        &title.unwrap().as_node()
    ));
}

#[test]
fn richtext_replace_changes_text() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let rich_texts = doc.get_child_nodes::<RichText>();
    assert!(!rich_texts.is_empty());

    let target = rich_texts
        .iter()
        .find(|rt| !rt.text().is_empty() && rt.text().contains(' '))
        .expect("a RichText with a replaceable space");

    let before = target.text();
    let returned = target.replace(" ", "  ", None);
    let after = target.text();
    assert_eq!(&returned, target);
    assert_ne!(before, after);
}

#[test]
fn richtext_text_runs_and_length() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let rich_text = doc
        .get_child_nodes::<RichText>()
        .into_iter()
        .find(|rt| !rt.text_runs().is_empty())
        .unwrap();
    assert!(!rich_text.text_runs().is_empty());
    assert_eq!(rich_text.length(), rich_text.text().chars().count());
}

#[test]
fn page_clone_duplicates_content_tree() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let page = doc.get_child_nodes::<Page>().remove(0);
    let cloned = page.deep_clone(false);

    assert_ne!(cloned, page);
    assert_eq!(cloned.child_nodes().len(), page.child_nodes().len());
    assert_eq!(
        cloned.get_child_nodes::<RichText>().len(),
        page.get_child_nodes::<RichText>().len()
    );
    assert!(cloned.parent_node().is_none());
}

#[test]
fn page_clone_accepts_clone_history_flag() {
    let doc = Document::open(fixture("FormattedRichText.one")).unwrap();
    let page = doc.get_child_nodes::<Page>().remove(0);
    let cloned = page.deep_clone(true);
    assert_ne!(cloned, page);
}

#[derive(Default)]
struct CountingVisitor {
    pages: u32,
    doc_start: u32,
}

impl DocumentVisitor for CountingVisitor {
    fn visit_document_start(&mut self, _document: &Document) {
        self.doc_start += 1;
    }

    fn visit_page_start(&mut self, _page: &Page) {
        self.pages += 1;
    }
}

#[test]
fn accept_visits_document_and_pages() {
    let doc = Document::open(fixture("SimpleTable.one")).unwrap();
    let mut visitor = CountingVisitor::default();
    doc.accept(&mut visitor);
    assert_eq!(visitor.doc_start, 1);
    assert!(visitor.pages >= 1);
}

#[test]
fn document_keeps_subpages_from_multiple_page_spaces() {
    let doc = Document::open(fixture("Issue1Subpages.one")).unwrap();
    let pages = doc.get_child_nodes::<Page>();

    assert_eq!(doc.child_nodes().len(), 2);
    assert_eq!(pages.len(), 2);
    let title_texts: Vec<Option<String>> = pages
        .iter()
        .map(|page| page.title().and_then(|t| t.title_text()).map(|rt| rt.text()))
        .collect();
    assert_eq!(
        title_texts,
        vec![Some("Page1".to_string()), Some("Subpage2".to_string())]
    );
    assert_eq!(
        pages.iter().map(|p| p.level()).collect::<Vec<_>>(),
        vec![Some(1), Some(2)]
    );

    let mut page_texts: Vec<Vec<String>> = Vec::new();
    for page in &pages {
        let mut texts: Vec<String> = Vec::new();
        for outline in page.get_child_nodes::<Outline>() {
            for oe in outline.get_child_nodes::<OutlineElement>() {
                for rt in oe.get_child_nodes::<RichText>() {
                    if !rt.text().is_empty() {
                        texts.push(rt.text());
                    }
                }
            }
        }
        page_texts.push(texts);
    }
    assert_eq!(
        page_texts,
        vec![vec!["Content1".to_string()], vec!["Content2".to_string()]]
    );
}

#[test]
fn single_page_fixtures_do_not_create_synthetic_titles() {
    for name in ["SimpleTable.one", "3ImagesWithDifferentAlignment.one"] {
        let doc = Document::open(fixture(name)).unwrap();
        let pages = doc.get_child_nodes::<Page>();
        assert_eq!(pages.len(), 1, "{name}");
        assert!(pages[0].title().is_none(), "{name}");
        assert_eq!(pages[0].level(), Some(1), "{name}");
    }
}

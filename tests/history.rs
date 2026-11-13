use std::collections::HashSet;
use std::path::PathBuf;

use aspose_note::internal::onestore::parser::{parse_onestore_file, OneStoreSource};
use aspose_note::{Document, Page};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

#[test]
fn parser_keeps_default_and_history_roots_separate() {
    let parsed =
        parse_onestore_file(OneStoreSource::Path(fixture("SimpleHistory.one").as_path())).unwrap();
    let content_space = &parsed.object_spaces[1];

    let default_root = content_space.get_latest_root(1, None).expect("default root");
    assert_eq!(content_space.objects[&default_root].jcid_name, "PageManifestNode");

    let context_root_kinds: HashSet<String> = content_space
        .revisions
        .iter()
        .filter_map(|revision| {
            revision.context_id.as_ref()?;
            let root = revision.root_roles.get(&1)?;
            revision.objects.get(root).map(|obj| obj.jcid_name.clone())
        })
        .collect();
    assert!(context_root_kinds.contains("VersionHistoryContent"));
}

#[test]
fn parser_materializes_incremental_current_roots() {
    for name in ["SimpleTable.one", "TagSizes.one", "3ImagesWithDifferentAlignment.one"] {
        let parsed = parse_onestore_file(OneStoreSource::Path(fixture(name).as_path())).unwrap();
        let content_space = parsed
            .object_spaces
            .iter()
            .max_by_key(|space| space.revisions.len())
            .unwrap();

        let default_root = content_space.get_latest_root(1, None).expect("default root");
        assert!(content_space.objects.contains_key(&default_root), "{name}");
        assert_eq!(
            content_space.objects[&default_root].jcid_name, "PageManifestNode",
            "{name}"
        );
    }
}

#[test]
fn get_page_history_returns_current_only_in_stage_one() {
    let doc = Document::open(fixture("SimpleHistory.one")).unwrap();
    let page = doc.get_child_nodes::<Page>().remove(0);
    let history = doc.get_page_history(&page);
    assert_eq!(history.current(), page);
    assert_eq!(history.count(), 0);
}

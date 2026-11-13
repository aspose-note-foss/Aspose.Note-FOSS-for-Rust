use std::path::PathBuf;

use aspose_note::{
    AttachedFile, Document, HorizontalAlignment, Image, NoteTag, RichText, Table, TableCell,
    TableColumn, TableRow,
};

fn fixture(name: &str) -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("testfiles");
    path.push(name);
    path
}

#[test]
fn images_exposed_and_have_bytes() {
    let doc = Document::open(fixture("3ImagesWithDifferentAlignment.one")).unwrap();
    let images = doc.get_child_nodes::<Image>();
    assert_eq!(images.len(), 3);
    assert!(images.iter().all(|img| img.bytes().len() > 1024));
}

#[test]
fn images_are_identical_by_bytes() {
    let doc = Document::open(fixture("3ImagesWithDifferentAlignment.one")).unwrap();
    let images = doc.get_child_nodes::<Image>();
    assert_eq!(images.len(), 3);
    let first = images[0].bytes();
    assert!(images.iter().all(|img| img.bytes() == first));
}

#[test]
fn images_preserve_outline_alignment() {
    let doc = Document::open(fixture("3ImagesWithDifferentAlignment.one")).unwrap();
    let images = doc.get_child_nodes::<Image>();
    assert_eq!(
        images.iter().map(|img| img.alignment()).collect::<Vec<_>>(),
        vec![
            Some(HorizontalAlignment::Right),
            Some(HorizontalAlignment::Center),
            Some(HorizontalAlignment::Left),
        ]
    );
}

#[test]
fn images_expose_original_dimensions() {
    let doc = Document::open(fixture("3ImagesWithDifferentAlignment.one")).unwrap();
    let image = doc.get_child_nodes::<Image>().remove(0);
    assert_eq!(image.original_width(), image.width());
    assert_eq!(image.original_height(), image.height());
}

#[test]
fn image_constructed_with_alignment() {
    let image = Image::new()
        .with_bytes(b"img".to_vec())
        .with_width(Some(10.0))
        .with_height(Some(20.0))
        .with_alignment(Some(HorizontalAlignment::Center));
    assert_eq!(image.alignment(), Some(HorizontalAlignment::Center));
}

#[test]
fn image_tags_collection_is_mutable() {
    let image = Image::new()
        .with_file_name(Some("image.png".to_string()))
        .with_bytes(b"img".to_vec())
        .with_tags(vec![NoteTag::create_yellow_star(Some("Важно"))]);
    image.tags_mut().push(NoteTag::create_question_mark(Some("Вопрос")));
    assert_eq!(
        image.tags().iter().map(|t| t.label()).collect::<Vec<_>>(),
        vec![Some("Важно".to_string()), Some("Вопрос".to_string())]
    );
}

#[test]
fn tables_exposed_with_rows_and_cells() {
    let doc = Document::open(fixture("SimpleTable.one")).unwrap();
    let tables = doc.get_child_nodes::<Table>();
    assert!(!tables.is_empty());

    let rows = doc.get_child_nodes::<TableRow>();
    let cells = doc.get_child_nodes::<TableCell>();
    assert!(!rows.is_empty());
    assert!(!cells.is_empty());
    for row in rows.iter().take(10) {
        assert!(!row.child_nodes().is_empty());
    }
}

#[test]
fn simple_table_grid_values() {
    let doc = Document::open(fixture("SimpleTable.one")).unwrap();
    let page = doc.get_child_nodes::<Page>().remove(0);
    let table = page.get_child_nodes::<Table>().remove(0);
    let mut grid: Vec<Vec<String>> = Vec::new();
    for row in table.get_child_nodes::<TableRow>() {
        let mut values = Vec::new();
        for cell in row.get_child_nodes::<TableCell>() {
            let texts: Vec<String> = cell
                .get_child_nodes::<RichText>()
                .into_iter()
                .map(|rt| rt.text())
                .filter(|t| !t.is_empty())
                .collect();
            values.push(texts.last().cloned().unwrap_or_default());
        }
        grid.push(values);
    }
    assert_eq!(
        grid,
        vec![
            vec!["1", "2", "3"],
            vec!["6", "5", "4"],
            vec!["7", "8", "9"],
            vec!["b", "a", "0"],
        ]
    );
}

use aspose_note::Page;

#[test]
fn table_collections_are_mutable() {
    let table = Table::new()
        .with_tags(vec![NoteTag::create_yellow_star(Some("Важно"))])
        .with_columns(vec![TableColumn::new(Some(70.0))]);
    table.tags_mut().push(NoteTag::create_question_mark(Some("Вопрос")));
    table.columns_mut().push(TableColumn::new(Some(120.0)));

    assert_eq!(
        table.tags().iter().map(|t| t.label()).collect::<Vec<_>>(),
        vec![Some("Важно".to_string()), Some("Вопрос".to_string())]
    );
    assert_eq!(
        table.columns().iter().map(|c| c.width).collect::<Vec<_>>(),
        vec![Some(70.0), Some(120.0)]
    );
}

#[test]
fn attached_files_exposed_and_have_bytes() {
    let doc = Document::open(fixture("OnePageWithFile.one")).unwrap();
    let attachments = doc.get_child_nodes::<AttachedFile>();
    assert!(!attachments.is_empty());
    assert!(attachments
        .iter()
        .any(|a| !a.file_name().unwrap_or_default().trim().is_empty()));
}

#[test]
fn attached_file_tags_collection_is_mutable() {
    let attached = AttachedFile::new()
        .with_file_name(Some("doc.bin".to_string()))
        .with_bytes(b"abc".to_vec())
        .with_tags(vec![NoteTag::create_yellow_star(Some("Важно"))]);
    attached
        .tags_mut()
        .push(NoteTag::create_question_mark(Some("Вопрос")));
    assert_eq!(
        attached.tags().iter().map(|t| t.label()).collect::<Vec<_>>(),
        vec![Some("Важно".to_string()), Some("Вопрос".to_string())]
    );
}

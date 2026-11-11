//! MS-ONE object graph → public DOM, mirroring
//! `aspose.note._internal.ms_one.loader`.

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use crate::enums::HorizontalAlignment;
use crate::error::AsposeNoteError;
use crate::internal::onestore::parser::{
    decode_utf16le, parse_onestore_file, ExtendedGuid, OneStoreFile, OneStoreObject, OneStoreSource,
    PropMap, PropValue,
};
use crate::model::{
    coerce_note_tag_datetime, filetime_to_datetime, push_child, time32_to_datetime, AttachedFile,
    DateTime, Image, NoteTag, Node, NumberList, Outline, OutlineElement, Page, ParagraphStyle,
    RichText, Table, TableCell, TableColumn, TableRow, TextRun, TextStyle, Title, TypedNode,
};

use super::entities::{LogicalDocument, LogicalNode};

const ROOT_ROLE_DEFAULT_CONTENT: u32 = 0x0000_0001;

type ObjMap = HashMap<ExtendedGuid, OneStoreObject>;
type PageInfo = (ExtendedGuid, LogicalNode, Page);
type RevisionIndexMap = HashMap<ExtendedGuid, usize>;
type SnapshotCache = HashMap<(usize, i64), Rc<ObjMap>>;
type SectionOrders = HashMap<Vec<u8>, usize>;

const STRUCTURE_PROPERTIES: &[&str] = &[
    "ElementChildNodes",
    "ContentChildNodes",
    "StructureElementChildNodes",
    "ChildGraphSpaceElementNodes",
    "ListNodes",
    "TextRunDataObject",
];
const METADATA_PROPERTIES: &[&str] = &["MetaDataObjectsAboveGraphSpace"];
const PUBLIC_CONTAINER_KINDS: &[&str] = &[
    "SectionNode",
    "PageSeriesNode",
    "OutlineGroup",
    "PageManifestNode",
    "VersionHistoryContent",
    "VersionProxy",
];
const CONTENT_JCIDS: &[&str] = &[
    "PageNode",
    "OutlineNode",
    "RichTextOENode",
    "ImageNode",
    "TableNode",
    "EmbeddedFileNode",
];

/// The result of loading a `.one` document.
pub struct LoadedOneNoteDocument {
    pub display_name: Option<String>,
    pub logical_document: LogicalDocument,
    pub pages: Vec<Page>,
    pub page_histories: Vec<Vec<Page>>,
}

/// Parse and load a `.one` document into the public DOM.
pub fn load_onenote_document(
    source: OneStoreSource<'_>,
    include_page_history: bool,
) -> Result<LoadedOneNoteDocument, AsposeNoteError> {
    let parsed = parse_onestore_file(source)?;
    Ok(Builder::new(&parsed, include_page_history).build())
}

// ---------------------------------------------------------------------------
// Decoding helpers
// ---------------------------------------------------------------------------

fn truthy(value: Option<&PropValue>) -> bool {
    value.is_some_and(|value| value.truthy())
}

fn as_int(value: Option<&PropValue>) -> Option<i64> {
    value.and_then(|value| value.as_int())
}

fn decode_utf16(value: Option<&PropValue>) -> Option<String> {
    match value {
        Some(PropValue::Bytes(bytes)) => {
            let decoded = decode_utf16le(bytes);
            if decoded.is_empty() {
                None
            } else {
                Some(decoded)
            }
        }
        _ => None,
    }
}

/// Decode `TextExtendedAscii`: genuine UTF-8 first, else Windows-1252. Mirrors
/// `_decode_ascii_text` for the bytes case.
pub(crate) fn decode_ascii_text_bytes(raw: &[u8]) -> Option<String> {
    let decoded = match std::str::from_utf8(raw) {
        Ok(text) => text.to_string(),
        Err(_) => decode_cp1252(raw),
    };
    let trimmed = decoded.trim_end_matches('\u{0}');
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn decode_ascii_text(value: Option<&PropValue>) -> Option<String> {
    match value {
        Some(PropValue::Bytes(bytes)) => decode_ascii_text_bytes(bytes),
        _ => None,
    }
}

fn decode_cp1252(raw: &[u8]) -> String {
    let mut result = String::with_capacity(raw.len());
    for &byte in raw {
        match byte {
            0x80..=0x9F => {
                if let Some(mapped) = cp1252_high(byte) {
                    result.push(mapped);
                }
            }
            _ => result.push(byte as char),
        }
    }
    result
}

fn cp1252_high(byte: u8) -> Option<char> {
    let code = match byte {
        0x80 => 0x20AC,
        0x82 => 0x201A,
        0x83 => 0x0192,
        0x84 => 0x201E,
        0x85 => 0x2026,
        0x86 => 0x2020,
        0x87 => 0x2021,
        0x88 => 0x02C6,
        0x89 => 0x2030,
        0x8A => 0x0160,
        0x8B => 0x2039,
        0x8C => 0x0152,
        0x8E => 0x017D,
        0x91 => 0x2018,
        0x92 => 0x2019,
        0x93 => 0x201C,
        0x94 => 0x201D,
        0x95 => 0x2022,
        0x96 => 0x2013,
        0x97 => 0x2014,
        0x98 => 0x02DC,
        0x99 => 0x2122,
        0x9A => 0x0161,
        0x9B => 0x203A,
        0x9C => 0x0153,
        0x9E => 0x017E,
        0x9F => 0x0178,
        _ => return None,
    };
    char::from_u32(code)
}

fn u32_to_float(value: Option<&PropValue>) -> Option<f64> {
    as_int(value).map(|value| f32::from_bits(value as u32) as f64)
}

fn bytes_to_float_list(value: Option<&PropValue>) -> Vec<f64> {
    match value {
        Some(PropValue::Bytes(bytes)) => bytes
            .chunks_exact(4)
            .map(|chunk| f32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]) as f64)
            .collect(),
        _ => Vec::new(),
    }
}

fn decode_paragraph_alignment(value: Option<&PropValue>) -> Option<HorizontalAlignment> {
    match as_int(value) {
        Some(0) => Some(HorizontalAlignment::Left),
        Some(1) => Some(HorizontalAlignment::Center),
        Some(2) => Some(HorizontalAlignment::Right),
        _ => None,
    }
}

fn decode_layout_alignment(value: Option<&PropValue>) -> Option<HorizontalAlignment> {
    let value = as_int(value)?;
    if value == 0 {
        return None;
    }
    match value & 0x7 {
        1 => Some(HorizontalAlignment::Left),
        2 => Some(HorizontalAlignment::Center),
        3 => Some(HorizontalAlignment::Right),
        4 => Some(HorizontalAlignment::Left),
        5 => Some(HorizontalAlignment::Right),
        _ => None,
    }
}

fn flatten_refs(value: &PropValue, out: &mut Vec<ExtendedGuid>) {
    match value {
        PropValue::ObjectId(guid) => out.push(guid.clone()),
        PropValue::ObjectIds(guids) => out.extend(guids.iter().cloned()),
        PropValue::List(items) => {
            for item in items {
                flatten_refs(item, out);
            }
        }
        PropValue::Set(map) => {
            for item in map.values() {
                flatten_refs(item, out);
            }
        }
        _ => {}
    }
}

fn decode_number_list_format(value: Option<&PropValue>) -> Option<String> {
    if let Some(text) = decode_utf16(value) {
        return Some(text);
    }
    match value {
        Some(PropValue::Bytes(bytes)) => Some(
            bytes
                .iter()
                .map(|byte| format!("{:02x}", byte))
                .collect::<String>(),
        ),
        _ => None,
    }
}

type TagKey = (
    Option<String>,
    Option<i64>,
    Option<i64>,
    Option<i64>,
    Option<DateTime>,
    Option<DateTime>,
);

fn tag_key(tag: &NoteTag) -> TagKey {
    (
        tag.label(),
        tag.icon(),
        tag.font_color(),
        tag.highlight(),
        tag.creation_time(),
        tag.completed_time(),
    )
}

fn unique_tags(tags: Vec<NoteTag>) -> Vec<NoteTag> {
    let mut seen: HashSet<TagKey> = HashSet::new();
    let mut result = Vec::new();
    for tag in tags {
        if seen.insert(tag_key(&tag)) {
            result.push(tag);
        }
    }
    result
}

fn extract_guid(reference: &str) -> Option<String> {
    let chars: Vec<char> = reference.chars().collect();
    let pattern = [8usize, 4, 4, 4, 12];
    let total: usize = pattern.iter().sum::<usize>() + 4;
    if chars.len() < total {
        return None;
    }
    for start in 0..=(chars.len() - total) {
        if guid_matches(&chars[start..start + total]) {
            return Some(chars[start..start + total].iter().collect());
        }
    }
    None
}

fn guid_matches(window: &[char]) -> bool {
    let groups = [8usize, 4, 4, 4, 12];
    let mut index = 0;
    for (group_index, group_len) in groups.iter().enumerate() {
        if group_index > 0 {
            if window.get(index) != Some(&'-') {
                return false;
            }
            index += 1;
        }
        for _ in 0..*group_len {
            match window.get(index) {
                Some(c) if c.is_ascii_hexdigit() => index += 1,
                _ => return false,
            }
        }
    }
    index == window.len()
}

fn descendants(node: &LogicalNode) -> Vec<&LogicalNode> {
    let mut result = Vec::new();
    let mut stack = vec![node];
    while let Some(current) = stack.pop() {
        result.push(current);
        for child in current.children.iter().rev() {
            stack.push(child);
        }
    }
    result
}

fn notebook_guid_bytes(object: &OneStoreObject) -> Option<Vec<u8>> {
    match object.properties.get("NotebookManagementEntityGuid") {
        Some(PropValue::Bytes(bytes)) if !bytes.is_empty() => Some(bytes.clone()),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Builder
// ---------------------------------------------------------------------------

struct Builder<'a> {
    parsed: &'a OneStoreFile,
    include_page_history: bool,
    section_space_index: Option<usize>,
    content_space_indices: Vec<usize>,
    primary_space_index: Option<usize>,
    primary_revision_index: Option<usize>,
    snapshot_cache: RefCell<SnapshotCache>,
    revindex_cache: RefCell<HashMap<usize, Rc<RevisionIndexMap>>>,
    section_orders_cache: RefCell<Option<Rc<SectionOrders>>>,
}

impl<'a> Builder<'a> {
    fn new(parsed: &'a OneStoreFile, include_page_history: bool) -> Builder<'a> {
        let mut builder = Builder {
            parsed,
            include_page_history,
            section_space_index: None,
            content_space_indices: Vec::new(),
            primary_space_index: None,
            primary_revision_index: None,
            snapshot_cache: RefCell::new(HashMap::new()),
            revindex_cache: RefCell::new(HashMap::new()),
            section_orders_cache: RefCell::new(None),
        };
        builder.section_space_index = builder.select_section_space();
        builder.content_space_indices = builder.select_content_spaces();
        let primary = builder
            .content_space_indices
            .first()
            .copied()
            .or_else(|| builder.select_content_space());
        builder.primary_space_index = primary.or(builder.section_space_index);
        builder.primary_revision_index = builder
            .primary_space_index
            .and_then(|space_index| builder.select_current_revision_index(space_index));
        builder
    }

    fn revindex_map(&self, space_index: usize) -> Rc<RevisionIndexMap> {
        if let Some(cached) = self.revindex_cache.borrow().get(&space_index) {
            return cached.clone();
        }
        let space = &self.parsed.object_spaces[space_index];
        let map: RevisionIndexMap = space
            .revisions
            .iter()
            .enumerate()
            .map(|(index, revision)| (revision.rid.clone(), index))
            .collect();
        let rc = Rc::new(map);
        self.revindex_cache
            .borrow_mut()
            .insert(space_index, rc.clone());
        rc
    }

    fn parent_revision_index(&self, space_index: usize, revision_index: usize) -> Option<usize> {
        let space = &self.parsed.object_spaces[space_index];
        let dependent = space.revisions.get(revision_index)?.rid_dependent.as_ref()?;
        if dependent.is_zero() {
            return None;
        }
        self.revindex_map(space_index).get(dependent).copied()
    }

    fn get_snapshot_objects(&self, space_index: usize, revision_index: Option<usize>) -> Rc<ObjMap> {
        let key = (space_index, revision_index.map(|r| r as i64).unwrap_or(-1));
        if let Some(cached) = self.snapshot_cache.borrow().get(&key) {
            return cached.clone();
        }
        let space = &self.parsed.object_spaces[space_index];
        let result = match revision_index {
            Some(ri) if ri < space.revisions.len() => match self.parent_revision_index(space_index, ri) {
                None => space.revisions[ri].objects.clone(),
                Some(parent) => {
                    let mut merged = (*self.get_snapshot_objects(space_index, Some(parent))).clone();
                    for (oid, object) in &space.revisions[ri].objects {
                        merged.insert(oid.clone(), object.clone());
                    }
                    merged
                }
            },
            _ => space.objects.clone(),
        };
        let rc = Rc::new(result);
        self.snapshot_cache.borrow_mut().insert(key, rc.clone());
        rc
    }

    fn get_snapshot_root(&self, space_index: usize, revision_index: Option<usize>) -> Option<ExtendedGuid> {
        let space = &self.parsed.object_spaces[space_index];
        if let Some(revision_index) = revision_index {
            if revision_index < space.revisions.len() {
                let mut current = Some(revision_index);
                let mut seen: HashSet<usize> = HashSet::new();
                while let Some(ci) = current {
                    if !seen.insert(ci) {
                        break;
                    }
                    if let Some(root) = space.revisions[ci].root_roles.get(&ROOT_ROLE_DEFAULT_CONTENT) {
                        return Some(root.clone());
                    }
                    current = self.parent_revision_index(space_index, ci);
                }
            }
        }
        space.get_latest_root(ROOT_ROLE_DEFAULT_CONTENT, None)
    }

    fn is_page_revision(&self, space_index: usize, revision_index: Option<usize>) -> bool {
        let revision_index = match revision_index {
            Some(index) => index,
            None => return false,
        };
        let space = &self.parsed.object_spaces[space_index];
        if revision_index >= space.revisions.len() {
            return false;
        }
        let root = match self.get_snapshot_root(space_index, Some(revision_index)) {
            Some(root) => root,
            None => return false,
        };
        let objects = self.get_snapshot_objects(space_index, Some(revision_index));
        match objects.get(&root) {
            Some(record) => matches!(record.jcid_name.as_str(), "PageManifestNode" | "PageNode"),
            None => false,
        }
    }

    fn select_current_revision_index(&self, space_index: usize) -> Option<usize> {
        let space = &self.parsed.object_spaces[space_index];
        if space.revisions.is_empty() {
            return None;
        }
        let preferred = space.get_latest_revision_index(ROOT_ROLE_DEFAULT_CONTENT, None);
        if let Some(preferred) = preferred {
            if self.is_page_revision(space_index, Some(preferred)) {
                return Some(preferred);
            }
        }
        for revision_index in (0..space.revisions.len()).rev() {
            if !self.is_page_revision(space_index, Some(revision_index)) {
                continue;
            }
            if space.revisions[revision_index].context_id.is_none() {
                return Some(revision_index);
            }
        }
        for revision_index in (0..space.revisions.len()).rev() {
            if self.is_page_revision(space_index, Some(revision_index)) {
                return Some(revision_index);
            }
        }
        Some(space.revisions.len() - 1)
    }

    fn select_section_space(&self) -> Option<usize> {
        for (index, space) in self.parsed.object_spaces.iter().enumerate() {
            if let Some(root) = space.get_latest_root(ROOT_ROLE_DEFAULT_CONTENT, None) {
                if let Some(record) = space.objects.get(&root) {
                    if record.jcid_name == "SectionNode" {
                        return Some(index);
                    }
                }
            }
        }
        None
    }

    fn space_content_score(&self, space_index: usize) -> usize {
        let space = &self.parsed.object_spaces[space_index];
        let mut best = count_content(&space.objects);
        for revision in &space.revisions {
            best = best.max(count_content(&revision.objects));
        }
        best
    }

    fn select_content_space(&self) -> Option<usize> {
        let mut best_space = None;
        let mut best_score: i64 = -1;
        for index in 0..self.parsed.object_spaces.len() {
            let score = self.space_content_score(index) as i64;
            if score > best_score {
                best_space = Some(index);
                best_score = score;
            }
        }
        best_space.or(self.section_space_index)
    }

    fn select_content_spaces(&self) -> Vec<usize> {
        let mut candidates: Vec<(usize, usize)> = Vec::new();
        for index in 0..self.parsed.object_spaces.len() {
            if Some(index) == self.section_space_index {
                continue;
            }
            let revision_index = self.select_current_revision_index(index);
            if !self.is_page_revision(index, revision_index) {
                continue;
            }
            candidates.push((index, self.space_content_score(index)));
        }
        candidates.sort_by(|a, b| b.1.cmp(&a.1).then(a.0.cmp(&b.0)));
        candidates.into_iter().map(|(index, _)| index).collect()
    }

    fn build(&self) -> LoadedOneNoteDocument {
        let logical_root = self.build_logical_root();
        let pages = self.build_current_page_entries();
        // Full multi-revision history reconstruction arrives in a later stage;
        // the `include_page_history` flag is retained for that work. For now
        // every page exposes a current-only history.
        let _ = self.include_page_history;
        let page_histories: Vec<Vec<Page>> = pages.iter().map(|_| Vec::new()).collect();
        let display_name = self.compute_display_name(&pages);
        LoadedOneNoteDocument {
            display_name,
            logical_document: LogicalDocument { root: logical_root },
            pages,
            page_histories,
        }
    }

    fn compute_display_name(&self, pages: &[Page]) -> Option<String> {
        if let Some(index) = self.section_space_index {
            if let Some(name) = &self.parsed.object_spaces[index].section_display_name {
                return Some(name.clone());
            }
        }
        let content_space = self
            .content_space_indices
            .first()
            .copied()
            .or_else(|| self.select_content_space());
        if let Some(index) = content_space {
            if let Some(name) = &self.parsed.object_spaces[index].section_display_name {
                return Some(name.clone());
            }
        }
        if self.parsed.display_name.is_some() {
            return self.parsed.display_name.clone();
        }
        pages
            .first()
            .and_then(|page| page.title())
            .and_then(|title| title.title_text())
            .map(|rich_text| rich_text.text())
    }

    fn build_logical_root(&self) -> Option<LogicalNode> {
        let space_index = self.primary_space_index?;
        let revision_index = self.primary_revision_index;
        let objects = self.get_snapshot_objects(space_index, revision_index);
        let root = self.get_snapshot_root(space_index, revision_index)?;
        self.build_logical_node(&root, objects.as_ref(), &mut HashSet::new())
    }

    fn build_logical_node(
        &self,
        oid: &ExtendedGuid,
        objects: &ObjMap,
        stack: &mut HashSet<ExtendedGuid>,
    ) -> Option<LogicalNode> {
        let record = objects.get(oid)?;
        if stack.contains(oid) {
            return Some(LogicalNode::new(
                record.jcid_name.clone(),
                oid.clone(),
                record.properties.clone(),
            ));
        }
        stack.insert(oid.clone());
        let mut children = Vec::new();
        for reference in ordered_property_refs(record, STRUCTURE_PROPERTIES) {
            if let Some(child) = self.build_logical_node(&reference, objects, stack) {
                children.push(child);
            }
        }
        let mut metadata = Vec::new();
        for reference in ordered_property_refs(record, METADATA_PROPERTIES) {
            if let Some(child) = self.build_logical_node(&reference, objects, stack) {
                metadata.push(child);
            }
        }
        stack.remove(oid);
        Some(LogicalNode {
            kind: record.jcid_name.clone(),
            oid: oid.clone(),
            properties: record.properties.clone(),
            children,
            metadata,
        })
    }

    fn build_current_page_entries(&self) -> Vec<Page> {
        let section_orders = self.section_page_orders();
        let candidate_spaces: Vec<usize> = if !self.content_space_indices.is_empty() {
            self.content_space_indices.clone()
        } else {
            self.select_content_space().into_iter().collect()
        };

        let mut entries: Vec<PageEntry> = Vec::new();
        for (space_order, space_index) in candidate_spaces.into_iter().enumerate() {
            let revision_index = self.select_current_revision_index(space_index);
            let objects = self.get_snapshot_objects(space_index, revision_index);
            let objects = objects.as_ref();
            let root = self.get_snapshot_root(space_index, revision_index);
            let logical_root =
                root.and_then(|root| self.build_logical_node(&root, objects, &mut HashSet::new()));
            let page_infos = self.build_page_infos(&logical_root, objects);
            if page_infos.is_empty() {
                continue;
            }
            for (page_order, (_oid, logical_page, page)) in page_infos.iter().enumerate() {
                let notebook_guid = self.page_notebook_guid(logical_page, objects);
                let topology = self.page_topology_timestamp(logical_page, objects);
                let section_order = notebook_guid
                    .as_ref()
                    .and_then(|guid| section_orders.get(guid).copied());
                entries.push(PageEntry {
                    page: page.clone(),
                    topology,
                    section_order,
                    space_order,
                    page_order,
                });
            }
        }

        entries.sort_by_key(|entry| {
            (
                entry.section_order.is_none(),
                entry.section_order.unwrap_or(0),
                entry.topology.is_none(),
                entry.topology.unwrap_or(0),
                entry.space_order,
                entry.page_order,
            )
        });
        entries.into_iter().map(|entry| entry.page).collect()
    }

    fn build_page_infos(
        &self,
        logical_root: &Option<LogicalNode>,
        objects: &ObjMap,
    ) -> Vec<PageInfo> {
        let mut pages = Vec::new();
        if let Some(root) = logical_root {
            for node in descendants(root) {
                if node.kind == "PageNode" {
                    pages.push((node.oid.clone(), node.clone(), self.build_page(node, objects)));
                }
            }
        }
        if !pages.is_empty() {
            return pages;
        }
        let page_oids: Vec<ExtendedGuid> = objects
            .values()
            .filter(|object| object.jcid_name == "PageNode")
            .map(|object| object.oid.clone())
            .collect();
        for oid in page_oids {
            if let Some(logical_page) = self.build_logical_node(&oid, objects, &mut HashSet::new()) {
                let page = self.build_page(&logical_page, objects);
                pages.push((oid, logical_page, page));
            }
        }
        pages
    }

    fn build_page(&self, node: &LogicalNode, objects: &ObjMap) -> Page {
        let page = Page::new();
        let record = objects.get(&node.oid);
        let metadata = self.page_metadata_from_node(node, objects);
        if let Some(record) = record {
            page.set_author(decode_utf16(record.properties.get("Author")));
            page.set_creation_time(time32_to_datetime(as_int(record.properties.get("CreationTimeStamp"))));
            page.set_last_modified_time(
                filetime_to_datetime(as_int(record.properties.get("LastModifiedTimeStamp")))
                    .or_else(|| time32_to_datetime(as_int(record.properties.get("LastModifiedTime")))),
            );
            page.set_level(as_int(record.properties.get("PageLevel")));
        }
        if let Some(metadata) = metadata {
            if page.author().is_none() {
                page.set_author(decode_utf16(metadata.properties.get("Author")));
            }
            if page.creation_time().is_none() {
                page.set_creation_time(time32_to_datetime(as_int(
                    metadata.properties.get("CreationTimeStamp"),
                )));
            }
            if page.last_modified_time().is_none() {
                page.set_last_modified_time(
                    filetime_to_datetime(as_int(metadata.properties.get("LastModifiedTimeStamp")))
                        .or_else(|| time32_to_datetime(as_int(metadata.properties.get("LastModifiedTime")))),
                );
            }
            if page.level().is_none() {
                page.set_level(as_int(metadata.properties.get("PageLevel")));
            }
        }

        let title_node = node.children.iter().find(|child| child.kind == "TitleNode");
        let mut has_title = false;
        if let Some(title_node) = title_node {
            let title = self.build_title(title_node, objects);
            page.append_child_last(title.as_node());
            has_title = true;
        }

        for child in &node.children {
            if child.kind == "TitleNode" {
                continue;
            }
            if let Some(public_child) = self.build_public_node(child, objects) {
                page.append_child_last(public_child);
            }
        }

        if !has_title && self.should_synthesize_page_title(node, objects) {
            let title_text = metadata.and_then(|metadata| {
                decode_utf16(metadata.properties.get("CachedTitleString"))
                    .or_else(|| decode_utf16(metadata.properties.get("CachedTitleStringFromPage")))
            });
            if let Some(title_text) = title_text {
                let title = Title::new();
                title.set_title_text(Some(RichText::build(
                    title_text.clone(),
                    Some(vec![TextRun::new(title_text)]),
                    ParagraphStyle::default(),
                    None,
                    Vec::new(),
                )));
                page.append_child_first(title.as_node());
            }
        }

        page
    }

    fn build_title(&self, node: &LogicalNode, objects: &ObjMap) -> Title {
        let title = Title::new();
        let rich_texts: Vec<&LogicalNode> = descendants(node)
            .into_iter()
            .filter(|candidate| candidate.kind == "RichTextOENode")
            .collect();
        if let Some(node) = rich_texts.first() {
            title.set_title_text(Some(self.build_richtext(node, objects)));
        }
        if let Some(node) = rich_texts.get(1) {
            title.set_title_date(Some(self.build_richtext(node, objects)));
        }
        if let Some(node) = rich_texts.get(2) {
            title.set_title_time(Some(self.build_richtext(node, objects)));
        }
        title
    }

    fn build_public_node(&self, node: &LogicalNode, objects: &ObjMap) -> Option<Node> {
        if PUBLIC_CONTAINER_KINDS.contains(&node.kind.as_str()) {
            return None;
        }
        match node.kind.as_str() {
            "OutlineNode" => Some(self.build_outline(node, objects).as_node()),
            "OutlineElementNode" => Some(self.build_outline_element(node, objects).as_node()),
            "RichTextOENode" => Some(self.build_richtext(node, objects).as_node()),
            "ImageNode" => Some(self.build_image(node, objects).as_node()),
            "EmbeddedFileNode" => Some(self.build_attached_file(node, objects).as_node()),
            "TableNode" => Some(self.build_table(node, objects).as_node()),
            "TableRowNode" => Some(self.build_table_row(node, objects).as_node()),
            "TableCellNode" => Some(self.build_table_cell(node, objects).as_node()),
            _ => None,
        }
    }

    fn append_public_children(&self, parent: &Node, node: &LogicalNode, objects: &ObjMap) {
        for child in &node.children {
            if let Some(public_child) = self.build_public_node(child, objects) {
                push_child(parent, public_child);
            } else if PUBLIC_CONTAINER_KINDS.contains(&child.kind.as_str()) {
                self.append_public_children(parent, child, objects);
            }
        }
    }

    fn build_outline(&self, node: &LogicalNode, objects: &ObjMap) -> Outline {
        let outline = Outline::new();
        if let Some(record) = objects.get(&node.oid) {
            outline.set_horizontal_offset(u32_to_float(record.properties.get("OffsetFromParentHoriz")));
            outline.set_vertical_offset(u32_to_float(record.properties.get("OffsetFromParentVert")));
            outline.set_max_width(u32_to_float(record.properties.get("LayoutMaxWidth")));
        }
        self.append_public_children(&outline.as_node(), node, objects);
        outline
    }

    fn build_outline_element(&self, node: &LogicalNode, objects: &ObjMap) -> OutlineElement {
        let outline_element = OutlineElement::new();
        let outline_tags = objects
            .get(&node.oid)
            .map(|record| self.collect_tags(record, objects))
            .unwrap_or_default();

        if let Some(list_node) = node.children.iter().find(|child| child.kind == "NumberListNode") {
            outline_element.set_number_list(Some(self.build_number_list(list_node, objects)));
        }

        let mut ordered: Vec<&LogicalNode> = node
            .children
            .iter()
            .filter(|child| child.kind != "NumberListNode" && child.kind != "OutlineElementNode")
            .collect();
        ordered.extend(node.children.iter().filter(|child| child.kind == "OutlineElementNode"));

        for child in ordered {
            if let Some(public_child) = self.build_public_node(child, objects) {
                if let Some(rich_text) = RichText::from_node(&public_child) {
                    if rich_text.tags().is_empty() && !outline_tags.is_empty() {
                        rich_text.tags_mut().extend(outline_tags.iter().cloned());
                    }
                }
                push_child(&outline_element.as_node(), public_child);
            }
        }
        outline_element
    }

    fn build_richtext(&self, node: &LogicalNode, objects: &ObjMap) -> RichText {
        let record = match objects.get(&node.oid) {
            Some(record) => record,
            None => return RichText::new(),
        };
        let text = decode_utf16(record.properties.get("RichEditTextUnicode"))
            .or_else(|| decode_ascii_text(record.properties.get("TextExtendedAscii")))
            .unwrap_or_default();
        let base_style = text_style_from_propmap(&record.properties);
        let runs = if text.is_empty() {
            None
        } else {
            Some(vec![TextRun::with_style(text.clone(), base_style)])
        };
        let rich_text = RichText::build(
            text.clone(),
            runs,
            paragraph_style_from_propmap(&record.properties),
            decode_paragraph_alignment(record.properties.get("ParagraphAlignment")),
            self.collect_tags(record, objects),
        );

        let run_styles: Vec<PropValue> = match record.properties.get("TextRunFormatting") {
            Some(PropValue::List(items)) => items.clone(),
            Some(PropValue::ObjectIds(ids)) => ids.iter().cloned().map(PropValue::ObjectId).collect(),
            _ => Vec::new(),
        };
        if !run_styles.is_empty() {
            let text_chars: Vec<char> = text.chars().collect();
            let text_len = text_chars.len();
            let mut boundaries =
                decode_run_boundaries(record.properties.get("TextRunIndex"), run_styles.len(), text_len);
            if boundaries.is_empty() || boundaries[0] != 0 {
                boundaries.insert(0, 0);
            }
            let mut trimmed: Vec<usize> = boundaries.into_iter().take(run_styles.len()).collect();
            trimmed.push(text_len);
            trimmed.sort_unstable();
            let boundaries = trimmed;

            let mut text_runs = Vec::new();
            for (index, style_props) in run_styles.iter().enumerate() {
                let start = boundaries.get(index).copied().unwrap_or(0).min(text_len);
                let end = boundaries.get(index + 1).copied().unwrap_or(text_len).min(text_len);
                let run_text: String = if end > start {
                    text_chars[start..end].iter().collect()
                } else {
                    String::new()
                };
                let style = text_style_from_opt(self.resolve_style_properties(style_props, objects));
                text_runs.push(TextRun::with_style(run_text, style));
            }
            rich_text.replace_text_runs(text_runs);
        }

        rich_text
    }

    fn resolve_style_properties<'o>(
        &self,
        value: &'o PropValue,
        objects: &'o ObjMap,
    ) -> Option<&'o PropMap> {
        match value {
            PropValue::Set(map) => Some(map),
            PropValue::ObjectId(guid) => objects.get(guid).map(|object| &object.properties),
            _ => None,
        }
    }

    fn resolve_file_bytes(&self, record: &OneStoreObject, property: &str, objects: &ObjMap) -> Vec<u8> {
        let reference_guid = match record.properties.get(property) {
            Some(PropValue::ObjectId(guid)) => guid,
            _ => return Vec::new(),
        };
        let target = match objects.get(reference_guid) {
            Some(target) => target,
            None => return Vec::new(),
        };
        let reference = target.file_data_reference.clone().unwrap_or_default();
        if reference.to_ascii_lowercase().starts_with("<ifndf>") {
            if let Some(guid) = extract_guid(&reference) {
                return self
                    .parsed
                    .file_data_store
                    .get(&guid.to_ascii_lowercase())
                    .cloned()
                    .unwrap_or_default();
            }
        }
        Vec::new()
    }

    fn build_image(&self, node: &LogicalNode, objects: &ObjMap) -> Image {
        let record = match objects.get(&node.oid) {
            Some(record) => record,
            None => return Image::new(),
        };
        let width = u32_to_float(record.properties.get("PictureWidth"));
        let height = u32_to_float(record.properties.get("PictureHeight"));
        let mut bytes = self.resolve_file_bytes(record, "PictureContainer", objects);
        if bytes.is_empty() {
            bytes = self.resolve_file_bytes(record, "WebPictureContainer14", objects);
        }
        Image::new()
            .with_file_name(decode_utf16(record.properties.get("ImageFilename")))
            .with_bytes(bytes)
            .with_width(width)
            .with_height(height)
            .with_original_width(width)
            .with_original_height(height)
            .with_alignment(
                decode_layout_alignment(record.properties.get("LayoutAlignmentSelf"))
                    .or_else(|| decode_layout_alignment(record.properties.get("LayoutAlignmentInParent"))),
            )
            .with_alternative_text_description(decode_utf16(record.properties.get("ImageAltText")))
            .with_hyperlink_url(decode_utf16(record.properties.get("WzHyperlinkUrl")))
            .with_tags(self.collect_tags(record, objects))
    }

    fn build_attached_file(&self, node: &LogicalNode, objects: &ObjMap) -> AttachedFile {
        let record = match objects.get(&node.oid) {
            Some(record) => record,
            None => return AttachedFile::new(),
        };
        AttachedFile::new()
            .with_file_name(
                decode_utf16(record.properties.get("EmbeddedFileName"))
                    .or_else(|| decode_utf16(record.properties.get("SourceFilepath"))),
            )
            .with_bytes(self.resolve_file_bytes(record, "EmbeddedFileContainer", objects))
            .with_tags(self.collect_tags(record, objects))
    }

    fn build_table(&self, node: &LogicalNode, objects: &ObjMap) -> Table {
        let record = objects.get(&node.oid);
        let column_widths =
            record.map(|record| bytes_to_float_list(record.properties.get("TableColumnWidths")));
        let columns: Vec<TableColumn> = column_widths
            .unwrap_or_default()
            .into_iter()
            .map(|width| TableColumn::new(Some(width)))
            .collect();
        let borders_visible = match record.and_then(|record| record.properties.get("TableBordersVisible")) {
            Some(value) => value.truthy(),
            None => true,
        };
        let tags = record.map(|record| self.collect_tags(record, objects)).unwrap_or_default();
        let table = Table::new()
            .with_tags(tags)
            .with_columns(columns)
            .with_borders_visible(borders_visible);
        for child in &node.children {
            if child.kind == "TableRowNode" {
                push_child(&table.as_node(), self.build_table_row(child, objects).as_node());
            }
        }
        table
    }

    fn build_table_row(&self, node: &LogicalNode, objects: &ObjMap) -> TableRow {
        let row = TableRow::new();
        for child in &node.children {
            if child.kind == "TableCellNode" {
                push_child(&row.as_node(), self.build_table_cell(child, objects).as_node());
            }
        }
        row
    }

    fn build_table_cell(&self, node: &LogicalNode, objects: &ObjMap) -> TableCell {
        let cell = TableCell::new();
        self.append_public_children(&cell.as_node(), node, objects);
        cell
    }

    fn build_number_list(&self, node: &LogicalNode, objects: &ObjMap) -> NumberList {
        let record = objects.get(&node.oid);
        let format = record.and_then(|record| decode_number_list_format(record.properties.get("NumberListFormat")));
        let restart = record.and_then(|record| as_int(record.properties.get("ListRestart")));
        NumberList {
            format: format.clone(),
            number_format: format,
            restart,
            ..NumberList::default()
        }
    }

    fn collect_tags(&self, record: &OneStoreObject, objects: &ObjMap) -> Vec<NoteTag> {
        let mut tags = tag_from_record(Some(record));
        if let Some(PropValue::List(states)) = record.properties.get("NoteTagStates") {
            for state in states {
                tags.extend(self.tag_from_state(state, objects));
            }
        }
        for (name, value) in &record.raw_properties {
            if name != "NoteTagDefinitionOid"
                && name != "MetaDataObjectsAboveGraphSpace"
                && name != "NoteTagStates"
            {
                continue;
            }
            let mut refs = Vec::new();
            flatten_refs(value, &mut refs);
            for reference in refs {
                if let Some(target) = objects.get(&reference) {
                    tags.extend(tag_from_record(Some(target)));
                    if let Some(PropValue::ObjectId(nested)) = target.properties.get("NoteTagDefinitionOid") {
                        tags.extend(tag_from_record(objects.get(nested)));
                    }
                }
            }
        }
        unique_tags(tags)
    }

    fn tag_from_state(&self, state: &PropValue, objects: &ObjMap) -> Vec<NoteTag> {
        let map = match state {
            PropValue::Set(map) => map,
            _ => return Vec::new(),
        };
        let mut tags = Vec::new();
        if let Some(PropValue::ObjectId(definition)) = map.get("NoteTagDefinitionOid") {
            tags.extend(tag_from_record(objects.get(definition)));
        }
        let inline = NoteTag::create_internal(
            decode_utf16(map.get("NoteTagLabel")),
            as_int(map.get("NoteTagShape")),
            as_int(map.get("NoteTagHighlightColor")),
            as_int(map.get("NoteTagCreated")).and_then(coerce_note_tag_datetime),
            as_int(map.get("NoteTagCompleted")).and_then(coerce_note_tag_datetime),
            as_int(map.get("NoteTagTextColor")),
        );
        if inline.label().is_some() || inline.icon().is_some() {
            tags.push(inline);
        }
        tags
    }

    fn find_metadata_record<'o>(
        &self,
        node: &LogicalNode,
        kind: &str,
        objects: &'o ObjMap,
    ) -> Option<&'o OneStoreObject> {
        for metadata in &node.metadata {
            if let Some(record) = objects.get(&metadata.oid) {
                if record.jcid_name == kind {
                    return Some(record);
                }
            }
        }
        None
    }

    fn page_metadata_from_node<'o>(
        &self,
        node: &LogicalNode,
        objects: &'o ObjMap,
    ) -> Option<&'o OneStoreObject> {
        if let Some(metadata) = self.find_metadata_record(node, "PageMetaData", objects) {
            return Some(metadata);
        }
        let candidates: Vec<&OneStoreObject> = objects
            .values()
            .filter(|object| object.jcid_name == "PageMetaData")
            .collect();
        if candidates.len() == 1 {
            return Some(candidates[0]);
        }
        if let Some(record) = objects.get(&node.oid) {
            if let Some(notebook_guid) = notebook_guid_bytes(record) {
                for candidate in &candidates {
                    if notebook_guid_bytes(candidate) == Some(notebook_guid.clone()) {
                        return Some(candidate);
                    }
                }
            }
        }
        None
    }

    fn page_notebook_guid(&self, node: &LogicalNode, objects: &ObjMap) -> Option<Vec<u8>> {
        if let Some(metadata) = self.page_metadata_from_node(node, objects) {
            if let Some(guid) = notebook_guid_bytes(metadata) {
                return Some(guid);
            }
        }
        if let Some(record) = objects.get(&node.oid) {
            if let Some(guid) = notebook_guid_bytes(record) {
                return Some(guid);
            }
        }
        None
    }

    fn page_topology_timestamp(&self, node: &LogicalNode, objects: &ObjMap) -> Option<i64> {
        let metadata = self.page_metadata_from_node(node, objects)?;
        as_int(metadata.properties.get("TopologyCreationTimeStamp"))
    }

    fn should_synthesize_page_title(&self, node: &LogicalNode, objects: &ObjMap) -> bool {
        if self.find_metadata_record(node, "PageMetaData", objects).is_some() {
            return true;
        }
        self.section_page_orders().len() > 1
    }

    fn section_page_orders(&self) -> Rc<SectionOrders> {
        if let Some(cached) = self.section_orders_cache.borrow().as_ref().cloned() {
            return cached;
        }
        let mut result: SectionOrders = HashMap::new();
        if let Some(space_index) = self.section_space_index {
            let revision_index = self.select_current_revision_index(space_index);
            let objects = self.get_snapshot_objects(space_index, revision_index);
            let objects = objects.as_ref();
            let root = self.get_snapshot_root(space_index, revision_index);
            let logical_root =
                root.and_then(|root| self.build_logical_node(&root, objects, &mut HashSet::new()));
            if let Some(logical_root) = &logical_root {
                let mut next_order = 0;
                for node in descendants(logical_root) {
                    if node.kind != "PageSeriesNode" {
                        continue;
                    }
                    for metadata in &node.metadata {
                        if let Some(record) = objects.get(&metadata.oid) {
                            if record.jcid_name != "PageMetaData" {
                                continue;
                            }
                            if let Some(notebook_guid) = notebook_guid_bytes(record) {
                                if result.contains_key(&notebook_guid) {
                                    continue;
                                }
                                result.insert(notebook_guid, next_order);
                                next_order += 1;
                            }
                        }
                    }
                }
            }
        }
        let rc = Rc::new(result);
        *self.section_orders_cache.borrow_mut() = Some(rc.clone());
        rc
    }
}

fn count_content(objects: &ObjMap) -> usize {
    objects
        .values()
        .filter(|object| CONTENT_JCIDS.contains(&object.jcid_name.as_str()))
        .count()
}

struct PageEntry {
    page: Page,
    topology: Option<i64>,
    section_order: Option<usize>,
    space_order: usize,
    page_order: usize,
}

fn ordered_property_refs(record: &OneStoreObject, allowed: &[&str]) -> Vec<ExtendedGuid> {
    let mut refs = Vec::new();
    for (name, value) in &record.raw_properties {
        if allowed.contains(&name.as_str()) {
            flatten_refs(value, &mut refs);
        }
    }
    refs
}

fn le_uint_usize(bytes: &[u8]) -> usize {
    let mut value: usize = 0;
    for (index, byte) in bytes.iter().enumerate() {
        value |= (*byte as usize) << (8 * index);
    }
    value
}

/// Decode per-run start offsets from `TextRunIndex`. Mirrors
/// `_decode_run_boundaries`.
fn decode_run_boundaries(raw: Option<&PropValue>, run_count: usize, text_length: usize) -> Vec<usize> {
    let bytes = match raw {
        Some(PropValue::Bytes(bytes)) if run_count > 0 => bytes,
        _ => return vec![0; run_count],
    };
    for width in [4usize, 2] {
        if bytes.len() % width != 0 {
            continue;
        }
        let values: Vec<usize> = bytes.chunks_exact(width).map(le_uint_usize).collect();
        if values.len() != run_count - 1 && values.len() != run_count {
            continue;
        }
        let ascending = values.windows(2).all(|pair| pair[0] <= pair[1]);
        if !values.is_empty() && values.iter().all(|&value| value <= text_length) && ascending {
            return values.into_iter().take(run_count).collect();
        }
    }
    vec![0; run_count]
}

fn text_style_from_propmap(props: &PropMap) -> TextStyle {
    TextStyle {
        is_hyperlink: truthy(props.get("Hyperlink")) || truthy(props.get("WzHyperlinkUrl")),
        hyperlink_address: decode_utf16(props.get("WzHyperlinkUrl")),
        font_name: decode_utf16(props.get("Font")),
        font_size: as_int(props.get("FontSize")).map(|value| value as f64 / 2.0),
        font_color: as_int(props.get("FontColor")),
        highlight: as_int(props.get("Highlight")),
        language: as_int(props.get("LanguageID")),
        is_bold: truthy(props.get("Bold")),
        is_italic: truthy(props.get("Italic")),
        is_underline: truthy(props.get("Underline")),
        is_strikethrough: truthy(props.get("Strikethrough")),
        is_superscript: truthy(props.get("Superscript")),
        is_subscript: truthy(props.get("Subscript")),
        is_hidden: false,
        is_math_formatting: false,
    }
}

fn text_style_from_opt(props: Option<&PropMap>) -> TextStyle {
    props.map(text_style_from_propmap).unwrap_or_default()
}

fn paragraph_style_from_propmap(props: &PropMap) -> ParagraphStyle {
    ParagraphStyle {
        font_name: decode_utf16(props.get("Font")),
        font_size: as_int(props.get("FontSize")).map(|value| value as f64 / 2.0),
        font_color: as_int(props.get("FontColor")),
        highlight: as_int(props.get("Highlight")),
        is_bold: truthy(props.get("Bold")),
        is_italic: truthy(props.get("Italic")),
        is_underline: truthy(props.get("Underline")),
        is_strikethrough: truthy(props.get("Strikethrough")),
        is_superscript: truthy(props.get("Superscript")),
        is_subscript: truthy(props.get("Subscript")),
    }
}

fn tag_from_record(record: Option<&OneStoreObject>) -> Vec<NoteTag> {
    let record = match record {
        Some(record) => record,
        None => return Vec::new(),
    };
    let props = &record.properties;
    let tag = NoteTag::create_internal(
        decode_utf16(props.get("NoteTagLabel")),
        as_int(props.get("NoteTagShape")),
        as_int(props.get("NoteTagHighlightColor")),
        as_int(props.get("NoteTagCreated")).and_then(coerce_note_tag_datetime),
        as_int(props.get("NoteTagCompleted")).and_then(coerce_note_tag_datetime),
        as_int(props.get("NoteTagTextColor")),
    );
    if tag.label().is_some() || tag.icon().is_some() {
        vec![tag]
    } else {
        Vec::new()
    }
}

#[cfg(test)]
mod tests {
    use super::decode_ascii_text_bytes;

    #[test]
    fn decodes_cp1252_umlauts() {
        assert_eq!(
            decode_ascii_text_bytes(b"caf\xe9 \xc4\xf6\xfc\xdf").as_deref(),
            Some("café Äöüß")
        );
    }

    #[test]
    fn decodes_cp1252_typography() {
        assert_eq!(
            decode_ascii_text_bytes(b"\x93x\x94 \x97 \x80").as_deref(),
            Some("“x” — €")
        );
    }

    #[test]
    fn prefers_genuine_utf8() {
        assert_eq!(decode_ascii_text_bytes("café".as_bytes()).as_deref(), Some("café"));
    }

    #[test]
    fn strips_trailing_nul() {
        assert_eq!(decode_ascii_text_bytes(b"caf\xe9\x00").as_deref(), Some("café"));
    }

    #[test]
    fn empty_bytes_returns_none() {
        assert_eq!(decode_ascii_text_bytes(b""), None);
    }
}

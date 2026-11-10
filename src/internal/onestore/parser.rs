//! MS-ONESTORE binary container parser, mirroring
//! `aspose.note._internal.onestore.parser`.
//!
//! This parses the `.one` revision-store file into object spaces, each carrying
//! the `OneStoreObject`s (JCID + decoded property set) that the MS-ONE loader
//! turns into the public DOM.

use crate::error::AsposeNoteError;
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

const FILE_NODE_LIST_MAGIC: u64 = 0xA456_7AB1_F5F7_F4C4;
const FILE_NODE_LIST_FOOTER_MAGIC: u64 = 0x8BC2_15C3_8233_BA4B;
const SECTION_FILE_TYPE: &str = "7b5c52e4-d88c-4da7-aeb1-5378d02996d3";

fn stp_size(format: u8) -> usize {
    match format {
        0 => 8,
        1 => 4,
        2 => 2,
        3 => 4,
        _ => 0,
    }
}

fn cb_size(format: u8) -> usize {
    match format {
        0 => 4,
        1 => 8,
        2 => 1,
        3 => 2,
        _ => 0,
    }
}

/// Known property ids → symbolic names (`PROPERTY_NAMES`). Unknown ids fall back
/// to a lower-case hex string, matching Python's `hex(property_id)`.
fn property_name(property_id: u32) -> String {
    let known = match property_id {
        0x1C01 => "PageWidth",
        0x1C02 => "PageHeight",
        0x1C03 => "OutlineElementChildLevel",
        0x1C04 => "Bold",
        0x1C05 => "Italic",
        0x1C06 => "Underline",
        0x1C07 => "Strikethrough",
        0x1C08 => "Superscript",
        0x1C09 => "Subscript",
        0x1C0A => "Font",
        0x1C0B => "FontSize",
        0x1C0C => "FontColor",
        0x1C0D => "Highlight",
        0x1C14 => "OffsetFromParentHoriz",
        0x1C15 => "OffsetFromParentVert",
        0x1C1A => "NumberListFormat",
        0x1C1B => "LayoutMaxWidth",
        0x1C1C => "LayoutMaxHeight",
        0x1C1F => "ContentChildNodes",
        0x1C20 => "ElementChildNodes",
        0x1C22 => "RichEditTextUnicode",
        0x1C26 => "ListNodes",
        0x1C30 => "NotebookManagementEntityGuid",
        0x1C3B => "LanguageID",
        0x1C3E => "LayoutAlignmentInParent",
        0x1C3F => "PictureContainer",
        0x1C65 => "TopologyCreationTimeStamp",
        0x1C84 => "LayoutAlignmentSelf",
        0x1C87 => "IsTitleTime",
        0x1C88 => "IsBoilerText",
        0x1CB4 => "IsTitleText",
        0x1CB5 => "IsTitleDate",
        0x1CB7 => "ListRestart",
        0x1CF3 => "CachedTitleString",
        0x1CFE => "RichEditTextLangID",
        0x1D09 => "CreationTimeStamp",
        0x1D3C => "CachedTitleStringFromPage",
        0x1D57 => "RowCount",
        0x1D58 => "ColumnCount",
        0x1D5E => "TableBordersVisible",
        0x1D5F => "StructureElementChildNodes",
        0x1D63 => "ChildGraphSpaceElementNodes",
        0x1D66 => "TableColumnWidths",
        0x1D75 => "Author",
        0x1D77 => "LastModifiedTimeStamp",
        0x1D78 => "AuthorOriginal",
        0x1D79 => "AuthorMostRecent",
        0x1D7A => "LastModifiedTime",
        0x1D9B => "EmbeddedFileContainer",
        0x1D9C => "EmbeddedFileName",
        0x1D9D => "SourceFilepath",
        0x1DD7 => "ImageFilename",
        0x1DFF => "PageLevel",
        0x1E12 => "TextRunIndex",
        0x1E13 => "TextRunFormatting",
        0x1E14 => "Hyperlink",
        0x1E20 => "WzHyperlinkUrl",
        0x1E58 => "ImageAltText",
        0x342C => "ParagraphStyle",
        0x3442 => "MetaDataObjectsAboveGraphSpace",
        0x3458 => "TextRunDataObject",
        0x3498 => "TextExtendedAscii",
        0x3499 => "TextRunData",
        0x345A => "ParagraphStyleId",
        0x3462 => "HasVersionPages",
        0x3463 => "ActionItemType",
        0x3464 => "NoteTagShape",
        0x3465 => "NoteTagHighlightColor",
        0x3466 => "NoteTagTextColor",
        0x3467 => "NoteTagPropertyStatus",
        0x3468 => "NoteTagLabel",
        0x346E => "NoteTagCreated",
        0x346F => "NoteTagCompleted",
        0x3470 => "ActionItemStatus",
        0x3473 => "ActionItemSchemaVersion",
        0x3477 => "ParagraphAlignment",
        0x347B => "VersionHistoryGraphSpaceContextNodes",
        0x3480 => "DisplayedPageNumber",
        0x3488 => "NoteTagDefinitionOid",
        0x3489 => "NoteTagStates",
        0x349B => "SectionDisplayName",
        0x34C8 => "WebPictureContainer14",
        0x34CD => "PictureWidth",
        0x34CE => "PictureHeight",
        _ => "",
    };
    if known.is_empty() {
        format!("{:#x}", property_id)
    } else {
        known.to_string()
    }
}

fn property_type_override(raw_property_id: u32) -> Option<u8> {
    match raw_property_id {
        0x3400_347B => Some(0x11),
        _ => None,
    }
}

/// Known JCID index → symbolic names (`JCID_NAMES`). Unknown indices fall back
/// to a lower-case hex string.
fn jcid_name(index: u32) -> String {
    let known = match index {
        0x0001 => "ReadOnlyPersistablePropertyContainerForAuthor",
        0x0007 => "SectionNode",
        0x0008 => "PageSeriesNode",
        0x000B => "PageNode",
        0x000C => "OutlineNode",
        0x000D => "OutlineElementNode",
        0x000E => "RichTextOENode",
        0x0011 => "ImageNode",
        0x0012 => "NumberListNode",
        0x0019 => "OutlineGroup",
        0x0022 => "TableNode",
        0x0023 => "TableRowNode",
        0x0024 => "TableCellNode",
        0x002C => "TitleNode",
        0x0030 => "PageMetaData",
        0x0031 => "SectionMetaData",
        0x0035 => "EmbeddedFileNode",
        0x0037 => "PageManifestNode",
        0x0038 => "ConflictPageMetaData",
        0x003C => "VersionHistoryContent",
        0x003D => "VersionProxy",
        0x0043 => "NoteTagSharedDefinitionContainer",
        0x0044 => "RevisionMetaData",
        0x0046 => "VersionHistoryMetaData",
        0x004D => "ParagraphStyleObject",
        _ => "",
    };
    if known.is_empty() {
        format!("{:#x}", index)
    } else {
        known.to_string()
    }
}

/// An MS-ONESTORE ExtendedGUID (GUID + 32-bit index). Used as an object key.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ExtendedGuid {
    pub guid: String,
    pub n: u32,
}

impl ExtendedGuid {
    pub fn is_zero(&self) -> bool {
        self.guid == "00000000-0000-0000-0000-000000000000" && self.n == 0
    }
}

/// A decoded JCID (object type id + flags).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Jcid {
    pub index: u32,
    pub is_binary: bool,
    pub is_property_set: bool,
    pub is_graph_node: bool,
    pub is_file_data: bool,
    pub is_read_only: bool,
}

/// A decoded property value. Mirrors the dynamically typed values the Python
/// parser stores in `OneStoreObject.properties`.
#[derive(Debug, Clone, PartialEq)]
pub enum PropValue {
    /// Python `None` (e.g. an unresolved object id or unknown value type).
    Null,
    Bool(bool),
    /// Any 1/2/4/8-byte integer value.
    Int(u64),
    Bytes(Vec<u8>),
    ObjectId(ExtendedGuid),
    ObjectIds(Vec<ExtendedGuid>),
    /// A scalar array (type 0x10) or a property merged from duplicate ids.
    List(Vec<PropValue>),
    /// A nested property set (type 0x11).
    Set(PropMap),
}

impl PropValue {
    /// Python truthiness of a property value.
    pub fn truthy(&self) -> bool {
        match self {
            PropValue::Null => false,
            PropValue::Bool(value) => *value,
            PropValue::Int(value) => *value != 0,
            PropValue::Bytes(value) => !value.is_empty(),
            PropValue::ObjectId(_) => true,
            PropValue::ObjectIds(value) => !value.is_empty(),
            PropValue::List(value) => !value.is_empty(),
            PropValue::Set(value) => !value.is_empty(),
        }
    }

    pub fn as_int(&self) -> Option<i64> {
        match self {
            PropValue::Int(value) => Some(*value as i64),
            _ => None,
        }
    }

    pub fn as_bytes(&self) -> Option<&[u8]> {
        match self {
            PropValue::Bytes(value) => Some(value),
            _ => None,
        }
    }

    pub fn as_object_id(&self) -> Option<&ExtendedGuid> {
        match self {
            PropValue::ObjectId(value) => Some(value),
            _ => None,
        }
    }
}

pub type PropMap = HashMap<String, PropValue>;

/// A parsed OneStore object: a JCID plus its decoded property set.
#[derive(Debug, Clone)]
pub struct OneStoreObject {
    pub oid: ExtendedGuid,
    pub jcid: Jcid,
    pub jcid_name: String,
    pub properties: PropMap,
    pub raw_properties: Vec<(String, PropValue)>,
    pub file_data_reference: Option<String>,
    pub file_extension: Option<String>,
}

#[derive(Debug, Clone)]
struct FileNode {
    file_node_id: u16,
    payload: Vec<u8>,
    stp_format: u8,
    cb_format: u8,
    #[allow(dead_code)]
    base_type: u8,
}

/// A single revision manifest and the objects it introduces.
#[derive(Debug, Clone)]
pub struct RevisionManifestRecord {
    pub rid: ExtendedGuid,
    pub rid_dependent: Option<ExtendedGuid>,
    pub context_id: Option<ExtendedGuid>,
    pub root_roles: HashMap<u32, ExtendedGuid>,
    pub objects: HashMap<ExtendedGuid, OneStoreObject>,
    pub section_display_name: Option<String>,
    pub order_index: usize,
}

#[derive(Debug, Clone)]
pub struct RevisionRoleBinding {
    pub rid: ExtendedGuid,
    pub context_id: Option<ExtendedGuid>,
    pub role: u32,
    pub order_index: usize,
}

/// The state of a single object space: its revisions, role bindings, and the
/// materialized "current" object snapshot.
#[derive(Debug, Clone, Default)]
pub struct ObjectSpaceState {
    pub gosid: ExtendedGuid,
    pub root_roles: HashMap<u32, ExtendedGuid>,
    pub objects: HashMap<ExtendedGuid, OneStoreObject>,
    pub section_display_name: Option<String>,
    pub revisions: Vec<RevisionManifestRecord>,
    pub role_bindings: Vec<RevisionRoleBinding>,
    pub latest_roots_by_context_role: HashMap<(Option<ExtendedGuid>, u32), ExtendedGuid>,
    pub latest_revision_indices_by_context_role: HashMap<(Option<ExtendedGuid>, u32), usize>,
}

impl Default for ExtendedGuid {
    fn default() -> Self {
        ExtendedGuid {
            guid: "00000000-0000-0000-0000-000000000000".to_string(),
            n: 0,
        }
    }
}

impl ObjectSpaceState {
    fn new(gosid: ExtendedGuid) -> Self {
        ObjectSpaceState {
            gosid,
            ..Default::default()
        }
    }

    pub fn get_latest_root(&self, role: u32, context_id: Option<&ExtendedGuid>) -> Option<ExtendedGuid> {
        if !self.latest_roots_by_context_role.is_empty() {
            return self
                .latest_roots_by_context_role
                .get(&(context_id.cloned(), role))
                .cloned();
        }
        if context_id.is_none() {
            return self.root_roles.get(&role).cloned();
        }
        None
    }

    pub fn iter_latest_roots(&self, role: u32) -> Vec<(Option<ExtendedGuid>, ExtendedGuid)> {
        if self.latest_roots_by_context_role.is_empty() {
            return match self.root_roles.get(&role) {
                Some(root) => vec![(None, root.clone())],
                None => Vec::new(),
            };
        }
        let mut result = Vec::new();
        for ((context_id, current_role), root) in &self.latest_roots_by_context_role {
            if *current_role == role {
                result.push((context_id.clone(), root.clone()));
            }
        }
        result
    }

    pub fn get_latest_revision_index(&self, role: u32, context_id: Option<&ExtendedGuid>) -> Option<usize> {
        if !self.latest_revision_indices_by_context_role.is_empty() {
            return self
                .latest_revision_indices_by_context_role
                .get(&(context_id.cloned(), role))
                .copied();
        }
        for index in (0..self.revisions.len()).rev() {
            let revision = &self.revisions[index];
            if revision.context_id.as_ref() != context_id {
                continue;
            }
            if resolve_revision_root(self, revision, role).is_some() {
                return Some(index);
            }
        }
        None
    }
}

/// A fully parsed OneStore file.
#[derive(Debug, Clone)]
pub struct OneStoreFile {
    pub data: Vec<u8>,
    pub object_spaces: Vec<ObjectSpaceState>,
    pub file_data_store: HashMap<String, Vec<u8>>,
    pub display_name: Option<String>,
}

struct PropertyState {
    object_ids: Vec<ExtendedGuid>,
    oid_cursor: usize,
}

// ---------------------------------------------------------------------------
// Low-level readers
// ---------------------------------------------------------------------------

fn rd_u16(data: &[u8], offset: usize) -> Option<u16> {
    data.get(offset..offset + 2)
        .map(|s| u16::from_le_bytes([s[0], s[1]]))
}

fn rd_u32(data: &[u8], offset: usize) -> Option<u32> {
    data.get(offset..offset + 4)
        .map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]]))
}

fn rd_u64(data: &[u8], offset: usize) -> Option<u64> {
    data.get(offset..offset + 8)
        .map(|s| u64::from_le_bytes([s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7]]))
}

/// Format 16 GUID bytes (little-endian layout) as a canonical lower-case string,
/// zero-padding a truncated payload. Mirrors `_read_guid`.
fn format_guid_le(chunk: &[u8]) -> String {
    let mut bytes = [0u8; 16];
    let take = chunk.len().min(16);
    bytes[..take].copy_from_slice(&chunk[..take]);
    let d1 = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let d2 = u16::from_le_bytes([bytes[4], bytes[5]]);
    let d3 = u16::from_le_bytes([bytes[6], bytes[7]]);
    format!(
        "{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        d1, d2, d3, bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
        bytes[15]
    )
}

fn read_guid(data: &[u8], offset: usize) -> String {
    if offset >= data.len() {
        return format_guid_le(&[]);
    }
    let end = (offset + 16).min(data.len());
    format_guid_le(&data[offset..end])
}

fn read_extended_guid(data: &[u8], offset: usize) -> ExtendedGuid {
    ExtendedGuid {
        guid: read_guid(data, offset),
        n: rd_u32(data, offset + 16).unwrap_or(0),
    }
}

fn read_compact_id(data: &[u8], offset: usize) -> Option<(u32, u32)> {
    let raw = rd_u32(data, offset)?;
    Some((raw >> 8, raw & 0xFF))
}

fn read_jcid(raw: u32) -> Jcid {
    Jcid {
        index: raw & 0xFFFF,
        is_binary: (raw >> 16) & 0x1 != 0,
        is_property_set: (raw >> 17) & 0x1 != 0,
        is_graph_node: (raw >> 18) & 0x1 != 0,
        is_file_data: (raw >> 19) & 0x1 != 0,
        is_read_only: (raw >> 20) & 0x1 != 0,
    }
}

fn decode_chunk_reference(
    data: &[u8],
    offset: usize,
    stp_format: u8,
    cb_format: u8,
) -> (u64, u64, usize) {
    let stp_sz = stp_size(stp_format);
    let cb_sz = cb_size(cb_format);
    let stp_raw = le_uint(data, offset, stp_sz);
    let cb_raw = le_uint(data, offset + stp_sz, cb_sz);
    let stp = if stp_format >= 2 { stp_raw * 8 } else { stp_raw };
    let cb = if cb_format >= 2 { cb_raw * 8 } else { cb_raw };
    (stp, cb, stp_sz + cb_sz)
}

/// Read a little-endian unsigned integer of `size` bytes, clamped to the buffer.
fn le_uint(data: &[u8], offset: usize, size: usize) -> u64 {
    let mut value: u64 = 0;
    for index in 0..size {
        if let Some(byte) = data.get(offset + index) {
            value |= (*byte as u64) << (8 * index);
        }
    }
    value
}

fn merge_property(target: &mut PropMap, name: String, value: PropValue) {
    match target.get_mut(&name) {
        None => {
            target.insert(name, value);
        }
        Some(PropValue::List(list)) => list.push(value),
        Some(slot) => {
            let existing = std::mem::replace(slot, PropValue::Null);
            *slot = PropValue::List(vec![existing, value]);
        }
    }
}

fn resolve_object_id(compact_id: (u32, u32), gid_table: &HashMap<u32, String>) -> Option<ExtendedGuid> {
    let guid = gid_table.get(&compact_id.0)?;
    Some(ExtendedGuid {
        guid: guid.clone(),
        n: compact_id.1,
    })
}

// ---------------------------------------------------------------------------
// Property-set decoding
// ---------------------------------------------------------------------------

fn parse_scalar_from_array(data: &[u8], offset: &mut usize, property_type: u8) -> Option<PropValue> {
    match property_type {
        0x3 => {
            let value = *data.get(*offset)?;
            *offset += 1;
            Some(PropValue::Int(value as u64))
        }
        0x4 => {
            let value = rd_u16(data, *offset)?;
            *offset += 2;
            Some(PropValue::Int(value as u64))
        }
        0x5 => {
            let value = rd_u32(data, *offset)?;
            *offset += 4;
            Some(PropValue::Int(value as u64))
        }
        0x6 => {
            let value = rd_u64(data, *offset)?;
            *offset += 8;
            Some(PropValue::Int(value))
        }
        0x7 => {
            let cb = rd_u32(data, *offset)? as usize;
            *offset += 4;
            let end = (*offset + cb).min(data.len());
            let slice = data[*offset..end].to_vec();
            *offset += cb;
            Some(PropValue::Bytes(slice))
        }
        _ => Some(PropValue::Null),
    }
}

/// Read a single property value of `selected_type`. Returns `None` to signal a
/// failed (out-of-bounds) read that the caller must recover from; a genuine
/// Python `None` value is returned as `Some(PropValue::Null)`.
fn read_value(
    data: &[u8],
    data_offset: &mut usize,
    state: &mut PropertyState,
    raw_property_id: u32,
    selected_type: u8,
) -> Option<PropValue> {
    match selected_type {
        0x2 => Some(PropValue::Bool((raw_property_id >> 31) & 0x1 != 0)),
        0x3 => {
            let value = *data.get(*data_offset)?;
            *data_offset += 1;
            Some(PropValue::Int(value as u64))
        }
        0x4 => {
            let value = rd_u16(data, *data_offset)?;
            *data_offset += 2;
            Some(PropValue::Int(value as u64))
        }
        0x5 => {
            let value = rd_u32(data, *data_offset)?;
            *data_offset += 4;
            Some(PropValue::Int(value as u64))
        }
        0x6 => {
            let value = rd_u64(data, *data_offset)?;
            *data_offset += 8;
            Some(PropValue::Int(value))
        }
        0x7 => {
            let cb = rd_u32(data, *data_offset)? as usize;
            *data_offset += 4;
            let end = (*data_offset + cb).min(data.len());
            let slice = data[*data_offset..end].to_vec();
            *data_offset += cb;
            Some(PropValue::Bytes(slice))
        }
        0x8 => {
            let value = state.object_ids.get(state.oid_cursor).cloned();
            state.oid_cursor += 1;
            Some(value.map_or(PropValue::Null, PropValue::ObjectId))
        }
        0x9 => {
            let count = rd_u32(data, *data_offset)? as usize;
            *data_offset += 4;
            let end = (state.oid_cursor + count).min(state.object_ids.len());
            let items = state.object_ids[state.oid_cursor..end].to_vec();
            state.oid_cursor += count;
            Some(PropValue::ObjectIds(items))
        }
        0x10 => {
            let count = rd_u32(data, *data_offset)? as usize;
            *data_offset += 4;
            let mut items: Vec<PropValue> = Vec::new();
            if count > 0 {
                let array_property_id = rd_u32(data, *data_offset)?;
                *data_offset += 4;
                let array_type =
                    property_type_override(array_property_id).unwrap_or(((array_property_id >> 26) & 0x1F) as u8);
                if array_type == 0x11 {
                    for _ in 0..count {
                        let (child_props, _, consumed) = parse_property_set(data, *data_offset, state);
                        items.push(PropValue::Set(child_props));
                        *data_offset += consumed;
                    }
                } else {
                    for _ in 0..count {
                        let item = parse_scalar_from_array(data, data_offset, array_type)?;
                        items.push(item);
                    }
                }
            }
            Some(PropValue::List(items))
        }
        0x11 => {
            let (child_props, _, consumed) = parse_property_set(data, *data_offset, state);
            *data_offset += consumed;
            Some(PropValue::Set(child_props))
        }
        _ => Some(PropValue::Null),
    }
}

fn parse_property_set(
    data: &[u8],
    mut offset: usize,
    state: &mut PropertyState,
) -> (PropMap, Vec<(String, PropValue)>, usize) {
    let start = offset;
    if offset + 2 > data.len() {
        return (PropMap::new(), Vec::new(), 0);
    }
    let mut property_count = rd_u16(data, offset).unwrap_or(0) as usize;
    offset += 2;
    let available_property_ids = data.len().saturating_sub(offset) / 4;
    if property_count > available_property_ids {
        property_count = available_property_ids;
    }
    let mut property_ids = Vec::with_capacity(property_count);
    for index in 0..property_count {
        property_ids.push(rd_u32(data, offset + index * 4).unwrap_or(0));
    }
    offset += property_count * 4;
    let mut data_offset = offset;
    let mut properties = PropMap::new();
    let mut raw_properties: Vec<(String, PropValue)> = Vec::new();

    for raw_property_id in property_ids {
        let property_id = raw_property_id & 0x03FF_FFFF;
        let default_property_type = ((raw_property_id >> 26) & 0x1F) as u8;
        let property_type = property_type_override(raw_property_id).unwrap_or(default_property_type);
        let name = property_name(property_id);

        let snapshot_offset = data_offset;
        let snapshot_oid_cursor = state.oid_cursor;
        let value = match read_value(data, &mut data_offset, state, raw_property_id, property_type) {
            Some(value) => value,
            None => {
                data_offset = snapshot_offset;
                state.oid_cursor = snapshot_oid_cursor;
                if property_type == default_property_type {
                    break;
                }
                match read_value(data, &mut data_offset, state, raw_property_id, default_property_type) {
                    Some(value) => value,
                    None => {
                        data_offset = snapshot_offset;
                        state.oid_cursor = snapshot_oid_cursor;
                        break;
                    }
                }
            }
        };

        merge_property(&mut properties, name.clone(), value.clone());
        raw_properties.push((name, value));
    }

    (properties, raw_properties, data_offset - start)
}

fn parse_object_blob(blob: &[u8], gid_table: &HashMap<u32, String>) -> (PropMap, Vec<(String, PropValue)>) {
    let header = match rd_u32(blob, 0) {
        Some(value) => value,
        None => return (PropMap::new(), Vec::new()),
    };
    let mut offset = 4usize;
    let oid_count = (header & 0x00FF_FFFF) as usize;
    let no_osid_stream = (header >> 31) & 0x1 != 0;

    let mut object_ids: Vec<ExtendedGuid> = Vec::new();
    for index in 0..oid_count {
        if let Some(compact_id) = read_compact_id(blob, offset + index * 4) {
            if let Some(resolved) = resolve_object_id(compact_id, gid_table) {
                object_ids.push(resolved);
            }
        }
    }
    offset += oid_count * 4;

    if !no_osid_stream {
        if let Some(osid_header) = rd_u32(blob, offset) {
            offset += 4;
            let osid_count = (osid_header & 0x00FF_FFFF) as usize;
            let osid_has_extended_streams = (osid_header >> 30) & 0x1 != 0;
            offset += osid_count * 4;
            if osid_has_extended_streams {
                if let Some(context_header) = rd_u32(blob, offset) {
                    offset += 4;
                    let context_count = (context_header & 0x00FF_FFFF) as usize;
                    offset += context_count * 4;
                }
            }
        }
    }

    let mut state = PropertyState {
        object_ids,
        oid_cursor: 0,
    };
    let (properties, raw_properties, _) = parse_property_set(blob, offset, &mut state);
    (properties, raw_properties)
}

fn decode_storage_string(data: &[u8], offset: usize) -> (String, usize) {
    let char_count = rd_u32(data, offset).unwrap_or(0) as usize;
    let mut offset = offset + 4;
    let mut byte_count = char_count * 2;
    if offset + byte_count > data.len() {
        byte_count = char_count;
    }
    let end = (offset + byte_count).min(data.len());
    let raw = &data[offset.min(data.len())..end];
    offset += byte_count;
    (decode_utf16le(raw), offset)
}

pub(crate) fn decode_utf16le(raw: &[u8]) -> String {
    let units: Vec<u16> = raw
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let decoded: String = char::decode_utf16(units)
        .map(|r| r.unwrap_or('\u{FFFD}'))
        .collect();
    decoded.trim_end_matches('\u{0}').to_string()
}

fn normalize_guid_token(value: &str) -> String {
    value
        .trim()
        .trim_matches(|c| c == '{' || c == '}')
        .to_ascii_lowercase()
}

fn parse_file_data_store_object(blob: &[u8]) -> Vec<u8> {
    if blob.len() < 52 {
        return Vec::new();
    }
    let length = rd_u64(blob, 16).unwrap_or(0) as usize;
    let data_offset = 36usize;
    let end = (data_offset + length).min(blob.len());
    blob[data_offset.min(blob.len())..end].to_vec()
}

// ---------------------------------------------------------------------------
// File-node list parsing
// ---------------------------------------------------------------------------

fn parse_file_node_list(data: &[u8], start_offset: u64, size: u64) -> Result<Vec<FileNode>, AsposeNoteError> {
    let mut result: Vec<FileNode> = Vec::new();
    let mut current_offset = start_offset as usize;
    let mut current_size = size as usize;
    let mut visited: std::collections::HashSet<usize> = std::collections::HashSet::new();

    while current_size != 0 {
        if !visited.insert(current_offset) {
            break;
        }
        let frag_end = current_offset
            .saturating_add(current_size)
            .min(data.len());
        if current_offset >= data.len() {
            break;
        }
        let fragment = &data[current_offset..frag_end];
        if fragment.len() < 24 {
            break;
        }
        if rd_u64(fragment, 0) != Some(FILE_NODE_LIST_MAGIC) {
            return Err(AsposeNoteError::file_corrupted("Invalid FileNodeListFragment magic"));
        }

        let mut offset = 16usize;
        while offset + 4 <= fragment.len().saturating_sub(20) {
            let header = match rd_u32(fragment, offset) {
                Some(value) => value,
                None => break,
            };
            let file_node_id = (header & 0x3FF) as u16;
            let node_size = ((header >> 10) & 0x1FFF) as usize;
            let stp_format = ((header >> 23) & 0x3) as u8;
            let cb_format = ((header >> 25) & 0x3) as u8;
            let base_type = ((header >> 27) & 0xF) as u8;
            if node_size < 4 || offset + node_size > fragment.len().saturating_sub(8) {
                break;
            }
            result.push(FileNode {
                file_node_id,
                payload: fragment[offset + 4..offset + node_size].to_vec(),
                stp_format,
                cb_format,
                base_type,
            });
            offset += node_size;
            if file_node_id == 0x0FF {
                break;
            }
        }

        if rd_u64(fragment, fragment.len() - 8) != Some(FILE_NODE_LIST_FOOTER_MAGIC) {
            break;
        }
        let footer = fragment.len();
        current_offset = le_uint(fragment, footer - 20, 8) as usize;
        current_size = le_uint(fragment, footer - 12, 4) as usize;
    }

    Ok(result)
}

fn parse_object_declaration(
    node: &FileNode,
    gid_table: &HashMap<u32, String>,
    data: &[u8],
) -> Option<OneStoreObject> {
    let payload = &node.payload;

    if matches!(node.file_node_id, 0x0A4 | 0x0A5 | 0x0C4 | 0x0C5) {
        let (blob_offset, blob_size, ref_size) =
            decode_chunk_reference(payload, 0, node.stp_format, node.cb_format);
        let oid = resolve_object_id(read_compact_id(payload, ref_size)?, gid_table)?;
        let jcid = read_jcid(rd_u32(payload, ref_size + 4)?);
        let blob_end = (blob_offset as usize)
            .saturating_add(blob_size as usize)
            .min(data.len());
        let blob = data
            .get((blob_offset as usize).min(data.len())..blob_end)
            .unwrap_or(&[]);
        let (properties, raw_properties) = parse_object_blob(blob, gid_table);
        return Some(OneStoreObject {
            oid,
            jcid: jcid.clone(),
            jcid_name: jcid_name(jcid.index),
            properties,
            raw_properties,
            file_data_reference: None,
            file_extension: None,
        });
    }

    if matches!(node.file_node_id, 0x072 | 0x073) {
        let mut offset = 0usize;
        let oid = resolve_object_id(read_compact_id(payload, offset)?, gid_table)?;
        offset += 4;
        let jcid = read_jcid(rd_u32(payload, offset)?);
        offset += 4;
        offset += if node.file_node_id == 0x073 { 4 } else { 1 };
        let (file_data_reference, next) = decode_storage_string(payload, offset);
        offset = next;
        let (file_extension, _) = decode_storage_string(payload, offset);
        return Some(OneStoreObject {
            oid,
            jcid: jcid.clone(),
            jcid_name: jcid_name(jcid.index),
            properties: PropMap::new(),
            raw_properties: Vec::new(),
            file_data_reference: Some(file_data_reference),
            file_extension: Some(file_extension),
        });
    }

    None
}

fn process_group(
    data: &[u8],
    state: &mut ObjectSpaceState,
    group_ref: (u64, u64),
    gid_table: &HashMap<u32, String>,
) -> Result<(), AsposeNoteError> {
    let mut local_gid_table = gid_table.clone();
    for node in parse_file_node_list(data, group_ref.0, group_ref.1)? {
        match node.file_node_id {
            0x024 => {
                if let Some(index) = rd_u32(&node.payload, 0) {
                    local_gid_table.insert(index, read_guid(&node.payload, 4));
                }
                continue;
            }
            0x025 => {
                if let (Some(source), Some(target)) =
                    (rd_u32(&node.payload, 0), rd_u32(&node.payload, 4))
                {
                    if let Some(guid) = local_gid_table.get(&source).cloned() {
                        local_gid_table.insert(target, guid);
                    }
                }
                continue;
            }
            0x026 => {
                if let (Some(source_start), Some(count), Some(target_start)) = (
                    rd_u32(&node.payload, 0),
                    rd_u32(&node.payload, 4),
                    rd_u32(&node.payload, 8),
                ) {
                    for index in 0..count {
                        let source = source_start + index;
                        if let Some(guid) = local_gid_table.get(&source).cloned() {
                            local_gid_table.insert(target_start + index, guid);
                        }
                    }
                }
                continue;
            }
            0x0A4 | 0x0A5 | 0x0C4 | 0x0C5 | 0x072 | 0x073 => {}
            _ => continue,
        }
        if let Some(parsed) = parse_object_declaration(&node, &local_gid_table, data) {
            if parsed.jcid_name == "SectionMetaData" {
                if let Some(PropValue::Bytes(value)) = parsed.properties.get("SectionDisplayName") {
                    let decoded = decode_utf16le(value);
                    if !decoded.is_empty() {
                        state.section_display_name = Some(decoded);
                    }
                }
            }
            state.objects.insert(parsed.oid.clone(), parsed);
        }
    }
    Ok(())
}

fn process_revision_manifest(
    data: &[u8],
    gosid: ExtendedGuid,
    nodes: &[FileNode],
    order_index: usize,
) -> Result<Option<RevisionManifestRecord>, AsposeNoteError> {
    let mut gid_table: HashMap<u32, String> = HashMap::new();
    let mut temporary_state = ObjectSpaceState::new(gosid);
    let mut revision: Option<RevisionManifestRecord> = None;

    for node in nodes {
        match node.file_node_id {
            0x01B | 0x01E | 0x01F => {
                let rid = read_extended_guid(&node.payload, 0);
                let rid_dependent = if node.payload.len() >= 40 {
                    Some(read_extended_guid(&node.payload, 20))
                } else {
                    None
                };
                let context_id = if node.file_node_id == 0x01F && node.payload.len() >= 60 {
                    Some(read_extended_guid(&node.payload, 40))
                } else {
                    None
                };
                revision = Some(RevisionManifestRecord {
                    rid,
                    rid_dependent,
                    context_id,
                    root_roles: HashMap::new(),
                    objects: HashMap::new(),
                    section_display_name: None,
                    order_index,
                });
            }
            0x024 => {
                if let Some(index) = rd_u32(&node.payload, 0) {
                    gid_table.insert(index, read_guid(&node.payload, 4));
                }
            }
            0x025 => {
                if let (Some(source), Some(target)) =
                    (rd_u32(&node.payload, 0), rd_u32(&node.payload, 4))
                {
                    if let Some(guid) = gid_table.get(&source).cloned() {
                        gid_table.insert(target, guid);
                    }
                }
            }
            0x026 => {
                if let (Some(source_start), Some(count), Some(target_start)) = (
                    rd_u32(&node.payload, 0),
                    rd_u32(&node.payload, 4),
                    rd_u32(&node.payload, 8),
                ) {
                    for index in 0..count {
                        let source = source_start + index;
                        if let Some(guid) = gid_table.get(&source).cloned() {
                            gid_table.insert(target_start + index, guid);
                        }
                    }
                }
            }
            0x05A => {
                if let Some(revision) = revision.as_mut() {
                    let root = read_extended_guid(&node.payload, 0);
                    if let Some(role) = rd_u32(&node.payload, 20) {
                        revision.root_roles.insert(role, root);
                    }
                }
            }
            0x059 => {
                if let Some(revision) = revision.as_mut() {
                    if let Some(compact_id) = read_compact_id(&node.payload, 0) {
                        let root = resolve_object_id(compact_id, &gid_table);
                        if let (Some(root), Some(role)) = (root, rd_u32(&node.payload, 4)) {
                            revision.root_roles.insert(role, root);
                        }
                    }
                }
            }
            0x0B0 => {
                if revision.is_some() {
                    let (group_offset, group_size, _) =
                        decode_chunk_reference(&node.payload, 0, node.stp_format, node.cb_format);
                    process_group(data, &mut temporary_state, (group_offset, group_size), &gid_table)?;
                }
            }
            _ => {}
        }
    }

    let mut revision = match revision {
        Some(revision) => revision,
        None => return Ok(None),
    };
    revision.objects = temporary_state.objects;
    revision.section_display_name = temporary_state.section_display_name;
    Ok(Some(revision))
}

// ---------------------------------------------------------------------------
// Revision-chain resolution
// ---------------------------------------------------------------------------

fn iter_revision_chain<'a>(
    state: &'a ObjectSpaceState,
    revision: &'a RevisionManifestRecord,
) -> Vec<&'a RevisionManifestRecord> {
    let revision_by_rid: HashMap<&ExtendedGuid, &RevisionManifestRecord> =
        state.revisions.iter().map(|item| (&item.rid, item)).collect();
    let mut chain: Vec<&RevisionManifestRecord> = Vec::new();
    let mut seen: std::collections::HashSet<&ExtendedGuid> = std::collections::HashSet::new();
    let mut current: Option<&RevisionManifestRecord> = Some(revision);
    while let Some(item) = current {
        if seen.contains(&item.rid) {
            break;
        }
        seen.insert(&item.rid);
        chain.push(item);
        match &item.rid_dependent {
            Some(dependent) if !dependent.is_zero() => {
                current = revision_by_rid.get(dependent).copied();
            }
            _ => break,
        }
    }
    chain.reverse();
    chain
}

fn effective_revision_roots(
    state: &ObjectSpaceState,
    revision: &RevisionManifestRecord,
) -> HashMap<u32, ExtendedGuid> {
    let mut roots: HashMap<u32, ExtendedGuid> = HashMap::new();
    for item in iter_revision_chain(state, revision) {
        for (role, root) in &item.root_roles {
            roots.insert(*role, root.clone());
        }
    }
    roots
}

fn resolve_revision_root(
    state: &ObjectSpaceState,
    revision: &RevisionManifestRecord,
    role: u32,
) -> Option<ExtendedGuid> {
    effective_revision_roots(state, revision).get(&role).cloned()
}

fn materialize_latest_roots(state: &mut ObjectSpaceState) {
    let mut latest: HashMap<(Option<ExtendedGuid>, u32), ExtendedGuid> = HashMap::new();
    let mut latest_revision_indices: HashMap<(Option<ExtendedGuid>, u32), usize> = HashMap::new();
    let revision_index_by_rid: HashMap<ExtendedGuid, usize> = state
        .revisions
        .iter()
        .enumerate()
        .map(|(index, revision)| (revision.rid.clone(), index))
        .collect();

    enum Event<'a> {
        Revision(&'a RevisionManifestRecord),
        Binding(&'a RevisionRoleBinding),
    }
    let mut ordered_events: Vec<(usize, usize, Event)> = Vec::new();
    // The secondary key preserves insertion order for equal order_index values,
    // matching Python's stable sort over a single appended event list.
    let mut sequence = 0usize;
    for revision in &state.revisions {
        ordered_events.push((revision.order_index, sequence, Event::Revision(revision)));
        sequence += 1;
    }
    for binding in &state.role_bindings {
        ordered_events.push((binding.order_index, sequence, Event::Binding(binding)));
        sequence += 1;
    }
    ordered_events.sort_by_key(|item| (item.0, item.1));

    let revision_by_rid: HashMap<&ExtendedGuid, &RevisionManifestRecord> =
        state.revisions.iter().map(|item| (&item.rid, item)).collect();

    for (_, _, event) in &ordered_events {
        match event {
            Event::Revision(revision) => {
                let revision_index = revision_index_by_rid[&revision.rid];
                for (role, root) in effective_revision_roots(state, revision) {
                    latest.insert((revision.context_id.clone(), role), root);
                    latest_revision_indices.insert((revision.context_id.clone(), role), revision_index);
                }
            }
            Event::Binding(binding) => {
                if let Some(revision) = revision_by_rid.get(&binding.rid).copied() {
                    if let Some(root) = resolve_revision_root(state, revision, binding.role) {
                        latest.insert((binding.context_id.clone(), binding.role), root);
                        if let Some(index) = revision_index_by_rid.get(&binding.rid) {
                            latest_revision_indices
                                .insert((binding.context_id.clone(), binding.role), *index);
                        }
                    }
                }
            }
        }
    }

    state.latest_roots_by_context_role = latest;
    state.latest_revision_indices_by_context_role = latest_revision_indices;
}

fn build_revision_snapshot(
    state: &ObjectSpaceState,
    revision: &RevisionManifestRecord,
) -> HashMap<ExtendedGuid, OneStoreObject> {
    let mut snapshot: HashMap<ExtendedGuid, OneStoreObject> = HashMap::new();
    for item in iter_revision_chain(state, revision) {
        for (oid, object) in &item.objects {
            snapshot.insert(oid.clone(), object.clone());
        }
    }
    snapshot
}

fn process_revision_manifest_list(
    data: &[u8],
    state: &mut ObjectSpaceState,
    nodes: &[FileNode],
) -> Result<(), AsposeNoteError> {
    let mut current_manifest: Option<Vec<FileNode>> = None;
    let mut order_index = 0usize;

    for node in nodes {
        if matches!(node.file_node_id, 0x01B | 0x01E | 0x01F) {
            current_manifest = Some(vec![node.clone()]);
            continue;
        }

        if let Some(manifest) = current_manifest.as_mut() {
            manifest.push(node.clone());
            if node.file_node_id == 0x01C {
                let manifest = current_manifest.take().unwrap();
                if let Some(revision) =
                    process_revision_manifest(data, state.gosid.clone(), &manifest, order_index)?
                {
                    if state.section_display_name.is_none() {
                        if let Some(name) = &revision.section_display_name {
                            state.section_display_name = Some(name.clone());
                        }
                    }
                    state.revisions.push(revision);
                }
                order_index += 1;
            }
            continue;
        }

        if node.file_node_id == 0x05C {
            let rid = read_extended_guid(&node.payload, 0);
            if let Some(role) = rd_u32(&node.payload, 20) {
                state.role_bindings.push(RevisionRoleBinding {
                    rid,
                    context_id: None,
                    role,
                    order_index,
                });
                order_index += 1;
            }
        } else if node.file_node_id == 0x05D {
            let rid = read_extended_guid(&node.payload, 0);
            if let Some(role) = rd_u32(&node.payload, 20) {
                let context_id = read_extended_guid(&node.payload, 24);
                state.role_bindings.push(RevisionRoleBinding {
                    rid,
                    context_id: Some(context_id),
                    role,
                    order_index,
                });
                order_index += 1;
            }
        }
    }

    materialize_latest_roots(state);
    let mut preferred_revision_index = state.get_latest_revision_index(1, None);
    if preferred_revision_index.is_none() && !state.revisions.is_empty() {
        preferred_revision_index = Some(state.revisions.len() - 1);
    }
    if let Some(index) = preferred_revision_index {
        let preferred = state.revisions[index].clone();
        state.objects = build_revision_snapshot(state, &preferred);
    }
    Ok(())
}

fn parse_object_space(
    data: &[u8],
    gosid: ExtendedGuid,
    list_ref: (u64, u64),
) -> Result<ObjectSpaceState, AsposeNoteError> {
    let mut state = ObjectSpaceState::new(gosid);
    let manifest_nodes = parse_file_node_list(data, list_ref.0, list_ref.1)?;
    for node in &manifest_nodes {
        if node.file_node_id != 0x010 {
            continue;
        }
        let (revision_offset, revision_size, _) =
            decode_chunk_reference(&node.payload, 0, node.stp_format, node.cb_format);
        let revision_nodes = parse_file_node_list(data, revision_offset, revision_size)?;
        process_revision_manifest_list(data, &mut state, &revision_nodes)?;
    }
    Ok(state)
}

fn process_file_data_store_list(
    data: &[u8],
    file_data_store: &mut HashMap<String, Vec<u8>>,
    reference: (u64, u64),
) -> Result<(), AsposeNoteError> {
    for node in parse_file_node_list(data, reference.0, reference.1)? {
        if node.file_node_id != 0x094 {
            continue;
        }
        let (blob_offset, blob_size, ref_size) =
            decode_chunk_reference(&node.payload, 0, node.stp_format, node.cb_format);
        let guid_reference = normalize_guid_token(&read_guid(&node.payload, ref_size));
        let blob_end = (blob_offset as usize)
            .saturating_add(blob_size as usize)
            .min(data.len());
        let blob = data
            .get((blob_offset as usize).min(data.len())..blob_end)
            .unwrap_or(&[]);
        file_data_store.insert(guid_reference, parse_file_data_store_object(blob));
    }
    Ok(())
}

/// Source for [`parse_onestore_file`]: a filesystem path or in-memory bytes.
pub enum OneStoreSource<'a> {
    Path(&'a Path),
    Bytes(Vec<u8>),
}

/// Parse an entire `.one` section file.
pub fn parse_onestore_file(source: OneStoreSource<'_>) -> Result<OneStoreFile, AsposeNoteError> {
    let (data, display_name) = match source {
        OneStoreSource::Path(path) => {
            let data = std::fs::read(path)?;
            let display_name = path
                .file_stem()
                .map(|stem| stem.to_string_lossy().into_owned());
            (data, display_name)
        }
        OneStoreSource::Bytes(data) => (data, None),
    };

    if data.len() < 1024 {
        return Err(AsposeNoteError::file_corrupted(
            "File is too small to be a OneNote section",
        ));
    }

    let file_type_guid = read_guid(&data, 0);
    if file_type_guid.to_ascii_lowercase() != SECTION_FILE_TYPE {
        return Err(AsposeNoteError::unsupported_file_format(
            "Unsupported file format",
            Some(&file_type_guid),
        ));
    }

    let root_offset = rd_u64(&data, 172).unwrap_or(0);
    let root_size = rd_u32(&data, 180).unwrap_or(0) as u64;
    let root_nodes = parse_file_node_list(&data, root_offset, root_size)?;

    let mut file_data_store: HashMap<String, Vec<u8>> = HashMap::new();
    let mut object_spaces: Vec<ObjectSpaceState> = Vec::new();

    for node in &root_nodes {
        if node.file_node_id == 0x090 {
            let (list_offset, list_size, _) =
                decode_chunk_reference(&node.payload, 0, node.stp_format, node.cb_format);
            process_file_data_store_list(&data, &mut file_data_store, (list_offset, list_size))?;
        } else if node.file_node_id == 0x008 {
            let (list_offset, list_size, ref_size) =
                decode_chunk_reference(&node.payload, 0, node.stp_format, node.cb_format);
            let gosid = read_extended_guid(&node.payload, ref_size);
            object_spaces.push(parse_object_space(&data, gosid, (list_offset, list_size))?);
        }
    }

    Ok(OneStoreFile {
        data,
        object_spaces,
        file_data_store,
        display_name,
    })
}

/// Convenience wrapper that reads an entire stream and parses it.
pub fn parse_onestore_reader<R: Read>(reader: &mut R) -> Result<OneStoreFile, AsposeNoteError> {
    let mut data = Vec::new();
    reader.read_to_end(&mut data)?;
    parse_onestore_file(OneStoreSource::Bytes(data))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn uint32_property(property_id: u32) -> u32 {
        (0x5 << 26) | property_id
    }

    #[test]
    fn read_guid_zero_pads_truncated_payload() {
        let mut payload = 0u32.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0x11, 0x22, 0x33]);
        let result = read_guid(&payload, 4);
        assert_eq!(result, format_guid_le(&[0x11, 0x22, 0x33]));
        assert_eq!(result, "00332211-0000-0000-0000-000000000000");
    }

    #[test]
    fn read_guid_preserves_full_payload() {
        let guid_bytes: Vec<u8> = (0u8..16).collect();
        let mut payload = 0u32.to_le_bytes().to_vec();
        payload.extend_from_slice(&guid_bytes);
        assert_eq!(read_guid(&payload, 4), format_guid_le(&guid_bytes));
    }

    #[test]
    fn parse_object_blob_handles_truncated_array_count() {
        let mut blob = 0x8000_0000u32.to_le_bytes().to_vec();
        blob.extend_from_slice(&1u16.to_le_bytes());
        blob.extend_from_slice(&0x2400_0001u32.to_le_bytes());
        let (properties, raw_properties) = parse_object_blob(&blob, &HashMap::new());
        assert!(properties.is_empty());
        assert!(raw_properties.is_empty());
    }

    #[test]
    fn parse_object_blob_keeps_properties_before_truncation() {
        let mut blob = 0x8000_0000u32.to_le_bytes().to_vec();
        blob.extend_from_slice(&2u16.to_le_bytes());
        blob.extend_from_slice(&uint32_property(0x1C01).to_le_bytes());
        blob.extend_from_slice(&uint32_property(0x1C02).to_le_bytes());
        blob.extend_from_slice(&42u32.to_le_bytes());
        blob.extend_from_slice(&7u16.to_le_bytes());
        let (properties, raw_properties) = parse_object_blob(&blob, &HashMap::new());
        assert_eq!(properties.get("PageWidth"), Some(&PropValue::Int(42)));
        assert_eq!(properties.len(), 1);
        assert_eq!(
            raw_properties,
            vec![("PageWidth".to_string(), PropValue::Int(42))]
        );
    }

    #[test]
    fn parse_object_blob_clamps_overlong_property_count() {
        let mut blob = 0x8000_0000u32.to_le_bytes().to_vec();
        blob.extend_from_slice(&1000u16.to_le_bytes());
        let (properties, raw_properties) = parse_object_blob(&blob, &HashMap::new());
        assert!(properties.is_empty());
        assert!(raw_properties.is_empty());
    }
}

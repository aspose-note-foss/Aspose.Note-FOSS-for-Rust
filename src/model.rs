//! Public DOM model and the `Document` entry point, mirroring
//! `aspose.note.model`.
//!
//! The DOM is a reference-counted tree: every node is an `Rc<RefCell<NodeInner>>`
//! so that nodes keep upward `parent` links, shared identity, and in-place
//! mutation — matching the reference semantics of the Python classes. Public
//! node types (`Page`, `RichText`, ...) are thin typed handles over that shared
//! node.

use std::cell::{Ref, RefCell, RefMut};
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;
use std::rc::{Rc, Weak};

use crate::enums::{FileFormat, HorizontalAlignment, NodeType, SaveFormat, TagStatus};
use crate::error::AsposeNoteError;
use crate::saving::options::PdfSaveOptions;

// ---------------------------------------------------------------------------
// Date / time
// ---------------------------------------------------------------------------

/// A UTC instant, stored as microseconds since the Unix epoch. This is the Rust
/// stand-in for the timezone-aware `datetime` values the Python model exposes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DateTime {
    micros: i64,
}

impl DateTime {
    pub fn from_unix_micros(micros: i64) -> DateTime {
        DateTime { micros }
    }

    pub fn unix_micros(self) -> i64 {
        self.micros
    }

    /// Build a UTC instant from civil calendar fields.
    pub fn from_ymd_hms_micro(
        year: i64,
        month: i64,
        day: i64,
        hour: i64,
        minute: i64,
        second: i64,
        micro: i64,
    ) -> DateTime {
        let days = days_from_civil(year, month, day);
        let seconds = days * 86_400 + hour * 3_600 + minute * 60 + second;
        DateTime {
            micros: seconds * 1_000_000 + micro,
        }
    }
}

/// Days from 1970-01-01 for a civil date (Howard Hinnant's algorithm).
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let y = if month <= 2 { year - 1 } else { year };
    let era = (if y >= 0 { y } else { y - 399 }) / 400;
    let yoe = y - era * 400;
    let mp = (month + 9) % 12;
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// Convert an MS-ONE 32-bit "time32" value (seconds since 1980-01-01) to a
/// [`DateTime`]. Mirrors `_time32_to_datetime`.
pub(crate) fn time32_to_datetime(value: Option<i64>) -> Option<DateTime> {
    let value = value?;
    let base = DateTime::from_ymd_hms_micro(1980, 1, 1, 0, 0, 0, 0).micros;
    Some(DateTime::from_unix_micros(base + value * 1_000_000))
}

/// Convert a Windows FILETIME (100 ns units since 1601-01-01) to a [`DateTime`].
/// Mirrors `_filetime_to_datetime`.
pub(crate) fn filetime_to_datetime(value: Option<i64>) -> Option<DateTime> {
    let value = value?;
    let base = DateTime::from_ymd_hms_micro(1601, 1, 1, 0, 0, 0, 0).micros;
    Some(DateTime::from_unix_micros(base + value / 10))
}

/// Coerce a raw MS-ONE note-tag timestamp integer into a [`DateTime`]. Mirrors
/// `_coerce_note_tag_datetime` for the integer case.
pub(crate) fn coerce_note_tag_datetime(value: i64) -> Option<DateTime> {
    if value.unsigned_abs() > 0xFFFF_FFFF {
        filetime_to_datetime(Some(value))
    } else {
        time32_to_datetime(Some(value))
    }
}

// ---------------------------------------------------------------------------
// Visitor
// ---------------------------------------------------------------------------

/// Base visitor for traversing the DOM via `visit_*_start` / `visit_*_end`
/// callbacks. Mirrors `DocumentVisitor`.
pub trait DocumentVisitor {
    fn visit_document_start(&mut self, _document: &Document) {}
    fn visit_document_end(&mut self, _document: &Document) {}
    fn visit_page_start(&mut self, _page: &Page) {}
    fn visit_page_end(&mut self, _page: &Page) {}
    fn visit_title_start(&mut self, _title: &Title) {}
    fn visit_title_end(&mut self, _title: &Title) {}
    fn visit_outline_start(&mut self, _outline: &Outline) {}
    fn visit_outline_end(&mut self, _outline: &Outline) {}
    fn visit_outline_element_start(&mut self, _outline_element: &OutlineElement) {}
    fn visit_outline_element_end(&mut self, _outline_element: &OutlineElement) {}
    fn visit_rich_text_start(&mut self, _rich_text: &RichText) {}
    fn visit_rich_text_end(&mut self, _rich_text: &RichText) {}
    fn visit_image_start(&mut self, _image: &Image) {}
    fn visit_image_end(&mut self, _image: &Image) {}
}

// ---------------------------------------------------------------------------
// Licensing / load options (compatibility surface)
// ---------------------------------------------------------------------------

/// Compatibility surface for the commercial licensing API; a no-op here.
#[derive(Debug, Default, Clone)]
pub struct License;

impl License {
    pub fn new() -> License {
        License
    }

    pub fn set_license(&self, _license_path_or_stream: &str) {}
}

/// Compatibility surface for the commercial metered-licensing API; a no-op here.
#[derive(Debug, Default, Clone)]
pub struct Metered;

impl Metered {
    pub fn new() -> Metered {
        Metered
    }

    pub fn set_metered_key(&self, _public_key: &str, _private_key: &str) {}
}

/// Options passed to [`Document::open_with_options`]. Mirrors `LoadOptions`.
#[derive(Debug, Default, Clone)]
pub struct LoadOptions {
    pub document_password: Option<String>,
    pub load_history: bool,
}

impl LoadOptions {
    pub fn new() -> LoadOptions {
        LoadOptions::default()
    }
}

// ---------------------------------------------------------------------------
// Character / paragraph styles
// ---------------------------------------------------------------------------

/// Character-level formatting. Mirrors `TextStyle`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TextStyle {
    pub is_hyperlink: bool,
    pub hyperlink_address: Option<String>,
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<i64>,
    pub highlight: Option<i64>,
    pub language: Option<i64>,
    pub is_bold: bool,
    pub is_italic: bool,
    pub is_underline: bool,
    pub is_strikethrough: bool,
    pub is_superscript: bool,
    pub is_subscript: bool,
    pub is_hidden: bool,
    pub is_math_formatting: bool,
}

impl TextStyle {
    pub fn new() -> TextStyle {
        TextStyle::default()
    }

    /// The .NET `FontStyle` bit field (bold=1, italic=2, underline=4, strike=8).
    pub fn font_style(&self) -> i32 {
        let mut style = 0;
        if self.is_bold {
            style |= 1;
        }
        if self.is_italic {
            style |= 2;
        }
        if self.is_underline {
            style |= 4;
        }
        if self.is_strikethrough {
            style |= 8;
        }
        style
    }

    /// Mirrors `TextStyle.Default` (note: language is `1033`, unlike `new()`).
    pub fn default_style() -> TextStyle {
        TextStyle {
            language: Some(1033),
            ..TextStyle::default()
        }
    }

    pub fn default_ms_one_note_title_text_style() -> TextStyle {
        TextStyle {
            font_size: Some(20.0),
            is_bold: true,
            language: Some(1033),
            ..TextStyle::default()
        }
    }

    pub fn default_ms_one_note_title_date_style() -> TextStyle {
        TextStyle {
            font_size: Some(11.0),
            language: Some(1033),
            ..TextStyle::default()
        }
    }

    pub fn default_ms_one_note_title_time_style() -> TextStyle {
        TextStyle {
            font_size: Some(11.0),
            language: Some(1033),
            ..TextStyle::default()
        }
    }
}

/// Default paragraph-level text formatting. Mirrors `ParagraphStyle`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphStyle {
    pub font_name: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<i64>,
    pub highlight: Option<i64>,
    pub is_bold: bool,
    pub is_italic: bool,
    pub is_underline: bool,
    pub is_strikethrough: bool,
    pub is_superscript: bool,
    pub is_subscript: bool,
}

impl ParagraphStyle {
    pub fn new() -> ParagraphStyle {
        ParagraphStyle::default()
    }

    pub fn font_style(&self) -> i32 {
        let mut style = 0;
        if self.is_bold {
            style |= 1;
        }
        if self.is_italic {
            style |= 2;
        }
        if self.is_underline {
            style |= 4;
        }
        if self.is_strikethrough {
            style |= 8;
        }
        style
    }

    /// Mirrors `ParagraphStyle.Default`.
    pub fn default_style() -> ParagraphStyle {
        ParagraphStyle::default()
    }
}

fn text_style_from_paragraph_style(style: &ParagraphStyle) -> TextStyle {
    TextStyle {
        font_name: style.font_name.clone(),
        font_size: style.font_size,
        font_color: style.font_color,
        highlight: style.highlight,
        is_bold: style.is_bold,
        is_italic: style.is_italic,
        is_underline: style.is_underline,
        is_strikethrough: style.is_strikethrough,
        is_superscript: style.is_superscript,
        is_subscript: style.is_subscript,
        ..TextStyle::default()
    }
}

/// A single formatted run within [`RichText`]. Mirrors `TextRun`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TextRun {
    pub text: String,
    pub style: TextStyle,
}

impl TextRun {
    pub fn new(text: impl Into<String>) -> TextRun {
        TextRun {
            text: text.into(),
            style: TextStyle::default(),
        }
    }

    pub fn with_style(text: impl Into<String>, style: TextStyle) -> TextRun {
        TextRun {
            text: text.into(),
            style,
        }
    }
}

/// Numbered-list formatting applied to an [`OutlineElement`]. Mirrors `NumberList`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct NumberList {
    pub format: Option<String>,
    pub number_format: Option<String>,
    pub font: Option<String>,
    pub font_size: Option<f64>,
    pub font_color: Option<i64>,
    pub is_bold: bool,
    pub is_italic: bool,
    pub last_modified_time: Option<DateTime>,
    pub restart: Option<i64>,
}

impl NumberList {
    pub fn get_numbered_list_header(&self, number: i64) -> String {
        match &self.format {
            None => String::new(),
            Some(format) if format.contains("{0}") => format.replace("{0}", &number.to_string()),
            Some(format) => format.clone(),
        }
    }
}

/// A table column's width and lock state. Mirrors `TableColumn`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TableColumn {
    pub width: Option<f64>,
    pub locked_width: bool,
}

impl TableColumn {
    pub fn new(width: Option<f64>) -> TableColumn {
        TableColumn {
            width,
            locked_width: false,
        }
    }
}

// ---------------------------------------------------------------------------
// Note tags
// ---------------------------------------------------------------------------

/// A OneNote tag (yellow star, question mark, musical note, or custom) attached
/// to text, images, or tables. Mirrors `NoteTag`; use the `create_*` factories.
#[derive(Debug, Clone, PartialEq)]
pub struct NoteTag {
    label: Option<String>,
    icon: Option<i64>,
    status: Option<TagStatus>,
    highlight: Option<i64>,
    creation_time: Option<DateTime>,
    completed_time: Option<DateTime>,
    font_color: Option<i64>,
}

impl NoteTag {
    pub(crate) fn create_internal(
        label: Option<String>,
        icon: Option<i64>,
        highlight: Option<i64>,
        creation_time: Option<DateTime>,
        completed_time: Option<DateTime>,
        font_color: Option<i64>,
    ) -> NoteTag {
        NoteTag {
            label,
            icon,
            status: None,
            highlight,
            creation_time,
            completed_time,
            font_color,
        }
    }

    pub fn create_yellow_star(label: Option<&str>) -> NoteTag {
        NoteTag::create_internal(
            Some(label.unwrap_or("Yellow Star").to_string()),
            Some(13),
            None,
            None,
            None,
            None,
        )
    }

    pub fn create_question_mark(label: Option<&str>) -> NoteTag {
        NoteTag::create_internal(
            Some(label.unwrap_or("Question Mark").to_string()),
            Some(15),
            None,
            None,
            None,
            None,
        )
    }

    pub fn create_musical_note(label: Option<&str>) -> NoteTag {
        NoteTag::create_internal(
            Some(label.unwrap_or("Musical Note").to_string()),
            Some(121),
            None,
            None,
            None,
            None,
        )
    }

    pub fn label(&self) -> Option<String> {
        self.label.clone()
    }

    pub fn set_label(&mut self, value: Option<String>) {
        self.label = value;
    }

    pub fn icon(&self) -> Option<i64> {
        self.icon
    }

    pub fn set_icon(&mut self, value: Option<i64>) {
        self.icon = value;
    }

    /// Mirrors `NoteTag.Status`: the explicit status, else `Completed` when a
    /// completion time is present, else `Open`.
    pub fn status(&self) -> TagStatus {
        if let Some(status) = self.status {
            return status;
        }
        if self.completed_time.is_some() {
            TagStatus::Completed
        } else {
            TagStatus::Open
        }
    }

    pub fn highlight(&self) -> Option<i64> {
        self.highlight
    }

    pub fn set_highlight(&mut self, value: Option<i64>) {
        self.highlight = value;
    }

    pub fn creation_time(&self) -> Option<DateTime> {
        self.creation_time
    }

    pub fn set_creation_time(&mut self, value: Option<DateTime>) {
        self.creation_time = value;
    }

    pub fn completed_time(&self) -> Option<DateTime> {
        self.completed_time
    }

    pub fn font_color(&self) -> Option<i64> {
        self.font_color
    }

    pub fn set_font_color(&mut self, value: Option<i64>) {
        self.font_color = value;
    }
}

// ---------------------------------------------------------------------------
// Node tree infrastructure
// ---------------------------------------------------------------------------

type Link = Rc<RefCell<NodeInner>>;
type WeakLink = Weak<RefCell<NodeInner>>;

struct NodeInner {
    parent: WeakLink,
    children: Vec<Node>,
    data: NodeData,
}

enum NodeData {
    Document(DocumentData),
    Page(PageData),
    Title(TitleData),
    Outline(OutlineData),
    OutlineElement(OutlineElementData),
    RichText(RichTextData),
    Image(ImageData),
    AttachedFile(AttachedFileData),
    Table(TableData),
    TableRow(TableRowData),
    TableCell(TableCellData),
}

#[derive(Debug, Clone, Default)]
struct DocumentData {
    display_name: Option<String>,
    creation_time: Option<DateTime>,
    page_histories: HashMap<usize, PageHistory>,
}

#[derive(Debug, Clone, Default)]
struct PageData {
    author: Option<String>,
    creation_time: Option<DateTime>,
    last_modified_time: Option<DateTime>,
    level: Option<i64>,
    is_conflict_page: bool,
}

#[derive(Debug, Clone, Default)]
struct TitleData {
    text: Option<Node>,
    date: Option<Node>,
    time: Option<Node>,
}

#[derive(Debug, Clone, Default)]
struct OutlineData {
    horizontal_offset: Option<f64>,
    vertical_offset: Option<f64>,
    max_width: Option<f64>,
    max_height: Option<f64>,
    min_width: Option<f64>,
    reserved_width: Option<f64>,
    indent_position: Option<f64>,
    descendants_cannot_be_moved: bool,
    last_modified_time: Option<DateTime>,
}

#[derive(Debug, Clone, Default)]
struct OutlineElementData {
    number_list: Option<NumberList>,
}

#[derive(Debug, Clone, Default)]
struct RichTextData {
    text: String,
    text_runs: Vec<TextRun>,
    paragraph_style: ParagraphStyle,
    alignment: Option<HorizontalAlignment>,
    default_text_style: TextStyle,
    tags: Vec<NoteTag>,
    last_modified_time: Option<DateTime>,
    space_before: Option<f64>,
    space_after: Option<f64>,
    line_spacing: Option<f64>,
}

#[derive(Debug, Clone, Default)]
struct ImageData {
    file_name: Option<String>,
    file_path: Option<String>,
    format: Option<String>,
    bytes: Vec<u8>,
    width: Option<f64>,
    height: Option<f64>,
    original_width: Option<f64>,
    original_height: Option<f64>,
    horizontal_offset: Option<f64>,
    vertical_offset: Option<f64>,
    alignment: Option<HorizontalAlignment>,
    is_background: bool,
    last_modified_time: Option<DateTime>,
    alternative_text_title: Option<String>,
    alternative_text_description: Option<String>,
    hyperlink_url: Option<String>,
    tags: Vec<NoteTag>,
}

#[derive(Debug, Clone, Default)]
struct AttachedFileData {
    file_name: Option<String>,
    bytes: Vec<u8>,
    tags: Vec<NoteTag>,
}

#[derive(Debug, Clone)]
struct TableData {
    tags: Vec<NoteTag>,
    columns: Vec<TableColumn>,
    is_borders_visible: bool,
    last_modified_time: Option<DateTime>,
}

impl Default for TableData {
    fn default() -> Self {
        TableData {
            tags: Vec::new(),
            columns: Vec::new(),
            is_borders_visible: true,
            last_modified_time: None,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct TableRowData;

#[derive(Debug, Clone, Default)]
struct TableCellData;

fn new_node(data: NodeData) -> Link {
    Rc::new(RefCell::new(NodeInner {
        parent: Weak::new(),
        children: Vec::new(),
        data,
    }))
}

/// A generic handle to any DOM node. Mirrors `Node`.
#[derive(Clone)]
pub struct Node(Link);

impl Node {
    /// True when both handles refer to the same underlying node (Python `is`).
    pub fn ptr_eq(a: &Node, b: &Node) -> bool {
        Rc::ptr_eq(&a.0, &b.0)
    }

    pub fn parent_node(&self) -> Option<Node> {
        node_parent(&self.0)
    }

    pub fn document(&self) -> Option<Document> {
        node_document(&self.0)
    }

    pub fn accept<V: DocumentVisitor>(&self, visitor: &mut V) {
        accept_node(self, visitor);
    }

    pub fn node_type(&self) -> Option<NodeType> {
        match &self.0.borrow().data {
            NodeData::Document(_) => Some(NodeType::Document),
            NodeData::Page(_) => Some(NodeType::Page),
            NodeData::Outline(_) => Some(NodeType::Outline),
            NodeData::OutlineElement(_) => Some(NodeType::OutlineElement),
            NodeData::RichText(_) => Some(NodeType::RichText),
            NodeData::Image(_) => Some(NodeType::Image),
            NodeData::Table(_) => Some(NodeType::Table),
            NodeData::AttachedFile(_) => Some(NodeType::AttachedFile),
            _ => None,
        }
    }
}

impl PartialEq for Node {
    fn eq(&self, other: &Node) -> bool {
        Rc::ptr_eq(&self.0, &other.0)
    }
}

impl std::fmt::Debug for Node {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Node({:?})", self.node_type())
    }
}

/// Implemented by every typed node handle, enabling `get_child_nodes::<T>()`.
pub trait TypedNode: Sized {
    fn from_node(node: &Node) -> Option<Self>;
}

fn node_parent(link: &Link) -> Option<Node> {
    link.borrow().parent.upgrade().map(Node)
}

fn node_document(link: &Link) -> Option<Document> {
    let mut current = Some(link.clone());
    while let Some(l) = current {
        if matches!(l.borrow().data, NodeData::Document(..)) {
            return Some(Document(l));
        }
        let parent = l.borrow().parent.upgrade();
        current = parent;
    }
    None
}

/// The logical children used for traversal: `children` for composites, the
/// three rich-text slots for a `Title`, and nothing for leaf nodes. Mirrors
/// `_iter_node_children`.
fn logical_children(node: &Node) -> Vec<Node> {
    let inner = node.0.borrow();
    match &inner.data {
        NodeData::Title(title) => {
            let mut result = Vec::new();
            for child in [&title.text, &title.date, &title.time].into_iter().flatten() {
                result.push(child.clone());
            }
            result
        }
        NodeData::RichText(_) | NodeData::AttachedFile(_) => Vec::new(),
        _ => inner.children.clone(),
    }
}

fn collect_descendants<T: TypedNode>(node: &Node, out: &mut Vec<T>) {
    for child in logical_children(node) {
        if let Some(found) = T::from_node(&child) {
            out.push(found);
        }
        collect_descendants::<T>(&child, out);
    }
}

fn comp_append_last(link: &Link, child: Node) -> Node {
    child.0.borrow_mut().parent = Rc::downgrade(link);
    link.borrow_mut().children.push(child.clone());
    child
}

fn comp_append_first(link: &Link, child: Node) -> Node {
    child.0.borrow_mut().parent = Rc::downgrade(link);
    link.borrow_mut().children.insert(0, child.clone());
    child
}

/// Attach `child` as the last child of `parent`. Used by the loader to build
/// the DOM for arbitrary composite parents.
pub(crate) fn push_child(parent: &Node, child: Node) {
    comp_append_last(&parent.0, child);
}

fn comp_insert(link: &Link, index: usize, child: Node) -> Node {
    child.0.borrow_mut().parent = Rc::downgrade(link);
    link.borrow_mut().children.insert(index, child.clone());
    child
}

fn comp_remove(link: &Link, child: &Node) {
    let removed = {
        let mut inner = link.borrow_mut();
        if let Some(pos) = inner.children.iter().position(|c| Rc::ptr_eq(&c.0, &child.0)) {
            inner.children.remove(pos);
            true
        } else {
            false
        }
    };
    if removed {
        child.0.borrow_mut().parent = Weak::new();
    }
}

/// Recursively duplicate a node subtree, rebinding parents. Mirrors
/// `deepcopy` + `_rebind_tree_parents` used by `Page.Clone`.
fn deep_clone_node(node: &Node, parent: Option<&Link>) -> Node {
    let (data_clone, child_links, title_slots) = {
        let inner = node.0.borrow();
        let title_slots = if let NodeData::Title(title) = &inner.data {
            Some((title.text.clone(), title.date.clone(), title.time.clone()))
        } else {
            None
        };
        (clone_scalar_data(&inner.data), inner.children.clone(), title_slots)
    };

    let new_link = new_node(data_clone);
    new_link.borrow_mut().parent = parent.map_or(Weak::new(), Rc::downgrade);

    let mut new_children = Vec::with_capacity(child_links.len());
    for child in &child_links {
        new_children.push(deep_clone_node(child, Some(&new_link)));
    }
    new_link.borrow_mut().children = new_children;

    if let Some((text, date, time)) = title_slots {
        let text = text.as_ref().map(|n| deep_clone_node(n, Some(&new_link)));
        let date = date.as_ref().map(|n| deep_clone_node(n, Some(&new_link)));
        let time = time.as_ref().map(|n| deep_clone_node(n, Some(&new_link)));
        if let NodeData::Title(title) = &mut new_link.borrow_mut().data {
            title.text = text;
            title.date = date;
            title.time = time;
        }
    }

    Node(new_link)
}

fn clone_scalar_data(data: &NodeData) -> NodeData {
    match data {
        NodeData::Document(d) => NodeData::Document(d.clone()),
        NodeData::Page(d) => NodeData::Page(d.clone()),
        NodeData::Title(_) => NodeData::Title(TitleData::default()),
        NodeData::Outline(d) => NodeData::Outline(d.clone()),
        NodeData::OutlineElement(d) => NodeData::OutlineElement(d.clone()),
        NodeData::RichText(d) => NodeData::RichText(d.clone()),
        NodeData::Image(d) => NodeData::Image(d.clone()),
        NodeData::AttachedFile(d) => NodeData::AttachedFile(d.clone()),
        NodeData::Table(d) => NodeData::Table(d.clone()),
        NodeData::TableRow(d) => NodeData::TableRow(d.clone()),
        NodeData::TableCell(d) => NodeData::TableCell(d.clone()),
    }
}

fn accept_node<V: DocumentVisitor>(node: &Node, visitor: &mut V) {
    let kind = {
        let inner = node.0.borrow();
        std::mem::discriminant(&inner.data);
        match &inner.data {
            NodeData::Document(_) => 0u8,
            NodeData::Page(_) => 1,
            NodeData::Title(_) => 2,
            NodeData::Outline(_) => 3,
            NodeData::OutlineElement(_) => 4,
            NodeData::RichText(_) => 5,
            NodeData::Image(_) => 6,
            _ => 255,
        }
    };
    match kind {
        0 => {
            let handle = Document(node.0.clone());
            visitor.visit_document_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_document_end(&handle);
        }
        1 => {
            let handle = Page(node.0.clone());
            visitor.visit_page_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_page_end(&handle);
        }
        2 => {
            let handle = Title(node.0.clone());
            visitor.visit_title_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_title_end(&handle);
        }
        3 => {
            let handle = Outline(node.0.clone());
            visitor.visit_outline_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_outline_end(&handle);
        }
        4 => {
            let handle = OutlineElement(node.0.clone());
            visitor.visit_outline_element_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_outline_element_end(&handle);
        }
        5 => {
            let handle = RichText(node.0.clone());
            visitor.visit_rich_text_start(&handle);
            visitor.visit_rich_text_end(&handle);
        }
        6 => {
            let handle = Image(node.0.clone());
            visitor.visit_image_start(&handle);
            for child in logical_children(node) {
                accept_node(&child, visitor);
            }
            visitor.visit_image_end(&handle);
        }
        _ => {}
    }
}

macro_rules! typed_handle {
    ($name:ident, $variant:ident) => {
        /// Typed handle over a shared DOM node.
        #[derive(Clone)]
        pub struct $name(Link);

        impl $name {
            pub fn as_node(&self) -> Node {
                Node(self.0.clone())
            }

            pub fn parent_node(&self) -> Option<Node> {
                node_parent(&self.0)
            }

            pub fn document(&self) -> Option<Document> {
                node_document(&self.0)
            }

            pub fn accept<V: DocumentVisitor>(&self, visitor: &mut V) {
                accept_node(&self.as_node(), visitor);
            }
        }

        impl TypedNode for $name {
            fn from_node(node: &Node) -> Option<$name> {
                if matches!(node.0.borrow().data, NodeData::$variant(..)) {
                    Some($name(node.0.clone()))
                } else {
                    None
                }
            }
        }

        impl PartialEq for $name {
            fn eq(&self, other: &$name) -> bool {
                Rc::ptr_eq(&self.0, &other.0)
            }
        }

        impl std::fmt::Debug for $name {
            fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                write!(f, "{}({:p})", stringify!($name), Rc::as_ptr(&self.0))
            }
        }
    };
}

macro_rules! composite_methods {
    ($name:ident) => {
        impl $name {
            pub fn first_child(&self) -> Option<Node> {
                self.0.borrow().children.first().cloned()
            }

            pub fn last_child(&self) -> Option<Node> {
                self.0.borrow().children.last().cloned()
            }

            pub fn child_nodes(&self) -> Vec<Node> {
                self.0.borrow().children.clone()
            }

            pub fn append_child_last(&self, node: Node) -> Node {
                comp_append_last(&self.0, node)
            }

            pub fn append_child_first(&self, node: Node) -> Node {
                comp_append_first(&self.0, node)
            }

            pub fn insert_child(&self, index: usize, node: Node) -> Node {
                comp_insert(&self.0, index, node)
            }

            pub fn remove_child(&self, node: &Node) {
                comp_remove(&self.0, node);
            }

            pub fn get_child_nodes<T: TypedNode>(&self) -> Vec<T> {
                let mut out = Vec::new();
                collect_descendants::<T>(&self.as_node(), &mut out);
                out
            }
        }
    };
}

// ---------------------------------------------------------------------------
// RichText
// ---------------------------------------------------------------------------

typed_handle!(RichText, RichText);

fn chars_of(text: &str) -> Vec<char> {
    text.chars().collect()
}

impl RichText {
    fn data(&self) -> Ref<'_, RichTextData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::RichText(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, RichTextData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::RichText(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> RichText {
        RichText::build(String::new(), None, ParagraphStyle::default(), None, Vec::new())
    }

    pub fn from_text(text: impl Into<String>) -> RichText {
        RichText::build(text.into(), None, ParagraphStyle::default(), None, Vec::new())
    }

    /// General constructor mirroring the Python `RichText(...)` keyword form.
    pub fn build(
        text: String,
        text_runs: Option<Vec<TextRun>>,
        paragraph_style: ParagraphStyle,
        alignment: Option<HorizontalAlignment>,
        tags: Vec<NoteTag>,
    ) -> RichText {
        let mut default_text_style = text_style_from_paragraph_style(&paragraph_style);
        let mut runs = text_runs.clone().unwrap_or_default();
        if let Some(first) = runs.first() {
            default_text_style = first.style.clone();
        }
        let mut text = if text_runs.is_none() || !text.is_empty() {
            text
        } else {
            runs.iter().map(|run| run.text.clone()).collect()
        };
        if runs.is_empty() && !text.is_empty() {
            runs = vec![TextRun::with_style(text.clone(), default_text_style.clone())];
        }
        let _ = &mut text;
        let data = RichTextData {
            text,
            text_runs: runs,
            paragraph_style,
            alignment,
            default_text_style,
            tags,
            last_modified_time: None,
            space_before: None,
            space_after: None,
            line_spacing: None,
        };
        RichText(new_node(NodeData::RichText(data)))
    }

    pub(crate) fn replace_text_runs(&self, runs: Vec<TextRun>) {
        let mut data = self.data_mut();
        data.text_runs = runs;
        if let Some(first) = data.text_runs.first() {
            data.default_text_style = first.style.clone();
        }
        data.text = data.text_runs.iter().map(|run| run.text.clone()).collect();
    }

    pub fn text(&self) -> String {
        let data = self.data();
        if data.text_runs.is_empty() {
            data.text.clone()
        } else {
            data.text_runs.iter().map(|run| run.text.as_str()).collect()
        }
    }

    pub fn set_text(&self, value: impl Into<String>) {
        let value = value.into();
        let mut data = self.data_mut();
        data.text = value;
        if !data.text.is_empty() {
            let style = data
                .text_runs
                .first()
                .map(|run| run.style.clone())
                .unwrap_or_else(|| data.default_text_style.clone());
            data.default_text_style = style.clone();
            let text = data.text.clone();
            data.text_runs = vec![TextRun::with_style(text, style)];
        } else {
            data.text_runs.clear();
        }
    }

    pub fn text_runs(&self) -> Vec<TextRun> {
        self.data().text_runs.clone()
    }

    pub fn paragraph_style(&self) -> ParagraphStyle {
        self.data().paragraph_style.clone()
    }

    pub fn set_paragraph_style(&self, value: ParagraphStyle) {
        let mut data = self.data_mut();
        data.default_text_style = text_style_from_paragraph_style(&value);
        data.paragraph_style = value;
    }

    pub fn tags(&self) -> Vec<NoteTag> {
        self.data().tags.clone()
    }

    pub fn tags_mut(&self) -> RefMut<'_, Vec<NoteTag>> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::RichText(data) => &mut data.tags,
            _ => unreachable!(),
        })
    }

    pub fn length(&self) -> usize {
        self.text().chars().count()
    }

    pub fn alignment(&self) -> Option<HorizontalAlignment> {
        self.data().alignment
    }

    pub fn set_alignment(&self, value: Option<HorizontalAlignment>) {
        self.data_mut().alignment = value;
    }

    pub fn last_modified_time(&self) -> Option<DateTime> {
        self.data().last_modified_time
    }

    pub fn space_before(&self) -> Option<f64> {
        self.data().space_before
    }

    pub fn space_after(&self) -> Option<f64> {
        self.data().space_after
    }

    pub fn line_spacing(&self) -> Option<f64> {
        self.data().line_spacing
    }

    pub fn is_title_text(&self) -> bool {
        self.is_title_slot(|title| title.title_text())
    }

    pub fn is_title_date(&self) -> bool {
        self.is_title_slot(|title| title.title_date())
    }

    pub fn is_title_time(&self) -> bool {
        self.is_title_slot(|title| title.title_time())
    }

    fn is_title_slot(&self, slot: impl Fn(&Title) -> Option<RichText>) -> bool {
        if let Some(parent) = self.parent_node() {
            if let Some(title) = Title::from_node(&parent) {
                if let Some(target) = slot(&title) {
                    return Rc::ptr_eq(&target.0, &self.0);
                }
            }
        }
        false
    }

    pub fn append(&self, text: &str, style: Option<TextStyle>) -> RichText {
        let effective = match style {
            Some(style) => style,
            None => self.data().default_text_style.clone(),
        };
        {
            let mut data = self.data_mut();
            data.default_text_style = effective.clone();
            if !data.text_runs.is_empty() {
                data.text_runs.push(TextRun::with_style(text.to_string(), effective));
                data.text = data.text_runs.iter().map(|run| run.text.clone()).collect();
            } else {
                let combined = format!("{}{}", data.text, text);
                data.text = combined.clone();
                data.text_runs = if combined.is_empty() {
                    Vec::new()
                } else {
                    vec![TextRun::with_style(combined, effective)]
                };
            }
        }
        self.clone()
    }

    pub fn append_front(&self, text: &str, style: Option<TextStyle>) -> RichText {
        self.insert(0, text, style)
    }

    pub fn clear(&self) -> RichText {
        let mut data = self.data_mut();
        data.text.clear();
        data.text_runs.clear();
        drop(data);
        self.clone()
    }

    pub fn insert(&self, index: usize, text: &str, style: Option<TextStyle>) -> RichText {
        let current = self.text();
        let current_chars = chars_of(&current);
        let index = index.min(current_chars.len());
        let mut new_text: String = current_chars[..index].iter().collect();
        new_text.push_str(text);
        new_text.extend(&current_chars[index..]);
        let mut data = self.data_mut();
        let effective = style.unwrap_or_else(|| {
            data.text_runs
                .first()
                .map(|run| run.style.clone())
                .unwrap_or_else(|| data.default_text_style.clone())
        });
        data.default_text_style = effective.clone();
        data.text = new_text.clone();
        data.text_runs = if new_text.is_empty() {
            Vec::new()
        } else {
            vec![TextRun::with_style(new_text, effective)]
        };
        drop(data);
        self.clone()
    }

    pub fn remove(&self, start: usize, count: Option<usize>) -> RichText {
        let current = self.text();
        let current_chars = chars_of(&current);
        let start = start.min(current_chars.len());
        let end = match count {
            None => current_chars.len(),
            Some(count) => (start + count).min(current_chars.len()),
        };
        let mut new_text: String = current_chars[..start].iter().collect();
        new_text.extend(&current_chars[end..]);
        let mut data = self.data_mut();
        if !new_text.is_empty() {
            let style = data
                .text_runs
                .first()
                .map(|run| run.style.clone())
                .unwrap_or_else(|| data.default_text_style.clone());
            data.default_text_style = style.clone();
            data.text = new_text.clone();
            data.text_runs = vec![TextRun::with_style(new_text, style)];
        } else {
            data.text.clear();
            data.text_runs.clear();
        }
        drop(data);
        self.clone()
    }

    pub fn replace(&self, old_value: &str, new_value: &str, style: Option<TextStyle>) -> RichText {
        if old_value.is_empty() {
            return self.clone();
        }
        let replaced = self.text().replace(old_value, new_value);
        let mut data = self.data_mut();
        if !replaced.is_empty() {
            let effective = style.unwrap_or_else(|| {
                data.text_runs
                    .first()
                    .map(|run| run.style.clone())
                    .unwrap_or_else(|| data.default_text_style.clone())
            });
            data.default_text_style = effective.clone();
            data.text = replaced.clone();
            data.text_runs = vec![TextRun::with_style(replaced, effective)];
        } else {
            data.text.clear();
            data.text_runs.clear();
        }
        drop(data);
        self.clone()
    }

    pub fn trim(&self) -> RichText {
        let current = self.text();
        let trimmed = current.trim();
        if trimmed.len() != current.len() {
            self.set_text(trimmed.to_string());
        }
        self.clone()
    }

    pub fn trim_start(&self) -> RichText {
        let current = self.text();
        let trimmed = current.trim_start();
        if trimmed.len() != current.len() {
            self.set_text(trimmed.to_string());
        }
        self.clone()
    }

    pub fn trim_end(&self) -> RichText {
        let current = self.text();
        let trimmed = current.trim_end();
        if trimmed.len() != current.len() {
            self.set_text(trimmed.to_string());
        }
        self.clone()
    }

    pub fn index_of(
        &self,
        value: &str,
        start_index: usize,
        count: Option<usize>,
        ignore_case: bool,
    ) -> isize {
        let mut haystack = self.text();
        let mut needle = value.to_string();
        if ignore_case {
            haystack = haystack.to_lowercase();
            needle = needle.to_lowercase();
        }
        let haystack_chars = chars_of(&haystack);
        let start = start_index.min(haystack_chars.len());
        let end = match count {
            None => haystack_chars.len(),
            Some(count) => (start + count).min(haystack_chars.len()),
        };
        let needle_chars = chars_of(&needle);
        if needle_chars.is_empty() {
            return start as isize;
        }
        if end < start || needle_chars.len() > end - start {
            return -1;
        }
        for position in start..=(end - needle_chars.len()) {
            if haystack_chars[position..position + needle_chars.len()] == needle_chars[..] {
                return position as isize;
            }
        }
        -1
    }
}

impl Default for RichText {
    fn default() -> Self {
        RichText::new()
    }
}

// ---------------------------------------------------------------------------
// Title
// ---------------------------------------------------------------------------

typed_handle!(Title, Title);

impl Title {
    fn data(&self) -> Ref<'_, TitleData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Title(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, TitleData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Title(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> Title {
        Title(new_node(NodeData::Title(TitleData::default())))
    }

    pub fn title_text(&self) -> Option<RichText> {
        self.data().text.as_ref().map(|node| RichText(node.0.clone()))
    }

    pub fn title_date(&self) -> Option<RichText> {
        self.data().date.as_ref().map(|node| RichText(node.0.clone()))
    }

    pub fn title_time(&self) -> Option<RichText> {
        self.data().time.as_ref().map(|node| RichText(node.0.clone()))
    }

    pub fn set_title_text(&self, value: Option<RichText>) {
        self.set_slot(0, value);
    }

    pub fn set_title_date(&self, value: Option<RichText>) {
        self.set_slot(1, value);
    }

    pub fn set_title_time(&self, value: Option<RichText>) {
        self.set_slot(2, value);
    }

    fn set_slot(&self, slot: u8, value: Option<RichText>) {
        let current = {
            let data = self.data();
            match slot {
                0 => data.text.clone(),
                1 => data.date.clone(),
                _ => data.time.clone(),
            }
        };
        let new_node_opt = value.as_ref().map(|rt| rt.as_node());
        if let (Some(current), Some(new_node)) = (&current, &new_node_opt) {
            if Rc::ptr_eq(&current.0, &new_node.0) {
                return;
            }
        }
        if let Some(current) = &current {
            if let Some(parent) = current.0.borrow().parent.upgrade() {
                if Rc::ptr_eq(&parent, &self.0) {
                    current.0.borrow_mut().parent = Weak::new();
                }
            }
        }
        if let Some(new_node) = &new_node_opt {
            new_node.0.borrow_mut().parent = Rc::downgrade(&self.0);
        }
        let mut data = self.data_mut();
        match slot {
            0 => data.text = new_node_opt,
            1 => data.date = new_node_opt,
            _ => data.time = new_node_opt,
        }
    }

    pub fn child_nodes(&self) -> Vec<Node> {
        logical_children(&self.as_node())
    }

    pub fn get_child_nodes<T: TypedNode>(&self) -> Vec<T> {
        let mut out = Vec::new();
        collect_descendants::<T>(&self.as_node(), &mut out);
        out
    }
}

impl Default for Title {
    fn default() -> Self {
        Title::new()
    }
}

// ---------------------------------------------------------------------------
// Outline / OutlineElement
// ---------------------------------------------------------------------------

typed_handle!(Outline, Outline);
composite_methods!(Outline);

impl Outline {
    fn data(&self) -> Ref<'_, OutlineData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Outline(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, OutlineData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Outline(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> Outline {
        Outline(new_node(NodeData::Outline(OutlineData::default())))
    }

    pub fn horizontal_offset(&self) -> Option<f64> {
        self.data().horizontal_offset
    }

    pub fn set_horizontal_offset(&self, value: Option<f64>) {
        self.data_mut().horizontal_offset = value;
    }

    pub fn vertical_offset(&self) -> Option<f64> {
        self.data().vertical_offset
    }

    pub fn set_vertical_offset(&self, value: Option<f64>) {
        self.data_mut().vertical_offset = value;
    }

    pub fn max_width(&self) -> Option<f64> {
        self.data().max_width
    }

    pub fn set_max_width(&self, value: Option<f64>) {
        self.data_mut().max_width = value;
    }

    pub fn max_height(&self) -> Option<f64> {
        self.data().max_height
    }

    pub fn set_max_height(&self, value: Option<f64>) {
        self.data_mut().max_height = value;
    }

    pub fn min_width(&self) -> Option<f64> {
        self.data().min_width
    }

    pub fn reserved_width(&self) -> Option<f64> {
        self.data().reserved_width
    }

    pub fn indent_position(&self) -> Option<f64> {
        self.data().indent_position
    }

    pub fn descendants_cannot_be_moved(&self) -> bool {
        self.data().descendants_cannot_be_moved
    }

    pub fn last_modified_time(&self) -> Option<DateTime> {
        self.data().last_modified_time
    }
}

impl Default for Outline {
    fn default() -> Self {
        Outline::new()
    }
}

typed_handle!(OutlineElement, OutlineElement);
composite_methods!(OutlineElement);

impl OutlineElement {
    pub fn new() -> OutlineElement {
        OutlineElement(new_node(NodeData::OutlineElement(OutlineElementData::default())))
    }

    pub fn number_list(&self) -> Option<NumberList> {
        match &self.0.borrow().data {
            NodeData::OutlineElement(data) => data.number_list.clone(),
            _ => unreachable!(),
        }
    }

    pub fn set_number_list(&self, value: Option<NumberList>) {
        match &mut self.0.borrow_mut().data {
            NodeData::OutlineElement(data) => data.number_list = value,
            _ => unreachable!(),
        }
    }
}

impl Default for OutlineElement {
    fn default() -> Self {
        OutlineElement::new()
    }
}

// ---------------------------------------------------------------------------
// Image
// ---------------------------------------------------------------------------

typed_handle!(Image, Image);
composite_methods!(Image);

impl Image {
    fn data(&self) -> Ref<'_, ImageData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Image(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, ImageData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Image(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> Image {
        Image(new_node(NodeData::Image(ImageData::default())))
    }

    pub fn with_file_name(self, value: Option<String>) -> Image {
        self.data_mut().file_name = value;
        self
    }

    pub fn with_file_path(self, value: Option<String>) -> Image {
        self.data_mut().file_path = value;
        self
    }

    pub fn with_format(self, value: Option<String>) -> Image {
        self.data_mut().format = value;
        self
    }

    pub fn with_bytes(self, value: Vec<u8>) -> Image {
        self.data_mut().bytes = value;
        self
    }

    pub fn with_width(self, value: Option<f64>) -> Image {
        self.data_mut().width = value;
        self
    }

    pub fn with_height(self, value: Option<f64>) -> Image {
        self.data_mut().height = value;
        self
    }

    pub fn with_original_width(self, value: Option<f64>) -> Image {
        self.data_mut().original_width = value;
        self
    }

    pub fn with_original_height(self, value: Option<f64>) -> Image {
        self.data_mut().original_height = value;
        self
    }

    pub fn with_alignment(self, value: Option<HorizontalAlignment>) -> Image {
        self.data_mut().alignment = value;
        self
    }

    pub fn with_alternative_text_title(self, value: Option<String>) -> Image {
        self.data_mut().alternative_text_title = value;
        self
    }

    pub fn with_alternative_text_description(self, value: Option<String>) -> Image {
        self.data_mut().alternative_text_description = value;
        self
    }

    pub fn with_hyperlink_url(self, value: Option<String>) -> Image {
        self.data_mut().hyperlink_url = value;
        self
    }

    pub fn with_tags(self, value: Vec<NoteTag>) -> Image {
        self.data_mut().tags = value;
        self
    }

    pub fn file_name(&self) -> Option<String> {
        self.data().file_name.clone()
    }

    pub fn file_path(&self) -> Option<String> {
        self.data().file_path.clone()
    }

    pub fn format(&self) -> Option<String> {
        self.data().format.clone()
    }

    pub fn bytes(&self) -> Vec<u8> {
        self.data().bytes.clone()
    }

    pub fn width(&self) -> Option<f64> {
        self.data().width
    }

    pub fn height(&self) -> Option<f64> {
        self.data().height
    }

    pub fn original_width(&self) -> Option<f64> {
        self.data().original_width
    }

    pub fn original_height(&self) -> Option<f64> {
        self.data().original_height
    }

    pub fn alignment(&self) -> Option<HorizontalAlignment> {
        self.data().alignment
    }

    pub fn set_alignment(&self, value: Option<HorizontalAlignment>) {
        self.data_mut().alignment = value;
    }

    pub fn is_background(&self) -> bool {
        self.data().is_background
    }

    pub fn horizontal_offset(&self) -> Option<f64> {
        self.data().horizontal_offset
    }

    pub fn vertical_offset(&self) -> Option<f64> {
        self.data().vertical_offset
    }

    pub fn last_modified_time(&self) -> Option<DateTime> {
        self.data().last_modified_time
    }

    pub fn alternative_text_title(&self) -> Option<String> {
        self.data().alternative_text_title.clone()
    }

    pub fn alternative_text_description(&self) -> Option<String> {
        self.data().alternative_text_description.clone()
    }

    pub fn hyperlink_url(&self) -> Option<String> {
        self.data().hyperlink_url.clone()
    }

    pub fn tags(&self) -> Vec<NoteTag> {
        self.data().tags.clone()
    }

    pub fn tags_mut(&self) -> RefMut<'_, Vec<NoteTag>> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Image(data) => &mut data.tags,
            _ => unreachable!(),
        })
    }

    pub fn replace(&self, image: &Image) {
        let source = image.data().clone();
        *self.data_mut() = source;
    }
}

impl Default for Image {
    fn default() -> Self {
        Image::new()
    }
}

// ---------------------------------------------------------------------------
// AttachedFile
// ---------------------------------------------------------------------------

typed_handle!(AttachedFile, AttachedFile);

impl AttachedFile {
    fn data(&self) -> Ref<'_, AttachedFileData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::AttachedFile(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, AttachedFileData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::AttachedFile(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> AttachedFile {
        AttachedFile(new_node(NodeData::AttachedFile(AttachedFileData::default())))
    }

    pub fn with_file_name(self, value: Option<String>) -> AttachedFile {
        self.data_mut().file_name = value;
        self
    }

    pub fn with_bytes(self, value: Vec<u8>) -> AttachedFile {
        self.data_mut().bytes = value;
        self
    }

    pub fn with_tags(self, value: Vec<NoteTag>) -> AttachedFile {
        self.data_mut().tags = value;
        self
    }

    pub fn file_name(&self) -> Option<String> {
        self.data().file_name.clone()
    }

    pub fn bytes(&self) -> Vec<u8> {
        self.data().bytes.clone()
    }

    pub fn tags(&self) -> Vec<NoteTag> {
        self.data().tags.clone()
    }

    pub fn tags_mut(&self) -> RefMut<'_, Vec<NoteTag>> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::AttachedFile(data) => &mut data.tags,
            _ => unreachable!(),
        })
    }
}

impl Default for AttachedFile {
    fn default() -> Self {
        AttachedFile::new()
    }
}

// ---------------------------------------------------------------------------
// Table / TableRow / TableCell
// ---------------------------------------------------------------------------

typed_handle!(Table, Table);
composite_methods!(Table);

impl Table {
    fn data(&self) -> Ref<'_, TableData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Table(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, TableData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Table(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> Table {
        Table(new_node(NodeData::Table(TableData::default())))
    }

    pub fn with_tags(self, value: Vec<NoteTag>) -> Table {
        self.data_mut().tags = value;
        self
    }

    pub fn with_columns(self, value: Vec<TableColumn>) -> Table {
        self.data_mut().columns = value;
        self
    }

    pub fn with_borders_visible(self, value: bool) -> Table {
        self.data_mut().is_borders_visible = value;
        self
    }

    pub fn tags(&self) -> Vec<NoteTag> {
        self.data().tags.clone()
    }

    pub fn tags_mut(&self) -> RefMut<'_, Vec<NoteTag>> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Table(data) => &mut data.tags,
            _ => unreachable!(),
        })
    }

    pub fn columns(&self) -> Vec<TableColumn> {
        self.data().columns.clone()
    }

    pub fn columns_mut(&self) -> RefMut<'_, Vec<TableColumn>> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Table(data) => &mut data.columns,
            _ => unreachable!(),
        })
    }

    pub fn is_borders_visible(&self) -> bool {
        self.data().is_borders_visible
    }

    pub fn last_modified_time(&self) -> Option<DateTime> {
        self.data().last_modified_time
    }
}

impl Default for Table {
    fn default() -> Self {
        Table::new()
    }
}

typed_handle!(TableRow, TableRow);
composite_methods!(TableRow);

impl TableRow {
    pub fn new() -> TableRow {
        TableRow(new_node(NodeData::TableRow(TableRowData)))
    }
}

impl Default for TableRow {
    fn default() -> Self {
        TableRow::new()
    }
}

typed_handle!(TableCell, TableCell);
composite_methods!(TableCell);

impl TableCell {
    pub fn new() -> TableCell {
        TableCell(new_node(NodeData::TableCell(TableCellData)))
    }
}

impl Default for TableCell {
    fn default() -> Self {
        TableCell::new()
    }
}

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

typed_handle!(Page, Page);
composite_methods!(Page);

impl Page {
    fn data(&self) -> Ref<'_, PageData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Page(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, PageData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Page(data) => data,
            _ => unreachable!(),
        })
    }

    pub fn new() -> Page {
        Page(new_node(NodeData::Page(PageData::default())))
    }

    /// The page title, which is always the first child when present.
    pub fn title(&self) -> Option<Title> {
        self.first_child().and_then(|node| Title::from_node(&node))
    }

    pub fn author(&self) -> Option<String> {
        self.data().author.clone()
    }

    pub fn set_author(&self, value: Option<String>) {
        self.data_mut().author = value;
    }

    pub fn creation_time(&self) -> Option<DateTime> {
        self.data().creation_time
    }

    pub fn set_creation_time(&self, value: Option<DateTime>) {
        self.data_mut().creation_time = value;
    }

    pub fn last_modified_time(&self) -> Option<DateTime> {
        self.data().last_modified_time
    }

    pub fn set_last_modified_time(&self, value: Option<DateTime>) {
        self.data_mut().last_modified_time = value;
    }

    pub fn level(&self) -> Option<i64> {
        self.data().level
    }

    pub fn set_level(&self, value: Option<i64>) {
        self.data_mut().level = value;
    }

    pub fn is_conflict_page(&self) -> bool {
        self.data().is_conflict_page
    }

    /// Deep-copy the page subtree, detaching it from any parent. This is the
    /// Rust equivalent of the Python `Page.Clone()` (named `deep_clone` to avoid
    /// colliding with the shallow handle `Clone`).
    pub fn deep_clone(&self, _clone_history: bool) -> Page {
        let cloned = deep_clone_node(&self.as_node(), None);
        Page::from_node(&cloned).expect("clone of a page node is a page")
    }
}

impl Default for Page {
    fn default() -> Self {
        Page::new()
    }
}

// ---------------------------------------------------------------------------
// PageHistory
// ---------------------------------------------------------------------------

/// The ordered revision history for a [`Page`], returned by
/// [`Document::get_page_history`]. Mirrors `PageHistory`.
#[derive(Debug, Clone)]
pub struct PageHistory {
    current: Node,
    versions: Vec<Node>,
}

impl PageHistory {
    pub fn new(current: Page) -> PageHistory {
        PageHistory {
            current: current.as_node(),
            versions: Vec::new(),
        }
    }

    pub fn count(&self) -> usize {
        self.versions.len()
    }

    pub fn current(&self) -> Page {
        Page::from_node(&self.current).expect("page history current is a page")
    }

    pub fn is_read_only(&self) -> bool {
        false
    }

    pub fn add(&mut self, page: Page) {
        self.versions.push(page.as_node());
    }

    pub fn add_range(&mut self, pages: Vec<Page>) {
        self.versions.extend(pages.into_iter().map(|page| page.as_node()));
    }

    pub fn clear(&mut self) {
        self.versions.clear();
    }

    pub fn contains(&self, page: &Page) -> bool {
        self.versions.iter().any(|node| Node::ptr_eq(node, &page.as_node()))
    }

    pub fn index_of(&self, page: &Page) -> isize {
        match self
            .versions
            .iter()
            .position(|node| Node::ptr_eq(node, &page.as_node()))
        {
            Some(index) => index as isize,
            None => -1,
        }
    }

    pub fn insert(&mut self, index: usize, page: Page) {
        self.versions.insert(index, page.as_node());
    }

    pub fn remove(&mut self, page: &Page) -> bool {
        match self
            .versions
            .iter()
            .position(|node| Node::ptr_eq(node, &page.as_node()))
        {
            Some(index) => {
                self.versions.remove(index);
                true
            }
            None => false,
        }
    }

    pub fn remove_at(&mut self, index: usize) {
        self.versions.remove(index);
    }

    pub fn remove_range(&mut self, index: usize, count: usize) {
        let end = (index + count).min(self.versions.len());
        self.versions.drain(index..end);
    }

    pub fn get(&self, index: usize) -> Page {
        Page::from_node(&self.versions[index]).expect("page history item is a page")
    }

    pub fn set(&mut self, index: usize, value: Page) {
        self.versions[index] = value.as_node();
    }

    pub fn len(&self) -> usize {
        self.versions.len()
    }

    pub fn is_empty(&self) -> bool {
        self.versions.is_empty()
    }

    pub fn iter(&self) -> impl Iterator<Item = Page> + '_ {
        self.versions
            .iter()
            .map(|node| Page::from_node(node).expect("page history item is a page"))
    }
}

// ---------------------------------------------------------------------------
// Document
// ---------------------------------------------------------------------------

typed_handle!(Document, Document);
composite_methods!(Document);

fn link_ptr(page: &Page) -> usize {
    Rc::as_ptr(&page.0) as *const () as usize
}

fn infer_save_format(path: &Path) -> Result<SaveFormat, AsposeNoteError> {
    let suffix = path
        .extension()
        .map(|ext| ext.to_string_lossy().to_ascii_lowercase())
        .unwrap_or_default();
    if suffix == "pdf" {
        Ok(SaveFormat::Pdf)
    } else {
        Err(AsposeNoteError::unsupported_save_format(
            "Only .pdf file targets are supported for save operations",
        ))
    }
}

impl Document {
    fn data(&self) -> Ref<'_, DocumentData> {
        Ref::map(self.0.borrow(), |inner| match &inner.data {
            NodeData::Document(data) => data,
            _ => unreachable!(),
        })
    }

    fn data_mut(&self) -> RefMut<'_, DocumentData> {
        RefMut::map(self.0.borrow_mut(), |inner| match &mut inner.data {
            NodeData::Document(data) => data,
            _ => unreachable!(),
        })
    }

    /// Create an empty document (mirrors `Document()` with no source).
    pub fn new() -> Document {
        Document(new_node(NodeData::Document(DocumentData::default())))
    }

    /// Load a `.one` file from a path.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Document, AsposeNoteError> {
        Document::open_with_options(path, &LoadOptions::default())
    }

    pub fn open_with_options<P: AsRef<Path>>(
        path: P,
        options: &LoadOptions,
    ) -> Result<Document, AsposeNoteError> {
        Self::check_password(options)?;
        let source = crate::internal::onestore::parser::OneStoreSource::Path(path.as_ref());
        let loaded =
            crate::internal::ms_one::loader::load_onenote_document(source, options.load_history)?;
        Ok(Self::from_loaded(loaded))
    }

    /// Load a `.one` document from any byte stream.
    pub fn from_reader<R: Read>(reader: &mut R) -> Result<Document, AsposeNoteError> {
        Document::from_reader_with_options(reader, &LoadOptions::default())
    }

    pub fn from_reader_with_options<R: Read>(
        reader: &mut R,
        options: &LoadOptions,
    ) -> Result<Document, AsposeNoteError> {
        Self::check_password(options)?;
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        let source = crate::internal::onestore::parser::OneStoreSource::Bytes(data);
        let loaded =
            crate::internal::ms_one::loader::load_onenote_document(source, options.load_history)?;
        Ok(Self::from_loaded(loaded))
    }

    fn check_password(options: &LoadOptions) -> Result<(), AsposeNoteError> {
        if matches!(&options.document_password, Some(password) if !password.is_empty()) {
            return Err(AsposeNoteError::incorrect_password(
                "Password-protected documents are not supported",
            ));
        }
        Ok(())
    }

    fn from_loaded(
        loaded: crate::internal::ms_one::loader::LoadedOneNoteDocument,
    ) -> Document {
        let document = Document::new();
        document.set_display_name(loaded.display_name);
        for (page, history) in loaded.pages.into_iter().zip(loaded.page_histories.into_iter()) {
            document.append_child_last(page.as_node());
            let ordered: Vec<Page> = if history.is_empty() {
                vec![page.clone()]
            } else {
                history
            };
            let current = ordered.last().cloned().unwrap_or_else(|| page.clone());
            let mut page_history = PageHistory::new(current);
            if ordered.len() > 1 {
                page_history.add_range(ordered[..ordered.len() - 1].to_vec());
            }
            document
                .data_mut()
                .page_histories
                .insert(link_ptr(&page), page_history);
        }
        document
    }

    pub fn display_name(&self) -> Option<String> {
        self.data().display_name.clone()
    }

    pub fn set_display_name(&self, value: Option<String>) {
        self.data_mut().display_name = value;
    }

    pub fn creation_time(&self) -> Option<DateTime> {
        self.data().creation_time
    }

    /// Always reports [`FileFormat::OneNote2010`], matching the Python edition.
    pub fn file_format(&self) -> FileFormat {
        FileFormat::OneNote2010
    }

    /// Compatibility stub; performs no layout computation.
    pub fn detect_layout_changes(&self) {}

    pub fn get_page_history(&self, page: &Page) -> PageHistory {
        let key = link_ptr(page);
        self.data()
            .page_histories
            .get(&key)
            .cloned()
            .unwrap_or_else(|| PageHistory::new(page.clone()))
    }

    /// The pages directly under this document.
    pub fn child_pages(&self) -> Vec<Page> {
        self.child_nodes()
            .iter()
            .filter_map(Page::from_node)
            .collect()
    }

    /// Save the document. PDF export is implemented in a later stage; this
    /// currently validates the requested format and reports
    /// [`AsposeNoteError::UnsupportedSaveFormat`].
    pub fn save_to_path<P: AsRef<Path>>(
        &self,
        path: P,
        format: Option<SaveFormat>,
    ) -> Result<(), AsposeNoteError> {
        let save_format = match format {
            Some(format) => format,
            None => infer_save_format(path.as_ref())?,
        };
        self.save_resolved(save_format)
    }

    pub fn save_with_options<P: AsRef<Path>>(
        &self,
        _path: P,
        options: &PdfSaveOptions,
    ) -> Result<(), AsposeNoteError> {
        self.save_resolved(options.save_format())
    }

    pub fn save_to_writer<W: std::io::Write>(
        &self,
        _writer: &mut W,
        format: SaveFormat,
    ) -> Result<(), AsposeNoteError> {
        self.save_resolved(format)
    }

    fn save_resolved(&self, format: SaveFormat) -> Result<(), AsposeNoteError> {
        match format {
            SaveFormat::Pdf => Err(AsposeNoteError::unsupported_save_format(
                "PDF export is not implemented in this build (stage 1 opens files only)",
            )),
        }
    }
}

impl Default for Document {
    fn default() -> Self {
        Document::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rich_text_append_and_replace() {
        let rich_text = RichText::from_text("hello world");
        assert_eq!(rich_text.text(), "hello world");
        assert_eq!(rich_text.length(), 11);

        let returned = rich_text.replace("world", "rust", None);
        assert!(RichText::ptr_eq_handle(&returned, &rich_text));
        assert_eq!(rich_text.text(), "hello rust");
    }

    #[test]
    fn rich_text_append_builds_runs() {
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
        assert_eq!(runs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>(), ["First", "Second"]);
        assert!(runs.iter().all(|run| run.style.is_italic && run.style.font_name.as_deref() == Some("Courier New")));
    }

    #[test]
    fn note_tag_factories() {
        let star = NoteTag::create_yellow_star(Some("Важно"));
        assert_eq!(star.label().as_deref(), Some("Важно"));
        assert_eq!(star.icon(), Some(13));
        assert_eq!(star.status(), TagStatus::Open);
        assert!(star.completed_time().is_none());

        let default_star = NoteTag::create_yellow_star(None);
        assert_eq!(default_star.label().as_deref(), Some("Yellow Star"));
    }

    #[test]
    fn page_deep_clone_detaches_parent() {
        let document = Document::new();
        let page = Page::new();
        let title = Title::new();
        title.set_title_text(Some(RichText::from_text("Title")));
        page.append_child_last(title.as_node());
        document.append_child_last(page.as_node());

        assert!(page.parent_node().is_some());
        let cloned = page.deep_clone(false);
        assert!(cloned.parent_node().is_none());
        assert_eq!(cloned.child_nodes().len(), page.child_nodes().len());
        assert_eq!(
            cloned.get_child_nodes::<RichText>().len(),
            page.get_child_nodes::<RichText>().len()
        );
    }

    impl RichText {
        fn ptr_eq_handle(a: &RichText, b: &RichText) -> bool {
            Rc::ptr_eq(&a.0, &b.0)
        }
    }
}

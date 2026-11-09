//! # Aspose.Note FOSS for Rust
//!
//! A free, open-source Rust library for reading Microsoft OneNote `.one` files.
//! This is a Rust port of the `aspose-note-foss` Python package, preserving the
//! same module structure, public API surface, and test behavior (with idiomatic
//! `snake_case` naming).
//!
//! The primary entry point is [`Document`], which loads a `.one` file and
//! exposes its pages as a tree of [`Page`], [`Outline`], [`OutlineElement`],
//! [`RichText`], [`Image`], [`Table`], and [`AttachedFile`] nodes.

pub mod enums;
pub mod error;
pub mod internal;
pub mod model;
pub mod saving;

pub use enums::{FileFormat, HorizontalAlignment, NodeType, SaveFormat, TagStatus};
pub use error::AsposeNoteError;
pub use model::{
    AttachedFile, DateTime, Document, DocumentVisitor, Image, License, LoadOptions, Metered, Node,
    NoteTag, NumberList, Outline, OutlineElement, Page, PageHistory, ParagraphStyle, RichText,
    Table, TableCell, TableColumn, TableRow, TextRun, TextStyle, Title, TypedNode,
};

//! MS-ONE revision store → public DOM.

pub mod entities;
pub mod loader;

pub use entities::{LogicalDocument, LogicalNode};
pub use loader::{load_onenote_document, LoadedOneNoteDocument};

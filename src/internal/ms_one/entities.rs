//! Intermediate logical tree shared between the MS-ONE loader stages. Mirrors
//! `aspose.note._internal.ms_one.entities`.

use crate::internal::onestore::parser::{ExtendedGuid, PropMap};

/// A node in the intermediate logical tree, built from the parsed OneStore
/// object graph before it is turned into the public DOM.
#[derive(Debug, Clone)]
pub struct LogicalNode {
    pub kind: String,
    pub oid: ExtendedGuid,
    pub properties: PropMap,
    pub children: Vec<LogicalNode>,
    pub metadata: Vec<LogicalNode>,
}

impl LogicalNode {
    pub fn new(kind: String, oid: ExtendedGuid, properties: PropMap) -> LogicalNode {
        LogicalNode {
            kind,
            oid,
            properties,
            children: Vec::new(),
            metadata: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct LogicalDocument {
    pub root: Option<LogicalNode>,
}

//! MS-ONESTORE container parsing.

pub mod parser;

pub use parser::{
    parse_onestore_file, parse_onestore_reader, ExtendedGuid, Jcid, ObjectSpaceState,
    OneStoreFile, OneStoreObject, OneStoreSource, PropMap, PropValue, RevisionManifestRecord,
};

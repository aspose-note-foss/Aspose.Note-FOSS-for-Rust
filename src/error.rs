//! Error type, mirroring `aspose.note.exceptions`.
//!
//! The Python package exposes a base `AsposeNoteError` with five concrete
//! subclasses. In Rust these map to the variants of a single error enum; the
//! whole enum plays the role of the `AsposeNoteError` base.

use crate::enums::FileFormat;
use std::fmt;

/// Base error for the library. Fallible entry points return
/// `Result<_, AsposeNoteError>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AsposeNoteError {
    /// The source file's OneStore container is malformed.
    FileCorrupted(String),
    /// The parsed OneStore object graph doesn't match the expected structure.
    IncorrectDocumentStructure(String),
    /// A password was supplied; password-protected documents are not supported.
    IncorrectPassword(String),
    /// The source isn't a recognized OneNote format; carries the detected format.
    UnsupportedFileFormat {
        message: String,
        file_format: FileFormat,
    },
    /// `Document::save` was called with anything other than `SaveFormat::Pdf`.
    UnsupportedSaveFormat(String),
}

impl AsposeNoteError {
    pub fn file_corrupted(message: impl Into<String>) -> AsposeNoteError {
        AsposeNoteError::FileCorrupted(message.into())
    }

    pub fn incorrect_document_structure(message: impl Into<String>) -> AsposeNoteError {
        AsposeNoteError::IncorrectDocumentStructure(message.into())
    }

    pub fn incorrect_password(message: impl Into<String>) -> AsposeNoteError {
        AsposeNoteError::IncorrectPassword(message.into())
    }

    pub fn unsupported_save_format(message: impl Into<String>) -> AsposeNoteError {
        AsposeNoteError::UnsupportedSaveFormat(message.into())
    }

    /// Build an [`AsposeNoteError::UnsupportedFileFormat`], coercing the detected
    /// token/GUID into a [`FileFormat`] just like the Python constructor.
    pub fn unsupported_file_format(
        message: impl Into<String>,
        detected: Option<&str>,
    ) -> AsposeNoteError {
        let file_format = match detected {
            Some(value) => FileFormat::from_detected(value),
            None => FileFormat::Unknown,
        };
        AsposeNoteError::UnsupportedFileFormat {
            message: message.into(),
            file_format,
        }
    }

    /// The detected source format for an [`AsposeNoteError::UnsupportedFileFormat`],
    /// mirroring `UnsupportedFileFormatException.FileFormat`.
    pub fn file_format(&self) -> Option<FileFormat> {
        match self {
            AsposeNoteError::UnsupportedFileFormat { file_format, .. } => Some(*file_format),
            _ => None,
        }
    }
}

impl fmt::Display for AsposeNoteError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            AsposeNoteError::FileCorrupted(message)
            | AsposeNoteError::IncorrectDocumentStructure(message)
            | AsposeNoteError::IncorrectPassword(message)
            | AsposeNoteError::UnsupportedSaveFormat(message) => f.write_str(message),
            AsposeNoteError::UnsupportedFileFormat { message, .. } => f.write_str(message),
        }
    }
}

impl std::error::Error for AsposeNoteError {}

impl From<std::io::Error> for AsposeNoteError {
    fn from(error: std::io::Error) -> Self {
        AsposeNoteError::FileCorrupted(error.to_string())
    }
}

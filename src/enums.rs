//! Public enumerations, mirroring `aspose.note.enums`.

/// The output format accepted by [`crate::Document::save`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SaveFormat {
    Pdf,
}

/// The detected OneNote source format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FileFormat {
    Unknown,
    OneNote2010,
    OneNoteOnline,
    OneNote2007,
}

impl FileFormat {
    /// Lower-cased canonical token, matching the Python enum `value`.
    fn token(self) -> &'static str {
        match self {
            FileFormat::Unknown => "unknown",
            FileFormat::OneNote2010 => "onenote2010",
            FileFormat::OneNoteOnline => "onenoteonline",
            FileFormat::OneNote2007 => "onenote2007",
        }
    }

    /// Coerce a detected value (format token, variant name, or section file-type
    /// GUID) into a [`FileFormat`]. Mirrors `exceptions._coerce_file_format`.
    pub fn from_detected(value: &str) -> FileFormat {
        let normalized = value.trim().to_ascii_lowercase();
        if normalized.is_empty() {
            return FileFormat::Unknown;
        }
        for candidate in [
            FileFormat::Unknown,
            FileFormat::OneNote2010,
            FileFormat::OneNoteOnline,
            FileFormat::OneNote2007,
        ] {
            // The Python enum `name` and `value` lower-case to the same token.
            if normalized == candidate.token() {
                return candidate;
            }
        }
        if normalized == "7b5c52e4-d88c-4da7-aeb1-5378d02996d3" {
            return FileFormat::OneNote2010;
        }
        FileFormat::Unknown
    }
}

/// Horizontal text/image alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HorizontalAlignment {
    Left,
    Center,
    Right,
}

/// A [`crate::NoteTag`]'s completion state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TagStatus {
    Open = 0,
    Completed = 1,
    Disabled = 2,
}

impl TagStatus {
    /// Build a [`TagStatus`] from its integer discriminant, mirroring the
    /// `TagStatus(int)` conversion in the Python model.
    pub fn from_i64(value: i64) -> Option<TagStatus> {
        match value {
            0 => Some(TagStatus::Open),
            1 => Some(TagStatus::Completed),
            2 => Some(TagStatus::Disabled),
            _ => None,
        }
    }
}

/// The DOM node kind used by [`crate::CompositeNode::get_child_nodes`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NodeType {
    Document,
    Page,
    Outline,
    OutlineElement,
    RichText,
    Image,
    Table,
    AttachedFile,
}

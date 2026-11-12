//! Save options, mirroring `aspose.note.saving.options`.

use crate::enums::SaveFormat;

/// Abstract base for save-options types. Mirrors `SaveOptions`; construct a
/// concrete options type such as [`PdfSaveOptions`] instead.
#[derive(Debug, Clone)]
pub struct SaveOptions {
    save_format: SaveFormat,
    pub page_index: i32,
    pub page_count: Option<i32>,
    pub fonts_subsystem: Option<String>,
}

impl SaveOptions {
    pub(crate) fn new(save_format: SaveFormat) -> SaveOptions {
        SaveOptions {
            save_format,
            page_index: 0,
            page_count: None,
            fonts_subsystem: None,
        }
    }

    pub fn save_format(&self) -> SaveFormat {
        self.save_format
    }
}

/// Options for PDF export — image compression, JPEG quality, page settings, and
/// page range. Mirrors `PdfSaveOptions`.
#[derive(Debug, Clone)]
pub struct PdfSaveOptions {
    base: SaveOptions,
    pub image_compression: Option<String>,
    pub jpeg_quality: Option<i32>,
    pub page_settings: Option<String>,
    pub page_splitting_algorithm: Option<String>,
}

impl PdfSaveOptions {
    pub fn new() -> PdfSaveOptions {
        PdfSaveOptions {
            base: SaveOptions::new(SaveFormat::Pdf),
            image_compression: None,
            jpeg_quality: None,
            page_settings: None,
            page_splitting_algorithm: None,
        }
    }

    pub fn save_format(&self) -> SaveFormat {
        self.base.save_format()
    }

    pub fn page_index(&self) -> i32 {
        self.base.page_index
    }

    pub fn set_page_index(&mut self, value: i32) {
        self.base.page_index = value;
    }

    pub fn page_count(&self) -> Option<i32> {
        self.base.page_count
    }

    pub fn set_page_count(&mut self, value: Option<i32>) {
        self.base.page_count = value;
    }

    pub fn fonts_subsystem(&self) -> Option<String> {
        self.base.fonts_subsystem.clone()
    }

    pub fn set_fonts_subsystem(&mut self, value: Option<String>) {
        self.base.fonts_subsystem = value;
    }
}

impl Default for PdfSaveOptions {
    fn default() -> Self {
        PdfSaveOptions::new()
    }
}

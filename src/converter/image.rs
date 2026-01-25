//! Image extractor - handles image extraction from DOCX.

use crate::{error::Error, Result};
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine;
use rs_docx::document::Drawing;
use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};

/// Extractor for images embedded in DOCX.
pub struct ImageExtractor {
    mode: ImageMode,
    source: ImageSource,
    counter: usize,
}

enum ImageMode {
    SaveToDir(PathBuf),
    Inline,
    Skip,
}

enum ImageSource {
    Path(PathBuf),
    Bytes(Vec<u8>),
    None,
}

impl ImageExtractor {
    /// Creates an extractor that saves images to a directory (from file).
    pub fn new_with_dir<P: AsRef<Path>>(docx_path: P, output_dir: PathBuf) -> Result<Self> {
        // Ensure output directory exists
        fs::create_dir_all(&output_dir)?;

        Ok(Self {
            mode: ImageMode::SaveToDir(output_dir),
            source: ImageSource::Path(docx_path.as_ref().to_path_buf()),
            counter: 0,
        })
    }

    /// Creates an extractor that saves images to a directory (from bytes).
    pub fn new_with_dir_from_bytes(bytes: &[u8], output_dir: PathBuf) -> Result<Self> {
        // Ensure output directory exists
        fs::create_dir_all(&output_dir)?;

        Ok(Self {
            mode: ImageMode::SaveToDir(output_dir),
            source: ImageSource::Bytes(bytes.to_vec()),
            counter: 0,
        })
    }

    /// Creates an extractor that embeds images as base64 (from file).
    pub fn new_inline<P: AsRef<Path>>(docx_path: P) -> Result<Self> {
        Ok(Self {
            mode: ImageMode::Inline,
            source: ImageSource::Path(docx_path.as_ref().to_path_buf()),
            counter: 0,
        })
    }

    /// Creates an extractor that embeds images as base64 (from bytes).
    pub fn new_inline_from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            mode: ImageMode::Inline,
            source: ImageSource::Bytes(bytes.to_vec()),
            counter: 0,
        })
    }

    /// Creates an extractor that skips all images.
    pub fn new_skip() -> Self {
        Self {
            mode: ImageMode::Skip,
            source: ImageSource::None,
            counter: 0,
        }
    }

    /// Extracts image from a Drawing element and returns Markdown.
    pub fn extract_from_drawing(
        &mut self,
        drawing: &Drawing,
        rels: &HashMap<String, String>,
    ) -> Result<Option<String>> {
        if matches!(self.mode, ImageMode::Skip) {
            return Ok(None);
        }

        // Try to find blip (image reference) in drawing
        let blip_id = self.find_blip_id(drawing);

        let Some(rel_id) = blip_id else {
            return Ok(None);
        };

        // Get image path from relationships
        let Some(image_path) = rels.get(&rel_id) else {
            return Ok(None);
        };

        // Extract and process image
        self.process_image(image_path)
    }

    fn find_blip_id(&self, drawing: &Drawing) -> Option<String> {
        // Try inline first (most common for embedded images)
        if let Some(inline) = &drawing.inline {
            if let Some(graphic) = &inline.graphic {
                if let Some(pic) = graphic.data.children.first() {
                    let embed = &pic.fill.blip.embed;
                    if !embed.is_empty() {
                        return Some(embed.to_string());
                    }
                }
            }
        }

        // Try anchor (for floating images)
        if let Some(anchor) = &drawing.anchor {
            if let Some(graphic) = &anchor.graphic {
                if let Some(pic) = graphic.data.children.first() {
                    let embed = &pic.fill.blip.embed;
                    if !embed.is_empty() {
                        return Some(embed.to_string());
                    }
                }
            }
        }

        None
    }

    /// Extracts image from a Pict element (VML).
    pub fn extract_from_pict(
        &mut self,
        pict: &rs_docx::document::Pict,
        rels: &HashMap<String, String>,
    ) -> Result<Option<String>> {
        if matches!(self.mode, ImageMode::Skip) {
            return Ok(None);
        }

        // Try to find image ID in shape or rect
        let blip_id = self.find_pict_blip_id(pict);

        let Some(rel_id) = blip_id else {
            return Ok(None);
        };

        // Get image path from relationships
        let Some(image_path) = rels.get(&rel_id) else {
            return Ok(None);
        };

        // Extract and process image
        self.process_image(image_path)
    }

    fn find_pict_blip_id(&self, pict: &rs_docx::document::Pict) -> Option<String> {
        // Check shape -> imagedata
        if let Some(shape) = &pict.shape {
            if let Some(img_data) = &shape.image_data {
                if let Some(id) = &img_data.id {
                    return Some(id.to_string());
                }
            }
        }

        // Check rect -> imagedata
        if let Some(rect) = &pict.rect {
            if let Some(img_data) = &rect.image_data {
                if let Some(id) = &img_data.id {
                    return Some(id.to_string());
                }
            }
        }

        None
    }

    fn process_image(&mut self, image_path: &str) -> Result<Option<String>> {
        // Read image from DOCX archive
        let image_data = self.read_image_from_docx(image_path)?;

        self.counter += 1;

        // Determine extension
        let ext = Path::new(image_path)
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("png");

        match &self.mode {
            ImageMode::SaveToDir(dir) => {
                let filename = format!("image_{}.{}", self.counter, ext);
                let output_path = dir.join(&filename);

                fs::write(&output_path, &image_data)?;

                // Return relative path
                Ok(Some(format!("![image]({})", output_path.display())))
            }
            ImageMode::Inline => {
                let mime_type = match ext.to_lowercase().as_str() {
                    "png" => "image/png",
                    "jpg" | "jpeg" => "image/jpeg",
                    "gif" => "image/gif",
                    "webp" => "image/webp",
                    "svg" => "image/svg+xml",
                    _ => "application/octet-stream",
                };

                let b64 = BASE64.encode(&image_data);
                Ok(Some(format!(
                    "<img src=\"data:{};base64,{}\" alt=\"image\" />",
                    mime_type, b64
                )))
            }
            ImageMode::Skip => Ok(None),
        }
    }

    fn read_image_from_docx(&self, image_path: &str) -> Result<Vec<u8>> {
        match &self.source {
            ImageSource::Path(path) => {
                let file = File::open(path)?;
                self.extract_from_zip(file, image_path)
            }
            ImageSource::Bytes(bytes) => {
                let cursor = Cursor::new(bytes);
                self.extract_from_zip(cursor, image_path)
            }
            ImageSource::None => Ok(Vec::new()),
        }
    }

    fn extract_from_zip<R: Read + Seek>(&self, reader: R, image_path: &str) -> Result<Vec<u8>> {
        let mut archive = zip::ZipArchive::new(reader)
            .map_err(|e| Error::DocxParse(format!("Failed to open DOCX as ZIP: {}", e)))?;

        // Image path is relative to word/ directory typically
        let full_path = if image_path.starts_with("media/") {
            format!("word/{}", image_path)
        } else if !image_path.starts_with("word/") {
            format!("word/{}", image_path)
        } else {
            image_path.to_string()
        };

        // Try full path first, then original
        let paths_to_try = [full_path.as_str(), image_path];

        for path in paths_to_try {
            if let Ok(mut entry) = archive.by_name(path) {
                let mut data = Vec::new();
                entry.read_to_end(&mut data)?;
                return Ok(data);
            }
        }

        Err(Error::MediaNotFound(image_path.to_string()))
    }
}

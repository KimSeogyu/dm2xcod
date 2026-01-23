//! Converter modules for DOCX to Markdown transformation.

mod hyperlink;
mod image;
mod numbering;
mod paragraph;
mod run;

mod styles;
mod table;

use crate::localization::{KoreanLocalization, LocalizationStrategy};
use crate::{error::Error, ConvertOptions, ImageHandling, Result};
use rs_docx::document::BodyContent;
use rs_docx::DocxFile;
use std::collections::HashMap;
use std::path::Path;

pub use self::hyperlink::resolve_hyperlink;
pub use self::image::ImageExtractor;
pub use self::numbering::NumberingResolver;
pub use self::paragraph::ParagraphConverter;
pub use self::run::RunConverter;
pub use self::styles::StyleResolver;
pub use self::table::TableConverter;

/// Main converter struct that orchestrates DOCX to Markdown conversion.
pub struct DocxToMarkdown {
    options: ConvertOptions,
}

impl DocxToMarkdown {
    /// Creates a new converter with the given options.
    pub fn new(options: ConvertOptions) -> Self {
        Self { options }
    }

    /// Creates a new converter with default options.
    pub fn with_defaults() -> Self {
        Self::new(ConvertOptions::default())
    }

    /// Converts a DOCX file to Markdown.
    ///
    /// # Arguments
    /// * `path` - Path to the DOCX file
    ///
    /// # Returns
    /// The converted Markdown content as a String.
    pub fn convert<P: AsRef<Path>>(&self, path: P) -> Result<String> {
        let path = path.as_ref();

        // Parse DOCX file
        let docx_file =
            DocxFile::from_file(path).map_err(|e| Error::DocxParse(format!("{:?}", e)))?;
        let docx = docx_file
            .parse()
            .map_err(|e| Error::DocxParse(format!("{:?}", e)))?;

        // Build relationship map for hyperlinks
        let rels = self.build_relationship_map(&docx);

        // Initialize numbering resolver
        let mut numbering_resolver = NumberingResolver::new(&docx);

        // Initialize style resolver
        let style_resolver = StyleResolver::new(&docx.styles);

        // Initialize image extractor based on options
        let mut image_extractor = match &self.options.image_handling {
            ImageHandling::SaveToDir(dir) => ImageExtractor::new_with_dir(path, dir.clone())?,
            ImageHandling::Inline => ImageExtractor::new_inline(path)?,
            ImageHandling::Skip => ImageExtractor::new_skip(),
        };

        // Select localization strategy (currently hardcoded to Korean as per plan for default)
        // TODO: Make this configurable via options
        let localization_strategy = KoreanLocalization;

        // Convert body content
        let mut output = String::new();
        let mut context = ConversionContext {
            rels: &rels,
            numbering: &mut numbering_resolver,
            image_extractor: &mut image_extractor,
            options: &self.options,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            comments: Vec::new(),
            docx_comments: docx.comments.as_ref(),
            docx_footnotes: docx.footnotes.as_ref(),
            docx_endnotes: docx.endnotes.as_ref(),
            styles: &docx.styles,
            style_resolver: &style_resolver,
            localization: &localization_strategy,
        };

        for content in &docx.document.body.content {
            output.push_str(&Self::convert_content(content, &mut context)?);
        }

        // Add footnotes/endnotes/comments if any
        if !context.footnotes.is_empty()
            || !context.endnotes.is_empty()
            || !context.comments.is_empty()
        {
            output.push_str("\n\n---\n\n");
            for (i, note) in context.footnotes.iter().enumerate() {
                output.push_str(&format!("[^{}]: {}\n", i + 1, note));
            }
            for (i, note) in context.endnotes.iter().enumerate() {
                output.push_str(&format!("[^en{}]: {}\n", i + 1, note));
            }
            for (id, text) in context.comments.iter() {
                output.push_str(&format!("[^c{}]: {}\n", id, text));
            }
        }

        Ok(output)
    }

    fn convert_content(content: &BodyContent, context: &mut ConversionContext) -> Result<String> {
        let mut output = String::new();
        match content {
            BodyContent::Paragraph(para) => {
                let converted = ParagraphConverter::convert(para, context)?;
                if !converted.is_empty() {
                    output.push_str(&converted);
                    output.push_str("\n\n");
                }
            }
            BodyContent::Table(table) => {
                let converted = TableConverter::convert(table, context)?;
                output.push_str(&converted);
                output.push_str("\n\n");
            }
            BodyContent::Sdt(sdt) => {
                if let Some(sdt_content) = &sdt.content {
                    for child in &sdt_content.content {
                        output.push_str(&Self::convert_content(child, context)?);
                    }
                }
            }
            _ => {}
        }
        Ok(output)
    }

    fn build_relationship_map<'a>(&self, docx: &'a rs_docx::Docx) -> HashMap<String, String> {
        let mut rels = HashMap::new();

        if let Some(doc_rels) = &docx.document_rels {
            for rel in &doc_rels.relationships {
                rels.insert(rel.id.to_string(), rel.target.to_string());
            }
        }

        rels
    }
}

/// Context passed through conversion for shared state.
pub struct ConversionContext<'a> {
    /// Relationship map (rId -> target URL/path)
    pub rels: &'a HashMap<String, String>,
    /// Numbering resolver for lists
    pub numbering: &'a mut NumberingResolver<'a>,
    /// Image extractor
    pub image_extractor: &'a mut ImageExtractor,
    /// Conversion options
    pub options: &'a ConvertOptions,
    /// Collected footnotes
    pub footnotes: Vec<String>,
    /// Collected endnotes
    pub endnotes: Vec<String>,
    /// Collected comments (id, text)
    pub comments: Vec<(String, String)>,
    /// Document comments reference
    pub docx_comments: Option<&'a rs_docx::document::Comments<'a>>,
    /// Document footnotes reference
    pub docx_footnotes: Option<&'a rs_docx::document::FootNotes<'a>>,
    /// Document endnotes reference
    pub docx_endnotes: Option<&'a rs_docx::document::EndNotes<'a>>,
    /// Document styles
    pub styles: &'a rs_docx::styles::Styles<'a>,
    /// Style resolver
    pub style_resolver: &'a StyleResolver<'a>,
    /// Localization strategy
    pub localization: &'a dyn LocalizationStrategy,
}

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

        // Initialize image extractor based on options
        let mut image_extractor = match &self.options.image_handling {
            ImageHandling::SaveToDir(dir) => ImageExtractor::new_with_dir(path, dir.clone())?,
            ImageHandling::Inline => ImageExtractor::new_inline(path)?,
            ImageHandling::Skip => ImageExtractor::new_skip(),
        };

        self.convert_inner(&docx, &mut image_extractor)
    }

    /// Converts a DOCX file from bytes to Markdown.
    ///
    /// # Arguments
    /// * `bytes` - The DOCX file content as bytes
    ///
    /// # Returns
    /// The converted Markdown content as a String.
    pub fn convert_from_bytes(&self, bytes: &[u8]) -> Result<String> {
        let reader = std::io::Cursor::new(bytes);
        let docx_file =
            DocxFile::from_reader(reader).map_err(|e| Error::DocxParse(format!("{:?}", e)))?;
        let docx = docx_file
            .parse()
            .map_err(|e| Error::DocxParse(format!("{:?}", e)))?;

        // Initialize image extractor based on options
        let mut image_extractor = match &self.options.image_handling {
            ImageHandling::SaveToDir(dir) => {
                ImageExtractor::new_with_dir_from_bytes(bytes, dir.clone())?
            }
            ImageHandling::Inline => ImageExtractor::new_inline_from_bytes(bytes)?,
            ImageHandling::Skip => ImageExtractor::new_skip(),
        };

        self.convert_inner(&docx, &mut image_extractor)
    }

    fn convert_inner<'a>(
        &'a self,
        docx: &'a rs_docx::Docx,
        image_extractor: &'a mut ImageExtractor,
    ) -> Result<String> {
        // Build relationship map for hyperlinks
        let rels = self.build_relationship_map(&docx);

        // Initialize numbering resolver
        let mut numbering_resolver = NumberingResolver::new(&docx);

        // Initialize style resolver
        let style_resolver = StyleResolver::new(&docx.styles);

        // Select localization strategy (currently hardcoded to Korean as per plan for default)
        // TODO: Make this configurable via options
        let localization_strategy = KoreanLocalization;

        // Convert body content
        let mut output = String::new();
        let mut context = ConversionContext {
            rels: &rels,
            numbering: &mut numbering_resolver,
            image_extractor,
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
            BodyContent::BookmarkStart(bookmark) => {
                if let Some(name) = &bookmark.name {
                    output.push_str(&format!("<a id=\"{}\"></a>", name));
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

#[cfg(test)]
mod tests {
    use super::*;
    use rs_docx::document::{BodyContent, BookmarkStart, Paragraph, SDTContent, SDT};
    use std::borrow::Cow;
    use std::collections::HashMap;

    #[test]
    fn test_convert_content_sdt_with_bookmark() {
        // Setup mock docx parts
        let styles = rs_docx::styles::Styles::new();
        let docx = rs_docx::Docx::default();

        let mut numbering_resolver = NumberingResolver::new(&docx);
        let mut image_extractor = ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let rels = HashMap::new();
        let style_resolver = StyleResolver::new(&styles);
        let localization = KoreanLocalization;

        let mut context = ConversionContext {
            rels: &rels,
            numbering: &mut numbering_resolver,
            image_extractor: &mut image_extractor,
            options: &options,
            footnotes: Vec::new(),
            endnotes: Vec::new(),
            comments: Vec::new(),
            docx_comments: None,
            docx_footnotes: None,
            docx_endnotes: None,
            styles: &styles,
            style_resolver: &style_resolver,
            localization: &localization,
        };

        // Construct SDT with nested BookmarkStart and Paragraph
        let mut sdt = SDT::default();
        let mut sdt_content = SDTContent::default();

        // Add BookmarkStart
        let mut bookmark = BookmarkStart::default();
        bookmark.name = Some(Cow::Borrowed("TestAnchor"));
        sdt_content
            .content
            .push(BodyContent::BookmarkStart(bookmark));

        // Add Paragraph
        let mut para = Paragraph::default();
        use rs_docx::document::{ParagraphContent, Run, RunContent, Text};
        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Content".into(),
            ..Default::default()
        }));
        para.content.push(ParagraphContent::Run(run));

        sdt_content.content.push(BodyContent::Paragraph(para));

        sdt.content = Some(sdt_content);

        // Convert
        let result = DocxToMarkdown::convert_content(&BodyContent::Sdt(sdt), &mut context).unwrap();

        // Verify
        assert!(result.contains("<a id=\"TestAnchor\"></a>"));
        assert!(result.contains("Content"));
    }
}

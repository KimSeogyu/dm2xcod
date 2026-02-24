//! Converter modules for DOCX to Markdown transformation.

mod hyperlink;
mod image;
mod numbering;
mod paragraph;
mod run;

mod context;
mod styles;
mod table;
mod table_grid;

use crate::adapters::docx::{AstExtractor, DocxExtractor};
#[cfg(test)]
use crate::render::escape_html_attr;
use crate::render::{MarkdownRenderer, Renderer};
use crate::{error::Error, ConvertOptions, ImageHandling, Result};
#[cfg(test)]
use rs_docx::document::BodyContent;
use rs_docx::DocxFile;
use std::collections::HashMap;
use std::path::Path;

pub use self::context::ConversionContext;
pub use self::hyperlink::resolve_hyperlink;
pub use self::image::ImageExtractor;
pub use self::numbering::NumberingResolver;
pub use self::paragraph::ParagraphConverter;
pub use self::run::RunConverter;
pub use self::styles::StyleResolver;
pub use self::table::TableConverter;

/// Main converter struct that orchestrates DOCX to Markdown conversion.
pub struct DocxToMarkdown<E = DocxExtractor, R = MarkdownRenderer> {
    options: ConvertOptions,
    extractor: E,
    renderer: R,
}

impl DocxToMarkdown<DocxExtractor, MarkdownRenderer> {
    /// Creates a new converter with the given options.
    pub fn new(options: ConvertOptions) -> Self {
        Self {
            options,
            extractor: DocxExtractor,
            renderer: MarkdownRenderer,
        }
    }

    /// Creates a new converter with default options.
    pub fn with_defaults() -> Self {
        Self::new(ConvertOptions::default())
    }
}

impl<E, R> DocxToMarkdown<E, R>
where
    E: AstExtractor,
    R: Renderer,
{
    /// Creates a converter with custom extractor and renderer components.
    ///
    /// This constructor is intended for advanced integrations where the default
    /// DOCX AST extraction and Markdown rendering pipeline needs to be replaced
    /// or decorated (for example, custom telemetry, alternate output formats,
    /// or strict test doubles).
    ///
    /// The conversion lifecycle remains the same:
    /// 1. Parse DOCX.
    /// 2. Build conversion context and resolve references.
    /// 3. Delegate block extraction to `extractor`.
    /// 4. Apply strict reference validation (when enabled).
    /// 5. Delegate final output generation to `renderer`.
    pub fn with_components(options: ConvertOptions, extractor: E, renderer: R) -> Self {
        Self {
            options,
            extractor,
            renderer,
        }
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
        let rels = self.build_relationship_map(docx);

        // Initialize numbering resolver
        let mut numbering_resolver = NumberingResolver::new(docx);

        // Initialize style resolver
        let style_resolver = StyleResolver::new(&docx.styles);

        let mut context = ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            image_extractor,
            &self.options,
            docx.comments.as_ref(),
            docx.footnotes.as_ref(),
            docx.endnotes.as_ref(),
            &style_resolver,
        );

        let mut document = self
            .extractor
            .extract(&docx.document.body.content, &mut context)?;
        document.references = context.reference_definitions();

        if self.options.strict_reference_validation {
            let missing = context.take_missing_references();
            if !missing.is_empty() {
                return Err(Error::MissingReference(missing.join(", ")));
            }
        }

        self.renderer.render(&document)
    }

    #[cfg(test)]
    fn convert_content<'a>(
        content: &BodyContent<'a>,
        context: &mut ConversionContext<'a>,
    ) -> Result<String> {
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
            BodyContent::Run(run) => {
                let converted = RunConverter::convert(run, context, None)?;
                if !converted.is_empty() {
                    output.push_str(&converted);
                    output.push_str("\n\n");
                }
            }
            BodyContent::TableCell(cell) => {
                for item in &cell.content {
                    match item {
                        rs_docx::document::TableCellContent::Paragraph(para) => {
                            let converted = ParagraphConverter::convert(para, context)?;
                            if !converted.is_empty() {
                                output.push_str(&converted);
                                output.push_str("\n\n");
                            }
                        }
                        rs_docx::document::TableCellContent::Table(table) => {
                            let converted = TableConverter::convert(table, context)?;
                            output.push_str(&converted);
                            output.push_str("\n\n");
                        }
                    }
                }
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
                    output.push_str(&format!("<a id=\"{}\"></a>", escape_html_attr(name)));
                }
            }
            _ => {}
        }
        Ok(output)
    }

    fn build_relationship_map(&self, docx: &rs_docx::Docx) -> HashMap<String, String> {
        let mut rels = HashMap::new();

        if let Some(doc_rels) = &docx.document_rels {
            for rel in &doc_rels.relationships {
                rels.insert(rel.id.to_string(), rel.target.to_string());
            }
        }

        rels
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::ast::{BlockNode, DocumentAst};
    use rs_docx::document::{
        BodyContent, BookmarkStart, EndNote, EndNotes, FootNote, FootNotes, Paragraph, Run,
        RunContent, SDTContent, SDT, TableCell, Text,
    };
    use std::borrow::Cow;
    use std::collections::HashMap;
    use std::path::PathBuf;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_docx_path(prefix: &str) -> PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time must be after UNIX_EPOCH")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "dm2xcod_converter_{}_{}_{}.docx",
            prefix,
            std::process::id(),
            nanos
        ))
    }

    #[derive(Debug, Default, Clone, Copy)]
    struct FakeExtractor;

    impl AstExtractor for FakeExtractor {
        fn extract<'a>(
            &self,
            _body: &[BodyContent<'a>],
            context: &mut ConversionContext<'a>,
        ) -> Result<DocumentAst> {
            let _ = context.register_footnote_reference(1);
            Ok(DocumentAst {
                blocks: vec![BlockNode::Paragraph("custom block".to_string())],
                references: Default::default(),
            })
        }
    }

    #[derive(Debug, Default, Clone, Copy)]
    struct MissingRefExtractor;

    impl AstExtractor for MissingRefExtractor {
        fn extract<'a>(
            &self,
            _body: &[BodyContent<'a>],
            context: &mut ConversionContext<'a>,
        ) -> Result<DocumentAst> {
            let _ = context.register_footnote_reference(999);
            Ok(DocumentAst::default())
        }
    }

    #[derive(Debug, Default, Clone, Copy)]
    struct MissingCommentExtractor;

    impl AstExtractor for MissingCommentExtractor {
        fn extract<'a>(
            &self,
            _body: &[BodyContent<'a>],
            context: &mut ConversionContext<'a>,
        ) -> Result<DocumentAst> {
            let _ = context.register_comment_reference("404");
            Ok(DocumentAst::default())
        }
    }

    #[derive(Debug, Default, Clone, Copy)]
    struct MissingEndnoteExtractor;

    impl AstExtractor for MissingEndnoteExtractor {
        fn extract<'a>(
            &self,
            _body: &[BodyContent<'a>],
            context: &mut ConversionContext<'a>,
        ) -> Result<DocumentAst> {
            let _ = context.register_endnote_reference(404);
            Ok(DocumentAst::default())
        }
    }

    #[derive(Debug, Default, Clone, Copy)]
    struct FakeRenderer;

    impl Renderer for FakeRenderer {
        fn render(&self, document: &DocumentAst) -> Result<String> {
            let first = document
                .references
                .footnotes
                .first()
                .map(String::as_str)
                .unwrap_or("");
            Ok(format!(
                "blocks={};footnotes={};first={}",
                document.blocks.len(),
                document.references.footnotes.len(),
                first
            ))
        }
    }

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

        let mut context = ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        // Construct SDT with nested BookmarkStart and Paragraph
        let mut sdt = SDT::default();
        let mut sdt_content = SDTContent::default();

        // Add BookmarkStart
        let bookmark = BookmarkStart {
            name: Some(Cow::Borrowed("TestAnchor")),
            ..Default::default()
        };
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
        let result = DocxToMarkdown::<DocxExtractor, MarkdownRenderer>::convert_content(
            &BodyContent::Sdt(sdt),
            &mut context,
        )
        .unwrap();

        // Verify
        assert!(result.contains("<a id=\"TestAnchor\"></a>"));
        assert!(result.contains("Content"));
    }

    #[test]
    fn test_reference_registration_deduplicates_ids() {
        let styles = rs_docx::styles::Styles::new();
        let docx = rs_docx::Docx::default();

        let mut numbering_resolver = NumberingResolver::new(&docx);
        let mut image_extractor = ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let rels = HashMap::new();
        let style_resolver = StyleResolver::new(&styles);

        let mut context = ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        assert_eq!(context.register_footnote_reference(42), "[^1]");
        assert_eq!(context.register_footnote_reference(42), "[^1]");
        assert_eq!(context.footnote_count(), 1);

        assert_eq!(context.register_endnote_reference(7), "[^en1]");
        assert_eq!(context.register_endnote_reference(7), "[^en1]");
        assert_eq!(context.endnote_count(), 1);

        assert_eq!(context.register_comment_reference("3"), "[^c3]");
        assert_eq!(context.register_comment_reference("3"), "[^c3]");
        assert_eq!(context.comment_count(), 1);
    }

    #[test]
    fn test_with_components_uses_custom_extractor_and_renderer() {
        let docx = rs_docx::Docx {
            footnotes: Some(FootNotes {
                content: vec![FootNote {
                    id: Some(1),
                    content: vec![BodyContent::Paragraph(
                        Paragraph::default().push_text("Injected note"),
                    )],
                    ..Default::default()
                }],
            }),
            ..Default::default()
        };

        let options = ConvertOptions::default();
        let converter = DocxToMarkdown::with_components(options, FakeExtractor, FakeRenderer);
        let mut image_extractor = ImageExtractor::new_skip();

        let rendered = converter
            .convert_inner(&docx, &mut image_extractor)
            .expect("conversion should succeed");

        assert_eq!(rendered, "blocks=1;footnotes=1;first=Injected note");
    }

    #[test]
    fn test_with_components_respects_strict_reference_validation() {
        let docx = rs_docx::Docx::default();
        let options = ConvertOptions {
            strict_reference_validation: true,
            ..Default::default()
        };
        let converter = DocxToMarkdown::with_components(options, MissingRefExtractor, FakeRenderer);
        let mut image_extractor = ImageExtractor::new_skip();

        let err = converter
            .convert_inner(&docx, &mut image_extractor)
            .expect_err("strict validation should fail on missing references");

        match err {
            Error::MissingReference(msg) => assert!(msg.contains("footnote:999")),
            other => panic!("unexpected error: {:?}", other),
        }
    }

    #[test]
    fn test_with_components_convert_from_bytes_uses_custom_pipeline() {
        let path = temp_docx_path("bytes");
        rs_docx::Docx::default()
            .write_file(&path)
            .expect("failed to write generated docx");
        let bytes = std::fs::read(&path).expect("failed to read generated docx");
        let _ = std::fs::remove_file(&path);

        let converter =
            DocxToMarkdown::with_components(ConvertOptions::default(), FakeExtractor, FakeRenderer);
        let rendered = converter
            .convert_from_bytes(&bytes)
            .expect("conversion from bytes should succeed");

        assert_eq!(rendered, "blocks=1;footnotes=1;first=");
    }

    #[test]
    fn test_with_components_strict_validation_fails_for_missing_comment() {
        let docx = rs_docx::Docx::default();
        let options = ConvertOptions {
            strict_reference_validation: true,
            ..Default::default()
        };
        let converter =
            DocxToMarkdown::with_components(options, MissingCommentExtractor, FakeRenderer);
        let mut image_extractor = ImageExtractor::new_skip();

        let err = converter
            .convert_inner(&docx, &mut image_extractor)
            .expect_err("strict validation should fail on missing comment");

        match err {
            Error::MissingReference(msg) => assert!(msg.contains("comment:404")),
            other => panic!("unexpected error: {:?}", other),
        }
    }

    #[test]
    fn test_with_components_strict_validation_fails_for_missing_endnote() {
        let docx = rs_docx::Docx {
            endnotes: Some(EndNotes {
                content: vec![EndNote {
                    id: Some(1),
                    content: vec![BodyContent::Paragraph(
                        Paragraph::default().push_text("existing endnote"),
                    )],
                    ..Default::default()
                }],
            }),
            ..Default::default()
        };
        let options = ConvertOptions {
            strict_reference_validation: true,
            ..Default::default()
        };
        let converter =
            DocxToMarkdown::with_components(options, MissingEndnoteExtractor, FakeRenderer);
        let mut image_extractor = ImageExtractor::new_skip();

        let err = converter
            .convert_inner(&docx, &mut image_extractor)
            .expect_err("strict validation should fail on missing endnote");

        match err {
            Error::MissingReference(msg) => assert!(msg.contains("endnote:404")),
            other => panic!("unexpected error: {:?}", other),
        }
    }

    #[test]
    fn test_convert_content_body_run_is_rendered() {
        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "loose run".into(),
            ..Default::default()
        }));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = NumberingResolver::new(&docx);
        let mut image_extractor = ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let style_resolver = StyleResolver::new(&docx.styles);
        let mut context = ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let output =
            DocxToMarkdown::<DocxExtractor, MarkdownRenderer>::convert_content(
                &BodyContent::Run(run),
                &mut context,
            )
            .expect("conversion failed");
        assert_eq!(output, "loose run\n\n");
    }

    #[test]
    fn test_convert_content_body_table_cell_is_rendered() {
        let cell = TableCell::paragraph(Paragraph::default().push_text("cell text"));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = NumberingResolver::new(&docx);
        let mut image_extractor = ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let style_resolver = StyleResolver::new(&docx.styles);
        let mut context = ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let output =
            DocxToMarkdown::<DocxExtractor, MarkdownRenderer>::convert_content(
                &BodyContent::TableCell(cell),
                &mut context,
            )
            .expect("conversion failed");
        assert_eq!(output, "cell text\n\n");
    }
}

//! Paragraph converter - handles paragraph elements and their structure.

use super::{ConversionContext, RunConverter};
use crate::Result;
use rs_docx::document::{Hyperlink, Paragraph, ParagraphContent};

/// Converter for Paragraph elements.
pub struct ParagraphConverter;

/// Segment of formatted text with consistent styling.
#[derive(Debug, Clone, PartialEq)]
struct FormattedSegment {
    text: String,
    is_bold: bool,
    is_italic: bool,
    has_underline: bool,
    has_strike: bool,
}

impl ParagraphConverter {
    /// Converts a Paragraph to Markdown.
    pub fn convert(para: &Paragraph, context: &mut ConversionContext) -> Result<String> {
        // Collect all formatted segments from runs
        let segments = Self::collect_segments(para, context)?;

        // Merge adjacent segments with same formatting
        let merged = Self::merge_segments(segments);

        // Convert merged segments to markdown
        let text = Self::segments_to_markdown(&merged, context);

        if text.trim().is_empty() {
            return Ok(String::new());
        }

        // Apply paragraph-level formatting
        Self::apply_paragraph_formatting(para, text, context)
    }

    /// Collects formatted segments from paragraph content.
    fn collect_segments(
        para: &Paragraph,
        context: &mut ConversionContext,
    ) -> Result<Vec<FormattedSegment>> {
        let mut segments = Vec::new();
        // Field state: 0 = normal, 1 = after begin (skip instrText), 2 = after separate (visible TOC text)
        let mut field_state = 0;

        // Get paragraph style ID for inheritance
        let para_style_id = para
            .property
            .as_ref()
            .and_then(|p| p.style_id.as_ref())
            .map(|s| s.value.as_ref());

        for content in &para.content {
            match content {
                ParagraphContent::Run(run) => {
                    // Check for field characters (TOC, page refs, etc.)
                    for rc in &run.content {
                        if let rs_docx::document::RunContent::FieldChar(fc) = rc {
                            if let Some(char_type) = &fc.ty {
                                match char_type {
                                    rs_docx::document::CharType::Begin => field_state = 1,
                                    rs_docx::document::CharType::Separate => field_state = 2,
                                    rs_docx::document::CharType::End => field_state = 0,
                                }
                            }
                        }
                    }

                    // Skip field codes (instrText) in state 1
                    if field_state == 1 {
                        continue;
                    }

                    // In state 2 (after separate) or 0 (normal), extract visible text

                    let text = Self::extract_text(run, context);
                    if !text.is_empty() {
                        let seg = Self::run_to_segment(run, &text, context, para_style_id);
                        segments.push(seg);
                    }
                }
                ParagraphContent::Link(hyperlink) => {
                    let link_md = Self::convert_hyperlink(hyperlink, context, para_style_id)?;
                    if !link_md.is_empty() {
                        // Hyperlinks are treated as plain text segments
                        segments.push(FormattedSegment {
                            text: link_md,
                            is_bold: false,
                            is_italic: false,
                            has_underline: false,
                            has_strike: false,
                        });
                    }
                }
                ParagraphContent::SDT(sdt) => {
                    // Structured document tags (TOC, etc.) - extract inner content
                    if let Some(sdt_content) = &sdt.content {
                        for bc in &sdt_content.content {
                            if let rs_docx::document::BodyContent::Paragraph(inner_para) = bc {
                                let inner_segs = Self::collect_segments(inner_para, context)?;
                                segments.extend(inner_segs);
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        Ok(segments)
    }

    /// Extracts text from a run, excluding field codes.
    fn extract_text(run: &rs_docx::document::Run, context: &mut ConversionContext) -> String {
        let mut text = String::new();
        for content in &run.content {
            match content {
                rs_docx::document::RunContent::Text(t) => {
                    text.push_str(&t.text);
                }
                rs_docx::document::RunContent::Tab(_) => {
                    text.push('\t');
                }
                rs_docx::document::RunContent::Break(br) => match br.ty {
                    Some(rs_docx::document::BreakType::Page) => text.push_str("\n\n---\n\n"),
                    _ => text.push('\n'),
                },
                rs_docx::document::RunContent::CarriageReturn(_) => {
                    text.push('\n');
                }
                rs_docx::document::RunContent::Drawing(drawing) => {
                    if let Ok(Some(img_md)) = context
                        .image_extractor
                        .extract_from_drawing(drawing, context.rels)
                    {
                        text.push_str(&img_md);
                    }
                }
                rs_docx::document::RunContent::Pict(pict) => {
                    if let Ok(Some(img_md)) = context
                        .image_extractor
                        .extract_from_pict(pict, context.rels)
                    {
                        text.push_str(&img_md);
                    }
                }
                // Skip InstrText (field codes like TOC, PAGEREF)
                rs_docx::document::RunContent::InstrText(_) => {}
                rs_docx::document::RunContent::DelInstrText(_) => {}
                _ => {}
            }
        }
        text
    }

    /// Creates a formatted segment from a run.
    fn run_to_segment(
        run: &rs_docx::document::Run,
        text: &str,
        context: &mut ConversionContext,
        para_style_id: Option<&str>,
    ) -> FormattedSegment {
        // Resolve run style ID
        let mut run_style_id = None;
        if let Some(props) = &run.property {
            if let Some(style) = &props.style_id {
                run_style_id = Some(style.value.as_ref());
            }
        }

        // Resolve effective properties
        let props = context.style_resolver.resolve_run_property(
            run.property.as_ref(),
            run_style_id,
            para_style_id,
        );

        let is_bold = props
            .bold
            .as_ref()
            .map(|b| b.value.unwrap_or(true))
            .unwrap_or(false);
        let is_italic = props
            .italics
            .as_ref()
            .map(|i| i.value.unwrap_or(true))
            .unwrap_or(false);
        let has_underline = props.underline.is_some();
        let has_strike = props
            .strike
            .as_ref()
            .map(|s| s.value.unwrap_or(true))
            .unwrap_or(false);

        FormattedSegment {
            text: text.to_string(),
            is_bold,
            is_italic,
            has_underline,
            has_strike,
        }
    }

    /// Merges adjacent segments with identical formatting.
    fn merge_segments(segments: Vec<FormattedSegment>) -> Vec<FormattedSegment> {
        let mut merged: Vec<FormattedSegment> = Vec::new();

        for seg in segments {
            if let Some(last) = merged.last_mut() {
                // Check if formatting matches
                if last.is_bold == seg.is_bold
                    && last.is_italic == seg.is_italic
                    && last.has_underline == seg.has_underline
                    && last.has_strike == seg.has_strike
                {
                    // Merge text
                    last.text.push_str(&seg.text);
                    continue;
                }
            }
            merged.push(seg);
        }

        merged
    }

    /// Converts segments to markdown text.
    fn segments_to_markdown(segments: &[FormattedSegment], context: &ConversionContext) -> String {
        let mut result = String::new();

        for seg in segments {
            let mut text = seg.text.clone();

            // Apply formatting in order
            if seg.has_underline && context.options.html_underline {
                text = format!("<u>{}</u>", text);
            }

            if seg.has_strike {
                if context.options.html_strikethrough {
                    text = format!("<s>{}</s>", text);
                } else {
                    text = format!("~~{}~~", text);
                }
            }

            if seg.is_bold && seg.is_italic {
                text = format!("***{}***", text);
            } else if seg.is_bold {
                text = format!("**{}**", text);
            } else if seg.is_italic {
                text = format!("*{}*", text);
            }

            result.push_str(&text);
        }

        result
    }

    /// Applies paragraph-level formatting (heading, list, alignment).
    /// Applies paragraph-level formatting (heading, list, alignment).
    fn apply_paragraph_formatting(
        para: &Paragraph,
        text: String,
        context: &mut ConversionContext,
    ) -> Result<String> {
        let para_style_id = para
            .property
            .as_ref()
            .and_then(|p| p.style_id.as_ref())
            .map(|s| s.value.as_ref());

        // Resolve effective paragraph properties
        let effective_props = context
            .style_resolver
            .resolve_paragraph_property(para.property.as_ref(), para_style_id);

        let mut prefix = String::new();
        let mut is_heading = false;

        // Check for heading via pStyle (check effective style_id or original para_style_id)
        // Note: For headings, we usually rely on the style name/ID directly applied.
        // If we want to support basedOn inheritance for headings, we would need to check the chain.
        // For now, let's check the effective style_id which might be set by direct formatting or preserved.
        if let Some(style) = &effective_props.style_id {
            if let Some(heading_level) = context.localization.parse_heading_style(&style.value) {
                // Don't generate heading for empty text (e.g., TOC entries with only field codes)
                if text.trim().is_empty() {
                    return Ok(String::new());
                }
                prefix.push_str(&"#".repeat(heading_level));
                prefix.push(' ');
                is_heading = true;
            }
        }

        // Check for numbering (list items)
        // Use effective properties which include inherited numbering
        if let Some(num_pr) = &effective_props.numbering {
            if let (Some(num_id), Some(ilvl)) = (&num_pr.id, &num_pr.level) {
                let num_id_val = num_id.value as i32;
                let ilvl_val = ilvl.value as i32;
                let mut marker = context.numbering.next_marker(num_id_val, ilvl_val);

                // Handle localization-specific numbering (e.g., Korean "제N조" -> Heading)
                if let Some(formatted_prefix) = context.localization.handle_numbering(&marker) {
                    if !is_heading {
                        prefix = formatted_prefix;
                        prefix.push(' ');
                        is_heading = true;
                        marker = String::new(); // Clear to prevent double-append below
                    }
                }

                if is_heading {
                    prefix.push_str(&marker);
                    prefix.push(' ');
                } else {
                    let indent = context.numbering.get_indent(num_id_val, ilvl_val);

                    // Cap indentation to avoid code block trigger (4 spaces = indent 2)
                    // Apply consistently to all markers for uniform behavior
                    let effective_indent = indent.min(1);

                    let indent_str = "  ".repeat(effective_indent);
                    prefix.push_str(&indent_str);
                    prefix.push_str(&marker);
                    prefix.push(' ');
                }
            }
        }

        let final_text = format!("{}{}", prefix, text.trim());

        // Check for text alignment (only if not heading)
        if !is_heading {
            if let Some(jc) = &effective_props.justification {
                match &jc.value {
                    rs_docx::formatting::JustificationVal::Center => {
                        return Ok(format!(
                            "<div style=\"text-align: center;\">{}</div>",
                            final_text
                        ));
                    }
                    rs_docx::formatting::JustificationVal::Right => {
                        return Ok(format!(
                            "<div style=\"text-align: right;\">{}</div>",
                            final_text
                        ));
                    }
                    _ => {}
                }
            }
        }

        Ok(final_text)
    }

    /// Converts a hyperlink to Markdown format.
    fn convert_hyperlink(
        hyperlink: &Hyperlink,
        context: &mut ConversionContext,
        para_style_id: Option<&str>,
    ) -> Result<String> {
        let mut link_text = String::new();
        let mut field_state = 0; // 0=normal, 1=instrText, 2=visible

        for run in &hyperlink.content {
            // Check field char
            for rc in &run.content {
                if let rs_docx::document::RunContent::FieldChar(fc) = rc {
                    if let Some(char_type) = &fc.ty {
                        match char_type {
                            rs_docx::document::CharType::Begin => field_state = 1,
                            rs_docx::document::CharType::Separate => field_state = 2,
                            rs_docx::document::CharType::End => field_state = 0,
                        }
                    }
                }
            }

            if field_state == 1 {
                continue;
            }

            let text = RunConverter::convert(run, context, para_style_id)?;
            link_text.push_str(&text);
        }

        // Get target URL from relationship
        let url = hyperlink
            .id
            .as_ref()
            .and_then(|id| context.rels.get(&id.to_string()))
            .cloned()
            .unwrap_or_else(|| "#".to_string());

        if link_text.is_empty() {
            Ok(url)
        } else {
            Ok(format!("[{}]({})", link_text, url))
        }
    }
}

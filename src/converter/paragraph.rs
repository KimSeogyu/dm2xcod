//! Paragraph converter - handles paragraph elements and their structure.

use super::{ConversionContext, RunConverter};
use crate::render::{
    escape_html_attr, escape_markdown_link_destination, escape_markdown_link_text,
};
use crate::Result;
use rs_docx::document::{Hyperlink, Paragraph, ParagraphContent};

/// Converter for Paragraph elements.
pub struct ParagraphConverter;

/// Segment of formatted text with consistent styling.
#[derive(Debug, Clone, PartialEq, Default)]
struct FormattedSegment {
    text: String,
    is_bold: bool,
    is_italic: bool,
    has_underline: bool,
    has_strike: bool,
    is_insertion: bool,
    is_deletion: bool,
    anchor: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FieldPhase {
    Instruction,
    Result,
}

impl ParagraphConverter {
    /// Filters a run so only field-visible content remains, updating field stack.
    fn filter_run_by_field_state<'a>(
        run: &rs_docx::document::Run<'a>,
        field_stack: &mut Vec<FieldPhase>,
    ) -> rs_docx::document::Run<'a> {
        let mut filtered = run.clone();
        filtered.content.clear();

        for content in &run.content {
            match content {
                rs_docx::document::RunContent::FieldChar(fc) => {
                    if let Some(char_type) = &fc.ty {
                        match char_type {
                            rs_docx::document::CharType::Begin => {
                                field_stack.push(FieldPhase::Instruction)
                            }
                            rs_docx::document::CharType::Separate => {
                                if let Some(last) = field_stack.last_mut() {
                                    *last = FieldPhase::Result;
                                }
                            }
                            rs_docx::document::CharType::End => {
                                let _ = field_stack.pop();
                            }
                        }
                    }
                }
                // Keep existing behavior: field instructions are never rendered.
                rs_docx::document::RunContent::InstrText(_)
                | rs_docx::document::RunContent::DelInstrText(_) => {}
                _ => {
                    // Skip non-instruction payload while inside field instruction section.
                    if field_stack.last() != Some(&FieldPhase::Instruction) {
                        filtered.content.push(content.clone());
                    }
                }
            }
        }

        filtered
    }

    /// Converts a Paragraph to Markdown.
    pub fn convert<'a>(
        para: &Paragraph<'a>,
        context: &mut ConversionContext<'a>,
    ) -> Result<String> {
        // Collect all formatted segments from runs
        let segments = Self::collect_segments(para, context)?;

        // Merge adjacent segments with same formatting
        let merged = Self::merge_segments(segments);

        // Separate leading anchors (anchors at the start with empty text) from the rest
        let mut leading_anchors = Vec::new();
        let mut content_segments = Vec::new();
        let mut looking_for_anchors = true;

        for seg in merged {
            if looking_for_anchors && seg.text.is_empty() && seg.anchor.is_some() {
                if let Some(anchor) = &seg.anchor {
                    // Use id attribute instead of name for better compatibility (VS Code etc.)
                    leading_anchors.push(format!("<a id=\"{}\"></a>", escape_html_attr(anchor)));
                }
            } else {
                looking_for_anchors = false;
                content_segments.push(seg);
            }
        }

        // Convert merged segments to markdown
        let text = Self::segments_to_markdown(&content_segments, context);

        let anchor_tags = leading_anchors.join("");

        let is_effectively_empty = if context.preserve_whitespace() {
            text.is_empty()
        } else {
            text.trim().is_empty()
        };

        if is_effectively_empty {
            // If there is no content but there are anchors, return just the anchors
            return Ok(anchor_tags);
        }

        // Apply paragraph-level formatting
        let formatted_text = Self::apply_paragraph_formatting(para, text, context)?;

        if !anchor_tags.is_empty() {
            // Place anchors on the line BEFORE the paragraph
            // This ensures scrolling lands above the header/list item
            // and maintains valid Markdown syntax for headers (e.g. ### Title)
            Ok(format!("{}\n{}", anchor_tags, formatted_text))
        } else {
            Ok(formatted_text)
        }
    }

    /// Collects formatted segments from paragraph content.
    fn collect_segments<'a>(
        para: &Paragraph<'a>,
        context: &mut ConversionContext<'a>,
    ) -> Result<Vec<FormattedSegment>> {
        let mut segments = Vec::new();
        let mut field_stack = Vec::new();

        // Get paragraph style ID for inheritance
        let para_style_id = para
            .property
            .as_ref()
            .and_then(|p| p.style_id.as_ref())
            .map(|s| s.value.as_ref());

        for content in &para.content {
            match content {
                ParagraphContent::Run(run) => {
                    let filtered_run = Self::filter_run_by_field_state(run, &mut field_stack);
                    if filtered_run.content.is_empty() {
                        continue;
                    }

                    // Extract visible text only (field instructions already filtered out).
                    let text = Self::extract_text(&filtered_run, context);
                    if !text.is_empty() {
                        let segs =
                            Self::run_to_segment(&filtered_run, &text, context, para_style_id);
                        segments.extend(segs);
                    }
                }
                ParagraphContent::Link(hyperlink) => {
                    let link_md = Self::convert_hyperlink(hyperlink, context, para_style_id)?;
                    if !link_md.is_empty() {
                        // Hyperlinks are treated as plain text segments
                        segments.push(FormattedSegment {
                            text: link_md,
                            ..Default::default()
                        });
                    }
                }
                ParagraphContent::BookmarkStart(bookmark) => {
                    if let Some(name) = &bookmark.name {
                        segments.push(FormattedSegment {
                            anchor: Some(name.to_string()),
                            ..Default::default()
                        });
                    }
                }
                ParagraphContent::BookmarkEnd(_) => {}
                ParagraphContent::CommentRangeStart(_) => {}
                ParagraphContent::CommentRangeEnd(_) => {}
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
                ParagraphContent::Insertion(ins) => {
                    // Handle inserted content (track changes)
                    for run in &ins.runs {
                        let text = Self::extract_text(run, context);
                        if !text.is_empty() {
                            let mut segs = Self::run_to_segment(run, &text, context, para_style_id);
                            for seg in &mut segs {
                                seg.is_insertion = true;
                            }
                            segments.extend(segs);
                        }
                    }
                }
                ParagraphContent::Deletion(del) => {
                    // Handle deleted content (track changes)
                    let text = Self::extract_deleted_text(del);
                    if !text.is_empty() {
                        segments.push(FormattedSegment {
                            text,
                            is_deletion: true,
                            ..Default::default()
                        });
                    }
                }
            }
        }

        Ok(segments)
    }

    /// Extracts deleted text from a Deletion element.
    fn extract_deleted_text(del: &rs_docx::document::Deletion) -> String {
        let mut text = String::new();
        for run in &del.runs {
            for content in &run.content {
                if let rs_docx::document::RunContent::DelText(del_text) = content {
                    text.push_str(&del_text.text);
                }
            }
        }
        text
    }

    /// Extracts text from a run, excluding field codes.
    fn extract_text<'a>(
        run: &rs_docx::document::Run<'a>,
        context: &mut ConversionContext<'a>,
    ) -> String {
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
                rs_docx::document::RunContent::NoBreakHyphen(_) => {
                    text.push('\u{2011}');
                }
                rs_docx::document::RunContent::SoftHyphen(_) => {
                    text.push('\u{00AD}');
                }
                rs_docx::document::RunContent::Sym(sym) => {
                    if let Some(char_code) = &sym.char {
                        if let Ok(code) = u32::from_str_radix(char_code, 16) {
                            if let Some(c) = char::from_u32(code) {
                                text.push(c);
                            }
                        }
                    }
                }
                rs_docx::document::RunContent::PTab(_) => {
                    text.push('\t');
                }
                rs_docx::document::RunContent::LastRenderedPageBreak(_) => {
                    text.push_str("\n\n---\n\n");
                }
                rs_docx::document::RunContent::PgNum(_) => {
                    text.push_str("{PAGE}");
                }
                rs_docx::document::RunContent::Drawing(drawing) => {
                    if let Ok(Some(img_md)) = context.extract_image_from_drawing(drawing) {
                        text.push_str(&img_md);
                    }
                }
                rs_docx::document::RunContent::Pict(pict) => {
                    if let Ok(Some(img_md)) = context.extract_image_from_pict(pict) {
                        text.push_str(&img_md);
                    }
                }
                // Skip InstrText (field codes like TOC, PAGEREF)
                rs_docx::document::RunContent::InstrText(_) => {}
                rs_docx::document::RunContent::DelInstrText(_) => {}
                rs_docx::document::RunContent::CommentReference(cref) => {
                    // Extract comment ID and look up comment text
                    if let Some(id) = &cref.id {
                        let marker = context.register_comment_reference(id.as_ref());
                        text.push_str(&marker);
                    }
                }
                rs_docx::document::RunContent::FootnoteReference(fnref) => {
                    // Extract footnote ID and look up footnote text
                    if let Some(ref id_str) = fnref.id {
                        if let Ok(id_num) = id_str.parse::<isize>() {
                            let marker = context.register_footnote_reference(id_num);
                            text.push_str(&marker);
                        }
                    }
                }
                rs_docx::document::RunContent::EndnoteReference(enref) => {
                    // Extract endnote ID and look up endnote text
                    if let Some(ref id_str) = enref.id {
                        if let Ok(id_num) = id_str.parse::<isize>() {
                            let marker = context.register_endnote_reference(id_num);
                            text.push_str(&marker);
                        }
                    }
                }
                rs_docx::document::RunContent::AnnotationRef(_)
                | rs_docx::document::RunContent::FootnoteRef(_)
                | rs_docx::document::RunContent::EndnoteRef(_)
                | rs_docx::document::RunContent::Separator(_)
                | rs_docx::document::RunContent::ContinuationSeparator(_) => {}
                _ => {}
            }
        }
        text
    }

    /// Creates formatted segments from a run, splitting on page breaks.
    fn run_to_segment<'a>(
        run: &rs_docx::document::Run<'a>,
        text: &str,
        context: &mut ConversionContext<'a>,
        para_style_id: Option<&str>,
    ) -> Vec<FormattedSegment> {
        // Resolve run style ID
        let mut run_style_id = None;
        if let Some(props) = &run.property {
            if let Some(style) = &props.style_id {
                run_style_id = Some(style.value.as_ref());
            }
        }

        // Resolve effective properties
        let props =
            context.resolve_run_property(run.property.as_ref(), run_style_id, para_style_id);

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

        let delimiter = "\n\n---\n\n";
        let parts: Vec<&str> = text.split(delimiter).collect();
        let mut segments = Vec::new();

        for (i, part) in parts.iter().enumerate() {
            if i > 0 {
                // Add the break segment with no formatting
                segments.push(FormattedSegment {
                    text: delimiter.to_string(),
                    is_bold: false,
                    is_italic: false,
                    has_underline: false,
                    has_strike: false,
                    is_insertion: false,
                    is_deletion: false,
                    anchor: None,
                });
            }
            if !part.is_empty() {
                segments.push(FormattedSegment {
                    text: part.to_string(),
                    is_bold,
                    is_italic,
                    has_underline,
                    has_strike,
                    is_insertion: false,
                    is_deletion: false,
                    anchor: None,
                });
            }
        }

        segments
    }

    /// Merges adjacent segments with identical formatting.
    fn merge_segments(segments: Vec<FormattedSegment>) -> Vec<FormattedSegment> {
        let mut merged: Vec<FormattedSegment> = Vec::new();

        for seg in segments {
            if let Some(last) = merged.last_mut() {
                // Check if formatting matches (including track changes flags)
                if last.is_bold == seg.is_bold
                    && last.is_italic == seg.is_italic
                    && last.has_underline == seg.has_underline
                    && last.has_strike == seg.has_strike
                    && last.is_insertion == seg.is_insertion
                    && last.is_deletion == seg.is_deletion
                    && last.anchor == seg.anchor
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

    /// Applies markdown formatting markers safely, handling edge cases.
    ///
    /// Handles:
    /// - Empty or whitespace-only text (skips formatting)
    /// - Text with newlines (applies formatting per line)
    /// - Leading/trailing whitespace (preserves outside markers)
    fn apply_format_safely(text: &str, open: &str, close: &str) -> String {
        // Skip if text is empty or whitespace-only
        if text.trim().is_empty() {
            return text.to_string();
        }

        // Handle leading/trailing whitespace - preserve it outside the markers
        let leading_ws: String = text
            .chars()
            .take_while(|c| c.is_whitespace() && *c != '\n')
            .collect();
        let trailing_ws: String = text
            .chars()
            .rev()
            .take_while(|c| c.is_whitespace() && *c != '\n')
            .collect::<String>()
            .chars()
            .rev()
            .collect();

        let content_start = leading_ws.len();
        let content_end = text.len() - trailing_ws.len();
        let content = &text[content_start..content_end];

        // If content contains newlines, apply formatting to each non-empty line
        if content.contains('\n') {
            let formatted: Vec<String> = content
                .split('\n')
                .map(|line| {
                    let line_trimmed = line.trim();
                    if line_trimmed.is_empty() {
                        line.to_string()
                    } else {
                        // Preserve line's own leading/trailing whitespace
                        let line_leading: String =
                            line.chars().take_while(|c| c.is_whitespace()).collect();
                        let line_trailing: String = line
                            .chars()
                            .rev()
                            .take_while(|c| c.is_whitespace())
                            .collect::<String>()
                            .chars()
                            .rev()
                            .collect();
                        format!(
                            "{}{}{}{}{}",
                            line_leading, open, line_trimmed, close, line_trailing
                        )
                    }
                })
                .collect();
            return format!("{}{}{}", leading_ws, formatted.join("\n"), trailing_ws);
        }

        // Normal case: wrap content with markers, preserve outer whitespace
        format!(
            "{}{}{}{}{}",
            leading_ws,
            open,
            content.trim(),
            close,
            trailing_ws
        )
    }

    /// Converts segments to markdown text.
    fn segments_to_markdown(
        segments: &[FormattedSegment],
        context: &ConversionContext<'_>,
    ) -> String {
        let mut result = String::new();

        for seg in segments {
            // Render anchor if present
            if let Some(anchor) = &seg.anchor {
                result.push_str(&format!("<a id=\"{}\"></a>", anchor));
            }

            let mut text = seg.text.clone();

            // Apply track changes formatting first
            if seg.is_deletion {
                // Deleted text: strikethrough
                text = Self::apply_format_safely(&text, "~~", "~~");
            }
            if seg.is_insertion {
                // Inserted text: HTML ins tag or underline
                text = format!("<ins>{}</ins>", text);
            }

            // Apply regular formatting
            if seg.has_underline && context.html_underline_enabled() && !seg.is_insertion {
                text = format!("<u>{}</u>", text);
            }

            if seg.has_strike && !seg.is_deletion {
                if context.html_strikethrough_enabled() {
                    text = format!("<s>{}</s>", text);
                } else {
                    text = Self::apply_format_safely(&text, "~~", "~~");
                }
            }

            if seg.is_bold && seg.is_italic {
                text = format!("<strong><em>{}</em></strong>", text);
            } else if seg.is_bold {
                text = format!("<strong>{}</strong>", text);
            } else if seg.is_italic {
                text = format!("<em>{}</em>", text);
            }

            result.push_str(&text);
        }

        result
    }

    /// Applies paragraph-level formatting (heading, list, alignment).
    fn apply_paragraph_formatting<'a>(
        para: &Paragraph<'a>,
        text: String,
        context: &mut ConversionContext<'a>,
    ) -> Result<String> {
        let para_style_id = para
            .property
            .as_ref()
            .and_then(|p| p.style_id.as_ref())
            .map(|s| s.value.as_ref());

        // Resolve effective paragraph properties
        let effective_props =
            context.resolve_paragraph_property(para.property.as_ref(), para_style_id);

        let mut prefix = String::new();
        let mut is_heading = false;

        // Check for heading via pStyle
        if let Some(style) = &effective_props.style_id {
            if let Some(heading_level) = crate::localization::parse_heading_style(&style.value) {
                // Don't generate heading for empty text
                if text.trim().is_empty() {
                    return Ok(String::new());
                }
                prefix.push_str(&"#".repeat(heading_level));
                prefix.push(' ');
                is_heading = true;
            }
        }

        // Check for numbering (list items)
        if let Some(num_pr) = &effective_props.numbering {
            if let (Some(num_id), Some(ilvl)) = (&num_pr.id, &num_pr.level) {
                let num_id_val = num_id.value as i32;
                let ilvl_val = ilvl.value as i32;
                let marker = context.next_list_marker(num_id_val, ilvl_val);

                if is_heading {
                    prefix.push_str(&marker);
                    if !marker.is_empty() {
                        prefix.push(' ');
                    }
                } else {
                    let indent = context.list_indent_level(num_id_val, ilvl_val);
                    let indent_str = "  ".repeat(indent);
                    prefix.push_str(&indent_str);
                    prefix.push_str(&marker);
                    prefix.push(' ');
                }
            }
        }

        let text_for_output = if context.preserve_whitespace() {
            text.as_str()
        } else {
            text.trim()
        };
        let final_text = format!("{}{}", prefix, text_for_output);

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
    fn convert_hyperlink<'a>(
        hyperlink: &Hyperlink<'a>,
        context: &mut ConversionContext<'a>,
        para_style_id: Option<&str>,
    ) -> Result<String> {
        let mut link_text = String::new();
        let mut field_stack = Vec::new();

        for run in &hyperlink.content {
            let filtered_run = Self::filter_run_by_field_state(run, &mut field_stack);
            if filtered_run.content.is_empty() {
                continue;
            }

            let text = RunConverter::convert(&filtered_run, context, para_style_id)?;
            link_text.push_str(&text);
        }

        // Get target URL from relationship or anchor
        let url = if let Some(anchor) = &hyperlink.anchor {
            // Internal bookmark link (used in TOC entries)
            format!("#{}", escape_markdown_link_destination(anchor))
        } else if let Some(id) = &hyperlink.id {
            // External link via relationship
            context
                .relationship_target(id.as_ref())
                .map(str::to_owned)
                .unwrap_or_else(|| "#".to_string())
        } else {
            "#".to_string()
        };

        if link_text.is_empty() {
            Ok(url)
        } else {
            Ok(format!(
                "[{}]({})",
                escape_markdown_link_text(&link_text),
                escape_markdown_link_destination(&url)
            ))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rs_docx::document::{Hyperlink, ParagraphContent, Run, RunContent, Text};
    use std::borrow::Cow;
    use std::collections::HashMap;

    #[test]
    fn test_toc_anchor_link() {
        // Create a paragraph with a hyperlink having an anchor
        let mut para = Paragraph::default();

        let mut hyperlink = Hyperlink {
            anchor: Some(Cow::Borrowed("_Toc123456789")),
            ..Default::default()
        };

        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Introduction".into(),
            ..Default::default()
        }));

        hyperlink.content.push(run);

        para.content.push(ParagraphContent::Link(hyperlink));

        // Setup minimal context
        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        // Convert
        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");

        // Verify
        assert_eq!(md, "[Introduction](#_Toc123456789)");
    }

    #[test]
    fn test_toc_anchor_target() {
        use rs_docx::document::BookmarkStart;

        // Create a paragraph with a bookmark start (anchor target)
        let mut para = Paragraph::default();

        let bookmark = BookmarkStart {
            name: Some(Cow::Borrowed("_Toc123456789")),
            ..Default::default()
        };

        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Chapter 1".into(),
            ..Default::default()
        }));

        para.content.push(ParagraphContent::BookmarkStart(bookmark));
        para.content.push(ParagraphContent::Run(run));

        // Setup minimal context
        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        // Convert
        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");

        // Verify that the anchor tag is generated BEFORE the text (on new line)
        assert_eq!(md, "<a id=\"_Toc123456789\"></a>\nChapter 1");
    }

    #[test]
    fn test_anchor_placement_header() {
        use rs_docx::document::BookmarkStart;

        // Create a paragraph with Heading 1 style and a bookmark
        let mut para = Paragraph::default();
        let props = rs_docx::formatting::ParagraphProperty {
            style_id: Some(rs_docx::formatting::ParagraphStyleId {
                value: "Heading1".into(),
            }),
            ..Default::default()
        };
        para.property = Some(props);

        let bookmark = BookmarkStart {
            name: Some(Cow::Borrowed("header_anchor")),
            ..Default::default()
        };

        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Header Title".into(),
            ..Default::default()
        }));

        para.content.push(ParagraphContent::BookmarkStart(bookmark));
        para.content.push(ParagraphContent::Run(run));

        // Setup mock context
        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        // Convert
        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");

        // Verify: Anchor should be on the line BEFORE the header
        // Expected: "<a id=\"header_anchor\"></a>\n# Header Title"
        assert_eq!(md, "<a id=\"header_anchor\"></a>\n# Header Title");
    }

    #[test]
    fn test_adjacent_anchors() {
        use rs_docx::document::BookmarkStart;

        // Create a paragraph with multiple adjacent bookmarks
        let mut para = Paragraph::default();

        let b1 = BookmarkStart {
            name: Some(Cow::Borrowed("anchor1")),
            ..Default::default()
        };
        let b2 = BookmarkStart {
            name: Some(Cow::Borrowed("anchor2")),
            ..Default::default()
        };

        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Content".into(),
            ..Default::default()
        }));

        para.content.push(ParagraphContent::BookmarkStart(b1));
        para.content.push(ParagraphContent::BookmarkStart(b2));
        para.content.push(ParagraphContent::Run(run));

        // Setup minimal context
        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        // Convert
        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");

        // Verify both anchors are present
        assert_eq!(md, "<a id=\"anchor1\"></a><a id=\"anchor2\"></a>\nContent");
    }

    #[test]
    fn test_preserve_whitespace_option() {
        let mut para = Paragraph::default();
        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "  Keep Surrounding Spaces  ".into(),
            ..Default::default()
        }));
        para.content.push(ParagraphContent::Run(run));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions {
            preserve_whitespace: true,
            ..Default::default()
        };
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "  Keep Surrounding Spaces  ");
    }

    #[test]
    fn test_deep_list_indentation_not_clamped() {
        use rs_docx::document::{
            AbstractNum, AbstractNumId, Level, LevelStart, LevelText, Num, NumFmt, Numbering,
        };

        let mut para = Paragraph::default();
        let props = rs_docx::formatting::ParagraphProperty {
            numbering: Some(rs_docx::formatting::NumberingProperty::from((
                2isize, 3isize,
            ))),
            ..Default::default()
        };
        para.property = Some(props);

        let mut run = Run::default();
        run.content.push(RunContent::Text(Text {
            text: "Deep Item".into(),
            ..Default::default()
        }));
        para.content.push(ParagraphContent::Run(run));

        let abstract_num = AbstractNum {
            abstract_num_id: Some(1),
            levels: vec![Level {
                i_level: Some(3),
                start: Some(LevelStart { value: Some(1) }),
                number_format: Some(NumFmt {
                    value: Cow::Borrowed("decimal"),
                }),
                level_text: Some(LevelText {
                    value: Some(Cow::Borrowed("%4.")),
                }),
                ..Default::default()
            }],
            ..Default::default()
        };
        let num = Num {
            num_id: Some(2),
            abstract_num_id: Some(AbstractNumId { value: Some(1) }),
            ..Default::default()
        };
        let docx = rs_docx::Docx {
            numbering: Some(Numbering {
                abstract_numberings: vec![abstract_num],
                numberings: vec![num],
            }),
            ..Default::default()
        };

        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "      1. Deep Item");
    }

    #[test]
    fn test_duplicate_footnote_references_reuse_index() {
        use rs_docx::document::{BodyContent, FootNote, FootNotes, FootnoteReference};

        let mut note_para = Paragraph::default();
        let mut note_run = Run::default();
        note_run.content.push(RunContent::Text(Text {
            text: "Same footnote text".into(),
            ..Default::default()
        }));
        note_para.content.push(ParagraphContent::Run(note_run));

        let docx = rs_docx::Docx {
            footnotes: Some(FootNotes {
                content: vec![FootNote {
                    id: Some(5),
                    content: vec![BodyContent::Paragraph(note_para)],
                    ..Default::default()
                }],
            }),
            ..Default::default()
        };

        let mut para = Paragraph::default();
        let mut run1 = Run::default();
        run1.content
            .push(RunContent::FootnoteReference(FootnoteReference {
                id: Some(Cow::Borrowed("5")),
                ..Default::default()
            }));
        para.content.push(ParagraphContent::Run(run1));
        let mut run2 = Run::default();
        run2.content
            .push(RunContent::FootnoteReference(FootnoteReference {
                id: Some(Cow::Borrowed("5")),
                ..Default::default()
            }));
        para.content.push(ParagraphContent::Run(run2));

        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            docx.footnotes.as_ref(),
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "[^1][^1]");
        assert_eq!(context.footnote_count(), 1);
    }

    #[test]
    fn test_duplicate_comment_references_reuse_definition() {
        use rs_docx::document::{Comment, CommentReference, Comments};

        let mut comment_para = Paragraph::default();
        let mut comment_run = Run::default();
        comment_run.content.push(RunContent::Text(Text {
            text: "Shared comment".into(),
            ..Default::default()
        }));
        comment_para
            .content
            .push(ParagraphContent::Run(comment_run));

        let docx = rs_docx::Docx {
            comments: Some(Comments {
                comments: vec![Comment {
                    id: Some(9),
                    author: Cow::Borrowed("tester"),
                    content: comment_para,
                }],
            }),
            ..Default::default()
        };

        let mut para = Paragraph::default();
        let mut run1 = Run::default();
        run1.content
            .push(RunContent::CommentReference(CommentReference {
                id: Some(Cow::Borrowed("9")),
            }));
        para.content.push(ParagraphContent::Run(run1));
        let mut run2 = Run::default();
        run2.content
            .push(RunContent::CommentReference(CommentReference {
                id: Some(Cow::Borrowed("9")),
            }));
        para.content.push(ParagraphContent::Run(run2));

        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            docx.comments.as_ref(),
            None,
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "[^c9][^c9]");
        assert_eq!(context.comment_count(), 1);
        assert_eq!(context.comment_at(0), Some(("9", "Shared comment")));
    }

    #[test]
    fn test_field_code_within_single_run_preserves_visible_text() {
        use hard_xml::XmlRead;

        let mut para = Paragraph::default();
        let run = Run::from_str(
            r#"<w:r>
                <w:t>prefix </w:t>
                <w:fldChar w:fldCharType="begin"/>
                <w:instrText>PAGEREF _Ref</w:instrText>
                <w:t>hidden </w:t>
                <w:fldChar w:fldCharType="separate"/>
                <w:t>Visible</w:t>
                <w:fldChar w:fldCharType="end"/>
                <w:t> suffix</w:t>
            </w:r>"#,
        )
        .expect("Failed to parse run XML");
        para.content.push(ParagraphContent::Run(run));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "prefix Visible suffix");
    }

    #[test]
    fn test_extended_run_content_is_preserved() {
        use hard_xml::XmlRead;

        let mut para = Paragraph::default();
        let run = Run::from_str(
            r#"<w:r>
                <w:t>A</w:t>
                <w:noBreakHyphen/>
                <w:t>B</w:t>
                <w:softHyphen/>
                <w:t>C</w:t>
                <w:sym w:char="2013"/>
                <w:ptab/>
                <w:lastRenderedPageBreak/>
                <w:pgNum/>
                <w:t>D</w:t>
            </w:r>"#,
        )
        .expect("Failed to parse run XML");
        para.content.push(ParagraphContent::Run(run));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = crate::ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);

        let mut context = super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let md = ParagraphConverter::convert(&para, &mut context).expect("Conversion failed");
        assert_eq!(md, "A\u{2011}B\u{00AD}C\u{2013}\t\n\n---\n\n{PAGE}D");
    }
}

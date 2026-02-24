//! Run element converter - handles text runs with formatting.

use super::ConversionContext;
use crate::Result;
use rs_docx::document::{BreakType, Run, RunContent};

/// Converter for Run elements.
pub struct RunConverter;

impl RunConverter {
    /// Converts a Run to Markdown text with formatting.
    pub fn convert<'a>(
        run: &Run<'a>,
        context: &mut ConversionContext<'a>,
        para_style_id: Option<&str>,
    ) -> Result<String> {
        let mut text = String::new();

        // Extract text from run content
        for content in &run.content {
            match content {
                RunContent::Text(t) => {
                    text.push_str(&t.text);
                }
                RunContent::Break(br) => match br.ty {
                    Some(BreakType::Page) => text.push_str("\n\n---\n\n"),
                    Some(BreakType::Column) => text.push_str("\n\n"),
                    _ => text.push('\n'),
                },
                RunContent::Tab(_) => {
                    text.push('\t');
                }
                RunContent::CarriageReturn(_) => {
                    text.push('\n');
                }
                RunContent::NoBreakHyphen(_) => {
                    text.push('\u{2011}');
                }
                RunContent::SoftHyphen(_) => {
                    text.push('\u{00AD}');
                }
                RunContent::Drawing(drawing) => {
                    // Handle inline images (DrawingML)
                    if let Some(img_md) = context.extract_image_from_drawing(drawing)? {
                        text.push_str(&img_md);
                    }
                }
                RunContent::Pict(pict) => {
                    // Handle legacy images (VML)
                    if let Some(img_md) = context.extract_image_from_pict(pict)? {
                        text.push_str(&img_md);
                    }
                }
                RunContent::Sym(sym) => {
                    // Symbol character - use Unicode if possible
                    if let Some(char_code) = &sym.char {
                        // Try to decode hex char code
                        if let Ok(code) = u32::from_str_radix(char_code, 16) {
                            if let Some(c) = char::from_u32(code) {
                                text.push(c);
                            }
                        }
                    }
                }
                RunContent::FootnoteReference(fnref) => {
                    if let Some(id_str) = &fnref.id {
                        if let Ok(id_num) = id_str.parse::<isize>() {
                            let marker = context.register_footnote_reference(id_num);
                            text.push_str(&marker);
                        }
                    }
                }
                RunContent::EndnoteReference(enref) => {
                    if let Some(id_str) = &enref.id {
                        if let Ok(id_num) = id_str.parse::<isize>() {
                            let marker = context.register_endnote_reference(id_num);
                            text.push_str(&marker);
                        }
                    }
                }
                RunContent::CommentReference(cref) => {
                    // Extract comment ID and look up comment text
                    if let Some(id) = &cref.id {
                        let marker = context.register_comment_reference(id.as_ref());
                        text.push_str(&marker);
                    }
                }
                RunContent::PTab(_) => {
                    text.push('\t');
                }
                RunContent::LastRenderedPageBreak(_) => {
                    text.push_str("\n\n---\n\n");
                }
                RunContent::PgNum(_) => {
                    text.push_str("{PAGE}");
                }
                RunContent::AnnotationRef(_)
                | RunContent::FootnoteRef(_)
                | RunContent::EndnoteRef(_)
                | RunContent::Separator(_)
                | RunContent::ContinuationSeparator(_) => {}
                _ => {}
            }
        }

        // Apply formatting if text is not empty
        if text.is_empty() {
            return Ok(text);
        }

        // Run Style ID
        let mut run_style_id = None;
        if let Some(props) = &run.property {
            if let Some(style) = &props.style_id {
                run_style_id = Some(style.value.as_ref());
            }
        }

        // Check formatting via resolver
        let effective_props =
            context.resolve_run_property(run.property.as_ref(), run_style_id, para_style_id);

        text = Self::apply_formatting(&text, &effective_props, context);

        Ok(text)
    }

    /// Applies text formatting based on run properties.
    fn apply_formatting(
        text: &str,
        props: &rs_docx::formatting::CharacterProperty<'_>,
        context: &ConversionContext<'_>,
    ) -> String {
        let mut result = text.to_string();

        // Check for bold
        let is_bold = props
            .bold
            .as_ref()
            .map(|b| b.value.unwrap_or(true))
            .unwrap_or(false);

        // Check for italic
        let is_italic = props
            .italics
            .as_ref()
            .map(|i| i.value.unwrap_or(true))
            .unwrap_or(false);

        // Check for underline
        let has_underline = props.underline.is_some();

        // Check for strikethrough
        let has_strike = props
            .strike
            .as_ref()
            .map(|s| s.value.unwrap_or(true))
            .unwrap_or(false);

        // Apply formatting in order: underline (HTML), strike, bold, italic
        if has_underline && context.html_underline_enabled() {
            result = format!("<u>{}</u>", result);
        }

        if has_strike {
            if context.html_strikethrough_enabled() {
                result = format!("<s>{}</s>", result);
            } else {
                result = format!("~~{}~~", result);
            }
        }

        if is_bold && is_italic {
            result = format!("<strong>*{}*</strong>", result);
        } else if is_bold {
            result = format!("<strong>{}</strong>", result);
        } else if is_italic {
            result = format!("*{}*", result);
        }

        result
    }
}

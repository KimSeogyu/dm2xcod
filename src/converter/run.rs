//! Run element converter - handles text runs with formatting.

use super::ConversionContext;
use crate::Result;
use rs_docx::document::{BreakType, Run, RunContent};

/// Converter for Run elements.
pub struct RunConverter;

impl RunConverter {
    /// Converts a Run to Markdown text with formatting.
    pub fn convert(
        run: &Run,
        context: &mut ConversionContext,
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
                RunContent::Drawing(drawing) => {
                    // Handle inline images (DrawingML)
                    if let Some(img_md) = context
                        .image_extractor
                        .extract_from_drawing(drawing, context.rels)?
                    {
                        text.push_str(&img_md);
                    }
                }
                RunContent::Pict(pict) => {
                    // Handle legacy images (VML)
                    if let Some(img_md) = context
                        .image_extractor
                        .extract_from_pict(pict, context.rels)?
                    {
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
                RunContent::FootnoteReference(_fnref) => {
                    let idx = context.footnotes.len() + 1;
                    text.push_str(&format!("[^{}]", idx));
                }
                RunContent::EndnoteReference(_enref) => {
                    let idx = context.endnotes.len() + 1;
                    text.push_str(&format!("[^en{}]", idx));
                }
                RunContent::CommentReference(cref) => {
                    // Extract comment ID and look up comment text
                    if let Some(id) = &cref.id {
                        let id_str = id.to_string();
                        // Look up comment content
                        if let Some(comments) = context.docx_comments {
                            if let Some(comment) = comments
                                .comments
                                .iter()
                                .find(|c| c.id.map(|i| i.to_string()) == Some(id_str.clone()))
                            {
                                // Extract text from comment paragraph
                                let comment_text = comment.content.text();
                                context.comments.push((id_str.clone(), comment_text));
                            }
                        }
                        text.push_str(&format!("[^c{}]", id_str));
                    }
                }
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
        let effective_props = context.style_resolver.resolve_run_property(
            run.property.as_ref(),
            run_style_id,
            para_style_id,
        );

        text = Self::apply_formatting(&text, &effective_props, context);

        Ok(text)
    }

    /// Applies text formatting based on run properties.
    fn apply_formatting(
        text: &str,
        props: &rs_docx::formatting::CharacterProperty,
        context: &ConversionContext,
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
        if has_underline && context.options.html_underline {
            result = format!("<u>{}</u>", result);
        }

        if has_strike {
            if context.options.html_strikethrough {
                result = format!("<s>{}</s>", result);
            } else {
                result = format!("~~{}~~", result);
            }
        }

        if is_bold && is_italic {
            result = format!("***{}***", result);
        } else if is_bold {
            result = format!("**{}**", result);
        } else if is_italic {
            result = format!("*{}*", result);
        }

        result
    }
}

//! Style resolver - handles style inheritance and property merging.

use rs_docx::formatting::{CharacterProperty, ParagraphProperty};
use rs_docx::styles::Style;
use std::collections::HashMap;

/// Resolver for DOCX styles and inheritance.
pub struct StyleResolver<'a> {
    styles: &'a rs_docx::styles::Styles<'a>,
    style_map: HashMap<&'a str, &'a Style<'a>>,
}

impl<'a> StyleResolver<'a> {
    pub fn new(styles: &'a rs_docx::styles::Styles<'a>) -> Self {
        let mut style_map = HashMap::new();
        for style in &styles.styles {
            style_map.insert(style.style_id.as_ref(), style);
        }
        Self { styles, style_map }
    }

    /// Resolves the effective character properties for a run.
    ///
    /// Hierarchy (highest priority first):
    /// 1. Direct formatting on the run (rPr)
    /// 2. Character style applied to the run (rStyle) and its ancestors
    /// 3. Paragraph style applied to the paragraph (pStyle) and its ancestors
    /// 4. Document defaults (docDefaults)
    pub fn resolve_run_property(
        &self,
        direct_props: Option<&CharacterProperty<'a>>,
        run_style_id: Option<&str>,
        para_style_id: Option<&str>,
    ) -> CharacterProperty<'a> {
        let mut merged = CharacterProperty::default();

        // 4. Document defaults
        if let Some(defaults) = &self.styles.default {
            if let Some(r_pr) = &defaults.character.inner {
                merge_char_props_mut(&mut merged, r_pr);
            }
        }

        // 3. Paragraph style chain (only if para_style_id is provided)
        if let Some(pid) = para_style_id {
            self.apply_style_chain_char(&mut merged, pid);
        }

        // 2. Character style chain (only if run_style_id is provided)
        if let Some(rid) = run_style_id {
            self.apply_style_chain_char(&mut merged, rid);
        }

        // 1. Direct formatting
        if let Some(direct) = direct_props {
            merge_char_props_mut(&mut merged, direct);
        }

        merged
    }

    /// Resolves the effective paragraph properties.
    pub fn resolve_paragraph_property(
        &self,
        direct_props: Option<&ParagraphProperty<'a>>,
        para_style_id: Option<&str>,
    ) -> ParagraphProperty<'a> {
        let mut merged = ParagraphProperty::default();

        // Defaults
        if let Some(defaults) = &self.styles.default {
            if let Some(p_pr) = &defaults.paragraph.inner {
                merge_para_props_mut(&mut merged, p_pr);
            }
        }

        // Style chain
        if let Some(pid) = para_style_id {
            self.apply_style_chain_para(&mut merged, pid);
        }

        // Direct
        if let Some(direct) = direct_props {
            merge_para_props_mut(&mut merged, direct);
        }

        merged
    }

    fn apply_style_chain_char(&self, target: &mut CharacterProperty<'a>, style_id: &str) {
        // Collect chain to apply from root to leaf (base -> derived)
        // because we want derived styles to override base styles.
        // However, the `merge_char_props` function assumes `target` is the accumulator (lower priority)
        // and overrides it with the new props (higher priority).
        // Wait, standard merge pattern: `base.merge(overlay)`.
        // So we should start with Defaults (base), then apply Base Style, then Derived Style, then Direct.
        // The `resolve_run_property` already initializes `target` with Defaults.
        // So here we should apply styles from Base to Derived (Leaf).
        // BUT, since we are doing `target = merge(target, new)`, where `new` overrides `target`,
        // we should apply base styles first, then derived styles.
        //
        // Let's implement an iterator or recursion to go up to the root, then unwind.

        let mut chain = Vec::new();
        let mut current_id = Some(style_id);

        while let Some(id) = current_id {
            if let Some(style) = self.style_map.get(id) {
                chain.push(style);
                current_id = style.base.as_ref().map(|b| b.value.as_ref());
            } else {
                break;
            }
        }

        // Apply from root (most generic) to leaf (most specific)
        for style in chain.into_iter().rev() {
            if let Some(r_pr) = &style.character {
                merge_char_props_mut(target, r_pr);
            }
        }
    }

    fn apply_style_chain_para(&self, target: &mut ParagraphProperty<'a>, style_id: &str) {
        let mut chain = Vec::new();
        let mut current_id = Some(style_id);

        while let Some(id) = current_id {
            if let Some(style) = self.style_map.get(id) {
                chain.push(style);
                current_id = style.base.as_ref().map(|b| b.value.as_ref());
            } else {
                break;
            }
        }

        for style in chain.into_iter().rev() {
            if let Some(p_pr) = &style.paragraph {
                merge_para_props_mut(target, p_pr);
            }
        }
    }
}

// Helper to merge character properties (in-place mutation)
fn merge_char_props_mut<'a>(target: &mut CharacterProperty<'a>, overlay: &CharacterProperty<'a>) {
    if overlay.bold.is_some() {
        target.bold = overlay.bold.clone();
    }
    if overlay.italics.is_some() {
        target.italics = overlay.italics.clone();
    }
    if overlay.strike.is_some() {
        target.strike = overlay.strike.clone();
    }
    if overlay.underline.is_some() {
        target.underline = overlay.underline.clone();
    }
}

// Helper to merge paragraph properties (in-place mutation)
fn merge_para_props_mut<'a>(target: &mut ParagraphProperty<'a>, overlay: &ParagraphProperty<'a>) {
    if overlay.justification.is_some() {
        target.justification = overlay.justification.clone();
    }
    if overlay.numbering.is_some() {
        target.numbering = overlay.numbering.clone();
    }
    if overlay.style_id.is_some() {
        target.style_id = overlay.style_id.clone();
    }
}

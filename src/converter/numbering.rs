//! Numbering resolver - handles list numbering and indentation.

use rs_docx::Docx;
use std::collections::HashMap;

/// Resolver for DOCX numbering definitions.
pub struct NumberingResolver<'a> {
    /// Maps numId -> abstractNumId
    num_instances: HashMap<i32, i32>,
    /// Maps abstractNumId -> level definitions
    abstract_nums: HashMap<i32, Vec<LevelDef>>,
    /// Maps (numId, ilvl) -> startOverride value
    overrides: HashMap<(i32, i32), i32>,
    /// Maps abstractNumId -> current counters (one counter per level 0..9)
    /// Using abstractNumId allows continuous numbering even if numId changes (e.g. broken lists)
    counters: HashMap<i32, Vec<i32>>,
    /// Maps abstractNumId -> base indentation level (shift)
    /// Used to normalize indentation for lists that start at high levels (e.g., Article at Level 4)
    level_shifts: HashMap<i32, i32>,
    _phantom: std::marker::PhantomData<&'a ()>,
}

#[derive(Clone, Debug)]
struct LevelDef {
    ilvl: i32,
    start: i32,
    num_fmt: String,
    lvl_text: Option<String>,
}

impl<'a> NumberingResolver<'a> {
    /// Creates a new numbering resolver from a parsed DOCX.
    pub fn new(docx: &'a Docx) -> Self {
        let mut num_instances = HashMap::new();
        let mut abstract_nums = HashMap::new();
        let mut overrides = HashMap::new();
        let mut level_shifts = HashMap::new();

        if let Some(numbering) = &docx.numbering {
            // Parse abstract numbering definitions
            for abs_num in &numbering.abstract_numberings {
                let abs_id = abs_num.abstract_num_id.map(|id| id as i32).unwrap_or(0);
                let mut levels = Vec::new();

                for lvl in &abs_num.levels {
                    let ilvl = lvl.i_level.map(|i| i as i32).unwrap_or(0);
                    let start = lvl
                        .start
                        .as_ref()
                        .and_then(|s| s.value)
                        .map(|v| v as i32)
                        .unwrap_or(1);
                    let num_fmt = lvl
                        .number_format
                        .as_ref()
                        .map(|f| f.value.to_string())
                        .unwrap_or_else(|| "decimal".to_string());
                    let lvl_text = lvl
                        .level_text
                        .as_ref()
                        .and_then(|t| t.value.as_ref())
                        .map(|v| v.to_string());

                    // Heuristic: If this level looks like an "Article" heading (제%1조),
                    // treat it as a base level (Level 0 equivalent) for indentation.
                    if let Some(text) = &lvl_text {
                        if text.contains("제") && text.contains("조") && text.contains("%") {
                            // Only set if not already set (prefer higher levels if multiple? No, prefer shallowest)
                            // But Article usually wraps everything.
                            level_shifts.entry(abs_id).or_insert(ilvl);
                        }
                    }

                    levels.push(LevelDef {
                        ilvl,
                        start,
                        num_fmt,
                        lvl_text,
                    });
                }

                levels.sort_by_key(|l| l.ilvl);
                abstract_nums.insert(abs_id, levels);
            }

            // Parse numbering instances
            for num in &numbering.numberings {
                if let (Some(num_id), Some(abs_ref)) = (num.num_id, &num.abstract_num_id) {
                    let nid = num_id as i32;
                    if let Some(abs_id) = abs_ref.value {
                        num_instances.insert(nid, abs_id as i32);
                    }

                    // Parse level overrides
                    for override_def in &num.level_overrides {
                        if let (Some(ilvl), Some(start_override)) =
                            (override_def.i_level, &override_def.start_override)
                        {
                            if let Some(val) = start_override.value {
                                overrides.insert((nid, ilvl as i32), val as i32);
                            }
                        }
                    }
                }
            }
        }

        Self {
            num_instances,
            abstract_nums,
            overrides,
            counters: HashMap::new(),
            level_shifts,
            _phantom: std::marker::PhantomData,
        }
    }

    /// Gets the indentation level for a list item.
    pub fn get_indent(&self, num_id: i32, ilvl: i32) -> usize {
        let mut indent = ilvl;

        // Apply shift if exists for this abstract numbering
        if let Some(&abs_id) = self.num_instances.get(&num_id) {
            if let Some(&base_level) = self.level_shifts.get(&abs_id) {
                indent = indent.saturating_sub(base_level);
            }
        }

        indent as usize
    }

    /// Gets the marker for a list item (e.g., "1.", "-", "a)").
    /// Updates the internal counter state.
    pub fn next_marker(&mut self, num_id: i32, ilvl: i32) -> String {
        let Some(&abs_id) = self.num_instances.get(&num_id) else {
            return "-".to_string();
        };

        let Some(levels) = self.abstract_nums.get(&abs_id) else {
            return "-".to_string();
        };

        // Initialize counters for this abstract_num_id if not present
        // Use abstract_id as key to share state across different num_ids for same style
        let counters = self.counters.entry(abs_id).or_insert_with(|| vec![0; 10]);

        // Find level definition
        let level_def = levels
            .iter()
            .find(|l| l.ilvl == ilvl)
            .or_else(|| levels.first());

        let Some(level) = level_def else {
            return "-".to_string();
        };

        // Increment current level
        let ilvl_idx = ilvl as usize;
        if counters.len() <= ilvl_idx {
            counters.resize(ilvl_idx + 1, 0);
        }

        // Determine start value (check override for specific instance first)
        let override_start = self.overrides.get(&(num_id, ilvl)).copied();

        // Update logic:
        // If counter is 0 (uninitialized), init it.
        // OR if there is an explicit OVERRIDE for this specific num_id instance, apply it.
        if counters[ilvl_idx] == 0 {
            counters[ilvl_idx] = override_start.unwrap_or(level.start);
        } else {
            counters[ilvl_idx] += 1;
        }

        // Reset lower levels
        for i in (ilvl_idx + 1)..counters.len() {
            counters[i] = 0;
        }

        // Use level text if available (substituting placeholders)
        if let Some(text) = &level.lvl_text {
            let mut marker = text.clone();
            // Replace %1, %2, etc. with formatted numbers
            for (i, count) in counters.iter().enumerate() {
                let level_num = i + 1; // %1 is index 0
                let placeholder = format!("%{}", level_num);
                if marker.contains(&placeholder) {
                    // Find formatting for this level
                    let fmt = levels
                        .iter()
                        .find(|l| l.ilvl == i as i32)
                        .map(|l| l.num_fmt.as_str())
                        .unwrap_or("decimal");

                    // If count is 0, it means it hasn't been initialized/incremented yet, so use start value
                    let val = if *count == 0 { 1 } else { *count };

                    let formatted_num = Self::format_num(fmt, val);
                    marker = marker.replace(&placeholder, &formatted_num);
                }
            }
            return marker;
        }

        // Fallback: if no lvlText, add dot for standard types
        let raw_num = Self::format_num(&level.num_fmt, counters[ilvl_idx]);
        match level.num_fmt.as_str() {
            "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" => {
                format!("{}.", raw_num)
            }
            _ => raw_num,
        }
    }

    /// Formats a number according to the format string.
    fn format_num(fmt: &str, val: i32) -> String {
        match fmt {
            "bullet" | "none" => "-".to_string(),
            "decimal" => format!("{}", val),
            "lowerLetter" => {
                if val >= 1 && val <= 26 {
                    char::from(b'a' + (val - 1) as u8).to_string()
                } else {
                    format!("{}", val)
                }
            }
            "upperLetter" => {
                if val >= 1 && val <= 26 {
                    char::from(b'A' + (val - 1) as u8).to_string()
                } else {
                    format!("{}", val)
                }
            }
            "lowerRoman" => Self::to_roman(val).to_lowercase(),
            "upperRoman" => Self::to_roman(val),
            "koreanCounting" | "korean" | "ganada" => Self::format_ganada(val),
            "chosung" => Self::format_chosung(val),
            "geonodeo" => Self::format_geonodeo(val),
            "decimalEnclosedCircle" => Self::format_circle_number(val),
            _ => format!("{}", val),
        }
    }

    /// Converts a number to circled number (①②③...).
    fn format_circle_number(val: i32) -> String {
        // Unicode circled numbers: ① = U+2460, ② = U+2461, ... ⑳ = U+2473
        // Extended: ㉑ = U+3251, ㉒ = U+3252, ... ㊿ = U+32BF (21-50)
        if val >= 1 && val <= 20 {
            char::from_u32(0x245F + val as u32)
                .map(|c| c.to_string())
                .unwrap_or_else(|| format!("{}", val))
        } else if val >= 21 && val <= 50 {
            char::from_u32(0x3250 + (val - 20) as u32)
                .map(|c| c.to_string())
                .unwrap_or_else(|| format!("{}", val))
        } else {
            format!("{}", val) // Fallback for numbers outside supported range
        }
    }

    /// Converts a number to Korean Ganada (가, 나, 다...).
    fn format_ganada(val: i32) -> String {
        let chars = [
            '가', '나', '다', '라', '마', '바', '사', '아', '자', '차', '카', '타', '파', '하',
        ];
        if val >= 1 && val as usize <= chars.len() {
            chars[(val - 1) as usize].to_string()
        } else {
            format!("{}", val) // Fallback
        }
    }

    /// Converts a number to Korean Geonodeo (거, 너, 더...).
    fn format_geonodeo(val: i32) -> String {
        let chars = [
            '거', '너', '더', '러', '머', '버', '서', '어', '저', '처', '커', '터', '퍼', '허',
        ];
        if val >= 1 && val as usize <= chars.len() {
            chars[(val - 1) as usize].to_string()
        } else {
            format!("{}", val)
        }
    }

    /// Converts a number to Korean Chosung (ㄱ, ㄴ, ㄷ...).
    fn format_chosung(val: i32) -> String {
        let chars = [
            'ㄱ', 'ㄴ', 'ㄷ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅅ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ',
        ];
        if val >= 1 && val as usize <= chars.len() {
            chars[(val - 1) as usize].to_string()
        } else {
            format!("{}", val) // Fallback
        }
    }

    /// Converts a number to Roman numeral.
    fn to_roman(mut num: i32) -> String {
        const ROMAN_TABLE: &[(i32, &str)] = &[
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ];

        if num <= 0 {
            return num.to_string();
        }
        let mut result = String::new();
        for &(v, s) in ROMAN_TABLE {
            while num >= v {
                result.push_str(s);
                num -= v;
            }
        }
        result
    }
}

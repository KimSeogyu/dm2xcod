//! Table converter - converts tables to HTML with merge support.

use super::table_grid;
use super::{ConversionContext, ParagraphConverter};
use crate::Result;
use rs_docx::document::{Table, TableCell, TableCellContent};

/// Converter for Table elements.
pub struct TableConverter;

impl TableConverter {
    /// Converts a Table to HTML format with correct merge handling.
    pub fn convert<'a>(table: &Table<'a>, context: &mut ConversionContext<'a>) -> Result<String> {
        let grid = table_grid::build_grid(table, |cell| Self::convert_cell_content(cell, context))?;
        Ok(table_grid::render_grid(grid))
    }

    fn convert_cell_content<'a>(
        cell: &TableCell<'a>,
        context: &mut ConversionContext<'a>,
    ) -> Result<String> {
        let mut content = String::new();
        for item in &cell.content {
            match item {
                TableCellContent::Paragraph(para) => {
                    let para_content = ParagraphConverter::convert(para, context)?;
                    if !para_content.is_empty() {
                        if !content.is_empty() {
                            content.push_str("<br/>");
                        }
                        content.push_str(&para_content);
                    }
                }
                TableCellContent::Table(table) => {
                    let table_content = TableConverter::convert(table, context)?;
                    content.push_str(&table_content);
                }
            }
        }
        Ok(content)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ConvertOptions;
    use rs_docx::document::{BodyContent, Paragraph, SDTContent, Table, TableCell, TableRow};
    use rs_docx::formatting::{GridSpan, TableCellProperty, VMerge, VMergeType};
    use std::collections::HashMap;

    #[test]
    fn test_vmerge_continuation_on_merged_left_column_increments_master_rowspan() {
        let top_master = TableCell::paragraph(Paragraph::default().push_text("TOP")).property(
            TableCellProperty::default()
                .grid_span(GridSpan { val: 2 })
                .v_merge(VMerge {
                    val: Some(VMergeType::Restart),
                }),
        );
        let left = TableCell::paragraph(Paragraph::default().push_text("L"));
        let cont = TableCell::paragraph(Paragraph::default()).property(
            TableCellProperty::default().v_merge(VMerge {
                val: Some(VMergeType::Continue),
            }),
        );

        let table = Table::default()
            .push_row(TableRow::default().push_cell(top_master))
            .push_row(TableRow::default().push_cell(left).push_cell(cont));

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);
        let mut context = super::super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let html = TableConverter::convert(&table, &mut context).expect("table conversion failed");
        assert!(html.contains("<td rowspan=\"2\" colspan=\"2\">TOP</td>"));
        assert!(html.contains("<td>L</td>"));
    }

    #[test]
    fn test_table_row_sdt_with_table_cell_is_rendered() {
        let mut sdt_content = SDTContent::default();
        sdt_content
            .content
            .push(BodyContent::TableCell(TableCell::paragraph(
                Paragraph::default().push_text("SDT-CELL"),
            )));
        let sdt = rs_docx::document::SDT::default().content(sdt_content);

        let mut row = TableRow::default();
        row.cells.push(rs_docx::document::TableRowContent::SDT(sdt));

        let table = Table::default().push_row(row);

        let docx = rs_docx::Docx::default();
        let rels = HashMap::new();
        let mut numbering_resolver = super::super::NumberingResolver::new(&docx);
        let mut image_extractor = super::super::ImageExtractor::new_skip();
        let options = ConvertOptions::default();
        let style_resolver = super::super::StyleResolver::new(&docx.styles);
        let mut context = super::super::ConversionContext::new(
            &rels,
            &mut numbering_resolver,
            &mut image_extractor,
            &options,
            None,
            None,
            None,
            &style_resolver,
        );

        let html = TableConverter::convert(&table, &mut context).expect("table conversion failed");
        assert!(html.contains("<td>SDT-CELL</td>"));
    }
}

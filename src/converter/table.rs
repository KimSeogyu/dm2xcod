//! Table converter - converts tables to HTML with merge support.

use super::{ConversionContext, ParagraphConverter};
use crate::Result;
use rs_docx::document::{Table, TableCell, TableCellContent};

/// Converter for Table elements.
pub struct TableConverter;

#[derive(Clone, Debug)]
enum CellStatus {
    /// Valid cell with content, rowspan, colspan
    Occupied {
        content: String,
        rowspan: usize,
        colspan: usize,
    },
    /// Cell merged horizontally (part of colspan)
    MergedLeft,
    /// Cell merged vertically (part of rowspan)
    MergedUp,
    /// Empty slot (should not happen after filling)
    Empty,
}

impl TableConverter {
    /// Converts a Table to HTML format with correct merge handling.
    pub fn convert(table: &Table, context: &mut ConversionContext) -> Result<String> {
        // 1. Build a 2D grid of cells
        let mut grid: Vec<Vec<CellStatus>> = Vec::new();

        for (row_idx, row) in table.rows.iter().enumerate() {
            // Ensure current row exists in grid
            if grid.len() <= row_idx {
                grid.push(Vec::new());
            }

            let mut col_idx = 0;
            for cell_content in &row.cells {
                // Skip if current slot is already occupied (e.g. by rowspan from above)
                while col_idx < grid[row_idx].len()
                    && !matches!(grid[row_idx][col_idx], CellStatus::Empty)
                {
                    col_idx += 1;
                }

                // Make sure row is large enough
                if grid[row_idx].len() <= col_idx {
                    grid[row_idx].resize(col_idx + 1, CellStatus::Empty);
                }

                match cell_content {
                    rs_docx::document::TableRowContent::TableCell(cell) => {
                        // Extract properties
                        let grid_span = cell
                            .property
                            .grid_span
                            .as_ref()
                            .map(|g| g.val as usize)
                            .unwrap_or(1);
                        let v_merge_val =
                            cell.property.v_merge.as_ref().and_then(|v| v.val.as_ref());

                        // Parse vMerge: "restart" or "continue".
                        let is_v_merge_restart =
                            matches!(v_merge_val, Some(rs_docx::formatting::VMergeType::Restart));

                        let is_v_merge_continue =
                            cell.property.v_merge.is_some() && !is_v_merge_restart;

                        let mut final_content = String::new();
                        // Only convert content if it's not a continuation (or if we want to preserve content? usually ignored)
                        // But for "restart", we capture content.
                        // For normal cells, we capture content.
                        if !is_v_merge_continue {
                            final_content = Self::convert_cell_content(cell, context)?;
                        }

                        // Determine Cell Status
                        if is_v_merge_continue {
                            // Find master cell above
                            Self::increment_rowspan(&mut grid, row_idx, col_idx);

                            // Mark current slots as MergedUp
                            // Note: If grid_span > 1, we mark all N columns as MergedUp?
                            // Yes, typically vMerge applies to the whole cell which might have gridSpan.
                            // But usually structure is rectangular.
                            for i in 0..grid_span {
                                Self::set_grid_cell(
                                    &mut grid,
                                    row_idx,
                                    col_idx + i,
                                    CellStatus::MergedUp,
                                );
                            }
                        } else {
                            // Occupied (New cell or Restart)
                            Self::set_grid_cell(
                                &mut grid,
                                row_idx,
                                col_idx,
                                CellStatus::Occupied {
                                    content: final_content,
                                    rowspan: 1,
                                    colspan: grid_span,
                                },
                            );

                            // Mark handled columns for colspan
                            for i in 1..grid_span {
                                Self::set_grid_cell(
                                    &mut grid,
                                    row_idx,
                                    col_idx + i,
                                    CellStatus::MergedLeft,
                                );
                            }
                        }

                        col_idx += grid_span;
                    }
                    rs_docx::document::TableRowContent::SDT(_) => {
                        // Ignore SDT for now, effectively skipping column?
                        // Or treat as empty cell? simpler to ignore.
                    }
                }
            }
        }

        // 2. Render HTML
        let mut html = String::from("<table>\n");
        for row in grid {
            html.push_str("  <tr>\n");
            for cell in row {
                match cell {
                    CellStatus::Occupied {
                        content,
                        rowspan,
                        colspan,
                    } => {
                        let mut attrs = String::new();
                        if rowspan > 1 {
                            attrs.push_str(&format!(" rowspan=\"{}\"", rowspan));
                        }
                        if colspan > 1 {
                            attrs.push_str(&format!(" colspan=\"{}\"", colspan));
                        }
                        html.push_str(&format!("    <td{}>{}</td>\n", attrs, content));
                    }
                    CellStatus::MergedLeft | CellStatus::MergedUp => {
                        // Skip
                    }
                    CellStatus::Empty => {
                        html.push_str("    <td></td>\n");
                    }
                }
            }
            html.push_str("  </tr>\n");
        }
        html.push_str("</table>");

        Ok(html)
    }

    fn set_grid_cell(grid: &mut Vec<Vec<CellStatus>>, row: usize, col: usize, status: CellStatus) {
        if grid.len() <= row {
            grid.resize(row + 1, Vec::new());
        }
        if grid[row].len() <= col {
            grid[row].resize(col + 1, CellStatus::Empty);
        }
        grid[row][col] = status;
    }

    fn increment_rowspan(grid: &mut Vec<Vec<CellStatus>>, current_row: usize, col: usize) {
        if current_row == 0 {
            return; // Cannot merge up at top row
        }

        // Trace back up to find the Occupied cell
        let mut target_row = current_row - 1;
        loop {
            // Ensure grid has this cell (it should)
            if grid.len() <= target_row || grid[target_row].len() <= col {
                break; // Should not happen in valid docx
            }

            match &mut grid[target_row][col] {
                CellStatus::Occupied { rowspan, .. } => {
                    *rowspan += 1;
                    return;
                }
                CellStatus::MergedUp => {
                    // Keep going up
                    if target_row == 0 {
                        break;
                    }
                    target_row -= 1;
                }
                _ => {
                    // MergedLeft or Empty?
                    // If MergedLeft, it means the master is to the left.
                    // But vMerge aligns columns. So [row][col] should align with [row-1][col].
                    // If [row-1][col] is MergedLeft, then the logic is complex.
                    // However, we increment the span of the *Cell* covering this column.
                    // If MergedLeft, find master to the left, and increment THAT?
                    // NO. vMerge happens column-wise.
                    // If a cell is gridSpan=2, vMerge must happen on BOTH columns usually?
                    // Or vMerge is specified on the first cell.
                    // Simplified: just try to find Occupied up.
                    break;
                }
            }
        }
    }

    fn convert_cell_content(cell: &TableCell, context: &mut ConversionContext) -> Result<String> {
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
            }
        }
        Ok(content)
    }
}

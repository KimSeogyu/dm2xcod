use crate::Result;
use rs_docx::document::{BodyContent, Table, TableCell};

#[derive(Clone, Debug)]
pub(crate) enum CellStatus {
    Occupied {
        content: String,
        rowspan: usize,
        colspan: usize,
    },
    MergedLeft,
    MergedUp,
    Empty,
}

pub(crate) fn build_grid<'a, F>(
    table: &Table<'a>,
    mut convert_cell: F,
) -> Result<Vec<Vec<CellStatus>>>
where
    F: FnMut(&TableCell<'a>) -> Result<String>,
{
    let mut grid: Vec<Vec<CellStatus>> = Vec::new();

    for (row_idx, row) in table.rows.iter().enumerate() {
        if grid.len() <= row_idx {
            grid.push(Vec::new());
        }

        let mut col_idx = 0;
        for cell_content in &row.cells {
            match cell_content {
                rs_docx::document::TableRowContent::TableCell(cell) => {
                    place_cell(&mut grid, row_idx, &mut col_idx, cell, &mut convert_cell)?;
                }
                rs_docx::document::TableRowContent::SDT(sdt) => {
                    if let Some(sdt_content) = &sdt.content {
                        for child in &sdt_content.content {
                            if let BodyContent::TableCell(cell) = child {
                                place_cell(
                                    &mut grid,
                                    row_idx,
                                    &mut col_idx,
                                    cell,
                                    &mut convert_cell,
                                )?;
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(grid)
}

fn place_cell<'a, F>(
    grid: &mut Vec<Vec<CellStatus>>,
    row_idx: usize,
    col_idx: &mut usize,
    cell: &TableCell<'a>,
    convert_cell: &mut F,
) -> Result<()>
where
    F: FnMut(&TableCell<'a>) -> Result<String>,
{
    while *col_idx < grid[row_idx].len() && !matches!(grid[row_idx][*col_idx], CellStatus::Empty) {
        *col_idx += 1;
    }

    if grid[row_idx].len() <= *col_idx {
        grid[row_idx].resize(*col_idx + 1, CellStatus::Empty);
    }

    let grid_span = cell
        .property
        .grid_span
        .as_ref()
        .map(|g| g.val as usize)
        .unwrap_or(1);
    let v_merge_val = cell.property.v_merge.as_ref().and_then(|v| v.val.as_ref());

    let is_v_merge_restart = matches!(v_merge_val, Some(rs_docx::formatting::VMergeType::Restart));
    let is_v_merge_continue = cell.property.v_merge.is_some() && !is_v_merge_restart;

    let content = if is_v_merge_continue {
        String::new()
    } else {
        convert_cell(cell)?
    };

    if is_v_merge_continue {
        increment_rowspan(grid, row_idx, *col_idx);
        for i in 0..grid_span {
            set_grid_cell(grid, row_idx, *col_idx + i, CellStatus::MergedUp);
        }
    } else {
        set_grid_cell(
            grid,
            row_idx,
            *col_idx,
            CellStatus::Occupied {
                content,
                rowspan: 1,
                colspan: grid_span,
            },
        );
        for i in 1..grid_span {
            set_grid_cell(grid, row_idx, *col_idx + i, CellStatus::MergedLeft);
        }
    }

    *col_idx += grid_span;
    Ok(())
}

pub(crate) fn render_grid(grid: Vec<Vec<CellStatus>>) -> String {
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
                CellStatus::MergedLeft | CellStatus::MergedUp => {}
                CellStatus::Empty => {
                    html.push_str("    <td></td>\n");
                }
            }
        }
        html.push_str("  </tr>\n");
    }
    html.push_str("</table>");
    html
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

fn increment_rowspan(grid: &mut [Vec<CellStatus>], current_row: usize, col: usize) {
    if current_row == 0 {
        return;
    }

    let mut target_row = current_row - 1;
    loop {
        if grid.len() <= target_row || grid[target_row].len() <= col {
            break;
        }

        match &grid[target_row][col] {
            CellStatus::Occupied { .. } => {
                if let CellStatus::Occupied { rowspan, .. } = &mut grid[target_row][col] {
                    *rowspan += 1;
                    return;
                }
            }
            CellStatus::MergedUp => {
                if target_row == 0 {
                    break;
                }
                target_row -= 1;
            }
            CellStatus::MergedLeft => {
                let mut master_col = col;
                let mut found_master = None;
                while master_col > 0 {
                    master_col -= 1;
                    match &grid[target_row][master_col] {
                        CellStatus::MergedLeft => continue,
                        CellStatus::Occupied { colspan, .. } => {
                            if master_col + *colspan > col {
                                found_master = Some(master_col);
                            }
                            break;
                        }
                        _ => break,
                    }
                }

                if let Some(master_col) = found_master {
                    if let CellStatus::Occupied { rowspan, .. } = &mut grid[target_row][master_col]
                    {
                        *rowspan += 1;
                        return;
                    }
                }
                break;
            }
            CellStatus::Empty => break,
        }
    }
}

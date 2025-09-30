use crate::document::model::*;
use crate::document::providers::DocumentProvider;
use calamine::{open_workbook_auto_from_rs, Data, Range, Reader};
use std::error::Error;
use std::io::Cursor;
use std::num::NonZeroU32;

pub struct ExcelProvider;

impl ExcelProvider {
  pub fn new() -> Self {
    Self
  }
}

impl DocumentProvider for ExcelProvider {
  fn parse_buffer(&self, data: &[u8]) -> Result<Document, Box<dyn Error + Send + Sync>> {
    let cursor = Cursor::new(data);
    let mut workbook = open_workbook_auto_from_rs(cursor)?;

    let mut blocks: Vec<Block> = Vec::new();
    let sheet_names = workbook.sheet_names();

    for sheet_name in sheet_names {
      if let Ok(range) = workbook.worksheet_range(&sheet_name) {
        if let Some(table) = parse_sheet_to_table(&range) {
          blocks.push(Block::Table(table));
        }
      }
    }

    let metadata = DocumentMetadata::default();

    Ok(Document {
      blocks,
      metadata,
      notes: Vec::new(),
      comments: Vec::new(),
    })
  }

  fn name(&self) -> &'static str {
    "excel"
  }
}

fn parse_sheet_to_table(range: &Range<Data>) -> Option<Table> {
  let rows = range.rows();
  let mut table_rows: Vec<TableRow> = Vec::new();
  let mut is_first_row = true;

  for row in rows {
    let cells: Vec<TableCell> = row
      .iter()
      .map(|cell| {
        let text = data_to_string(cell);
        let paragraph = Paragraph {
          kind: ParagraphKind::Normal,
          inlines: vec![Inline::Text(text)],
        };
        TableCell {
          blocks: vec![Block::Paragraph(paragraph)],
          colspan: NonZeroU32::new(1).unwrap(),
          rowspan: NonZeroU32::new(1).unwrap(),
        }
      })
      .collect();

    if !cells.is_empty() {
      let kind = if is_first_row && row_contains_text(row) {
        is_first_row = false;
        TableRowKind::Header
      } else {
        is_first_row = false;
        TableRowKind::Body
      };

      table_rows.push(TableRow { cells, kind });
    }
  }

  if table_rows.is_empty() {
    None
  } else {
    Some(Table { rows: table_rows })
  }
}

fn data_to_string(cell: &Data) -> String {
  match cell {
    Data::Int(i) => i.to_string(),
    Data::Float(f) => f.to_string(),
    Data::String(s) => s.clone(),
    Data::Bool(b) => b.to_string(),
    Data::DateTime(dt) => dt.to_string(),
    Data::DateTimeIso(dt) => dt.clone(),
    Data::DurationIso(d) => d.clone(),
    Data::Error(e) => format!("Error: {e:?}"),
    Data::Empty => String::new(),
  }
}

fn row_contains_text(row: &[Data]) -> bool {
  row.iter().any(|cell| matches!(cell, Data::String(_)))
}

#[cfg(test)]
mod tests {
  use super::*;

  #[test]
  fn test_data_to_string() {
    assert_eq!(data_to_string(&Data::Int(42)), "42");
    assert_eq!(data_to_string(&Data::Float(3.14)), "3.14");
    assert_eq!(data_to_string(&Data::String("test".to_string())), "test");
    assert_eq!(data_to_string(&Data::Bool(true)), "true");
    assert_eq!(data_to_string(&Data::Empty), "");
  }

  #[test]
  fn test_row_contains_text() {
    let row_with_text = vec![Data::String("Header".to_string()), Data::Int(42)];
    assert!(row_contains_text(&row_with_text));

    let row_without_text = vec![Data::Int(42), Data::Float(3.14)];
    assert!(!row_contains_text(&row_without_text));

    let empty_row: Vec<Data> = vec![];
    assert!(!row_contains_text(&empty_row));
  }
}

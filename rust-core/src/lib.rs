// use mimalloc::MiMalloc;

// #[global_allocator]
// static GLOBAL: MiMalloc = MiMalloc;

pub mod files_part;
mod polars_part;
mod read_part;
pub mod style;
mod test;
use std::{
    collections::HashMap, fs::File, io::Read, path::{Path, PathBuf}
};

use anyhow::{Context, Result, bail};
use quick_xml::{Reader, Writer, events::Event};

use crate::style::{AlignSpec, HorizAlignment, VertAlignment};
// use regex::Regex;
// use tempfile::NamedTempFile;
// use zip::{ZipArchive, ZipWriter, write::FileOptions};

/// `XlsxEditor` provides functionality to open, modify, and save XLSX files.
/// It allows appending rows and tables to a specified sheet within an XLSX file.

#[derive(Hash, Eq, PartialEq, Clone)]
struct FontKey {
    name: String,
    size_100: u32,
    bold: bool,
    italic: bool,
}
#[derive(Hash, Eq, PartialEq, Clone)]
struct StyleKey {
    num_fmt_id: u32,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
    align: Option<(Option<HorizAlignment>, Option<VertAlignment>, bool)>, // wrap
}

struct XfParts {
    num_fmt_id: u32,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
    align: Option<AlignSpec>,
}


struct StyleIndex {
    xfs: Vec<XfParts>, // index == style_id

    numfmt_by_code: HashMap<String, u32>,
    next_custom_numfmt: u32, // >=164

    font_by_key: HashMap<FontKey, u32>,
    fill_by_rgb: HashMap<String, u32>,    // RGB в верхнем регистре
    border_by_key: HashMap<String, u32>,  // единый style для всех сторон

    xf_by_key: HashMap<StyleKey, u32>,

    fonts_count: u32,
    fills_count: u32,
    borders_count: u32,
}

pub struct XlsxEditor {
    src_path: PathBuf,
    sheet_path: String,
    sheet_xml: Vec<u8>,
    last_row: u32,
    styles_xml: Vec<u8>,               // содержимое styles.xml
    workbook_xml: Vec<u8>,             // содержимое workbook.xml (может изменяться)
    rels_xml: Vec<u8>,                 // содержимое workbook.xml.rels
    new_files: Vec<(String, Vec<u8>)>, // новые или изменённые файлы для записи при save()
    styles_index: Option<StyleIndex>,
}

/// Polars

/// Main
impl XlsxEditor {
    /// Opens an XLSX file and prepares a specific sheet for editing by its name.
    ///
    /// This function first scans the workbook to find the sheet ID corresponding to the given sheet name,
    /// then calls `open_sheet` with the found ID.
    ///
    /// # Arguments
    /// * `src` - The path to the XLSX file.
    /// * `sheet_name` - The name of the sheet to open (e.g., "Sheet1").
    ///
    /// # Returns
    /// A `Result` containing an `XlsxEditor` instance if successful, or an `anyhow::Error` otherwise.
    pub fn open<P: AsRef<Path>>(src: P, sheet_name: &str) -> Result<Self> {
        let sheet_names = scan(src.as_ref())?;
        let sheet_id = sheet_names
            .iter()
            .position(|n| n == sheet_name)
            .context(format!("Sheet '{}' not found", sheet_name))?
            + 1;
        println!("Sheet ID: {} with name {}", sheet_id, sheet_name);
        Self::open_sheet(src, sheet_id)
    }

    /// Appends a single row of cells to the end of the current sheet.
    ///
    /// Each item in the `cells` iterator will be converted to a string and written as a cell.
    /// The cell type (number or inline string) is inferred based on whether the value can be parsed as a float.
    ///
    /// # Arguments
    /// * `cells` - An iterator over values that can be converted to strings, representing the cells in the new row.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_row<I, S>(&mut self, cells: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        self.last_row += 1;
        let row_num = self.last_row;
        let mut writer = Writer::new(Vec::new());

        // Create a new XML row element with the appropriate row number attribute.
        writer
            .create_element("row")
            .with_attribute(("r", row_num.to_string().as_str()))
            .write_inner_content(|w| {
                let mut col = b'A';
                for val in cells {
                    let coord = format!("{}{}", col as char, row_num);
                    let val_str = val.to_string();
                    let is_formula = val_str.starts_with('=');
                    let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                    {
                        let mut c_elem =
                            w.create_element("c").with_attribute(("r", coord.as_str()));
                        if !is_number && !is_formula {
                            c_elem = c_elem.with_attribute(("t", "inlineStr"));
                        }
                        c_elem.write_inner_content(|w2| {
                            use quick_xml::events::BytesText;
                            if is_formula {
                                w2.create_element("f")
                                    .write_text_content(BytesText::new(&val_str[1..]))?;
                            } else if !is_number {
                                w2.create_element("is").write_inner_content(|w3| {
                                    w3.create_element("t")
                                        .write_text_content(BytesText::new(&val_str))?;
                                    Ok(())
                                })?;
                            } else {
                                w2.create_element("v")
                                    .write_text_content(BytesText::new(&val_str))?;
                            }
                            Ok(())
                        })?;
                    }
                    col += 1;
                }
                Ok(())
            })?;

        let new_row_xml = writer.into_inner();

        // Find the closing </sheetData> tag and insert the new row before it.
        if let Some(pos) = self
            .sheet_xml
            .windows(12)
            .rposition(|w| w == b"</sheetData>")
        {
            self.sheet_xml.splice(pos..pos, new_row_xml);
            Ok(())
        } else {
            bail!("</sheetData> tag not found");
        }
    }

    /// Appends multiple rows (a table) to the end of the current sheet.
    ///
    /// This function iterates through the provided rows, and for each row, it iterates through its cells.
    /// Each cell's value is converted to a string, and its type (number or inline string) is inferred.
    /// The new rows are then appended to the sheet's XML content.
    ///
    /// # Arguments
    /// * `rows` - An iterator over iterators of values that can be converted to strings, representing the rows and cells of the table.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_table<R, I, S>(&mut self, rows: R) -> Result<()>
    where
        R: IntoIterator<Item = I>,
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        ensure_sheetdata_open_close(&mut self.sheet_xml)?;

        // Helper function to convert a 0-based column index to Excel column letters (e.g., 0 -> "A", 26 -> "AA").
        fn col_idx_to_letters(mut idx: usize) -> String {
            let mut s = String::new();
            loop {
                let rem = idx % 26;
                s.insert(0, (b'A' + rem as u8) as char);
                if idx < 26 {
                    break;
                }
                idx = idx / 26 - 1;
            }
            s
        }

        // Buffer to accumulate XML for all new rows.
        let mut bulk_rows_xml = Vec::<u8>::new();

        for row in rows {
            self.last_row += 1;
            let row_num = self.last_row;

            let mut writer = Writer::new(Vec::new());
            writer
                .create_element("row")
                .with_attribute(("r", row_num.to_string().as_str()))
                .write_inner_content(|w| {
                    for (col_idx, val) in row.into_iter().enumerate() {
                        let coord = format!("{}{}", col_idx_to_letters(col_idx), row_num);
                        let val_str = val.to_string();
                        let is_formula = val_str.starts_with('=');
                        let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                        let mut c_elem =
                            w.create_element("c").with_attribute(("r", coord.as_str()));
                        if !is_number && !is_formula {
                            c_elem = c_elem.with_attribute(("t", "inlineStr"));
                        }
                        c_elem.write_inner_content(|w2| {
                            use quick_xml::events::BytesText;
                            if is_formula {
                                w2.create_element("f")
                                    .write_text_content(BytesText::new(&val_str[1..]))?;
                            } else if !is_number {
                                w2.create_element("is").write_inner_content(|w3| {
                                    w3.create_element("t")
                                        .write_text_content(BytesText::new(&val_str))?;
                                    Ok(())
                                })?;
                            } else {
                                w2.create_element("v")
                                    .write_text_content(BytesText::new(&val_str))?;
                            }
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;

            bulk_rows_xml.extend_from_slice(&writer.into_inner());
        }

        // eprintln!(
        //     "rows appended: last_row={}, has_close_sheetdata={} path={}",
        //     self.last_row,
        //     self.sheet_xml
        //         .windows(12)
        //         .rposition(|w| w == b"</sheetData>")
        //         .is_some(),
        //     self.sheet_path
        // );

        // Find the closing </sheetData> tag and insert the new rows before it.
        if let Some(pos) = self
            .sheet_xml
            .windows(12)
            .rposition(|w| w == b"</sheetData>")
        {
            self.sheet_xml.splice(pos..pos, bulk_rows_xml);
            Ok(())
        } else {
            bail!("</sheetData> tag not found");
        }
    }

    /// Appends multiple rows (a table) starting at a specified coordinate in the current sheet.
    ///
    /// This function allows inserting a table at a specific cell coordinate (e.g., "A1", "C5").
    /// If the target rows already exist, their cells will be updated. If the target rows are beyond
    /// the current last row, new rows will be appended.
    ///
    /// # Arguments
    /// * `start_coord` - The starting cell coordinate (e.g., "A1") where the table should begin.
    /// * `rows` - An iterator over iterators of values that can be converted to strings, representing the rows and cells of the table.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_table_at<R, I, S>(&mut self, start_coord: &str, rows: R) -> Result<()>
    where
        R: IntoIterator<Item = I>,
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        ensure_sheetdata_open_close(&mut self.sheet_xml)?;

        // Helper function to convert a 0-based column index to Excel column letters (e.g., 0 -> "A", 26 -> "AA").
        fn col_idx_to_letters(mut idx: usize) -> String {
            let mut s = String::new();
            loop {
                let rem = idx % 26;
                s.insert(0, (b'A' + rem as u8) as char);
                if idx < 26 {
                    break;
                }
                idx = idx / 26 - 1;
            }
            s
        }
        // Helper function to convert Excel column letters (e.g., "A", "AA") to their corresponding 0-based column index.
        fn letters_to_col_idx(s: &str) -> usize {
            s.bytes().fold(0, |acc, b| {
                acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
            }) - 1
        }

        // Parse the starting coordinate to get the initial column index and row number.
        let row_start_pos = start_coord
            .find(|c: char| c.is_ascii_digit())
            .context("invalid start coordinate – no digits")?;
        let col_letters = &start_coord[..row_start_pos];
        let start_col_idx = letters_to_col_idx(col_letters);
        let current_row_num: u32 = start_coord[row_start_pos..]
            .parse()
            .context("invalid row in start coordinate")?;

        // Buffer to accumulate XML for new rows that need to be appended.
        let mut bulk_rows_xml = Vec::<u8>::new();
        let mut row_offset: usize = 0;

        for row in rows {
            let abs_row = current_row_num + row_offset as u32;
            if abs_row <= self.last_row {
                // If the row already exists, update cells within that row.
                for (col_offset, val) in row.into_iter().enumerate() {
                    let coord = format!(
                        "{}{}",
                        col_idx_to_letters(start_col_idx + col_offset),
                        abs_row
                    );
                    // Set the cell value using the existing set_cell method.
                    self.set_cell(&coord, val)?;
                }
            } else {
                // If the row does not exist, create a new row and append it.
                let mut writer = Writer::new(Vec::new());
                writer
                    .create_element("row")
                    .with_attribute(("r", abs_row.to_string().as_str()))
                    .write_inner_content(|w| {
                        for (col_offset, val) in row.into_iter().enumerate() {
                            let coord = format!(
                                "{}{}",
                                col_idx_to_letters(start_col_idx + col_offset),
                                abs_row
                            );
                            let val_str = val.to_string();
                            let is_formula = val_str.starts_with('=');
                            let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                            let mut c_elem =
                                w.create_element("c").with_attribute(("r", coord.as_str()));
                            if !is_number && !is_formula {
                                c_elem = c_elem.with_attribute(("t", "inlineStr"));
                            }
                            c_elem.write_inner_content(|w2| {
                                use quick_xml::events::BytesText;
                                if is_formula {
                                    w2.create_element("f")
                                        .write_text_content(BytesText::new(&val_str[1..]))?;
                                } else if !is_number {
                                    w2.create_element("is").write_inner_content(|w3| {
                                        w3.create_element("t")
                                            .write_text_content(BytesText::new(&val_str))?;
                                        Ok(())
                                    })?;
                                } else {
                                    w2.create_element("v")
                                        .write_text_content(BytesText::new(&val_str))?;
                                }
                                Ok(())
                            })?;
                        }
                        Ok(())
                    })?;

                bulk_rows_xml.extend_from_slice(&writer.into_inner());
                // Update the last row number if necessary.
                self.last_row = abs_row;
            }
            row_offset += 1;
        }
        // eprintln!(
        //     "rows appended: last_row={}, has_close_sheetdata={} path={}",
        //     self.last_row,
        //     self.sheet_xml
        //         .windows(12)
        //         .rposition(|w| w == b"</sheetData>")
        //         .is_some(),
        //     self.sheet_path
        // );

        // Find the closing </sheetData> tag and insert the new rows before it.
        if let Some(pos) = self
            .sheet_xml
            .windows(12)
            .rposition(|w| w == b"</sheetData>")
        {
            self.sheet_xml.splice(pos..pos, bulk_rows_xml);
            Ok(())
        } else {
            bail!("</sheetData> tag not found");
        }
    }

    /// Sets the value of a specific cell in the sheet.
    ///
    /// This function allows updating an existing cell or creating a new one if it doesn't exist.
    /// The cell type (number or inline string) is inferred based on whether the value can be parsed as a float.
    ///
    /// # Arguments
    /// * `coord` - The cell coordinate (e.g., "A1", "B2").
    /// * `value` - The value to set for the cell, which can be converted to a string.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn set_cell<S: ToString>(&mut self, coord: &str, value: S) -> Result<()> {
        // Extract row number from coordinate.
        let row_start = coord
            .find(|c: char| c.is_ascii_digit())
            .context("invalid cell coordinate – no digits found")?;
        let row_num: u32 = coord[row_start..]
            .parse()
            .context("invalid row number in cell coordinate")?;

        let val_str = value.to_string();
        let is_formula = val_str.starts_with('=');
        let is_number = !is_formula && val_str.parse::<f64>().is_ok();

        // Generate XML for the new cell.
        let mut cell_writer = Writer::new(Vec::new());
        // Create cell element with coordinate and type attributes.
        let mut c_elem = cell_writer.create_element("c").with_attribute(("r", coord));
        if !is_number && !is_formula {
            c_elem = c_elem.with_attribute(("t", "inlineStr"));
        }
        c_elem.write_inner_content(|w2| {
            use quick_xml::events::BytesText;
            if is_formula {
                w2.create_element("f")
                    .write_text_content(BytesText::new(&val_str[1..]))?;
            } else if !is_number {
                // For strings, use <is><t> tags.
                w2.create_element("is").write_inner_content(|w3| {
                    w3.create_element("t")
                        .write_text_content(BytesText::new(&val_str))?;
                    Ok(())
                })?;
            } else {
                // For numbers, use <v> tag.
                w2.create_element("v")
                    .write_text_content(BytesText::new(&val_str))?;
            }
            Ok(())
        })?;
        let cell_xml = cell_writer.into_inner();

        // Find the row containing the target cell.
        let row_marker = format!("<row r=\"{}\"", row_num);
        if let Some(row_start) = self
            .sheet_xml
            .windows(row_marker.len())
            .position(|w| w == row_marker.as_bytes())
        {
            // Find the end of the row.
            if let Some(rel_end) = self.sheet_xml[row_start..]
                .windows(6)
                .position(|w| w == b"</row>")
            {
                let row_end = row_start + rel_end + 6; // 6 is the length of "</row>"
                let mut row_slice = self.sheet_xml[row_start..row_end].to_vec();

                // Find the cell within the row and replace it.
                let cell_marker = format!("<c r=\"{}\"", coord);
                if let Some(cell_pos) = row_slice
                    .windows(cell_marker.len())
                    .position(|w| w == cell_marker.as_bytes())
                {
                    if let Some(cell_end_rel) =
                        row_slice[cell_pos..].windows(4).position(|w| w == b"</c>")
                    {
                        let cell_end = cell_pos + cell_end_rel + 4;
                        row_slice.drain(cell_pos..cell_end);
                    } else if let Some(cell_end_rel) =
                        row_slice[cell_pos..].windows(2).position(|w| w == b"/>")
                    {
                        let cell_end = cell_pos + cell_end_rel + 2;
                        row_slice.drain(cell_pos..cell_end);
                    }
                }

                // Insert the new cell at the correct position within the row.
                fn col_to_index(s: &str) -> u32 {
                    s.bytes()
                        .take_while(|b| b.is_ascii_alphabetic())
                        .fold(0, |acc, b| {
                            acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as u32
                        })
                }
                let target_col = col_to_index(coord);
                // Find the correct position to insert the new cell.
                let mut insert_pos = row_slice.len() - 6; // 6 is the length of "</row>"
                let mut i = 0;
                while let Some(c_pos) = row_slice[i..].windows(6).position(|w| w == b"<c r=\"") {
                    let abs = i + c_pos;
                    // Find the end of the cell's coordinate attribute.
                    if let Some(end_quote) = row_slice[abs + 6..].iter().position(|&b| b == b'"') {
                        let coord_bytes = &row_slice[abs + 6..abs + 6 + end_quote];
                        if let Ok(coord_str) = std::str::from_utf8(coord_bytes) {
                            let col_idx = col_to_index(coord_str);
                            if col_idx > target_col {
                                insert_pos = abs;
                                break;
                            }
                        }
                        i = abs + 6 + end_quote;
                    } else {
                        break;
                    }
                }
                row_slice.splice(insert_pos..insert_pos, cell_xml);

                // Replace the original row with the updated one.
                self.sheet_xml.splice(row_start..row_end, row_slice);
            }
        } else {
            // If the row does not exist, create a new row and insert it in the correct order so that
            // the `<row>` elements remain sorted by the `r` attribute.  Keeping the rows ordered
            // avoids Excel "recovered records" errors that occur when rows are out of sequence.
            let mut new_row_xml = Vec::new();
            new_row_xml.extend_from_slice(b"<row r=\"");
            new_row_xml.extend_from_slice(row_num.to_string().as_bytes());
            new_row_xml.extend_from_slice(b"\">");
            new_row_xml.extend_from_slice(&cell_xml);
            new_row_xml.extend_from_slice(b"</row>");

            // Try to find the first existing row whose `r` value is greater than the new row.
            // If found, we will insert the new row *before* it, otherwise we fall back to
            // inserting just before `</sheetData>` (the previous behaviour).
            let mut insert_pos: Option<usize> = None;
            let mut search_idx = 0;
            while let Some(rel) = self.sheet_xml[search_idx..]
                .windows(7)
                .position(|w| w == b"<row r=")
            {
                let abs = search_idx + rel;
                // Find the opening quote for the `r` attribute.
                if let Some(first_quote) = self.sheet_xml[abs..].iter().position(|&b| b == b'"') {
                    let num_start = abs + first_quote + 1;
                    // Find the closing quote for the `r` attribute.
                    if let Some(end_quote) =
                        self.sheet_xml[num_start..].iter().position(|&b| b == b'"')
                    {
                        let num_bytes = &self.sheet_xml[num_start..num_start + end_quote];
                        if let Ok(num_str) = std::str::from_utf8(num_bytes) {
                            if let Ok(existing_r) = num_str.parse::<u32>() {
                                if existing_r > row_num {
                                    insert_pos = Some(abs);
                                    break;
                                }
                            }
                        }
                        // Continue searching after this row tag.
                        search_idx = num_start + end_quote;
                    } else {
                        break; // Malformed XML (should not happen)
                    }
                } else {
                    break; // Malformed XML (should not happen)
                }
            }

            let pos = match insert_pos {
                Some(p) => p,
                None => self
                    .sheet_xml
                    .windows(12)
                    .rposition(|w| w == b"</sheetData>")
                    .context("</sheetData> tag not found")?,
            };

            self.sheet_xml.splice(pos..pos, new_row_xml);
        }

        if row_num > self.last_row {
            self.last_row = row_num;
        }
        Ok(())
    }
}

pub fn scan<P: AsRef<Path>>(src: P) -> Result<Vec<String>> {
    let mut zip = zip::ZipArchive::new(File::open(src)?)?;
    let mut wb = zip
        .by_name("xl/workbook.xml")
        .context("workbook.xml not found")?;

    let mut wb_xml = Vec::with_capacity(wb.size() as usize);
    wb.read_to_end(&mut wb_xml)?;

    let mut reader = Reader::from_reader(wb_xml.as_slice());
    reader.config_mut().trim_text(true);

    let mut names = Vec::new();

    while let Ok(ev) = reader.read_event() {
        match ev {
            Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                if let Some(n) = e.attributes().with_checks(false).flatten().find_map(|a| {
                    (a.key.as_ref() == b"name")
                        .then(|| String::from_utf8_lossy(&a.value).into_owned())
                }) {
                    names.push(n);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(names)
}

impl XlsxEditor {
    pub fn merge_cells(&mut self, range: &str) -> Result<()> {
        // 1. позиция после </sheetData>
        let sd_end = find_bytes(&self.sheet_xml, b"</sheetData>")
            .context("</sheetData> not found")?
            + "</sheetData>".len();

        let (insert_pos, created) = if let Some(pos) = find_bytes(&self.sheet_xml, b"<mergeCells") {
            // уже есть блок
            bump_count(&mut self.sheet_xml, b"<mergeCells", b"count=\"")?;
            let end = find_bytes_from(&self.sheet_xml, b"</mergeCells>", pos)
                .context("</mergeCells> not found")?;
            (end, false)
        } else {
            // нет блока – создаём
            let tpl = br#"<mergeCells count="0"></mergeCells>"#;
            self.sheet_xml.splice(sd_end..sd_end, tpl.iter().copied());
            (sd_end + tpl.len() - "</mergeCells>".len(), true)
        };

        // 2. сам <mergeCell>
        let tag = format!(r#"<mergeCell ref="{}"/>"#, range);
        self.sheet_xml
            .splice(insert_pos..insert_pos, tag.as_bytes().iter().copied());

        // 3. правим count (если блок создан только что)
        if created {
            bump_count(&mut self.sheet_xml, b"<mergeCells", b"count=\"")?;
        }
        Ok(())
    }
}

fn find_bytes(hay: &[u8], needle: &[u8]) -> Option<usize> {
    hay.windows(needle.len()).position(|w| w == needle)
}
fn find_bytes_from(hay: &[u8], needle: &[u8], start: usize) -> Option<usize> {
    hay[start..]
        .windows(needle.len())
        .position(|w| w == needle)
        .map(|p| p + start)
}

fn bump_count(xml: &mut Vec<u8>, tag: &[u8], attr: &[u8]) -> Result<()> {
    if let Some(pos) = find_bytes(xml, tag) {
        if let Some(a) = find_bytes_from(xml, attr, pos) {
            let start = a + attr.len();
            let end = find_bytes_from(xml, b"\"", start).unwrap();
            let mut num: u32 = std::str::from_utf8(&xml[start..end])?.parse()?;
            num += 1;
            xml.splice(start..end, num.to_string().as_bytes().iter().copied());
            return Ok(());
        }
    }
    Err(anyhow::anyhow!("attribute count not found"))
}

fn ensure_sheetdata_open_close(xml: &mut Vec<u8>) -> Result<()> {
    const SELF_CLOSING: &[u8] = b"<sheetData/>";
    if let Some(pos) = memchr::memmem::find(xml, SELF_CLOSING) {
        // заменяем на <sheetData></sheetData>
        let replacement = b"<sheetData></sheetData>";
        xml.splice(pos..pos + SELF_CLOSING.len(), replacement.iter().copied());
    }
    Ok(())
}

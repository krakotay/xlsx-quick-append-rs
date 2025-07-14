mod test;

use std::{
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, bail};
use quick_xml::{Reader, Writer, events::Event};
use tempfile::NamedTempFile;
use zip::{CompressionMethod, ZipWriter, read::ZipArchive, write::FileOptions};


/// `XlsxEditor` provides functionality to open, modify, and save XLSX files.
/// It allows appending rows and tables to a specified sheet within an XLSX file.
pub struct XlsxEditor {
    /// The path to the source XLSX file.
    src_path: PathBuf,
    /// The internal path to the sheet XML file within the XLSX archive (e.g., "xl/worksheets/sheet1.xml").
    sheet_path: String, 
    /// The raw XML content of the sheet being edited.
    sheet_xml: Vec<u8>, 
    /// The last row number identified in the sheet, used for appending new rows.
    last_row: u32,      
}


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
        let sheet_id = sheet_names.iter().position(|n| n == sheet_name).context(format!("Sheet '{}' not found", sheet_name))? + 1;
        println!("Sheet ID: {} with name {}", sheet_id, sheet_name);
        Self::open_sheet(src, sheet_id)
    }

    /// Opens an XLSX file and prepares a specific sheet for editing by its ID.
    ///
    /// This involves extracting the sheet's XML content and determining the last row number.
    ///
    /// # Arguments
    /// * `src` - The path to the XLSX file.
    /// * `sheet_id` - The 1-based ID of the sheet to open (e.g., 1 for the first sheet).
    ///
    /// # Returns
    /// A `Result` containing an `XlsxEditor` instance if successful, or an `anyhow::Error` otherwise.
    pub fn open_sheet<P: AsRef<Path>>(src: P, sheet_id: usize) -> Result<Self> {
        let src_path = src.as_ref().to_path_buf();
        let mut zip = ZipArchive::new(File::open(&src_path)?)?;
        
        // Read workbook.xml to get sheet information
        let mut wb = zip
            .by_name("xl/workbook.xml")
            .context("workbook.xml not found")?;
        let mut wb_xml = Vec::with_capacity(wb.size() as usize);
        wb.read_to_end(&mut wb_xml)?;
        drop(wb); 

        let mut reader = Reader::from_reader(wb_xml.as_slice());
        reader.config_mut().trim_text(true);
        
        // Construct the sheet XML path (e.g., "xl/worksheets/sheet1.xml")
        let sheet_path = format!("xl/worksheets/sheet{}.xml", sheet_id);

        // Read the sheet's XML content
        let mut sheet = zip
            .by_name(&sheet_path)
            .with_context(|| format!("{} not found", sheet_path))?;
        let mut sheet_xml = Vec::with_capacity(sheet.size() as usize);
        sheet.read_to_end(&mut sheet_xml)?;

        // Determine the last row number in the sheet
        let mut reader = Reader::from_reader(sheet_xml.as_slice());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut last_row = 0;

        while let Ok(ev) = reader.read_event_into(&mut buf) {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"row" => {
                    if let Some(r) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        last_row = r.parse::<u32>().unwrap_or(last_row);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(Self {
            src_path,
            sheet_path,
            sheet_xml,
            last_row,
        })
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
            while let Some(rel) = self.sheet_xml[search_idx..].windows(7).position(|w| w == b"<row r=") {
                let abs = search_idx + rel;
                // Find the opening quote for the `r` attribute.
                if let Some(first_quote) = self.sheet_xml[abs..].iter().position(|&b| b == b'"') {
                    let num_start = abs + first_quote + 1;
                    // Find the closing quote for the `r` attribute.
                    if let Some(end_quote) = self.sheet_xml[num_start..].iter().position(|&b| b == b'"') {
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

    /// Saves the modified XLSX file to a specified destination or overwrites the source file.
    ///
    /// This function creates a new ZIP archive, copying all original files from the source XLSX,
    /// but replacing the modified sheet's XML content with the updated content.
    ///
    /// # Arguments
    /// * `dest` - An optional path to save the modified file. If `None`, the original file will be overwritten.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the save operation fails.
    pub fn save<P: AsRef<Path>>(&self, dst: P) -> Result<()> {
        let mut zin = ZipArchive::new(File::open(&self.src_path)?)?;
        let mut tmp = NamedTempFile::new()?;
        {
            let mut zout = ZipWriter::new(&mut tmp);
            let opt: FileOptions<'_, ()> = FileOptions::default()
                .compression_method(CompressionMethod::Deflated)
                .unix_permissions(0o644);
            for i in 0..zin.len() {
                let mut f = zin.by_index(i)?;
                let name = f.name();
                if name == self.sheet_path.as_str() {
                    zout.start_file::<_, ()>(name, opt)?;
                    zout.write_all(&self.sheet_xml)?;
                } else {
                    zout.start_file::<_, ()>(name, opt)?;
                    std::io::copy(&mut f, &mut zout)?;
                }
            }
            zout.finish()?;
        }
        fs::rename(tmp.path(), dst)?;
        Ok(())
    }
}





pub fn scan<P: AsRef<Path>>(src: P) -> Result<Vec<String>> {
    
    let mut zip = ZipArchive::new(File::open(src)?)?;
    let mut wb = zip
        .by_name("xl/workbook.xml")
        .context("workbook.xml not found")?;

    let mut wb_xml = Vec::with_capacity(wb.size() as usize);
    wb.read_to_end(&mut wb_xml)?;

    
    let mut reader = Reader::from_reader(wb_xml.as_slice());
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut names = Vec::new();

    while let Ok(ev) = reader.read_event_into(&mut buf) {
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
        buf.clear();
    }
    Ok(names)
}

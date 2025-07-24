use mimalloc::MiMalloc;

#[global_allocator]
static GLOBAL: MiMalloc = MiMalloc;

mod test;

use quick_xml::{Reader, Writer, events::Event};
use regex::Regex;
use std::{
    fs::File,
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
};
use thiserror::Error;
// use tempfile::NamedTempFile;
// use zip::{ZipArchive, ZipWriter, write::FileOptions};
use ::zip as zip_crate;

/// Custom error type for XLSX editor operations
#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    
    #[error("ZIP error: {0}")]
    Zip(#[from] zip_crate::result::ZipError),
    
    #[error("XML parsing error: {0}")]
    Xml(#[from] quick_xml::Error),
    
    #[error("Parse error: {0}")]
    Parse(#[from] std::num::ParseIntError),
    
    #[error("UTF-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    
    #[error("Regex error: {0}")]
    Regex(#[from] regex::Error),
    
    #[error("File not found: {file}")]
    FileNotFound { file: String },
    
    #[error("XML tag not found: {tag}")]
    XmlTagNotFound { tag: String },
    
    #[error("Sheet not found: {name}")]
    SheetNotFound { name: String },
    
    #[error("Sheet already exists: {name}")]
    SheetAlreadyExists { name: String },
    
    #[error("Invalid coordinate: {coord}")]
    InvalidCoordinate { coord: String },
    
    #[error("Invalid range: {range}")]
    InvalidRange { range: String },
    
    #[error("Malformed XML tag: {tag}")]
    MalformedXmlTag { tag: String },
    
    #[error("No columns supplied")]
    NoColumnsSupplied,
    
    #[error("Invalid range order")]
    InvalidRangeOrder,
    
    #[error("Attribute not found: {attribute}")]
    AttributeNotFound { attribute: String },
}

pub type Result<T> = std::result::Result<T, XlsxError>;

#[cfg(feature = "polars")]
use polars_core::prelude::*;

// fn check_utf8<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<()> {
//     // Reader уже посмотрел декларацию `<?xml ... encoding="..."?>`
//     // и выбрал нужный декодер.
//     let enc = reader.decoder().encoding(); // -> &'static encoding_rs::Encoding

//     if enc.name() != "UTF-8" {
//         bail!("unsupported XML encoding {}", enc.name());
//     }
//     Ok(())
// }
/// `XlsxEditor` provides functionality to open, modify, and save XLSX files.
/// It allows appending rows and tables to a specified sheet within an XLSX file.
pub struct XlsxEditor {
    src_path: PathBuf,
    sheet_path: String,
    sheet_xml: Vec<u8>,
    last_row: u32,
    styles_xml: Vec<u8>, // содержимое styles.xml
    workbook_xml: Vec<u8>, // содержимое workbook.xml (может изменяться)
    rels_xml: Vec<u8>,     // содержимое workbook.xml.rels
    new_files: Vec<(String, Vec<u8>)>, // новые или изменённые файлы для записи при save()
}

/// Work with files
impl XlsxEditor {
    /// Открывает книгу и подготавливает лист `sheet_id` (1‑based).
    pub fn open_sheet<P: AsRef<Path>>(src: P, sheet_id: usize) -> Result<Self> {
        let src_path = src.as_ref().to_path_buf();
        let mut zip = zip_crate::ZipArchive::new(File::open(&src_path)?)?;

        // ── sheet#.xml ───────────────────────────────────────────────
        let sheet_path = format!("xl/worksheets/sheet{sheet_id}.xml");

        // читаем XML листа в отдельном блоке, чтобы `sheet` дропнулся,
        // и эксклюзивный займ `zip` освободился
        let sheet_xml: Vec<u8> = {
            let mut sheet = zip
                .by_name(&sheet_path)
                .map_err(|_| XlsxError::FileNotFound { 
                    file: sheet_path.clone() 
                })?;
            let mut buf = Vec::with_capacity(sheet.size() as usize);
            sheet.read_to_end(&mut buf)?;
            buf
        };

        // ── styles.xml ───────────────────────────────────────────────
        let styles_xml: Vec<u8> = {
            let mut styles = zip
                .by_name("xl/styles.xml")
                .map_err(|_| XlsxError::FileNotFound { 
                    file: "xl/styles.xml".to_string() 
                })?;
            let mut buf = Vec::with_capacity(styles.size() as usize);
            styles.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml ───────────────────────────────────────────────
        let workbook_xml: Vec<u8> = {
            let mut wb = zip
                .by_name("xl/workbook.xml")
                .map_err(|_| XlsxError::FileNotFound { 
                    file: "xl/workbook.xml".to_string() 
                })?;
            let mut buf = Vec::with_capacity(wb.size() as usize);
            wb.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml.rels ──────────────────────────────────────────
        let rels_xml: Vec<u8> = {
            let mut rels = zip
                .by_name("xl/_rels/workbook.xml.rels")
                .map_err(|_| XlsxError::FileNotFound { 
                    file: "xl/_rels/workbook.xml.rels".to_string() 
                })?;
            let mut buf = Vec::with_capacity(rels.size() as usize);
            rels.read_to_end(&mut buf)?;
            buf
        };

        // ── вычисляем last_row ───────────────────────────────────────
        let mut reader = Reader::from_reader(sheet_xml.as_slice());
        // check_utf8(&mut reader)?;
        reader.config_mut().trim_text(true);

        let mut last_row = 0;
        while let Ok(ev) = reader.read_event() {
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
        }

        Ok(Self {
            src_path,
            sheet_path,
            sheet_xml,
            last_row,
            styles_xml,
            workbook_xml,
            rels_xml,
            new_files: Vec::new(),
        })
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
    /// A `Result` indicating success or a `XlsxError` if the save operation fails.
    pub fn save<P: AsRef<Path>>(&self, dst: P) -> Result<()> {
        let mut zin = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;
        let mut zout = zip_crate::ZipWriter::new(File::create(dst)?);

        let opt: zip_crate::write::FileOptions<'_, ()> = zip_crate::write::FileOptions::default()
            .compression_method(zip_crate::CompressionMethod::Deflated)
            .compression_level(Some(1));

        use std::collections::HashSet;
        let mut written: HashSet<String> = HashSet::new();

        for i in 0..zin.len() {
            let file = zin.by_index_raw(i)?;
            let name = file.name();

            if let Some((_, content)) = self.new_files.iter().find(|(p, _)| p == name) {
                // файл был создан/изменён в памяти – записываем его
                zout.start_file(name, opt)?;
                zout.write_all(content)?;
                written.insert(name.to_string());
                continue;
            }

            match name {
                "xl/workbook.xml" => {
                    zout.start_file(name, opt)?;
                    zout.write_all(&self.workbook_xml)?;
                }
                "xl/_rels/workbook.xml.rels" => {
                    zout.start_file(name, opt)?;
                    zout.write_all(&self.rels_xml)?;
                }
                _ if name == self.sheet_path => {
                    zout.start_file(name, opt)?;
                    zout.write_all(&self.sheet_xml)?;
                }
                "xl/styles.xml" => {
                    zout.start_file(name, opt)?;
                    zout.write_all(&self.styles_xml)?;
                }
                _ => zout.raw_copy_file(file)?,
            }
        }

        // добавляем файлы, которые ещё не были записаны
        for (path, content) in &self.new_files {
            if !written.contains(path) {
                zout.start_file(path, opt)?;
                if path == &self.sheet_path {
                    zout.write_all(&self.sheet_xml)?;
                } else {
                    zout.write_all(content)?;
                }
                written.insert(path.clone());
            }
        }

        zout.finish()?;
        Ok(())
    }
    /// Добавляет новый пустой лист с именем `sheet_name`
    /// (он станет первым во вкладках).
    pub fn add_worksheet(&mut self, sheet_name: &str) -> Result<&mut Self> {
        // 1) читаем исходный архив
        let sheet_names = scan(&self.src_path)?;
        if sheet_names.contains(&sheet_name.to_owned()) {
            return Err(XlsxError::SheetAlreadyExists { 
                name: sheet_name.to_string() 
            });
        }
        let mut zin = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;

        // ── workbook.xml и workbook.xml.rels берем из текущего состояния, а не читаем заново
        let mut wb_xml = self.workbook_xml.clone();
        let mut rels_xml = self.rels_xml.clone();

        // 2) определяем свободные sheetId / rId / sheet#.xml
        let mut max_sheet_id = 0u32;
        let mut rdr = Reader::from_reader(wb_xml.as_slice());
        rdr.config_mut().trim_text(true);
        while let Ok(ev) = rdr.read_event() {
            if let Event::Empty(ref e) | Event::Start(ref e) = ev {
                if e.name().as_ref() == b"sheet" {
                    if let Some(id) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"sheetId")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        max_sheet_id = max_sheet_id.max(id.parse::<u32>().unwrap_or(0));
                    }
                }
            }
            if matches!(ev, Event::Eof) {
                break;
            }
        }
        let new_sheet_id = max_sheet_id + 1;

        let mut max_rid = 0u32;
        let mut rdr = Reader::from_reader(rels_xml.as_slice());
        rdr.config_mut().trim_text(true);
        while let Ok(ev) = rdr.read_event() {
            if let Event::Empty(ref e) | Event::Start(ref e) = ev {
                if e.name().as_ref() == b"Relationship" {
                    if let Some(id) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"Id")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        if let Some(num) = id.strip_prefix("rId") {
                            max_rid = max_rid.max(num.parse::<u32>().unwrap_or(0));
                        }
                    }
                }
            }
            if matches!(ev, Event::Eof) {
                break;
            }
        }
        let new_rid = max_rid + 1;

        // номер нового файла sheet#.xml
        let mut max_sheet_file = 0usize;
        for i in 0..zin.len() {
            let name = zin.by_index(i)?.name().to_owned();
            if let Some(n) = name
                .strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<usize>().ok())
            {
                max_sheet_file = max_sheet_file.max(n);
            }
        }
        // также учитываем ещё не сохранённые новые файлы
        for (path, _) in &self.new_files {
            if let Some(n) = path.strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<usize>().ok())
            {
                max_sheet_file = max_sheet_file.max(n);
            }
        }
        let new_sheet_file = max_sheet_file + 1;
        let new_sheet_path = format!("xl/worksheets/sheet{new}.xml", new = new_sheet_file);
        let new_sheet_target = format!("worksheets/sheet{new}.xml", new = new_sheet_file);

        // 3) формируем новые теги
        let sheet_tag = format!(
            r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
            sheet_name, new_sheet_id, new_rid
        );
        let rel_tag = format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{}"/>"#,
            new_rid, new_sheet_target
        );

        // 4) вставляем <sheet …/> перед закрывающим </sheets>
        if let Some(pos) = wb_xml
            .windows(9) // длина "</sheets>"
            .rposition(|w| w == b"</sheets>")
        {
            // небольшая косметика: перенос + два пробела, чтобы сохранить формат
            let mut tagged = Vec::with_capacity(sheet_tag.len() + 3);
            tagged.extend_from_slice(b"\n  "); // \n, отступ
            tagged.extend_from_slice(sheet_tag.as_bytes());

            wb_xml.splice(pos..pos, tagged);
        } else {
            return Err(XlsxError::XmlTagNotFound { 
                tag: "</sheets>".to_string() 
            });
        }

        // 5) вставляем Relationship перед </Relationships>
        if let Some(pos) = rels_xml.windows(16).rposition(|w| w == b"</Relationships>") {
            rels_xml.splice(pos..pos, rel_tag.as_bytes().iter().copied());
        } else {
            return Err(XlsxError::XmlTagNotFound { 
                tag: "</Relationships>".to_string() 
            });
        }

        // 6) минимальный XML нового листа
        const EMPTY_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData> </sheetData>
        </worksheet>"#;

        
        // обновляем внутреннее состояние
        self.workbook_xml = wb_xml;
        self.rels_xml = rels_xml;
        if let Some(pair) = self
            .new_files
            .iter_mut()
            .find(|(p, _)| p == &new_sheet_path)
        {
            pair.1 = EMPTY_SHEET.as_bytes().to_vec();
        } else {
            self
                .new_files
                .push((new_sheet_path.clone(), EMPTY_SHEET.as_bytes().to_vec()));
        }

        // перед переключением сохраняем изменённый текущий лист
        let cur_path = self.sheet_path.clone();
        let cur_xml = self.sheet_xml.clone();
        if let Some(pair) = self.new_files.iter_mut().find(|(p, _)| p == &cur_path) {
            pair.1 = cur_xml;
        } else {
            self.new_files.push((cur_path, cur_xml));
        }

        // переключаем редактор на новый лист
        self.sheet_path = new_sheet_path;
        self.sheet_xml = EMPTY_SHEET.as_bytes().to_vec();
        self.last_row = 0;

        Ok(self)
    }
}

/// Polars

impl XlsxEditor {
    /// Overwrites the specified `sheet_name` with the provided Polars `DataFrame`.
    ///
    /// The function opens the workbook, completely clears any existing data inside the
    /// sheet, then writes the column names followed by the DataFrame's rows.  Writing
    /// starts from the coordinate provided in `start_cell` or "A1" when `None`.
    ///
    /// Note: The workbook on disk is **overwritten**. If you need to keep the original
    #[cfg(feature = "polars")]
    pub fn with_polars(&mut self, df: &DataFrame, start_cell: Option<&str>) -> Result<()> {
        // Remove existing data so that the DataFrame overwrites the sheet.
        // self.clear_sheet()?;

        // Helper – convert the entire DataFrame into Vec<Vec<String>> where the first
        // row contains column headers.
        let mut rows: Vec<Vec<String>> = Vec::with_capacity(df.height() + 1);
        // 1. Header row.
        rows.push(
            df.get_columns()
                .iter()
                .map(|s| s.name().to_string())
                .collect(),
        );
        // 2. Data rows.
        for row_idx in 0..df.height() {
            let mut row = Vec::with_capacity(df.width());
            for series in df.get_columns() {
                let cell_val = series
                    .get(row_idx)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                row.push(cell_val.replace("\"", ""));
            }
            rows.push(row);
        }

        // Determine the starting coordinate.
        let start_coord = start_cell.unwrap_or("A1");
        // Write the table and save.
        self.append_table_at(start_coord, rows)?;
        Ok(())
    }
}

/// Testing

impl XlsxEditor {
    fn _save_in_ram(
        src: &Path,
        dst: &Path,
        sheet_path: &str,
        new_xml: &[u8],
    ) -> Result<()> {
        let mut zin = zip_crate::ZipArchive::new(File::open(src)?)?;

        // 1) Буфер‑growable в RAM
        let mut mem = Vec::with_capacity(10 * 1024 * 1024); // грубая оценка, чтобы меньше realloc
        {
            let cursor = Cursor::new(&mut mem);
            let mut zout = zip_crate::ZipWriter::new(cursor);
            let opt: zip_crate::write::FileOptions<'_, ()> =
                zip_crate::write::FileOptions::default()
                    .compression_method(zip_crate::CompressionMethod::Deflated);

            for i in 0..zin.len() {
                let file = zin.by_index_raw(i)?;
                let name = file.name().to_owned();
                if name == sheet_path {
                    zout.start_file(&name, opt)?;
                    zout.write_all(new_xml)?;
                } else {
                    zout.raw_copy_file(file)?;
                }
            }
            zout.finish()?; // важно закрыть writer, иначе central dir не допишется
        } // cursor → drop, но mem остаётся

        // 2) Одним системным вызовом кладём всё на диск
        std::fs::write(dst, &mem)?;
        Ok(())
    }

    /// Clears all existing rows from the currently opened sheet.
    ///
    /// This removes every `<row>` element inside the `<sheetData>` section and resets
    /// `self.last_row` to `0` so that subsequent writes start at the first row.
    fn _clear_sheet(&mut self) -> Result<()> {
        // Locate the opening <sheetData …> tag.
        let start_opt = self
            .sheet_xml
            .windows(10) // length of "<sheetData"
            .position(|w| w == b"<sheetData");
        let start_idx = match start_opt {
            Some(idx) => idx,
            None => return Err(XlsxError::XmlTagNotFound { 
                tag: "<sheetData>".to_string() 
            }),
        };

        // Find the end of the opening tag, i.e. the first '>' after the tag start.
        let open_tag_end = self.sheet_xml[start_idx..]
            .iter()
            .position(|&b| b == b'>')
            .map(|rel| start_idx + rel + 1)
            .ok_or_else(|| XlsxError::MalformedXmlTag { 
                tag: "<sheetData>".to_string() 
            })?;

        // Locate the closing </sheetData> tag.
        let end_opt = self
            .sheet_xml
            .windows(12) // length of "</sheetData>"
            .position(|w| w == b"</sheetData>");
        let end_idx = match end_opt {
            Some(idx) => idx,
            None => return Err(XlsxError::XmlTagNotFound { 
                tag: "</sheetData>".to_string() 
            }),
        };

        // Remove everything between the tags.
        self.sheet_xml.drain(open_tag_end..end_idx);
        // Reset state.
        self.last_row = 0;
        Ok(())
    }
}

/// Info Part

impl XlsxEditor {
    /// Returns the last non-empty row index for the specified column or columns.
    ///
    /// The `columns` argument can be a single column such as "B" or multiple comma–separated
    /// columns such as "B,D". The function scans the sheet for the highest populated row
    /// across all specified columns and returns that 1-based row index. If no data is found
    /// in those columns, `Ok(0)` is returned.
    pub fn get_last_row_index(&self, columns: &str) -> Result<u32> {
        // Local helper to split coordinate like "C12" -> ("C", 12)
        fn split_coord(coord: &str) -> (String, u32) {
            let pos = coord
                .find(|c: char| c.is_ascii_digit())
                .unwrap_or(coord.len());
            let col = coord[..pos].to_ascii_uppercase();
            let row: u32 = coord[pos..].parse().unwrap_or(0);
            (col, row)
        }

        let targets: std::collections::HashSet<String> = columns
            .split(',')
            .map(|s| s.trim().to_ascii_uppercase())
            .collect();
        if targets.is_empty() {
            return Err(XlsxError::NoColumnsSupplied);
        }

        let mut reader = Reader::from_reader(self.sheet_xml.as_slice());
        reader.config_mut().trim_text(true);
        let mut last_row: u32 = 0;

        while let Ok(ev) = reader.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"c" => {
                    // locate the coordinate attribute r="A1"
                    if let Some(coord) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        let (col, row) = split_coord(&coord);
                        if targets.contains(&col) {
                            if row > last_row {
                                last_row = row;
                            }
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(last_row)
    }

    /// Returns a vector with the last non-empty row indices for every column in the inclusive
    /// range like "A:E". The resulting vector has the same length as the number of columns in
    /// the range and is ordered left-to-right.
    ///
    /// Example: `get_last_roww_index("A:C")` might return `[10, 12, 7]`.
    pub fn get_last_roww_index(&self, range: &str) -> Result<Vec<u32>> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            return Err(XlsxError::InvalidRange { 
                range: range.to_string() 
            });
        }
        // Reuse helpers from outer function
        fn letters_to_col_idx(s: &str) -> usize {
            s.bytes().fold(0, |acc, b| {
                acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
            }) - 1
        }
        fn split_coord(coord: &str) -> (String, u32) {
            let pos = coord
                .find(|c: char| c.is_ascii_digit())
                .unwrap_or(coord.len());
            let col = coord[..pos].to_ascii_uppercase();
            let row: u32 = coord[pos..].parse().unwrap_or(0);
            (col, row)
        }

        let start = parts[0].trim().to_ascii_uppercase();
        let end = parts[1].trim().to_ascii_uppercase();

        let start_idx = letters_to_col_idx(&start);
        let end_idx = letters_to_col_idx(&end);
        if start_idx > end_idx {
            return Err(XlsxError::InvalidRangeOrder);
        }
        let mut per_col_last: Vec<u32> = vec![0; end_idx - start_idx + 1];

        let mut reader = Reader::from_reader(self.sheet_xml.as_slice());
        reader.config_mut().trim_text(true);

        while let Ok(ev) = reader.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"c" => {
                    if let Some(coord) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        let (col, row) = split_coord(&coord);
                        let idx = letters_to_col_idx(&col);
                        if idx >= start_idx && idx <= end_idx {
                            let vec_idx = idx - start_idx;
                            if row > per_col_last[vec_idx] {
                                per_col_last[vec_idx] = row;
                            }
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(per_col_last)
    }
}

///Style Part
impl XlsxEditor {
    pub fn set_number_format(&mut self, range: &str, fmt: &str) -> Result<()> {
        let style_id = self.ensure_style(Some(fmt), None, None, None)?;
        match parse_target(range)? {
            Target::Cell(c) => self.apply_style_to_cell(&c, style_id)?,
            Target::Rect { c0, r0, c1, r1 } => {
                for r in r0..=r1 {
                    for c in c0..=c1 {
                        let coord = format!("{}{}", col_letter(c), r);
                        self.apply_style_to_cell(&coord, style_id)?;
                    }
                }
            }
            Target::Col(_col) => self.force_column_number_format(&range.replace(":", ""), fmt)?,
            Target::Row(_row) => todo!(),
        }
        Ok(())
    }

    /// Гарантирует наличие стиля и возвращает его `styleId`
    ///
    /// * `num_fmt` – текст формата, например `"dd.mm.yyyy"`.  
    ///   `None`  →  стандартный «General» (`numFmtId = 0`).
    /// * `font_id`, `fill_id` – индексы существующих ресурсов  
    ///   (`<fonts>`, `<fills>`); можно передать `None`.
    /// * `align` – описание выравнивания; если `Some(_)`, всегда создаётся
    ///   новый `<xf>` (упрощение — иначе пришлось бы глубоко сравнивать).
    fn ensure_style(
        &mut self,
        num_fmt: Option<&str>,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        // ────────────────────────────────────────────────
        // 1.  numFmt  → получаем или добавляем  numFmtId
        // ────────────────────────────────────────────────
        let fmt_id: u32 = if let Some(code) = num_fmt {
            let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
            rdr.config_mut().trim_text(true);

            let mut found_id = None;
            let mut max_custom_id = 163u32; // 0‑163 — builtin

            while let Ok(ev) = rdr.read_event() {
                match ev {
                    Event::Start(ref e) | Event::Empty(ref e) if e.name().as_ref() == b"numFmt" => {
                        let mut id = None::<u32>;
                        let mut text = None::<String>;
                        for a in e.attributes().with_checks(false).flatten() {
                            match a.key.as_ref() {
                                b"numFmtId" => {
                                    id = Some(String::from_utf8_lossy(&a.value).parse::<u32>()?)
                                }
                                b"formatCode" => {
                                    text = Some(String::from_utf8_lossy(&a.value).into_owned())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(i), Some(t)) = (id, text) {
                            if t == code {
                                found_id = Some(i);
                            }
                            if i > max_custom_id {
                                max_custom_id = i;
                            }
                        }
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }

            if let Some(i) = found_id {
                i
            } else {
                // вставляем <numFmt …/>  + bump count
                let new_id = max_custom_id + 1;
                let tag = format!(r#"<numFmt numFmtId="{}" formatCode="{}"/>"#, new_id, code);

                if let Some(end) = find_bytes(&self.styles_xml, b"</numFmts>") {
                    self.styles_xml
                        .splice(end..end, tag.as_bytes().iter().copied());
                    bump_count(&mut self.styles_xml, b"<numFmts", b"count=\"")?;
                } else {
                    // блока ещё нет – создаём сразу после <styleSheet …>
                    let insert = find_bytes(&self.styles_xml, b">")
                        .ok_or_else(|| XlsxError::XmlTagNotFound { 
                            tag: "<styleSheet>".to_string() 
                        })?
                        + 1;
                    let block = format!(r#"<numFmts count="1">{}</numFmts>"#, tag);
                    self.styles_xml
                        .splice(insert..insert, block.as_bytes().iter().copied());
                }
                new_id
            }
        } else {
            0 // General
        };

        // ────────────────────────────────────────────────
        // 2.  ищем существующий <xf> с теми же id
        //     (если alignment == Some(_) — пропускаем поиск)
        // ────────────────────────────────────────────────
        if align.is_none() {
            let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
            rdr.config_mut().trim_text(true);

            let mut in_xfs = false;
            let mut idx: u32 = 0;

            while let Ok(ev) = rdr.read_event() {
                match ev {
                    Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                    Event::End(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = false,

                    Event::Start(ref e) | Event::Empty(ref e)
                        if in_xfs && e.name().as_ref() == b"xf" =>
                    {
                        let mut num = None::<u32>;
                        let mut fnt = None::<u32>;
                        let mut fil = None::<u32>;
                        for a in e.attributes().with_checks(false).flatten() {
                            match a.key.as_ref() {
                                b"numFmtId" => {
                                    num = Some(String::from_utf8_lossy(&a.value).parse()?)
                                }
                                b"fontId" => fnt = Some(String::from_utf8_lossy(&a.value).parse()?),
                                b"fillId" => fil = Some(String::from_utf8_lossy(&a.value).parse()?),
                                _ => {}
                            }
                        }
                        let num_ok = num.unwrap_or(0) == fmt_id;
                        let font_ok = font_id.map_or(true, |v| Some(v) == fnt);
                        let fill_ok = fill_id.map_or(true, |v| Some(v) == fil);

                        if num_ok && font_ok && fill_ok {
                            return Ok(idx); // ре‑юзаем стиль
                        }
                        idx += 1;
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }
        }

        // ────────────────────────────────────────────────
        // 3.  формируем новый  <xf …>  и добавляем
        // ────────────────────────────────────────────────
        let mut xf = String::from("<xf ");
        if let Some(fid) = font_id {
            xf += &format!(r#"fontId="{fid}" applyFont="1" "#);
        }
        if let Some(fid) = fill_id {
            xf += &format!(r#"fillId="{fid}" applyFill="1" "#);
        }

        xf += &format!(
            r#"numFmtId="{}"{} "#,
            fmt_id,
            if num_fmt.is_some() {
                r#" applyNumberFormat="1""#
            } else {
                ""
            }
        );

        if align.is_some() {
            xf += r#"applyAlignment="1" "#;
        }
        xf += r#"borderId="0" xfId="0">"#;

        if let Some(al) = align {
            xf += "<alignment";
            if let Some(h) = &al.horiz {
                xf += &format!(r#" horizontal="{}""#, h);
            }
            if let Some(v) = &al.vert {
                xf += &format!(r#" vertical="{}""#, v);
            }
            if al.wrap {
                xf += r#" wrapText="1""#;
            }
            xf += "/>";
        }
        xf += "</xf>";

        // вставляем перед </cellXfs>
        let pos = find_bytes(&self.styles_xml, b"</cellXfs>")
            .ok_or_else(|| XlsxError::XmlTagNotFound { 
                tag: "</cellXfs>".to_string() 
            })?;
        self.styles_xml
            .splice(pos..pos, xf.as_bytes().iter().copied());

        // обновляем счётчик
        bump_count(&mut self.styles_xml, b"<cellXfs", b"count=\"")?;

        // индекс нового стиля = старое количество xfs
        // (намеренно пересчитывать не нужно — bump_count ещё не меняет self.styles_xml в этой части)
        let new_id = {
            // очень дёшево: просто считаем, сколько раз встретили <xf …> в cellXfs
            let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
            rdr.config_mut().trim_text(true);
            let mut in_xfs = false;
            let mut cnt = 0u32;
            while let Ok(ev) = rdr.read_event() {
                match ev {
                    Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                    Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,
                    Event::Start(ref e) | Event::Empty(ref e)
                        if in_xfs && e.name().as_ref() == b"xf" =>
                    {
                        cnt += 1
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }
            cnt - 1 // последний добавленный
        };

        Ok(new_id)
    }

    // ──────────────────────────────────────────────────────────────────────
    // 3. ПРИМЕНИТЬ СТИЛЬ
    // ──────────────────────────────────────────────────────────────────────
    thread_local! { static REENTRY: std::cell::Cell<bool> = std::cell::Cell::new(false); }

    fn apply_style_to_cell(&mut self, coord: &str, style: u32) -> Result<()> {
        // ── вычисляем номер строки, чтобы найти <row …>
        let row_num = coord.trim_start_matches(|c: char| c.is_ascii_alphabetic());

        let row_tag = format!(r#"<row r="{row_num}""#);

        // если строки нет — создаём ячейку обычным set_cell и возвращаемся
        let row_pos = match find_bytes(&self.sheet_xml, row_tag.as_bytes()) {
            Some(p) => p,
            None => {
                // ── защита от бесконечной рекурсии ────────────────────
                let reentered = Self::REENTRY.with(|c| {
                    let old = c.get();
                    c.set(true);
                    old
                });
                if !reentered {
                    self.set_cell(coord, "")?;
                    Self::REENTRY.with(|c| c.set(false));
                    return self.apply_style_to_cell(coord, style);
                }
                return Ok(()); // второй заход – просто выходим
            }
        };

        // конец строки </row>
        let row_end = find_bytes_from(&self.sheet_xml, b"</row>", row_pos)
            .ok_or_else(|| XlsxError::XmlTagNotFound { 
                tag: "</row>".to_string() 
            })?;

        // ── ищем ячейку <c r="A1" …>
        let cell_tag = format!(r#"<c r="{coord}""#);
        let cpos = match find_bytes_from(&self.sheet_xml, cell_tag.as_bytes(), row_pos) {
            Some(p) => p, // есть — будем править
            None => {
                // нет — вставим пустую с нужным style перед </row>
                let new_cell = format!(r#"<c r="{coord}" s="{style}"/>"#);
                self.sheet_xml
                    .splice(row_end..row_end, new_cell.as_bytes().iter().copied());
                return Ok(());
            }
        };

        // граница открывающего тега ячейки '>'
        let ctag_end = find_bytes_from(&self.sheet_xml, b">", cpos)
            .ok_or_else(|| XlsxError::MalformedXmlTag { 
                tag: "<c>".to_string() 
            })?;

        // ── проверяем/ставим атрибут s="…"
        if let Some(sattr) = find_bytes_from(&self.sheet_xml, b" s=\"", cpos) {
            if sattr < ctag_end {
                // уже есть s="…" → заменить число
                let val_start = sattr + 4;
                let val_end = find_bytes_from(&self.sheet_xml, b"\"", val_start + 1).unwrap();
                self.sheet_xml.splice(
                    val_start..val_end,
                    style.to_string().as_bytes().iter().copied(),
                );
                return Ok(());
            }
        }
        // атрибута нет — вставляем перед '>'
        self.sheet_xml.splice(
            ctag_end..ctag_end,
            format!(r#" s="{style}""#).as_bytes().iter().copied(),
        );
        Ok(())
    }

    fn apply_style_to_row(&mut self, row: u32, style: u32) -> Result<()> {
        let row_tag = format!("<row r=\"{row}\"");
        if let Some(_pos) = find_bytes(&self.sheet_xml, row_tag.as_bytes()) {
            let _ = &self.insert_or_replace_attr(b"s", &style.to_string());
            let _ = &self.insert_or_replace_attr(b"customFormat", "1");
        } else {
            // строки нет — создаём пустую с атрибутами и вставляем по порядку
            let new_row = format!(r#"<row r="{row}" s="{style}" customFormat="1"></row>"#);
            // вставим перед </sheetData>
            let pos = find_bytes(&self.sheet_xml, b"</sheetData>")
                .ok_or_else(|| XlsxError::XmlTagNotFound { 
                    tag: "</sheetData>".to_string() 
                })?;
            self.sheet_xml
                .splice(pos..pos, new_row.as_bytes().iter().copied());
        }
        // дополнительно пройти по всем <c r="??row"> и проставить s
        let pat = format!(r#" r="([A-Z]+){}""#, row);
        let re = Regex::new(&pat).unwrap();
        for cap in re.find_iter(std::str::from_utf8(&self.sheet_xml.clone())?) {
            let start = cap.start();
            if let Some(_cpos) = self.sheet_xml[..start]
                .windows(3)
                .rposition(|w| w == b"<c ")
            {
                self.apply_style_to_cell(&cap.as_str()[3..], style)?; // cap = r="A12"
            }
        }
        Ok(())
    }

    fn insert_or_replace_attr(&mut self, key: &[u8], val: &str) {
        if let Some(p) = find_bytes(&self.sheet_xml, key) {
            let start = p + key.len() + 2; //  key + ="
            let end = find_bytes_from(&self.sheet_xml, b"\"", start).unwrap();
            self.sheet_xml
                .splice(start..end, val.as_bytes().iter().copied());
        } else {
            let end = self.sheet_xml.len() - 1; // '>'
            self.sheet_xml.splice(
                end..end,
                format!(" {}=\"{}\"", std::str::from_utf8(key).unwrap(), val)
                    .as_bytes()
                    .iter()
                    .copied(),
            );
        }
    }
}

// ────────────────────────────────────────────────────────────────
// Multiple columns – быстрая работа со стилями столбцов
// ────────────────────────────────────────────────────────────────
impl XlsxEditor {
    // ---------- вспомогалка -------------------------------------

    /// Вставляет или заменяет атрибут **внутри уже найденного тега**.
    /// `tag_start` и `tag_end` — абсолютные байтовые границы тега
    /// в `self.sheet_xml` (конец указывает сразу за `>` или `/>`).
    fn insert_or_replace_attr_in_tag(
        &mut self,
        tag_start: usize,
        tag_end: usize,
        key: &[u8],
        val: &str,
    ) {
        // попробуем найти существующий атрибут
        if let Some(p) = find_bytes_from(&self.sheet_xml, key, tag_start) {
            if p < tag_end {
                let start = p + key.len() + 2; // key="  → начало значения
                let end = find_bytes_from(&self.sheet_xml, b"\"", start).unwrap_or(tag_end);
                self.sheet_xml
                    .splice(start..end, val.as_bytes().iter().copied());
                return;
            }
        }
        // не нашли — добавляем перед закрывающей > или />
        let insert_at = tag_end
            - if self.sheet_xml[tag_end - 2] == b'/' {
                2
            } else {
                1
            };
        self.sheet_xml.splice(
            insert_at..insert_at,
            format!(" {}=\"{}\"", std::str::from_utf8(key).unwrap(), val)
                .as_bytes()
                .iter()
                .copied(),
        );
    }

    // ---------- ядро ---------------------------------------------

    /// «Колончный» способ: задаёт формат через `<col … style="…"/>`
    /// (Excel применит стиль даже к будущим ячейкам столбца).
    pub fn set_column_number_format(&mut self, col_letter: &str, fmt: &str) -> Result<()> {
        let style_id = self.ensure_style(Some(fmt), None, None, None)?;
        self.apply_style_to_column(col_index(col_letter) as u32, style_id)
    }

    /// Тот же формат, но **насильно** проставляет стиль
    /// всем существующим ячейкам столбца.
    pub fn force_column_number_format(&mut self, col_letter: &str, fmt: &str) -> Result<()> {
        self.set_column_number_format(col_letter, fmt)?;

        let style_id = self.ensure_style(Some(fmt), None, None, None)?;
        let sid_str = style_id.to_string();
        let pat = format!(
            r#"<c\b[^>]*\br="{}[0-9]+"[^>]*>"#,
            col_letter.to_ascii_uppercase()
        );
        let re = Regex::new(&pat)?;

        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len() + 512);
        let mut last = 0usize;

        for m in re.find_iter(std::str::from_utf8(&src)?) {
            dst.extend_from_slice(&src[last..m.start()]);

            let cell_start = m.start();
            let tag_end = find_bytes_from(&src, b">", cell_start).unwrap() + 1;
            let mut cell = src[cell_start..tag_end].to_vec();

            if let Some(p) = find_bytes(&cell, b" s=\"") {
                let v0 = p + 4;
                let v1 = find_bytes_from(&cell, b"\"", v0 + 1).unwrap();
                cell.splice(v0..v1, sid_str.as_bytes().iter().copied());
            } else {
                let ins = if cell[cell.len() - 2] == b'/' {
                    cell.len() - 2
                } else {
                    cell.len() - 1
                };
                cell.splice(
                    ins..ins,
                    format!(r#" s="{}""#, sid_str).as_bytes().iter().copied(),
                );
            }
            dst.extend_from_slice(&cell);
            last = tag_end;
        }
        dst.extend_from_slice(&src[last..]);
        self.sheet_xml = dst;
        Ok(())
    }

    // ---------- низкоуровневое: работа с блоком <cols> -----------

    /// Внутренняя работа со стилем столбца через тег `<col …/>`.
    fn apply_style_to_column(&mut self, col: u32, style: u32) -> Result<()> {
        let col_tag = format!(r#"<col min="{c}" max="{c}""#, c = col + 1);

        if let Some(tag_start) = find_bytes(&self.sheet_xml, col_tag.as_bytes()) {
            // ── тег существует, определяем его границу
            let tag_end = if let Some(p) = find_bytes_from(&self.sheet_xml, b"/>", tag_start) {
                p + 2 // "/>"  → длина 2
            } else {
                let p = find_bytes_from(&self.sheet_xml, b">", tag_start)
                    .ok_or_else(|| XlsxError::MalformedXmlTag { 
                        tag: "<col>".to_string() 
                    })?;
                p + 1 // ">"   → длина 1
            };
            // ставим / заменяем  style="…"
            self.insert_or_replace_attr_in_tag(tag_start, tag_end, b"style", &style.to_string());
        } else {
            // ‑‑ блока нет: создаём новый <col … style="…"/>
            let new_tag = format!(r#"<col min="{c}" max="{c}" style="{style}"/>"#, c = col + 1);
            if let Some(cols_start) = find_bytes(&self.sheet_xml, b"<cols") {
                // вставляем в существующий блок
                let insert = find_bytes_from(&self.sheet_xml, b">", cols_start).unwrap() + 1;
                self.sheet_xml
                    .splice(insert..insert, new_tag.as_bytes().iter().copied());
            } else {
                // создаём новый <cols> перед <sheetData>
                let sheetdata_pos =
                    find_bytes(&self.sheet_xml, b"<sheetData")
                        .ok_or_else(|| XlsxError::XmlTagNotFound { 
                            tag: "<sheetData>".to_string() 
                        })?;
                let block = format!("<cols>{}</cols>", new_tag);
                self.sheet_xml.splice(
                    sheetdata_pos..sheetdata_pos,
                    block.as_bytes().iter().copied(),
                );
            }
        }
        Ok(())
    }
}

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
    /// A `Result` containing an `XlsxEditor` instance if successful, or a `XlsxError` otherwise.
    pub fn open<P: AsRef<Path>>(src: P, sheet_name: &str) -> Result<Self> {
        let sheet_names = scan(src.as_ref())?;
        let sheet_id = sheet_names
            .iter()
            .position(|n| n == sheet_name)
            .ok_or_else(|| XlsxError::SheetNotFound { 
                name: sheet_name.to_string() 
            })?
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
    /// A `Result` indicating success or a `XlsxError` if the operation fails.
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
            Err(XlsxError::XmlTagNotFound { 
                tag: "</sheetData>".to_string() 
            })
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
    /// A `Result` indicating success or a `XlsxError` if the operation fails.
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
            Err(XlsxError::XmlTagNotFound { 
                tag: "</sheetData>".to_string() 
            })
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
    /// A `Result` indicating success or a `XlsxError` if the operation fails.
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
            .ok_or_else(|| XlsxError::InvalidCoordinate { 
                coord: "no digits found".to_string() 
            })?;
        let col_letters = &start_coord[..row_start_pos];
        let start_col_idx = letters_to_col_idx(col_letters);
        let current_row_num: u32 = start_coord[row_start_pos..]
            .parse()
            .map_err(|_| XlsxError::InvalidCoordinate { 
                coord: start_coord.to_string() 
            })?;

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
            Err(XlsxError::XmlTagNotFound { 
                tag: "</sheetData>".to_string() 
            })
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
    /// A `Result` indicating success or a `XlsxError` if the operation fails.
    pub fn set_cell<S: ToString>(&mut self, coord: &str, value: S) -> Result<()> {
        // Extract row number from coordinate.
        let row_start = coord
            .find(|c: char| c.is_ascii_digit())
            .ok_or_else(|| XlsxError::InvalidCoordinate {
                coord: coord.to_string()
            })?
        let row_num: u32 = coord[row_start..]
            .parse()
            .map_err(|_| XlsxError::InvalidCoordinate { 
                coord: coord.to_string() 
            })?;

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
                    .ok_or_else(|| XlsxError::XmlTagNotFound { 
                        tag: "</sheetData>".to_string() 
                    })?,
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
    let mut zip = zip_crate::ZipArchive::new(File::open(src)?)?;
    let mut wb = zip
        .by_name("xl/workbook.xml")
        .map_err(|_| XlsxError::FileNotFound { 
            file: "xl/workbook.xml".to_string() 
        })?;

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
            .ok_or_else(|| XlsxError::XmlTagNotFound { 
                tag: "</sheetData>".to_string() 
            })?
            + "</sheetData>".len();

        let (insert_pos, created) = if let Some(pos) = find_bytes(&self.sheet_xml, b"<mergeCells") {
            // уже есть блок
            bump_count(&mut self.sheet_xml, b"<mergeCells", b"count=\"")?;
            let end = find_bytes_from(&self.sheet_xml, b"</mergeCells>", pos)
                .ok_or_else(|| XlsxError::XmlTagNotFound { 
                    tag: "</mergeCells>".to_string() 
                })?;
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

impl XlsxEditor {
    /// Добавляет шрифт и возвращает его fontId
    fn ensure_font(&mut self, name: &str, size: f32, bold: bool, italic: bool) -> Result<u32> {
        // 1) сколько шрифтов уже есть?
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut fonts_cnt = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) | Event::Empty(ref e) if e.name().as_ref() == b"font" => {
                    fonts_cnt += 1
                }
                Event::Eof => break,
                _ => {}
            }
        }

        // 2) формируем <font> … и вставляем перед </fonts>
        let insert = find_bytes(&self.styles_xml, b"</fonts>")
            .ok_or_else(|| XlsxError::XmlTagNotFound { 
                tag: "</fonts>".to_string() 
            })?;
        let mut xml = String::from("<font>");
        if bold {
            xml.push_str("<b/>");
        }
        if italic {
            xml.push_str("<i/>");
        }
        xml.push_str(&format!(r#"<sz val="{}"/>"#, size));
        xml.push_str(&format!(r#"<name val="{}"/>"#, name));
        xml.push_str("</font>");
        self.styles_xml
            .splice(insert..insert, xml.as_bytes().iter().copied());

        // 3) bump count="…"
        bump_count(&mut self.styles_xml, b"<fonts", b"count=\"")?;

        Ok(fonts_cnt) // индекс нового = старое количество
    }

    /// Добавляет однотонную заливку и возвращает fillId
    fn ensure_fill(&mut self, rgb: &str) -> Result<u32> {
        // 1) текущее количество <fill>
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut fills_cnt = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) | Event::Empty(ref e) if e.name().as_ref() == b"fill" => {
                    fills_cnt += 1
                }
                Event::Eof => break,
                _ => {}
            }
        }

        // 2) вставляем перед </fills>
        let insert = find_bytes(&self.styles_xml, b"</fills>")
            .ok_or_else(|| XlsxError::XmlTagNotFound { 
                tag: "</fills>".to_string() 
            })?;
        let xml = format!(
            r#"<fill><patternFill patternType="solid"><fgColor rgb="{}"/><bgColor indexed="64"/></patternFill></fill>"#,
            rgb
        );
        self.styles_xml
            .splice(insert..insert, xml.as_bytes().iter().copied());

        // 3) bump count
        bump_count(&mut self.styles_xml, b"<fills", b"count=\"")?;

        Ok(fills_cnt)
    }
}
impl XlsxEditor {
    /// Устанавливает **шрифт** для указанного диапазона.
    ///
    /// * `range` –  "B2", "A1:C10", "D:", "5:" и т.п.  
    /// * `name`  –  название шрифта, например `"Calibri"`  
    /// * `size`  –  кегль в пунктах (`11.0`)  
    /// * `bold`, `italic` –  дополнительные атрибуты
    pub fn set_font(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
    ) -> Result<&mut Self> {
        let new_font = self.ensure_font(name, size, bold, italic)?;
        self._merge_and_apply(range, |_old_font, fill| (Some(new_font), fill))?;
        Ok(self)
    }

    /// Устанавливает однотонную **заливку** (`rgb` = "FFFF00" без `#`)
    pub fn set_fill(&mut self, range: &str, rgb: &str) -> Result<&mut Self> {
        let new_fill = self.ensure_fill(rgb)?;
        self._merge_and_apply(range, |font, _old_fill| (font, Some(new_fill)))?;
        Ok(self)
    }

    // // ------------------------------------------------------------------
    // // Вспомогательный приватный метод: применить styleId к любому Target
    // // ------------------------------------------------------------------
    // fn apply_style_range(&mut self, range: &str, style_id: u32) -> Result<()> {
    //     match parse_target(range)? {
    //         Target::Cell(c) => self.apply_style_to_cell(&c, style_id)?,
    //         Target::Rect { c0, r0, c1, r1 } => {
    //             for r in r0..=r1 {
    //                 for c in c0..=c1 {
    //                     let coord = format!("{}{}", col_letter(c), r);
    //                     self.apply_style_to_cell(&coord, style_id)?;
    //                 }
    //             }
    //         }
    //         Target::Col(col) => self.apply_style_to_column(col, style_id)?,
    //         Target::Row(row) => self.apply_style_to_row(row, style_id)?,
    //     }
    //     Ok(())
    // }

    // ────────────────────────────────────────────────────────────
    // ↓↓↓ ВНУТРЕННЕЕ — «слить» старый и новый атрибуты и выдать xf
    // ────────────────────────────────────────────────────────────

    /// Из существующего styleId достаём (fontId, fillId).
    fn xf_components(&self, style_id: u32) -> Result<(Option<u32>, Option<u32>)> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut idx = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    if idx == style_id {
                        let mut font = None;
                        let mut fill = None;
                        for a in e.attributes().with_checks(false).flatten() {
                            match a.key.as_ref() {
                                b"fontId" => {
                                    font = Some(String::from_utf8_lossy(&a.value).parse()?)
                                }
                                b"fillId" => {
                                    fill = Some(String::from_utf8_lossy(&a.value).parse()?)
                                }
                                _ => {}
                            }
                        }
                        return Ok((font, fill));
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        // если не нашли (ячейка без s или out‑of‑range) → None
        Ok((None, None))
    }

    /// Применяет к диапазону style, «сливая» его с тем, что уже есть.
    ///
    /// `f` — функция, которая получает `(old_font, old_fill)` и
    /// возвращает `(font_after, fill_after)`.
    fn _merge_and_apply<F>(&mut self, range: &str, mut f: F) -> Result<()>
    where
        F: FnMut(Option<u32>, Option<u32>) -> (Option<u32>, Option<u32>),
    {
        let style = self.ensure_style(
            None, // старые col/row стили
            None, None, None,
        )?;

        match parse_target(range)? {
            Target::Cell(c) => self._merge_one(&c, &mut f)?,

            Target::Row(r) => self.apply_style_to_row(
                r, // row‑level проще:
                style,
            )?, // игнорируем
            Target::Col(c) => {
                let style = self.ensure_style(None, None, None, None)?;
                self.apply_style_to_column(c, style)?
            }
            Target::Rect { c0, r0, c1, r1 } => {
                for rr in r0..=r1 {
                    for cc in c0..=c1 {
                        let coord = format!("{}{}", col_letter(cc), rr);
                        self._merge_one(&coord, &mut f)?;
                    }
                }
            }
        }
        Ok(())
    }

    fn _merge_one<F>(&mut self, coord: &str, f: &mut F) -> Result<()>
    where
        F: FnMut(Option<u32>, Option<u32>) -> (Option<u32>, Option<u32>),
    {
        // какой стиль уже стоит?
        let old_sid = self.cell_style_id(coord)?;
        let (old_font, old_fill) = if let Some(s) = old_sid {
            self.xf_components(s)?
        } else {
            (None, None)
        };

        let (new_font, new_fill) = f(old_font, old_fill);
        let new_sid = self.ensure_style(None, new_font, new_fill, None)?;

        self.apply_style_to_cell(coord, new_sid)
    }

    /// Выдёргивает styleId из атрибута `s="…"`, если есть.
    fn cell_style_id(&self, coord: &str) -> Result<Option<u32>> {
        let tag = format!(r#"<c r="{coord}""#);
        if let Some(pos) = find_bytes(&self.sheet_xml, tag.as_bytes()) {
            if let Some(spos) = find_bytes_from(&self.sheet_xml, b" s=\"", pos) {
                let val_start = spos + 4;
                let val_end =
                    find_bytes_from(&self.sheet_xml, b"\"", val_start + 1).unwrap_or(val_start);
                let id = std::str::from_utf8(&self.sheet_xml[val_start..val_end])?
                    .parse::<u32>()
                    .unwrap_or(0);
                return Ok(Some(id));
            }
        }
        Ok(None)
    }
}

#[derive(Clone)]
pub struct AlignSpec {
    pub horiz: Option<String>, // "center" | "left" | …
    pub vert: Option<String>,  // "center" | "top" | …
    pub wrap: bool,
}

// ──────────────────────────────────────────────────────────────────────
// 2. РАЗБОР ДИАПАЗОНА
// ──────────────────────────────────────────────────────────────────────
#[derive(Debug)]
enum Target {
    Cell(String),
    Rect { c0: u32, r0: u32, c1: u32, r1: u32 },
    Col(u32),
    Row(u32),
}
fn parse_target(s: &str) -> Result<Target> {
    let re_cell = Regex::new(r"^([A-Za-z]+)([0-9]+)$").unwrap();
    let re_rect = Regex::new(r"^([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)$").unwrap();
    let re_col = Regex::new(r"^([A-Za-z]+):$").unwrap();
    let re_row = Regex::new(r"^([0-9]+):$").unwrap();

    if let Some(_caps) = re_cell.captures(s) {
        return Ok(Target::Cell(s.to_owned()));
    }
    if let Some(caps) = re_rect.captures(s) {
        let (c0, r0) = split_coord(&caps[1]);
        let (c1, r1) = split_coord(&caps[2]);
        return Ok(Target::Rect { c0, r0, c1, r1 });
    }
    if let Some(caps) = re_col.captures(s) {
        return Ok(Target::Col(col_index(&caps[1]) as u32));
    }
    if let Some(caps) = re_row.captures(s) {
        return Ok(Target::Row(caps[1].parse::<u32>()?));
    }
    Err(XlsxError::InvalidRange { 
        range: s.to_string() 
    })
}
// ──────────────────────────────────────────────────────────────────────
// 4. ВСПОМОГАТЕЛЬНЫЕ  (буквы ↔ индекс, split_coord, splice‑утилиты)
// ──────────────────────────────────────────────────────────────────────
fn col_letter(mut n: u32) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}
fn col_index(s: &str) -> usize {
    s.bytes().fold(0, |acc, b| {
        acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
    }) - 1
}
fn split_coord(coord: &str) -> (u32, u32) {
    let p = coord.find(|c: char| c.is_ascii_digit()).unwrap();
    (
        col_index(&coord[..p]) as u32,
        coord[p..].parse::<u32>().unwrap(),
    )
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
    Err(XlsxError::AttributeNotFound { 
        attribute: "count".to_string() 
    })
}

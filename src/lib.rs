//! lib.rs  — ядро xlsx-append-rs
//! ✅  Minimal append: только inline-строки и числа, один лист.

use std::{
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
};

use anyhow::{bail, Context, Result};
use quick_xml::{events::Event, Reader, Writer};
use tempfile::NamedTempFile;
use zip::{read::ZipArchive, write::FileOptions, CompressionMethod, ZipWriter};

/// API-объект для дозаписи.
pub struct XlsxAppender {
    src_path: PathBuf,
    sheet_xml: Vec<u8>,      // исходный sheet1.xml
    last_row: u32,           // номер последней строки (r="")
}

impl XlsxAppender {
    /// Открыть книгу, считать sheet1.xml и запомнить номер последней строки.
    pub fn open<P: AsRef<Path>>(src: P) -> Result<Self> {
        let src_path = src.as_ref().to_path_buf();
        let mut zip = ZipArchive::new(File::open(&src_path)?)?;
        let mut sheet = zip
            .by_name("xl/worksheets/sheet1.xml")
            .context("sheet1.xml not found")?;
        let mut sheet_xml = Vec::with_capacity(sheet.size() as usize);
        sheet.read_to_end(&mut sheet_xml)?;

        // Быстро находим последний <row r="N">
        let mut reader = Reader::from_reader(sheet_xml.as_slice());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut last_row = 0;

        while let Ok(ev) = reader.read_event_into(&mut buf) {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"row" => {
                    if let Some(r) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r").then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        last_row = r.parse::<u32>().unwrap_or(last_row);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(Self { src_path, sheet_xml, last_row })
    }

    /// Добавить одну строку значений (строки/числа).
    pub fn append_row<I, S>(&mut self, cells: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        self.last_row += 1;
        let row_num = self.last_row;
        let mut writer = Writer::new(Vec::new());

        // <row r="N">
        writer
            .create_element("row")
            .with_attribute(("r", row_num.to_string().as_str()))
            .write_inner_content(|w| {
                let mut col = b'A';
                for val in cells {
                    let coord = format!("{}{}", col as char, row_num);
                    let val_str = val.to_string();
                    let is_number = val_str.parse::<f64>().is_ok();

                    {
                        let mut c_elem = w.create_element("c").with_attribute(("r", coord.as_str()));
                        if !is_number {
                            c_elem = c_elem.with_attribute(("t", "inlineStr"));
                        }
                        c_elem.write_inner_content(|w2| {
                            use quick_xml::events::BytesText;
                            if !is_number {
                                w2.create_element("is")
                                    .write_inner_content(|w3| {
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

        // Вставляем перед </sheetData>
        if let Some(pos) = self.sheet_xml.windows(12).rposition(|w| w == b"</sheetData>") {
            self.sheet_xml.splice(pos..pos, new_row_xml);
            Ok(())
        } else {
            bail!("</sheetData> tag not found");
        }
    }

    /// Сохранить в новый файл (не трогаем исходник).
    pub fn save<P: AsRef<Path>>(&self, dst: P) -> Result<()> {
        // 1) распаковываем исходный zip и создаём новый во временный файл
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
                if name == "xl/worksheets/sheet1.xml" {
                    zout.start_file::<_, ()>(name, opt)?;
                    zout.write_all(&self.sheet_xml)?;
                } else {
                    zout.start_file::<_, ()>(name, opt)?;
                    std::io::copy(&mut f, &mut zout)?;
                }
            }
            zout.finish()?;
        }
        // 2) переименовываем временный файл в целевой
        fs::rename(tmp.path(), dst)?;
        Ok(())
    }
}



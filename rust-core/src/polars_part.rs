#[cfg(feature = "polars")]
use crate::style::{col_letter, split_coord};
use crate::XlsxEditor;
#[cfg(feature = "polars")]
use anyhow::{Result, bail};
#[cfg(feature = "polars")]
use polars_core::prelude::*;
#[cfg(feature = "polars")]
use quick_xml::Writer;
#[cfg(feature = "polars")]
use quick_xml::events::BytesText;

impl XlsxEditor {
    #[cfg(feature = "polars")]
    pub fn with_polars(&mut self, df: &DataFrame, start_cell: Option<&str>) -> Result<()> {
        // ---------- 0.  Координаты диапазона, который будем затирать ----------
        let start_coord = start_cell.unwrap_or("A1");
        let (base_col, first_row) = {
            let split = start_coord.find(|c: char| c.is_ascii_digit()).unwrap();
            (
                split_coord(&start_coord[..]),
                start_coord[split..].parse::<u32>().unwrap(),
            )
        };
        let last_row = first_row + df.height() as u32 - 1;

        // ---------- 0‑bis.  Удаляем старые <row …> в нужном диапазоне ----------
        // ищем шаблон <row r="N" ...> … </row>
        // NB: линейный поиск по Vec<u8> дешевле, чем парсить всё XML quick‑xml‑ом
        let mut i = 0;
        while let Some(beg) = self.sheet_xml[i..]
            .windows(5)
            .position(|w| w == b"<row ")
            .map(|p| p + i)
        {
            // поиск атрибут r="..."
            if let Some(r_attr_pos) = self.sheet_xml[beg..]
                .windows(3)
                .position(|w| w == b"r=\"")
                .map(|p| p + beg + 3)
            {
                if let Some(quote_end) = self.sheet_xml[r_attr_pos..]
                    .iter()
                    .position(|&b| b == b'"')
                    .map(|p| p + r_attr_pos)
                {
                    let row_num: u32 = std::str::from_utf8(&self.sheet_xml[r_attr_pos..quote_end])?
                        .parse()
                        .unwrap_or(0);

                    if row_num >= first_row && row_num <= last_row {
                        // найти конец </row>
                        if let Some(end_tag) = self.sheet_xml[quote_end..]
                            .windows(6)
                            .position(|w| w == b"</row>")
                            .map(|p| p + quote_end + 6)
                        {
                            self.sheet_xml.splice(beg..end_tag, std::iter::empty());
                            // начинаем заново — буфер изменился
                            i = 0;
                            continue;
                        }
                    }
                }
            }
            // если не удалили, двигаемся далее
            i = beg + 5;
        }

        // ---------- 1.  Метаданные столбцов ----------
        struct ColMeta {
            is_number: bool,
            style_id: Option<u32>,
            conv: Box<dyn Fn(AnyValue) -> String>,
        }

        let mut cols = Vec::<ColMeta>::with_capacity(df.width());
        for s in df.get_columns() {
            match s.dtype() {
                // Текст — берём строку без кавычек.
                DataType::String => cols.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| match v {
                        AnyValue::String(s) => s.to_string(),
                        _ => {
                            // на всякий случай, вдруг что‑то ещё пролезет
                            let mut t = v.to_string();
                            if t.starts_with('"') && t.ends_with('"') {
                                t.truncate(t.len() - 1);
                                t.remove(0);
                            }
                            t
                        }
                    }),
                }),
                // Целые
                DataType::Int8
                | DataType::Int16
                | DataType::Int32
                | DataType::Int64
                | DataType::UInt8
                | DataType::UInt16
                | DataType::UInt32
                | DataType::UInt64 => cols.push(ColMeta {
                    is_number: true,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
                // Float
                DataType::Float32 | DataType::Float64 => cols.push(ColMeta {
                    is_number: true,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
                // Fallback
                _ => cols.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
            }
        }

        // ---------- 2.  Генерируем XML строк ----------
        let mut bulk_rows_xml = Vec::<u8>::new();
        let mut cur_row = first_row;
        for idx in 0..df.height() {
            let mut w = Writer::new(Vec::new());
            w.create_element("row")
                .with_attribute(("r", cur_row.to_string().as_str()))
                .write_inner_content(|wr| {
                    for (col_idx, s) in df.get_columns().iter().enumerate() {
                        let coord =
                            format!("{}{}", col_letter(base_col.0 + col_idx as u32), cur_row);
                        let val = s.get(idx).unwrap_or(AnyValue::Null);
                        let meta = &cols[col_idx];

                        let mut c = wr.create_element("c").with_attribute(("r", coord.as_str()));
                        if let Some(sid) = meta.style_id {
                            c = c.with_attribute(("s", sid.to_string().as_str()));
                        }
                        if !meta.is_number {
                            c = c.with_attribute(("t", "inlineStr"));
                        }

                        c.write_inner_content(|w2| {
                            let txt = (meta.conv)(val);
                            if meta.is_number {
                                w2.create_element("v")
                                    .write_text_content(BytesText::new(&txt))?;
                            } else {
                                w2.create_element("is").write_inner_content(|w3| {
                                    w3.create_element("t")
                                        .write_text_content(BytesText::new(&txt))?;
                                    Ok(())
                                })?;
                            }
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            bulk_rows_xml.extend_from_slice(&w.into_inner());
            cur_row += 1;
        }

        // ---------- 3.  Вставляем новые строки ----------
        if let Some(pos) = self
            .sheet_xml
            .windows(12)
            .rposition(|w| w == b"</sheetData>")
        {
            self.sheet_xml.splice(pos..pos, bulk_rows_xml);
            self.last_row = last_row;
            Ok(())
        } else {
            bail!("</sheetData> tag not found");
        }
    }
}

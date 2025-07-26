use crate::XlsxEditor;
#[cfg(feature = "polars")]
use crate::style::{col_letter, split_coord};
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
        // ---------- 0.  Координаты ----------
        let start_coord = start_cell.unwrap_or("A1");
        let (base_col, first_row) = {
            let split = start_coord.find(|c: char| c.is_ascii_digit()).unwrap();
            (
                split_coord(&start_coord[..]),
                start_coord[split..].parse::<u32>().unwrap(),
            )
        };

        // +1 строка на заголовок
        let last_row = first_row + df.height() as u32; // header + N строк данных

        // ---------- 0‑bis.  Сносим старые строки в диапазоне ----------
        let mut i = 0;
        while let Some(beg_rel) = self.sheet_xml[i..].windows(4).position(|w| w == b"<row") {
            let beg = i + beg_rel;

            // следующий символ после "<row" должен быть пробел или '>'
            let after = beg + 4;
            if after >= self.sheet_xml.len() {
                break;
            }
            let next = self.sheet_xml[after];
            if next != b' ' && next != b'>' {
                i = after;
                continue;
            }

            // конец открывающего тега <row ...>
            let Some(open_end_rel) = self.sheet_xml[after..].iter().position(|&b| b == b'>') else {
                break;
            };
            let open_end = after + open_end_rel + 1; // позиция сразу после '>'

            // конец всего блока </row>
            let Some(close_rel) = self.sheet_xml[open_end..]
                .windows(6)
                .position(|w| w == b"</row>")
            else {
                break;
            };
            let row_end = open_end + close_rel + 6; // позиция сразу после "</row>"

            // 1) пробуем достать r="N" только из открывающего тега <row ...>
            let mut row_num_opt = None;
            if let Some(r_pos_rel) = self.sheet_xml[beg..open_end]
                .windows(3)
                .position(|w| w == b"r=\"")
            {
                let r_pos = beg + r_pos_rel + 3;
                if let Some(q_end_rel) = self.sheet_xml[r_pos..open_end]
                    .iter()
                    .position(|&b| b == b'"')
                {
                    let q_end = r_pos + q_end_rel;
                    row_num_opt = std::str::from_utf8(&self.sheet_xml[r_pos..q_end])
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok());
                }
            }

            // 2) fallback: берем номер строки по первой ячейке внутри этого <row>
            if row_num_opt.is_none() {
                if let Some(c_r_rel) = self.sheet_xml[open_end..row_end]
                    .windows(3)
                    .position(|w| w == b"r=\"")
                {
                    let r_pos = open_end + c_r_rel + 3;
                    if let Some(q_end_rel) = self.sheet_xml[r_pos..row_end]
                        .iter()
                        .position(|&b| b == b'"')
                    {
                        let q_end = r_pos + q_end_rel;
                        let s = &self.sheet_xml[r_pos..q_end]; // типа b"A123"
                        // ищем начало хвоста с цифрами
                        let digits_start = s
                            .iter()
                            .rposition(|&b| !(b as char).is_ascii_digit())
                            .map(|p| p + 1)
                            .unwrap_or(0);
                        row_num_opt = std::str::from_utf8(&s[digits_start..])
                            .ok()
                            .and_then(|s| s.parse::<u32>().ok());
                    }
                }
            }

            if let Some(row_num) = row_num_opt {
                if row_num >= first_row && row_num <= last_row {
                    // вырезаем весь <row>...</row>, чтобы точно не было дублей
                    self.sheet_xml.splice(beg..row_end, std::iter::empty());
                    i = 0; // начинаем поиск заново с начала буфера
                    continue;
                }
            }

            // если не наш диапазон — перепрыгиваем за этот <row>
            i = row_end;
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
                DataType::String => cols.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| match v {
                        AnyValue::String(s) => s.to_string(),
                        _ => {
                            let mut t = v.to_string();
                            if t.starts_with('"') && t.ends_with('"') {
                                t.truncate(t.len() - 1);
                                t.remove(0);
                            }
                            t
                        }
                    }),
                }),
                DataType::Int8
                | DataType::Int16
                | DataType::Int32
                | DataType::Int64
                | DataType::UInt8
                | DataType::UInt16
                | DataType::UInt32
                | DataType::UInt64
                | DataType::Float32
                | DataType::Float64 => cols.push(ColMeta {
                    is_number: true,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
                _ => cols.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
            }
        }

        // ---------- 2.  Генерим XML: сначала заголовок, потом данные ----------
        let mut bulk_rows_xml = Vec::<u8>::new();

        // 2.1 Хедер
        let mut cur_row = first_row;
        {
            let mut w = Writer::new(Vec::new());
            w.create_element("row")
                .with_attribute(("r", cur_row.to_string().as_str()))
                .write_inner_content(|wr| {
                    for (col_idx, s) in df.get_columns().iter().enumerate() {
                        let coord =
                            format!("{}{}", col_letter(base_col.0 + col_idx as u32), cur_row);
                        let c = wr
                            .create_element("c")
                            .with_attribute(("r", coord.as_str()))
                            .with_attribute(("t", "inlineStr")); // всегда текст

                        c.write_inner_content(|w2| {
                            w2.create_element("is").write_inner_content(|w3| {
                                w3.create_element("t")
                                    .write_text_content(BytesText::new(s.name()))?;
                                Ok(())
                            })?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            bulk_rows_xml.extend_from_slice(&w.into_inner());
            cur_row += 1;
        }

        // 2.2 Данные
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

                        enum Kind {
                            Blank,
                            Num(String),
                            Str(String),
                        }
                        let kind = match val {
                            AnyValue::Null => Kind::Blank,
                            AnyValue::Float64(x) => {
                                if x.is_finite() {
                                    Kind::Num(x.to_string())
                                } else {
                                    Kind::Blank
                                }
                            }
                            AnyValue::Float32(x) => {
                                if x.is_finite() {
                                    Kind::Num(x.to_string())
                                } else {
                                    Kind::Blank
                                }
                            }
                            _ => {
                                if meta.is_number {
                                    Kind::Num(val.to_string())
                                } else {
                                    Kind::Str((meta.conv)(val))
                                }
                            }
                        };

                        let is_text = matches!(kind, Kind::Str(_));
                        let mut c = wr.create_element("c").with_attribute(("r", coord.as_str()));
                        if let Some(sid) = meta.style_id {
                            c = c.with_attribute(("s", sid.to_string().as_str()));
                        }
                        if is_text {
                            c = c.with_attribute(("t", "inlineStr"));
                        }

                        c.write_inner_content(|w2| {
                            match kind {
                                Kind::Blank => { /* пустая ячейка — ничего не пишем */
                                }
                                Kind::Num(txt) => {
                                    w2.create_element("v")
                                        .write_text_content(BytesText::new(&txt))?;
                                }
                                Kind::Str(txt) => {
                                    w2.create_element("is").write_inner_content(|w3| {
                                        w3.create_element("t")
                                            .write_text_content(BytesText::new(&txt))?;
                                        Ok(())
                                    })?;
                                }
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
        // 3. Вставляем новые строки в правильное место (сортировка по r)
        let sd_open = if let Some(p) = self.sheet_xml.windows(11).position(|w| w == b"<sheetData>")
        {
            p + 11
        } else {
            bail!("<sheetData> tag not found");
        };

        // по умолчанию — перед </sheetData>
        let mut insert_pos = self
            .sheet_xml
            .windows(12)
            .rposition(|w| w == b"</sheetData>")
            .ok_or_else(|| anyhow::anyhow!("</sheetData> tag not found"))?;

        // ищем первую <row> с r >= first_row
        let mut j = sd_open;
        while let Some(beg_rel) = self.sheet_xml[j..].windows(4).position(|w| w == b"<row") {
            let beg = j + beg_rel;
            let after = beg + 4;
            if after >= self.sheet_xml.len() {
                break;
            }
            let next = self.sheet_xml[after];
            if next != b' ' && next != b'>' {
                j = after;
                continue;
            }

            let Some(open_end_rel) = self.sheet_xml[after..].iter().position(|&b| b == b'>') else {
                break;
            };
            let open_end = after + open_end_rel + 1;

            let Some(close_rel) = self.sheet_xml[open_end..]
                .windows(6)
                .position(|w| w == b"</row>")
            else {
                break;
            };
            let row_end = open_end + close_rel + 6;

            let mut row_num_opt = None;
            if let Some(r_pos_rel) = self.sheet_xml[beg..open_end]
                .windows(3)
                .position(|w| w == b"r=\"")
            {
                let r_pos = beg + r_pos_rel + 3;
                if let Some(q_end_rel) = self.sheet_xml[r_pos..open_end]
                    .iter()
                    .position(|&b| b == b'"')
                {
                    let q_end = r_pos + q_end_rel;
                    row_num_opt = std::str::from_utf8(&self.sheet_xml[r_pos..q_end])
                        .ok()
                        .and_then(|s| s.parse::<u32>().ok());
                }
            }

            if let Some(n) = row_num_opt {
                if n >= first_row {
                    insert_pos = beg;
                    break;
                }
            }

            j = row_end;
        }

        self.sheet_xml.splice(insert_pos..insert_pos, bulk_rows_xml);
        self.last_row = last_row;
        if let Some(dim_beg) = self
            .sheet_xml
            .windows(16)
            .position(|w| w == b"<dimension ref=\"")
        {
            let start = dim_beg + 16;
            if let Some(q_end_rel) = self.sheet_xml[start..].iter().position(|&b| b == b'"') {
                let end = start + q_end_rel;
                let last_col = col_letter(base_col.0 + (df.width().saturating_sub(1) as u32));
                let dim = format!(
                    "{}{}:{}{}",
                    col_letter(base_col.0),
                    first_row,
                    last_col,
                    last_row
                );
                self.sheet_xml.splice(start..end, dim.into_bytes());
            }
        }

        Ok(())
    }
}

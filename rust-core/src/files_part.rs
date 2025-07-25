use crate::{XlsxEditor, scan};
use ::zip as zip_crate;
use anyhow::{Context, Result, bail};
use quick_xml::{Reader, events::Event};
use std::{
    fs::File,
    io::{Read, Write},
    path::Path,
};

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
                .with_context(|| format!("{sheet_path} not found"))?;
            let mut buf = Vec::with_capacity(sheet.size() as usize);
            sheet.read_to_end(&mut buf)?;
            buf
        };

        // ── styles.xml ───────────────────────────────────────────────
        let styles_xml: Vec<u8> = {
            let mut styles = zip
                .by_name("xl/styles.xml")
                .context("styles.xml not found")?;
            let mut buf = Vec::with_capacity(styles.size() as usize);
            styles.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml ───────────────────────────────────────────────
        let workbook_xml: Vec<u8> = {
            let mut wb = zip
                .by_name("xl/workbook.xml")
                .context("xl/workbook.xml not found")?;
            let mut buf = Vec::with_capacity(wb.size() as usize);
            wb.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml.rels ──────────────────────────────────────────
        let rels_xml: Vec<u8> = {
            let mut rels = zip
                .by_name("xl/_rels/workbook.xml.rels")
                .context("xl/_rels/workbook.xml.rels not found")?;
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

    fn flush_current_sheet(&mut self) {
        let cur_path = self.sheet_path.clone();
        let cur_xml  = self.sheet_xml.clone();
        if let Some((_, c)) = self.new_files.iter_mut().find(|(p, _)| p == &cur_path) {
            *c = cur_xml;
        } else {
            self.new_files.push((cur_path, cur_xml));
        }
    }
    
    pub fn save<P: AsRef<Path>>(&mut self, dst: P) -> Result<()> {
        self.flush_current_sheet();
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
        for (p, _) in &self.new_files {
            eprintln!("new_files has {}", p);
        }

        zout.finish()?;
        Ok(())
    }
}

impl XlsxEditor {
    /// Считает количество листов по текущему состоянию `workbook_xml`
    fn sheet_count(&self) -> usize {
        let mut rdr = Reader::from_reader(self.workbook_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut n = 0usize;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    n += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        n
    }

    /// Возвращает (позиция_начала_контента, позиция_конца_контента) для содержимого между
    /// `<sheets ...>` и `</sheets>` в `workbook_xml`.
    fn find_sheets_section(workbook_xml: &[u8]) -> Result<(usize, usize)> {
        let xml = workbook_xml;
        let open_tag =
            memchr::memmem::find(xml, b"<sheets").context("<sheets> not found in workbook.xml")?;
        let mut pos = open_tag;
        // ищем '>' у открывающего <sheets ...>
        while pos < xml.len() && xml[pos] != b'>' {
            pos += 1;
        }
        if pos >= xml.len() {
            bail!("Malformed workbook.xml: <sheets ...> not closed with '>'");
        }
        let content_start = pos + 1;

        let close_tag = memchr::memmem::rfind(xml, b"</sheets>")
            .context("</sheets> not found in workbook.xml")?;
        Ok((content_start, close_tag))
    }

    /// Добавляет новый пустой лист c именем `sheet_name` **на позицию `index` (0‑based)**,
    /// пересобирая порядок `<sheet/>` в workbook.xml.
    pub fn add_worksheet_at(&mut self, sheet_name: &str, mut index: usize) -> Result<&mut Self> {
        // -------- 0) валидации / подготовка ----------
        // 0.1) имя уже существует?
        let sheet_names = scan(&self.src_path)?;
        if sheet_names.contains(&sheet_name.to_owned()) {
            bail!("Sheet {} already exists", sheet_name);
        }

        // 0.2) текущее количество листов
        let cur_cnt = self.sheet_count();
        if index > cur_cnt {
            index = cur_cnt; // кладём в конец
        }

        // 0.3) читаем исходный архив (для поиска свободного sheet#.xml)
        let mut zin = zip::ZipArchive::new(File::open(&self.src_path)?)?;

        // 0.4) локальные (редактируемые) копии XML
        let mut wb_xml = self.workbook_xml.clone();
        let mut rels_xml = self.rels_xml.clone();

        // -------- 1) найдём max sheetId и max rId ----------
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
        // Новый sheetId нам не особо важен (мы потом все перенумеруем), но пусть будет > max_sheet_id
        let _new_sheet_id = max_sheet_id + 1;

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

        // -------- 2) найти свободный sheet#.xml ----------
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
        for (path, _) in &self.new_files {
            if let Some(n) = path
                .strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<usize>().ok())
            {
                max_sheet_file = max_sheet_file.max(n);
            }
        }
        let new_sheet_file = max_sheet_file + 1;
        let new_sheet_path = format!("xl/worksheets/sheet{new}.xml", new = new_sheet_file);
        let new_sheet_target = format!("worksheets/sheet{new}.xml", new = new_sheet_file);

        // -------- 3) распарсим текущие <sheet .../> из workbook.xml ----------
        #[derive(Debug, Clone)]
        struct SheetTag {
            name: String,
            rid: String,  // "rIdNN"
            path: String, // worksheets/sheet#.xml (нам нужно только для инфы; можно не хранить)
        }
        let (sheets_content_start, sheets_content_end) = Self::find_sheets_section(&wb_xml)?;
        let sheets_slice = &wb_xml[sheets_content_start..sheets_content_end];

        let mut rdr = Reader::from_reader(sheets_slice);
        rdr.config_mut().trim_text(true);
        let mut sheets: Vec<SheetTag> = Vec::new();

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    let mut name = None;
                    let mut rid = None;
                    // Target пути тут нет — он в rels, так что просто пустим.
                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"name" {
                            name = Some(v.clone());
                        }
                        if k == b"r:id" {
                            rid = Some(v);
                        }
                    }
                    sheets.push(SheetTag {
                        name: name.unwrap_or_default(),
                        rid: rid.unwrap_or_default(),
                        path: String::new(),
                    });
                }
                Event::Eof => break,
                _ => {}
            }
        }

        // -------- 4) формируем новый tag для нового листа ----------
        let new_sheet = SheetTag {
            name: sheet_name.to_string(),
            rid: format!("rId{}", new_rid),
            path: new_sheet_target.clone(),
        };

        // вставляем по индексу
        if index >= sheets.len() {
            sheets.push(new_sheet);
        } else {
            sheets.insert(index, new_sheet);
        }

        // -------- 5) перегенерируем <sheets>...</sheets> с новой нумерацией sheetId ----------
        let mut new_inner = Vec::new();
        // Сохраним форматирование: перенос строки + два пробела
        for (i, sh) in sheets.iter().enumerate() {
            let sheet_id = (i as u32) + 1; // «естественная» нумерация
            let line = format!(
                r#"\n  <sheet name="{}" sheetId="{}" r:id="{}"/>"#,
                xml_escape(&sh.name),
                sheet_id,
                sh.rid
            );
            new_inner.extend_from_slice(line.as_bytes());
        }

        // подменяем содержимое между <sheets ...> и </sheets>
        wb_xml.splice(
            sheets_content_start..sheets_content_end,
            new_inner.into_iter(),
        );

        // -------- 6) вставляем Relationship под конец </Relationships> ----------
        let rel_tag = format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{}"/>"#,
            new_rid, new_sheet_target
        );
        if let Some(pos) = rels_xml.windows(16).rposition(|w| w == b"</Relationships>") {
            rels_xml.splice(pos..pos, rel_tag.as_bytes().iter().copied());
        } else {
            bail!("</Relationships> not found in workbook.xml.rels");
        }

        // -------- 7) минимальный XML нового листа ----------
        const EMPTY_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData> </sheetData>
        </worksheet>"#;

        // Обновляем внутреннее состояние
        self.workbook_xml = wb_xml;
        self.rels_xml = rels_xml;

        // кладём текущий редактируемый лист в new_files (если ещё не лежит)
        {
            let cur_path = self.sheet_path.clone();
            let cur_xml = self.sheet_xml.clone();
            if let Some(pair) = self.new_files.iter_mut().find(|(p, _)| p == &cur_path) {
                pair.1 = cur_xml;
            } else {
                self.new_files.push((cur_path, cur_xml));
            }
        }

        // создаём запись для нового листа
        if let Some(pair) = self
            .new_files
            .iter_mut()
            .find(|(p, _)| p == &new_sheet_path)
        {
            pair.1 = EMPTY_SHEET.as_bytes().to_vec();
        } else {
            self.new_files
                .push((new_sheet_path.clone(), EMPTY_SHEET.as_bytes().to_vec()));
        }

        // переключаем редактор на новый лист
        self.sheet_path = new_sheet_path;
        self.sheet_xml = EMPTY_SHEET.as_bytes().to_vec();
        self.last_row = 0;

        Ok(self)
    }

    /// Старый API: просто добавляет в конец.
    pub fn add_worksheet(&mut self, sheet_name: &str) -> Result<&mut Self> {
        let last_idx = self.sheet_count(); // вставка в конец
        self.add_worksheet_at(sheet_name, last_idx)
    }
}

impl XlsxEditor {
    pub fn with_worksheet(&mut self, sheet_name: &str) -> Result<&mut Self> {
        // 0) Если уже на этом листе — просто вернуть себя (опционально).
        // У нас нет текущего имени, так что пропустим эту оптимизацию.

        // 1) Сохраним текущий лист в new_files (как в add_worksheet_at)
        {
            let cur_path = self.sheet_path.clone();
            let cur_xml = self.sheet_xml.clone();
            if !cur_path.is_empty() {
                if let Some(pair) = self.new_files.iter_mut().find(|(p, _)| p == &cur_path) {
                    pair.1 = cur_xml;
                } else {
                    self.new_files.push((cur_path, cur_xml));
                }
            }
        }

        // 2) Найти r:id по имени листа в workbook.xml
        let mut rdr = Reader::from_reader(self.workbook_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut target_rid: Option<String> = None;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    let mut name: Option<String> = None;
                    let mut rid: Option<String> = None;

                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"name" {
                            name = Some(v.clone());
                        }
                        if k == b"r:id" {
                            rid = Some(v);
                        }
                    }

                    if let (Some(n), Some(r)) = (name, rid) {
                        if n == sheet_name {
                            target_rid = Some(r);
                            break;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let target_rid = target_rid
            .with_context(|| format!("Sheet `{}` not found in workbook.xml", sheet_name))?;

        // 3) По r:id найти Target в workbook.xml.rels
        let mut rdr = Reader::from_reader(self.rels_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut target_rel: Option<String> = None;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e)
                    if e.name().as_ref() == b"Relationship" =>
                {
                    let mut id: Option<String> = None;
                    let mut target: Option<String> = None;

                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"Id" {
                            id = Some(v.clone());
                        }
                        if k == b"Target" {
                            target = Some(v);
                        }
                    }

                    if let (Some(idv), Some(t)) = (id, target) {
                        if idv == target_rid {
                            target_rel = Some(t);
                            break;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let target_rel = target_rel.with_context(|| {
            format!(
                "Relationship for `{}` not found in workbook.xml.rels",
                sheet_name
            )
        })?;

        // Собираем абсолютный путь внутри архива
        let new_sheet_path = if target_rel.starts_with("xl/") {
            target_rel.clone()
        } else {
            format!("xl/{}", target_rel)
        };

        // 4) Достаём XML листа: сперва смотрим в new_files, иначе читаем из ZIP
        let sheet_xml: Vec<u8> =
            if let Some((_, content)) = self.new_files.iter().find(|(p, _)| p == &new_sheet_path) {
                content.clone()
            } else {
                let mut zin = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;
                let mut f = zin
                    .by_name(&new_sheet_path)
                    .with_context(|| format!("{} not found in zip", new_sheet_path))?;
                let mut buf = Vec::with_capacity(f.size() as usize);
                f.read_to_end(&mut buf)?;
                buf
            };

        // 5) Пересчитываем last_row
        let last_row = calc_last_row(&sheet_xml);

        // 6) Переключаемся
        self.sheet_path = new_sheet_path;
        self.sheet_xml = sheet_xml;
        self.last_row = last_row;

        Ok(self)
    }
}

// маленький хелпер
fn calc_last_row(sheet_xml: &[u8]) -> u32 {
    let mut rdr = Reader::from_reader(sheet_xml);
    rdr.config_mut().trim_text(true);

    let mut last_row = 0u32;
    while let Ok(ev) = rdr.read_event() {
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
    }
    last_row
}

// Простейший экранировщик для XML-атрибутов.
fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

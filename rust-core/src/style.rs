//! style.rs – универсальный слой стилей + нормализация <cols>

use anyhow::{Context, Result, bail};
use quick_xml::{Reader, events::Event};
use regex::Regex;
use std::collections::{BTreeMap};
use std::{fmt, str::FromStr};

use crate::XlsxEditor;


// #[derive(Hash, Eq, PartialEq, Clone, Debug)]
// struct FontKey {
//     name: String,
//     size: u32, // храним как целое *100 (или округлённое), чтобы Hash работал стабильно
//     bold: bool,
//     italic: bool,
// }

/* ========================== ALIGNMENT API ================================= */

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HorizAlignment {
    Left,
    Center,
    Right,
    Fill,
    Justify,
}
impl fmt::Display for HorizAlignment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            HorizAlignment::Left => "left",
            HorizAlignment::Center => "center",
            HorizAlignment::Right => "right",
            HorizAlignment::Fill => "fill",
            HorizAlignment::Justify => "justify",
        })
    }
}
impl FromStr for HorizAlignment {
    type Err = anyhow::Error;
    fn from_str(s: &str) -> Result<Self> {
        Ok(match s {
            "left" => HorizAlignment::Left,
            "center" => HorizAlignment::Center,
            "right" => HorizAlignment::Right,
            "fill" => HorizAlignment::Fill,
            "justify" => HorizAlignment::Justify,
            _ => bail!("Unknown horizontal alignment: {s}"),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VertAlignment {
    Top,
    Center,
    Bottom,
    Justify,
}
impl fmt::Display for VertAlignment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            VertAlignment::Top => "top",
            VertAlignment::Center => "center",
            VertAlignment::Bottom => "bottom",
            VertAlignment::Justify => "justify",
        })
    }
}
impl FromStr for VertAlignment {
    type Err = anyhow::Error;
    fn from_str(s: &str) -> Result<Self> {
        Ok(match s {
            "top" => VertAlignment::Top,
            "center" => VertAlignment::Center,
            "bottom" => VertAlignment::Bottom,
            "justify" => VertAlignment::Justify,
            _ => bail!("Unknown vertical alignment: {s}"),
        })
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AlignSpec {
    pub horiz: Option<HorizAlignment>,
    pub vert: Option<VertAlignment>,
    pub wrap: bool,
}

/* ========================== CORE STYLE STRUCT ============================= */

#[derive(Debug, Clone, Default)]
struct StyleParts {
    pub num_fmt_code: Option<String>,
    pub font: Option<u32>,
    pub fill: Option<u32>,
    pub border: Option<u32>,
    pub align: Option<AlignSpec>,
}

/* ========================== TARGET PARSER ================================= */

#[derive(Debug)]
enum Target {
    Cell(String),
    Rect { c0: u32, r0: u32, c1: u32, r1: u32 },
    Col(u32), // 0-based
    Row(u32),
}

fn parse_target(s: &str) -> Result<Target> {
    let re_cell = Regex::new(r"^([A-Za-z]+)([0-9]+)$").unwrap();
    let re_rect = Regex::new(r"^([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)$").unwrap();
    let re_col = Regex::new(r"^([A-Za-z]+):$").unwrap();
    let re_row = Regex::new(r"^([0-9]+):$").unwrap();

    if re_cell.is_match(s) {
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
    bail!("invalid range syntax: {s}");
}

/* ========================== PUBLIC API ==================================== */

impl XlsxEditor {
    pub fn set_border(&mut self, range: &str, border_style: &str) -> Result<&mut Self> {
        let border_id = self.ensure_border(border_style)?;
        self.apply_patch(
            range,
            StyleParts {
                border: Some(border_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_font(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
    ) -> Result<&mut Self> {
        let font_id = self.ensure_font(name, size, bold, italic)?;
        self.apply_patch(
            range,
            StyleParts {
                font: Some(font_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_font_with_alignment(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
        align: &AlignSpec,
    ) -> Result<&mut Self> {
        let font_id = self.ensure_font(name, size, bold, italic)?;
        self.apply_patch(
            range,
            StyleParts {
                font: Some(font_id),
                align: Some(align.clone()),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_fill(&mut self, range: &str, rgb: &str) -> Result<&mut Self> {
        let fill_id = self.ensure_fill(rgb)?;
        self.apply_patch(
            range,
            StyleParts {
                fill: Some(fill_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_alignment(&mut self, range: &str, align: &AlignSpec) -> Result<&mut Self> {
        self.apply_patch(
            range,
            StyleParts {
                align: Some(align.clone()),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    /// Публичный API для числового формата.
    pub fn set_number_format(&mut self, range: &str, fmt: &str) -> Result<()> {
        let style_id = self.ensure_style(Some(fmt), None, None, None, None)?;
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
            Target::Col(c0) => self.force_column_number_format(c0, style_id)?,
            Target::Row(_row) => bail!("Row-level not implemented yet"),
        }
        Ok(())
    }

    pub fn set_column_width(&mut self, col_letter: &str, width: f64) -> Result<&mut Self> {
        let col0 = col_index(col_letter) as u32; // 0-based
        self.set_column_properties(col0, Some(width), None)?;
        Ok(self)
    }
}

/* ========================== CORE PATCH ENGINE ============================= */

impl XlsxEditor {
    fn apply_patch(&mut self, range: &str, patch: StyleParts) -> Result<()> {
        match parse_target(range)? {
            Target::Cell(cell) => {
                self.patch_one_cell(&cell, &patch)?;
            }
            Target::Rect { c0, r0, c1, r1 } => {
                for r in r0..=r1 {
                    for c in c0..=c1 {
                        let cell = format!("{}{}", col_letter(c), r);
                        self.patch_one_cell(&cell, &patch)?;
                    }
                }
            }
            Target::Row(_r) => bail!("Row-level styling is not implemented in this snippet"),
            Target::Col(_c) => bail!("Column-level styling is not implemented in this snippet"),
        }
        Ok(())
    }

    fn patch_one_cell(&mut self, coord: &str, patch: &StyleParts) -> Result<()> {
        let sid_opt = self.cell_style_id(coord)?;
        let old = self.read_style_parts(sid_opt)?;
        let merged = merge_style_parts(old, patch);
        let new_sid = self.ensure_style_from_parts(&merged)?;
        self.apply_style_to_cell(coord, new_sid)?;
        Ok(())
    }

    fn read_style_parts(&self, style_id: Option<u32>) -> Result<StyleParts> {
        if let Some(sid) = style_id {
            let (font, fill) = self.xf_components(sid)?;
            let border = self.xf_border(sid)?;
            let align = self.xf_alignment(sid)?;
            Ok(StyleParts {
                num_fmt_code: None,
                font,
                fill,
                border,
                align,
            })
        } else {
            Ok(StyleParts::default())
        }
    }

    fn ensure_style_from_parts(&mut self, parts: &StyleParts) -> Result<u32> {
        self.ensure_style(
            parts.num_fmt_code.as_deref(),
            parts.font,
            parts.fill,
            parts.border,
            parts.align.as_ref(),
        )
    }
}

fn merge_style_parts(mut base: StyleParts, patch: &StyleParts) -> StyleParts {
    if patch.num_fmt_code.is_some() {
        base.num_fmt_code = patch.num_fmt_code.clone();
    }
    if patch.font.is_some() {
        base.font = patch.font;
    }
    if patch.fill.is_some() {
        base.fill = patch.fill;
    }
    if patch.border.is_some() {
        base.border = patch.border;
    }
    if patch.align.is_some() {
        base.align = patch.align.clone();
    }
    base
}

/* ========================== LOW-LEVEL HELPERS ============================= */

impl XlsxEditor {
    fn ensure_style(
        &mut self,
        num_fmt: Option<&str>,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        let fmt_id: u32 = if let Some(code) = num_fmt {
            self.ensure_num_fmt(code)?
        } else {
            0
        };

        if align.is_none() {
            if let Some(id) = self.find_matching_xf(fmt_id, font_id, fill_id, border_id)? {
                return Ok(id);
            }
        }

        self.add_new_xf(fmt_id, font_id, fill_id, border_id, align)
    }

    fn ensure_num_fmt(&mut self, code: &str) -> Result<u32> {
        // если есть кэш

        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut found_id = None;
        let mut max_custom_id = 163u32;

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) | Event::Empty(ref e) if e.name().as_ref() == b"numFmt" => {
                    let mut id = None::<u32>;
                    let mut text = None::<String>;
                    for a in e.attributes().with_checks(false).flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => id = Some(String::from_utf8_lossy(&a.value).parse()?),
                            b"formatCode" => text = Some(String::from_utf8_lossy(&a.value).into()),
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

        let id = if let Some(i) = found_id {
            i
        } else {
            let new_id = max_custom_id + 1;
            let tag = format!(r#"<numFmt numFmtId="{new_id}" formatCode="{code}"/>"#);

            if let Some(end) = find_bytes(&self.styles_xml, b"</numFmts>") {
                self.styles_xml.splice(end..end, tag.bytes());
                bump_count(&mut self.styles_xml, b"<numFmts", b"count=\"")?;
            } else {
                let insert = find_bytes(&self.styles_xml, b">")
                    .context("<styleSheet> start tag not found")?
                    + 1;
                let block = format!(r#"<numFmts count="1">{tag}</numFmts>"#);
                self.styles_xml.splice(insert..insert, block.bytes());
            }
            new_id
        };

        Ok(id)
    }

    fn find_matching_xf(
        &self,
        fmt_id: u32,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
    ) -> Result<Option<u32>> {
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
                    // С xf с alignment мы не сравниваем — пропускаем
                    let mut has_alignment_child = false;
                    // Event::Start -> значит дальше внутри могут быть теги
                    if matches!(ev, Event::Start(_)) {
                        let mut depth = 1;
                        while depth > 0 {
                            match rdr.read_event()? {
                                Event::Start(ref ie) => {
                                    if ie.name().as_ref() == b"alignment" {
                                        has_alignment_child = true;
                                    }
                                    depth += 1;
                                }
                                Event::End(_) => depth -= 1,
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    if has_alignment_child {
                        idx += 1;
                        continue;
                    }

                    let mut num = None::<u32>;
                    let mut fnt = None::<u32>;
                    let mut fil = None::<u32>;
                    let mut bdr = None::<u32>;
                    for a in e.attributes().with_checks(false).flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => num = Some(String::from_utf8_lossy(&a.value).parse()?),
                            b"fontId" => fnt = Some(String::from_utf8_lossy(&a.value).parse()?),
                            b"fillId" => fil = Some(String::from_utf8_lossy(&a.value).parse()?),
                            b"borderId" => bdr = Some(String::from_utf8_lossy(&a.value).parse()?),
                            _ => {}
                        }
                    }
                    let num_ok = num.unwrap_or(0) == fmt_id;
                    let font_ok = font_id.map_or(true, |v| Some(v) == fnt);
                    let fill_ok = fill_id.map_or(true, |v| Some(v) == fil);
                    let border_ok = border_id.map_or(true, |v| Some(v) == bdr);

                    if num_ok && font_ok && fill_ok && border_ok {
                        return Ok(Some(idx));
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn add_new_xf(
        &mut self,
        fmt_id: u32,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        let mut xf = String::from("<xf xfId=\"0\" ");

        if let Some(fid) = font_id {
            xf.push_str(&format!(r#"fontId="{fid}" applyFont="1" "#));
        }
        if let Some(fid) = fill_id {
            xf.push_str(&format!(r#"fillId="{fid}" applyFill="1" "#));
        }
        if let Some(bid) = border_id {
            xf.push_str(&format!(r#"borderId="{bid}" applyBorder="1" "#));
        }
        xf.push_str(&format!(
            r#"numFmtId="{}"{} "#,
            fmt_id,
            if fmt_id != 0 { r#" applyNumberFormat="1""# } else { "" }
        ));
        if align.is_some() {
            xf.push_str(r#"applyAlignment="1" "#);
        }
        xf.pop();
        xf.push('>');

        if let Some(al) = align {
            if al.horiz.is_some() || al.vert.is_some() || al.wrap {
                xf.push_str("<alignment");
                if let Some(h) = &al.horiz {
                    xf.push_str(&format!(r#" horizontal="{}""#, h.to_string()));
                }
                if let Some(v) = &al.vert {
                    xf.push_str(&format!(r#" vertical="{}""#, v.to_string()));
                }
                if al.wrap {
                    xf.push_str(r#" wrapText="1""#);
                }
                xf.push_str("/>");
            }
        }
        xf.push_str("</xf>");

        let pos = find_bytes(&self.styles_xml, b"</cellXfs>")
            .context("styles.xml: </cellXfs> not found")?;
        self.styles_xml.splice(pos..pos, xf.bytes());
        bump_count(&mut self.styles_xml, b"<cellXfs", b"count=\"")?;

        // посчитать индекс нового
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
        Ok(cnt - 1)
    }

    fn ensure_font(&mut self, name: &str, size: f32, bold: bool, italic: bool) -> Result<u32> {
        // let key = FontKey {
        //     name: name.to_string(),
        //     size: (size * 100.0).round() as u32,
        //     bold,
        //     italic,
        // };

        // пройдёмся по существующим <font>, попробуем найти совпадение
        // (для простоты — без глубокого парсинга)
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut fonts_cnt = 0u32;
        let mut in_fonts_block = false;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"fonts" => in_fonts_block = true,
                Event::End(ref e) if e.name().as_ref() == b"fonts" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_fonts_block && e.name().as_ref() == b"font" =>
                {
                    fonts_cnt += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }

        // Добавляем новый
        let insert = find_bytes(&self.styles_xml, b"</fonts>")
            .context("<fonts> block not found in styles.xml")?;
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
        self.styles_xml.splice(insert..insert, xml.bytes());
        bump_count(&mut self.styles_xml, b"<fonts", b"count=\"")?;

        Ok(fonts_cnt)
    }


    fn ensure_fill(&mut self, rgb: &str) -> Result<u32> {

        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut fills_cnt = 0u32;
        let mut in_fills_block = false;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"fills" => in_fills_block = true,
                Event::End(ref e) if e.name().as_ref() == b"fills" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_fills_block && e.name().as_ref() == b"fill" =>
                {
                    fills_cnt += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let insert = find_bytes(&self.styles_xml, b"</fills>")
            .context("<fills> block not found in styles.xml")?;
        let xml = format!(
            r#"<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>"#
        );
        self.styles_xml.splice(insert..insert, xml.bytes());
        bump_count(&mut self.styles_xml, b"<fills", b"count=\"")?;

        Ok(fills_cnt)
    }


    fn ensure_border(&mut self, style: &str) -> Result<u32> {

        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut cnt: u32 = 0;
        let mut in_borders_block = false;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"borders" => in_borders_block = true,
                Event::End(ref e) if e.name().as_ref() == b"borders" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_borders_block && e.name().as_ref() == b"border" =>
                {
                    cnt += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        let new_id = cnt;

        let end_pos = find_bytes(&self.styles_xml, b"</borders>")
            .context("styles.xml: </borders> not found")?;
        let tag = format!(
            r#"<border><left style="{s}"/><right style="{s}"/><top style="{s}"/><bottom style="{s}"/><diagonal/></border>"#,
            s = style
        );
        self.styles_xml.splice(end_pos..end_pos, tag.bytes());
        bump_count(&mut self.styles_xml, b"<borders", b"count=\"")?;

        Ok(new_id)
    }


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
                                b"fontId" => font = Some(String::from_utf8_lossy(&a.value).parse()?),
                                b"fillId" => fill = Some(String::from_utf8_lossy(&a.value).parse()?),
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
        Ok((None, None))
    }

    fn xf_border(&self, style_id: u32) -> Result<Option<u32>> {
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
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() == b"borderId" {
                                let val: u32 = String::from_utf8_lossy(&a.value).parse()?;
                                return Ok(Some(val));
                            }
                        }
                        return Ok(None);
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn xf_alignment(&self, style_id: u32) -> Result<Option<AlignSpec>> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut xf_idx = 0u32;
        let mut depth = 0;

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,

                Event::Start(ref e) if in_xfs && e.name().as_ref() == b"xf" => {
                    if xf_idx == style_id {
                        depth = 1;
                        while depth > 0 {
                            match rdr.read_event()? {
                                Event::Start(ref ie) => {
                                    depth += 1;
                                    if ie.name().as_ref() == b"alignment" {
                                        let mut spec = AlignSpec::default();
                                        for attr in ie.attributes().with_checks(false).flatten() {
                                            let val = String::from_utf8_lossy(&attr.value).into_owned();
                                            match attr.key.as_ref() {
                                                b"horizontal" => spec.horiz = Some(val.parse()?),
                                                b"vertical" => spec.vert = Some(val.parse()?),
                                                b"wrapText" => if val == "1" { spec.wrap = true },
                                                _ => {}
                                            }
                                        }
                                        return Ok(Some(spec));
                                    }
                                }
                                Event::End(_) => depth -= 1,
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        return Ok(None);
                    }
                    xf_idx += 1;
                }
                Event::Empty(ref _e) if in_xfs => {
                    if xf_idx == style_id {
                        return Ok(None);
                    }
                    xf_idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn cell_style_id(&self, coord: &str) -> Result<Option<u32>> {
        let tag = format!(r#"<c r="{coord}""#);
        if let Some(pos) = find_bytes(&self.sheet_xml, tag.as_bytes()) {
            if let Some(spos) = find_bytes_from(&self.sheet_xml, b" s=\"", pos) {
                let val_start = spos + 4;
                let val_end = find_bytes_from(&self.sheet_xml, b"\"", val_start + 1)
                    .unwrap_or(val_start);
                let id = std::str::from_utf8(&self.sheet_xml[val_start..val_end])?
                    .parse::<u32>()
                    .unwrap_or(0);
                return Ok(Some(id));
            }
        }
        Ok(None)
    }

    fn apply_style_to_cell(&mut self, coord: &str, style: u32) -> Result<()> {
        let row_num = coord.trim_start_matches(|c: char| c.is_ascii_alphabetic());
        let row_tag = format!(r#"<row r="{row_num}""#);

        let row_pos = match find_bytes(&self.sheet_xml, row_tag.as_bytes()) {
            Some(p) => p,
            None => {
                self.set_cell(coord, "")?;
                return self.apply_style_to_cell(coord, style);
            }
        };

        let row_end = find_bytes_from(&self.sheet_xml, b"</row>", row_pos)
            .context("</row> not found")?;

        let cell_tag = format!(r#"<c r="{coord}""#);
        let cpos = match find_bytes_from(&self.sheet_xml, cell_tag.as_bytes(), row_pos) {
            Some(p) => p,
            None => {
                let new_cell = format!(r#"<c r="{coord}" s="{style}"/>"#);
                self.sheet_xml.splice(row_end..row_end, new_cell.bytes());
                return Ok(());
            }
        };

        let ctag_end = find_bytes_from(&self.sheet_xml, b">", cpos)
            .context("malformed <c> tag")?;

        if let Some(sattr) = find_bytes_from(&self.sheet_xml, b" s=\"", cpos) {
            if sattr < ctag_end {
                let val_start = sattr + 4;
                let val_end = find_bytes_from(&self.sheet_xml, b"\"", val_start + 1)
                    .context("attr closing '\"' not found")?;
                self.sheet_xml
                    .splice(val_start..val_end, style.to_string().bytes());
                return Ok(());
            }
        }
        self.sheet_xml.splice(
            ctag_end..ctag_end,
            format!(r#" s="{style}""#).bytes(),
        );
        Ok(())
    }
}

/* ========================== НОРМАЛИЗАЦИЯ <cols> =========================== */

#[derive(Clone, Debug, Default, PartialEq)]
struct ColProp {
    width: Option<f64>,
    style: Option<u32>,
    best_fit: bool,
    custom_width: bool,
    hidden: bool,
}

fn equal_props(a: &ColProp, b: &ColProp) -> bool {
    a.width == b.width
        && a.style == b.style
        && a.best_fit == b.best_fit
        && a.custom_width == b.custom_width
        && a.hidden == b.hidden
}

impl XlsxEditor {
    /// Главный публичный метод для столбца: точечное изменение + нормализация.
    fn set_column_properties(
        &mut self,
        col0: u32,                  // 0-based
        width: Option<f64>,
        style_id: Option<u32>,
    ) -> Result<()> {
        let (cols_start, cols_end) = self.ensure_cols_block()?;

        let mut cols_map = self.read_cols_map(cols_start, cols_end)?;
        let idx = col0 + 1; // храним в map 1-based для удобства
        let prop = cols_map.entry(idx).or_default();

        if let Some(w) = width {
            prop.width = Some(w);
            prop.custom_width = true;
        }
        if let Some(s) = style_id {
            prop.style = Some(s);
        }

        self.write_cols_map(cols_start, cols_end, &cols_map)
    }

    /// Более безопасный путь задания number format для столбца:
    /// 1) создаём style_id 1 раз
    /// 2) обновляем <cols> нормализованно
    /// 3) проставляем этот style_id во все существующие <c> в столбце
    fn force_column_number_format(&mut self, col0: u32, style_id: u32) -> Result<()> {
        self.set_column_properties(col0, None, Some(style_id))?;

        let col_letter = col_letter(col0);
        let sid_str = style_id.to_string();
        let pat = format!(r#"<c\b[^>]*\br="{}[0-9]+"[^>]*>"#, col_letter.to_ascii_uppercase());
        let re = Regex::new(&pat)?;

        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len() + 512);
        let mut last = 0usize;

        let utf = std::str::from_utf8(&src)?;
        for m in re.find_iter(utf) {
            dst.extend_from_slice(&src[last..m.start()]);

            let cell_start = m.start();
            let tag_end = find_bytes_from(&src, b">", cell_start).context("cell tag end")? + 1;
            let mut cell = src[cell_start..tag_end].to_vec();

            if let Some(p) = find_bytes(&cell, b" s=\"") {
                let v0 = p + 4;
                let v1 = find_bytes_from(&cell, b"\"", v0 + 1).context("closing quote")?;
                cell.splice(v0..v1, sid_str.bytes());
            } else {
                let ins = if cell[cell.len() - 2] == b'/' { cell.len() - 2 } else { cell.len() - 1 };
                cell.splice(ins..ins, format!(r#" s="{sid_str}""#).bytes());
            }
            dst.extend_from_slice(&cell);
            last = tag_end;
        }
        dst.extend_from_slice(&src[last..]);
        self.sheet_xml = dst;
        Ok(())
    }

    fn ensure_cols_block(&mut self) -> Result<(usize, usize)> {
        if let (Some(start), Some(end)) = (
            find_bytes(&self.sheet_xml, b"<cols>"),
            find_bytes(&self.sheet_xml, b"</cols>"),
        ) {
            return Ok((start, end + "</cols>".len()));
        }

        // куда вставлять: после </sheetFormatPr> если он есть, иначе перед <sheetData>
        let anchor_end = if let Some(p) = find_bytes(&self.sheet_xml, b"</sheetFormatPr>") {
            p + "</sheetFormatPr>".len()
        } else {
            find_bytes(&self.sheet_xml, b"<sheetData")
                .context("<sheetData> not found on the current sheet")?
        };

        let block = b"<cols></cols>";
        self.sheet_xml.splice(anchor_end..anchor_end, block.iter().copied());

        let start = anchor_end;
        let end = start + block.len();
        Ok((start, end))
    }

    fn read_cols_map(&self, cols_start: usize, cols_end: usize) -> Result<BTreeMap<u32, ColProp>> {
        let mut map: BTreeMap<u32, ColProp> = BTreeMap::new();
        let cols_xml = &self.sheet_xml[cols_start..cols_end];
        let text = std::str::from_utf8(cols_xml)?;
        let re = Regex::new(r#"<col\b[^>]*/>"#)?;
        let attrs_re = Regex::new(r#"([a-zA-Z:]+)\s*=\s*"([^"]*)""#)?;

        for m in re.find_iter(text) {
            let tag = &text[m.start()..m.end()];
            let mut attrs = BTreeMap::new();
            for cap in attrs_re.captures_iter(tag) {
                attrs.insert(cap[1].to_string(), cap[2].to_string());
            }

            let min: u32 = attrs.get("min").unwrap_or(&"1".to_string()).parse()?;
            let max: u32 = attrs.get("max").unwrap_or(&min.to_string()).parse()?;

            let style = attrs.get("style").and_then(|s| s.parse::<u32>().ok());
            let width = attrs.get("width").and_then(|s| s.parse::<f64>().ok());
            let best_fit = attrs.get("bestFit").map_or(false, |v| v == "1" || v == "true");
            let custom_width = attrs.get("customWidth").map_or(false, |v| v == "1" || v == "true");
            let hidden = attrs.get("hidden").map_or(false, |v| v == "1" || v == "true");

            let prop = ColProp {
                width,
                style,
                best_fit,
                custom_width,
                hidden,
            };
            for i in min..=max {
                map.insert(i, prop.clone());
            }
        }
        Ok(map)
    }

    fn write_cols_map(
        &mut self,
        cols_start: usize,
        cols_end: usize,
        map: &BTreeMap<u32, ColProp>,
    ) -> Result<()> {
        // Сжимаем одинаковые проперти в диапазоны
        let mut out = String::with_capacity(256);
        out.push_str("<cols>");

        let mut it = map.iter().peekable();
        while let Some((&i, prop)) = it.next() {
            let mut j = i;
            while let Some(&(&k, prop2)) = it.peek() {
                if k == j + 1 && equal_props(prop, prop2) {
                    j = k;
                    it.next();
                } else {
                    break;
                }
            }
            out.push_str(&build_one_col_tag(i, j, prop));
        }

        out.push_str("</cols>");

        // подменяем всё содержимое старого блока
        self.sheet_xml.splice(cols_start..cols_end, out.bytes());
        Ok(())
    }
}

fn build_one_col_tag(min: u32, max: u32, p: &ColProp) -> String {
    let mut s = format!(r#"<col min="{min}" max="{max}""#);
    if let Some(w) = p.width {
        s.push_str(&format!(r#" width="{w}""#));
        if p.custom_width {
            s.push_str(r#" customWidth="1""#);
        }
    }
    if let Some(st) = p.style {
        s.push_str(&format!(r#" style="{st}""#));
    }
    if p.best_fit {
        s.push_str(r#" bestFit="1""#);
    }
    if p.hidden {
        s.push_str(r#" hidden="1""#);
    }
    s.push_str("/>");
    s
}

/* ========================== BYTE/STRING HELPERS =========================== */

pub fn col_letter(mut n: u32) -> String {
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
    s.bytes().fold(0, |acc, b| acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize) - 1
}
pub fn split_coord(coord: &str) -> (u32, u32) {
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
            let end = find_bytes_from(xml, b"\"", start).context("closing quote not found")?;
            let mut num: u32 = std::str::from_utf8(&xml[start..end])?.parse()?;
            num += 1;
            xml.splice(start..end, num.to_string().bytes());
            return Ok(());
        }
    }
    Err(anyhow::anyhow!("attribute count not found"))
}

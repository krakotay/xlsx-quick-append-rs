#[cfg(test)]
use crate::{XlsxEditor, scan};
#[cfg(test)]
use anyhow::Result;
#[test]
#[cfg(test)]
fn test_insert_table_at() -> Result<()> {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let data = vec![
        ["Name", "Score", "Status", "Number"],
        ["Alice", "123", "OK", "1"],
        ["Bob", "456", "FAIL", "2"],
    ];

    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.append_table_at("A4", data)?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;

    Ok(())
}
#[test]
fn test_insert_cells() -> Result<()> {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.set_cell("A25", "Hello")?;
    app.set_cell("B25", "World")?;
    app.set_cell("C25", "!")?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;

    Ok(())
}
#[test]
fn test_get_last_row_index() -> Result<()> {
    let file_name = "../test/test_last_row_index.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let app = XlsxEditor::open(file_name, &sheet_names[0])?;
    assert_eq!(app.get_last_row_index("A")?, 4);
    assert_eq!(app.get_last_row_index("B")?, 5);
    assert_eq!(app.get_last_row_index("C")?, 8);
    assert_eq!(app.get_last_row_index("D")?, 8);
    Ok(())
}
#[test]
fn test_get_last_roww_index() -> Result<()> {
    let file_name = "../test/test_last_row_index.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let app = XlsxEditor::open(file_name, &sheet_names[0])?;
    assert_eq!(app.get_last_roww_index("A:D")?, vec![4, 5, 8, 8]);
    Ok(())
}
#[cfg(test)]
#[cfg(feature = "polars")]
use polars_core::prelude::*;
#[test]
#[cfg(feature = "polars")]
fn test_write_polars() -> Result<()>  {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    let s1 = Column::new("Fruit".into(), ["Apple", "Apple", "Pear"]);
    let s2 = Column::new("Color".into(), ["Red", "Yellow", "Green"]);

    let df: DataFrame = DataFrame::new(vec![s1, s2])?;
    app.with_polars(&df, None)?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;
    Ok(())
}

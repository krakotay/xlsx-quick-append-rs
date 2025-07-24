#[cfg(test)]
use crate::{XlsxEditor, scan, Result};
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
    let file_name = "../test/test_last_row_index.xlsx";
    let sheet_names: Vec<String> = scan(file_name)?;
    let app = XlsxEditor::open(file_name, &sheet_names[0])?;
    assert_eq!(app.get_last_roww_index("A:D")?, vec![4, 5, 8, 8]);
    Ok(())
}

#[test]
fn add_new_worksheet() -> Result<()> {
    let file_name = "../test/test_new_ws.xlsx"; // fixed
    let new_file_name = "../test/test_new_ws_out.xlsx";

    let mut app = XlsxEditor::open(file_name, &scan(file_name)?[0])?;
    app.append_table_at("A1", [["Name", "Score", "Status", "Number"]])?;
    app.add_worksheet("NewSheet")?.set_cell("A1", "123")?;
    app.add_worksheet("NewSheet2")?
        .append_table_at("A1", [["Name", "Score", "Status", "Number"]])?;
    app.save(new_file_name)?;
    let sheet_names: Vec<String> = scan(new_file_name)?;

    println!("Sheet names: {:#?}", sheet_names);
    assert!(sheet_names.contains(&"NewSheet".to_owned()));
    assert!(sheet_names.contains(&"NewSheet2".to_owned()));
    Ok(())
}

#[test]
fn set_number_format() -> Result<()> {
    let file_name = "../test/numeric_format_test.xlsx";
    let file_name_out = "../test/numeric_format_test_out.xlsx";
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.set_number_format("A9", "#,##0.00")?;
    app.set_number_format("B3:C5", "#,##0.00")?;
    app.save(file_name_out)?;
    Ok(())
}
#[test]
fn set_style() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let file_name_out = "../test/style_test_out.xlsx";

    let mut xl = XlsxEditor::open(file_name, "Sheet1")?;

    xl.set_fill("B14:B18", "FFFF00")?
        .set_font("D4:D8", "Arial", 12.0, true, false)?
        .set_fill("E4:E8", "FFCCCC")?
        .set_font("A1:C3", "Calibri", 10.0, false, true)?
        .set_fill("A1:C3", "FFFF00")?
        .merge_cells("B12:D12")?;

    xl.save(file_name_out)?;
    Ok(())
}
#[test]
fn set_column_number_format() -> Result<()> {
    let file_name = "../test/numeric_format_test.xlsx";
    let file_name_out = "../test/numeric_format_column_test.xlsx";

    let mut xl = XlsxEditor::open(file_name, "Sheet1")?;

    xl.set_number_format("A:", "#,##0.00")?;
    xl.set_number_format("B:", "#,##0.00")?;
    xl.set_number_format("C:", "#,##0.00")?;
    xl.set_number_format("G:", "#,##0.00")?;

    xl.save(file_name_out)?;
    Ok(())
}

#[cfg(test)]
#[cfg(feature = "polars")]
use polars_core::prelude::*;
#[test]
#[cfg(feature = "polars")]
fn test_write_polars() -> Result<()> {
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

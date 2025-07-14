#[cfg(test)]
use crate::{XlsxEditor, scan};
#[test]
#[cfg(test)]
fn test_insert_table_at() -> anyhow::Result<()> {
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
fn test_insert_cells() -> anyhow::Result<()> {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.set_cell("A25", "Hello")?;
    app.set_cell("B25", "World")?;
    app.set_cell("C25", "!")?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;

    Ok(())
}

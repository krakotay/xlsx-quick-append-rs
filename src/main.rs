use xlsx_append_rs::XlsxAppender;

fn main() -> anyhow::Result<()> {
    let mut app = XlsxAppender::open("Шаблон53. РД Выборка.xlsx")?;
    app.append_row(["Alice", "123", "OK"])?;
    app.append_row(["Bob", "456", "FAIL"])?;
    app.save("result_appended.xlsx")?;
    Ok(())
}

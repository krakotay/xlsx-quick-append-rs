# Usage Guide

This document describes how to use the `xlsx-quick-append-rs` project.

## Rust library

The core functionality resides in the `rust-core` crate. It exposes the
`XlsxEditor` type which can open an existing `.xlsx` file and modify its
contents.

### Opening a sheet
```rust
use rust_core::{XlsxEditor, scan};

let sheet_names = scan("test.xlsx")?;
let mut editor = XlsxEditor::open("test.xlsx", &sheet_names[0])?;
```

### Adding data
Append a row to the end of the current worksheet:
```rust
editor.append_row(["Hello", "World"])?;
```

Insert a table starting from a specific cell:
```rust
let rows = vec![vec!["1"], vec!["2"]];
editor.append_table_at("C1", rows)?;
```

Set the value of an individual cell:
```rust
editor.set_cell("A1", "Some text")?;
```

### Saving
Write the modified workbook to a new file:
```rust
editor.save("output.xlsx")?;
```

## Python bindings

Bindings are provided in the `python-bindings` crate. After building with
`maturin`, the package `xlsx_append_py` offers an interface similar to the
Rust API.

Example:
```python
from xlsx_append_py import scan_excel, PyXlsxEditor

sheets = scan_excel("tests/test_sum.xlsx")
editor = PyXlsxEditor("tests/test_sum.xlsx", sheets[0])
editor.append_row(["Hello", "World"])
editor.save("tests/result.xlsx")
```

Refer to `python-bindings/tests` for more examples.

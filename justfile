set shell := ["pwsh.exe", "-c"]

python-build:
    python-bindings\.venv\Scripts\activate 
    cd python-bindings && maturin build  --release

python-build-polars:
    python-bindings\.venv\Scripts\activate 
    cd python-bindings && maturin build  -F polars --release

python-develop:
    python-bindings\.venv\Scripts\activate 
    cd python-bindings && maturin develop  --release
python-develop-polars:
    python-bindings\.venv\Scripts\activate 
    cd python-bindings && maturin develop  --release  -F polars

# Document Loader

**Supported formats**
- pdf (`poppler`)
- doc (`antiword`/`catdoc`)
- docx (`docxpy`)
- rtf

**Installation**
```
apt-get update && apt-get install -y antiword catdoc git pkg-config libpoppler-private-dev libpoppler-cpp-dev \
pip install --upgrade setuptools cython
pip install git+https://github.com/izderadicka/pdfparser
pip install -r requirements.txt
```

**Examples**
- see `/test`


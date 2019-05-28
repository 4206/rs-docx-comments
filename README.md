Microsoft Word Document Docx Comment Extractor
==============================================

Simple Rust library to extract comments and the text ranges referenced by them
from Word Documents.
Currently only Microsoft Word (docx) is supported.

This is a fork of Dotext from https://github.com/anvie/dotext

Build
------

The executable is generated as an example to the library.

```bash
$ cargo test
```

This will generate the executable ```target/debug/examples/readdocx-comments```

Usage
------

```
readdocx-comments [-c] [-d] [-n] filename
```

Option ```-c``` extracts the text inside the comments.
Option ```-d``` extracts the text that the comments refer to.

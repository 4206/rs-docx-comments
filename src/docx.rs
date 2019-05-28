
use zip::ZipArchive;

use xml::reader::Reader;
use xml::events::Event;

use std::path::{Path, PathBuf};
use std::fs::File;
use std::io::prelude::*;
use std::io::Cursor;
use std::io;
use std::clone::Clone;
use zip::read::ZipFile;

use doc::{MsDoc, HasKind};

pub struct Docx {
    path: PathBuf,
    data: Cursor<String>
}

impl HasKind for Docx {
    fn kind(&self) -> &'static str {
        "Word Document"
    }

    fn ext(&self) -> &'static str {
        "docx"
    }
}

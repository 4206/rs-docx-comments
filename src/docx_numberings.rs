

use zip::ZipArchive;

use xml::reader::Reader;
use xml::events::{Event,BytesStart};
use std::io::{Error, ErrorKind};

use std::path::{Path, PathBuf};
use std::fs::File;
use std::io::prelude::*;
use std::io;
use std::clone::Clone;
use zip::read::ZipFile;

use ::Docx;

use std::collections::HashMap;
pub trait ReadNumbering<T> {
    fn open_numbering<P: AsRef<Path>>(path: P) -> io::Result<HashMap<numId,DocxNumbering>>;
}

pub struct DocxNumbering {
    num_id: usize,
    format: DocxNumFmt
}

// TODO: probably should be Item and Enumerate
pub enum DocxNumFmt {
    None,
    Bullet,
    Decimal,
    Other,
}

impl DocxNumFmt {
    fn read(num_fmt: &str) -> DocxNumFmt {
        match num_fmt {
            "bullet" => DocxNumFmt::Bullet
           ,"decimal" => DocxNumFmt::Decimal
           ,"none" => DocxNumFmt::None
           ,_ => DocxNumFmt::Other
        }
    }
}

/// numbering.xml maps "numId" to "abstractNumId"
/// and then "abstractNumId" to formats
impl ReadNumbering<Docx> for Docx {
    fn open_numbering<P: AsRef<Path>>(path: P) -> io::Result<HashMap<numId,DocxNumbering>> {

        let file = File::open(path.as_ref())?;
        let mut archive = ZipArchive::new(file)?;

        let mut xml_data = String::new();

        for i in 0..archive.len() {
            let mut c_file = archive.by_index(i).unwrap();
            if c_file.name() == "word/numbering.xml" {
                c_file.read_to_string(&mut xml_data);
                break
            }
        }
        // post: if no document.xml was found, then no data was read to buffer
        
        if xml_data.len() == 0 {
            return Err(io::Error::new(io::ErrorKind::NotFound,".docx invalid: did not contain word/numbering.xml".to_string()))
        }

        let xml_reader = Reader::from_str(xml_data.as_ref());

        let (con_abs_map,abs_fmt_map) = read_numbering(xml_reader)?;
        let con_fmt_map = join_numbering(con_abs_map, abs_fmt_map);
        
        Ok(con_fmt_map)
    }
}

#[allow(non_camel_case_types)]
pub type numId = usize;
#[allow(non_camel_case_types)]
type abstractNumId = usize;
#[allow(non_camel_case_types)]
type numFmt = String;

pub fn read_and_join_numbering<B: BufRead>(xml_reader: Reader<B>) -> io::Result<HashMap<numId,DocxNumbering>> {
    let (con_abs_map,abs_fmt_map) = read_numbering(xml_reader)?;
    Ok(join_numbering(con_abs_map, abs_fmt_map))
}

fn join_numbering(con_abs_map: HashMap<numId,abstractNumId>, abs_fmt_map: HashMap<abstractNumId,Vec<numFmt>>) -> HashMap<numId,DocxNumbering> {
    let mut res = HashMap::new();
        // join maps (sql style)
    for (con_id,abs_id) in con_abs_map.iter() {
        // select the first, ignore the rest
        let fmt_str_levels = abs_fmt_map.get(abs_id).expect(&format!("abstractNumId {} was defined in con-abs-mapping but not in abs-fmt",abs_id));
        let fmt_str = fmt_str_levels.get(0).expect(&format!("abstractNumId {} did not contain any formats ({})",abs_id,fmt_str_levels.len()));
        let fmt = DocxNumFmt::read(fmt_str);
        let r_entry = DocxNumbering { num_id: *con_id, format: fmt };
        res.insert(*con_id, r_entry);
    }
    res
}

fn read_numbering<B: BufRead>(mut xml_reader: Reader<B>) -> io::Result<(HashMap<numId,abstractNumId>,HashMap<abstractNumId,Vec<numFmt>>)> {
    
    let mut con_abs_map: HashMap<numId,abstractNumId> = HashMap::new(); // mapping from numId to abstractNumId at the end of the file
    let mut abs_fmt_map: HashMap<abstractNumId,Vec<numFmt>> = HashMap::new(); // mapping from abstractNumId to fmt at the beginning of the file

    let mut buf = Vec::new();
    let mut abstract_num_id_opt = None; // abstract num id from abs-fmt-part (also appears in con-abs-part)
    let mut num_id_opt = None; // concrete num id (from con-abs-part, only appears there)
    //let mut num_id = None;
    let mut fmt = Vec::new(); // format per indentation level (offset begins at "0") for current abstract_num_id

    fn attr_as_string(event: &BytesStart, key: &[u8]) -> String {
        let val_cstr = event.attributes().find(|a| a.as_ref().unwrap().key==key ).map(|a| a.unwrap().value ).expect(&format!("malformed word/comments.xml: missing attribute '{:?}' on tag 'comment'",key));
        String::from_utf8(val_cstr.to_vec()).unwrap()
    }
    fn attr_as_usize(event: &BytesStart, key: &[u8]) -> usize {
        let val_str = attr_as_string(event,key);
        val_str.parse::<usize>().expect(&format!("comment attribute '{:?}' was not a number",key))
    }

    loop {
        match xml_reader.read_event(&mut buf) {

            Ok(Event::Start(ref e)) => {
                match e.name() {
                      b"w:abstractNum" => { // begin-end may appear toplvl (inside w:num only as Empty)
                        abstract_num_id_opt = Some(attr_as_usize(e,b"w:abstractNumId"));
                        //fmt = Vec::new(); // assume this has been reset before
                    }
                    , b"w:lvl" => {} // ignore the w:ilvl attribute; assume they appear in order 0..n
                    , b"w:num" => {
                        num_id_opt = Some(attr_as_usize(e,b"w:numId"));
                    }
                    , _ => ()
                }
            }
            , Ok(Event::Empty(ref e)) => {
                match e.name() {
                      b"w:abstractNumId" => { // found leaf of con-abs-entry
                        let con_id = num_id_opt.unwrap();
                        let abs_id = attr_as_usize(e,b"w:val");
                        con_abs_map.insert(con_id, abs_id);
                        num_id_opt = None; // reset num id
                    }
                    , b"w:numFmt" => {
                        let lvl = attr_as_string(e,b"w:val");
                        fmt.push(lvl);
                    }
                    , _ => ()
                }
            }           
            , Ok(Event::End(ref e)) => {
                match e.name() {
                    b"w:abstractNum" => { // found head of abs-fmt-entry
                        let abstract_num_id = abstract_num_id_opt.unwrap();
                        abs_fmt_map.insert(abstract_num_id, fmt);
                        abstract_num_id_opt = None;
                        fmt = Vec::new();
                    }
                    , _ => ()
                }
            }
            /*, Ok(Event::Text(e)) => {
                if to_read {
                    txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                    to_read = false;
                }
            }*/
            , Ok(Event::Eof) => break
            , Ok(_) => ()
            , Err(e) => panic!("Error at position {}: {:?}", xml_reader.buffer_position(), e),
        }
    }

    Ok((con_abs_map, abs_fmt_map))
}
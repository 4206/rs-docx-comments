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

use std::collections::HashMap;

use ::Docx;
use get_attr::GetAttr;



// TODO: model as extension of a "identified range"
pub struct DocxHighlight {
  pub id: usize,  // hightlight color according to internal stringtable
  data: String
}

pub trait RangeText {
  fn text(&self) -> &str;
}

pub trait RangeId {
    fn id(&self) -> usize;
}

impl RangeText for DocxHighlight {
  fn text(&self) -> &str {
    self.data.as_str()
  }
}

impl RangeId for DocxHighlight {
    fn id(&self) -> usize {
        self.id
    }
}

pub trait ReadHighlights<T> {
  ///  extract all highlighted ranges from document
  fn open_highlighted<P: AsRef<Path>>(path: P) -> io::Result<(HashMap<usize,String>,Vec<DocxHighlight>)>;
}

impl ReadHighlights<Docx> for Docx {

  // take note that .docx only supports 16 colors
  fn open_highlighted<P: AsRef<Path>>(path: P) ->  io::Result<(HashMap<usize,String>,Vec<DocxHighlight>)> {
    let file = File::open(path.as_ref())?;
    let mut archive = ZipArchive::new(file)?;

    let mut xml_data = String::new();

    for i in 0..archive.len() {
        let mut c_file = archive.by_index(i).unwrap();
        if c_file.name() == "word/document.xml" {
            c_file.read_to_string(&mut xml_data);
            break
        }
    }
    // post: if no document.xml was found, then no data was read to buffer

    if xml_data.len() == 0 {
      return Err(io::Error::new(io::ErrorKind::NotFound,".docx invalid: did not contain word/document.xml".to_string()))
    }

    let xml_reader = Reader::from_str(xml_data.as_ref());
    read_highlighted(xml_reader)
  }
}

#[derive(PartialEq)]
enum ToRead {
  NoRead, NextWt, NextText
}

fn invert_hashmap<A: Clone,B: std::hash::Hash + Eq + Copy>(map: HashMap<A,B>) -> HashMap<B,A>
{
  let mut res = HashMap::with_capacity(map.len());
  for (key,val) in map.iter() {
    match res.insert((*val).clone(),(*key).clone()) {
      None => ()
    , Some(_) => eprintln!("Warning: invert_hashmap() key collision")
    }
  }
  return res;
}

fn read_highlighted<B: BufRead>(mut xml_reader: Reader<B>) -> io::Result<(HashMap<usize,String>,Vec<DocxHighlight>)> {
    let mut buf = Vec::new();
    let mut txt = Vec::new(); // a range
    let mut par = Vec::new();

    let mut stringtable: HashMap<String,usize> = HashMap::new(); // map color names to ints
    let mut highlight_ranges_open: HashMap<usize,String> = HashMap::new();
    let mut prev_highlight_id: Option<usize> = None;
    let mut cur_highlight_id : Option<usize> = None;

    let mut n_par = 0; // count paragraphs
    let mut n_r   = 0; // count ranges (total, spanning over paragraphs)

    let mut to_read = ToRead::NoRead;
    loop {
        // TODO: track highlights, in order to combine non-interleaved
        // ranges of same color
        match xml_reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"w:p" => {
                        n_par += 1;
                    }
                    , b"w:r" => { // any range 
                        n_r += 1;
                        // set state machine to look for <w:highlight w:val="..."/>
                        cur_highlight_id = None;
                        () // TODO
                    }
                    , b"w:t" => { // entered a text section
                        if to_read == ToRead::NextWt {
                          to_read = ToRead::NextText;
                        }
                    }
                    , b"w:highlight" => {
                      println!("<w:highlight>");
                    }
                    , _ => ()
                }
            }
            , Ok(Event::Empty(ref e)) => {
                match e.name() {
                    b"w:highlight" => { // expected within <w:r><w:rPr>
                        let cur_highlight_val = e.get_attr(b"w:val"); // e.g. yellow, red
                        // lookup color id or generate anew
                        let next_id = stringtable.len();
                        cur_highlight_id = Some(*stringtable.entry(cur_highlight_val).or_insert(next_id));

                        to_read = ToRead::NextWt;                      
                    }
                    , _ => ()
                }
            }
            , Ok(Event::End(ref e)) => {
                match e.name() {
                    b"w:rPr" => {
                        // flush if a new highlight color appears
                        if cur_highlight_id != prev_highlight_id {
                            match prev_highlight_id {
                                None => ()
                              , Some(x) => {
                                  // range ended, push to result
                                  let highlighted_range = DocxHighlight{ id: x, data: txt.join("") };
                                  par.push(highlighted_range);
                                  txt = Vec::new();
                                }
                            }
                        } // else do nothing
                    }
                    , b"w:r" => {
                        prev_highlight_id = cur_highlight_id;
                    }
                    , _ => ()
                }
            }
            , Ok(Event::Text(e)) => {
                if to_read == ToRead::NextText {
                    //print!("text:");
                    //println!("{}", e.unescape_and_decode(&xml_reader).unwrap());
                    txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                    to_read = ToRead::NoRead;
                }
            }
            , Ok(Event::Eof) => break
            , Ok(_) => ()
            , Err(e) => panic!("Error at position {}: {:?}", xml_reader.buffer_position(), e),
        }
    }

    // write buffer to result after last range
    match prev_highlight_id {
        None => ()
      , Some(x) => {
          // range ended, push to result
          let highlighted_range = DocxHighlight{ id: x, data: txt.join("") };
          par.push(highlighted_range);
          //txt = Vec::new();
        }
    }

    // invert stringtable
    let inverted_stringtable = invert_hashmap(stringtable);


    // could also panic
    //if txt.len() > 0 {
    //    eprintln!("After reading all comments, buffer still contained: {}", txt.join(""));
    //}

    Ok((inverted_stringtable,par))
}
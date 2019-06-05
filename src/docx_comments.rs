


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

pub struct DocxComment {
    id: usize,
    data: String
}

pub trait Comment {
    fn text(&self) -> &str;
}

pub trait RangeId {
    fn id(&self) -> usize;
}

impl Comment for DocxComment {
    fn text(&self) -> &str {
        self.data.as_str()
    }
}

impl RangeId for DocxComment {
    fn id(&self) -> usize {
        self.id
    }
}

pub trait ReadComments<T> {
    /// read the comment contents
    fn open_comments<P: AsRef<Path>>(path: P) -> io::Result<Vec<DocxComment>>;
    /// read the contents of the regions referenced by comments
    fn open_commented<P: AsRef<Path>>(path: P) -> io::Result<Vec<DocxComment>>;
}

impl ReadComments<Docx> for Docx {
    
    fn open_comments<P: AsRef<Path>>(path: P) -> io::Result<Vec<DocxComment>> {
        let file = File::open(path.as_ref())?;
        let mut archive = ZipArchive::new(file)?;

        let mut xml_data = String::new();

        for i in 0..archive.len() {
            let mut c_file = archive.by_index(i).unwrap();
            if c_file.name() == "word/comments.xml" {
                c_file.read_to_string(&mut xml_data);
                break
            }
        }
        // post: if no document.xml was found, then no data was read to buffer
        
        if xml_data.len() == 0 {
            //return Err(io::Error::new(io::ErrorKind::NotFound,".docx invalid: did not contain word/comments.xml".to_string()))
            // documents without comments may lack this file
            return Ok(vec![]);
        }

        let xml_reader = Reader::from_str(xml_data.as_ref());
        read_comments(xml_reader)
    }

    /// collect the commented areas per id
    /// consider that comment ranges may overlap
    fn open_commented<P: AsRef<Path>>(path: P) -> io::Result<Vec<DocxComment>> {    
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
        read_commented(xml_reader)
    }
}


/*
fn with_zipped_xml<P: AsRef<Path>,A,B>(path: P, xml_filename: &str, f: impl Fn(Reader<A>)->io::Result<B>) -> io::Result<B>
where A: for<'r> BufRead
{
    let file = File::open(path.as_ref())?;
        let mut archive = ZipArchive::new(file)?;

        let mut xml_data = String::new();

        for i in 0..archive.len() {
            let mut c_file = archive.by_index(i).unwrap();
            if c_file.name() == xml_filename {
                c_file.read_to_string(&mut xml_data);
                break
            }
        }
        // post: if no document.xml was found, then no data was read to buffer
        
        if xml_data.len() == 0 {
            return Err(io::Error::new(io::ErrorKind::NotFound,format!(".docx invalid: did not contain {}",xml_filename)));
        }

        let xml_reader = Reader::from_str(xml_data.as_ref());
        f(xml_reader)
}
*/
    
fn read_comments<B: BufRead>(mut xml_reader: Reader<B>) -> io::Result<Vec<DocxComment>> {

    let mut buf = Vec::new();
    let mut txt = Vec::new(); // a range
    let mut par = Vec::new();

    let mut a_id: isize = -1;

    let mut to_read = false;
    loop {
        match xml_reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"w:comment" => { // entered a paragraph
                        let id_cstr = e.attributes().find(|a| a.as_ref().unwrap().key==b"w:id" ).map(|a| a.unwrap().value ).expect("malformed word/comments.xml: missing attribute 'w:id' on tag 'comment'");
                        a_id = String::from_utf8(id_cstr.to_vec()).unwrap().parse::<isize>().expect("comment attribute 'w:id' was not a number");
                    }
                    , b"w:t" => { // entered a text section
                        to_read = true;
                    }
                    , _ => ()
                }
            }
            , Ok(Event::End(ref e)) => {
                match e.name() {
                    b"w:comment" => { // exited a paragraph
                        let comment = DocxComment { id: a_id as usize, data: txt.join("\n") }; // join ranges
                        par.push(comment);
                        a_id = -1;
                        txt = Vec::new();
                    }
                    , _ => ()
                }
            }
            , Ok(Event::Text(e)) => {
                if to_read {
                    txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                    to_read = false;
                }
            }
            , Ok(Event::Eof) => break
            , Ok(_) => ()
            , Err(e) => panic!("Error at position {}: {:?}", xml_reader.buffer_position(), e),
        }
    }

    // could also panic
    if txt.len() > 0 {
        eprintln!("After reading all comments, buffer still contained: {}", txt.join(""));
    }

    Ok(par)
}

/*
/// get an xml attribute by key and return is string value
trait GetAttr {
    fn get_attr(&self, key:&[u8]) -> String;
}

impl<'a> GetAttr for BytesStart<'a> {
    fn get_attr(&self, key:&[u8]) -> String {
        let cstr = self.attributes().find(|a| a.as_ref().unwrap().key==key ).map(|a| a.unwrap().value ).expect(&format!("malformed word/comments.xml: missing attribute '{:?}' on tag 'comment'",key));
        String::from_utf8(cstr.to_vec()).unwrap()
    }
}
*/

/// Word allows comments to overlap on the text,
/// this means that any given text can be quoted by multiple comments.
/// The 'comment_ranges_open' collects the text ranges
/// for all currently open comments while walking over the xml file
fn read_commented<B: BufRead>(mut xml_reader: Reader<B>) -> io::Result<Vec<DocxComment>> {

    let mut buf = Vec::new();
    let mut txt = Vec::new(); // collection of text ranges within a contiguous comment range
    let mut par = Vec::new();

    // map from comment_id -> buffer
    // used for collecting text in multiple open comments
    let mut comment_ranges_open: HashMap<usize,String> = HashMap::new();

    fn attr_id(event: &BytesStart) -> usize {
        //let id_cstr = event.attributes().find(|a| a.as_ref().unwrap().key==b"w:id" ).map(|a| a.unwrap().value ).expect("malformed word/comments.xml: missing attribute 'w:id' on tag 'comment'");
        //let id_str = String::from_utf8(id_cstr.to_vec()).unwrap();
        let id_str = event.get_attr(b"w:id");
        let a_id = id_str.parse::<usize>().expect("comment attribute 'w:id' was not a number");
        a_id
    }

    let mut n_par = 0; // count paragraphs
    let mut n_r   = 0; // count ranges (total, spanning over paragraphs)

    let mut to_read = false;
    loop {
        match xml_reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name() {
                    b"w:p" => {
                        n_par += 1;
                    }
                    , b"w:r" => {
                        n_r += 1;
                    }
                    , b"w:t" => { // entered a text section
                        to_read = comment_ranges_open.len() > 0;
                    }
                    , _ => ()
                }
            }
            , Ok(Event::Empty(ref e)) => {
                match e.name() {
                    b"w:commentRangeStart" => { // begin a new commented block
                        // push existing txt buffer to currently open ranges
                        comment_ranges_open.values_mut()
                            .for_each(|str_i| str_i.push_str(&txt.join("")) );
                        txt = Vec::new();
                        /*
                        for (_k,str_i) in comment_ranges_open.iter_mut() {
                            str_i.push_str(&txt.join("")); // append to string
                        }*/
                        // open a new comment range
                        let a_id = attr_id(e);
                        comment_ranges_open.insert(a_id, String::new());

                    }
                    , b"w:commentRangeEnd" =>  { // end one of the currently running comments
                        // push existing txt buffer to currently open ranges
                        for (_k,str_i) in comment_ranges_open.iter_mut() {
                            str_i.push_str(&txt.join("")); // append to string
                        }
                        txt = Vec::new();
                        // move a currently open range to result
                        let a_id = attr_id(e);
                        match comment_ranges_open.remove(&a_id) {
                            Some(rng) => {
                                let comment = DocxComment{ id: a_id, data: rng };
                                par.push(comment);
                            }
                            // we could also panic
                            , None => { eprintln!("Malformed word/document.xml: comment {} closed, but was not open", a_id); }
                        }
                        
                    }
                    ,_=>()
                }
            }
            , Ok(Event::Text(e)) => {
                if to_read {
                    txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                    to_read = false;
                }
            }
            , Ok(Event::Eof) => break
            , Ok(_) => ()
            , Err(e) => panic!("Error at position {}: {:?}", xml_reader.buffer_position(), e),
        }
    }

    // could also panic
    if txt.len() > 0 {
        eprintln!("After reading all comments, buffer still contained: {}", txt.join(""));
    }
    if comment_ranges_open.len() > 0 {
        eprintln!("After reading all comments, {} comments were not closed", comment_ranges_open.len() );
        // TODO: close the malformed comments
    }

    Ok(par)
}

// TODO: do analysis for "Dieser Begriff soll in Abschnitt 2 bis 5 erwähnt werden"
// - need access to reference word behind comment id
// - need access to sections by number
// - need to find string match in section

/*
#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};
    use super::*;

    #[test]
    fn read_with_newlines() {
        let mut f = Docx::open_comments(Path::new("samples/sample-with-comment.docx")).unwrap();

        let mut data = String::new();
        // TODO...
    }

    // TODO: write a test that reads the sample...docx and checks for expected comments and commented

    // TODO: write a test that reads a non-exisiting file and checks for the correct error message

    // TODO: write a test that reads a text file (wrong format) and checks for the correct error message

    // TODO: write a test for arabic comments with RTL text direction

    // TODO: write a test for extracting highlighted parts from a file
}
*/
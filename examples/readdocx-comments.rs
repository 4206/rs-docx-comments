/**
 * Copyright 2019 4206. All rights reserved.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software
 * and associated documentation files (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies
 * or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
 * IN THE SOFTWARE.
 *
 */

extern crate dotext;
extern crate getopts;


use dotext::*;
use dotext::docx_comments::*;
use dotext::docx_highlights::*;
use dotext::docx_highlights::RangeId as HightlightRangeId;

use getopts::Options;
use std::env;


/// Read the comments in a docx file
/// Print each comment as a CString, prefixed by its comment id
fn main(){
    let args: Vec<String> = env::args().collect();
    let program_name = args[0].clone();

    let mut opts = Options::new();
    //opts.optopt("o", "", "set output file name", "NAME");
    opts.optflag("c", "comments", "extract comments");
    opts.optflag("d", "commented", "extract ranges referenced by comments");
    opts.optflag("h", "highlighted", "extract highlighted ranges");
    opts.optflag("", "help", "print this help menu");
    let matches = match opts.parse(&args[1..]) {
        Ok(m) => { m }
        Err(f) => { panic!(f.to_string()) }
    };
    if matches.opt_present("help") {
        print_usage(&program_name, opts);
        return;
    }
    //let output = matches.opt_str("o");
    let input_path = if !matches.free.is_empty() {
        matches.free[0].clone()
    } else {
        print_usage(&program_name, opts);
        return;
    };

    let mut matched = false;

    if matches.opt_present("c") {
        matched = true;
        let comments = Docx::open_comments(&input_path).expect("Cannot open file");
        for (i,comment_i) in comments.iter().enumerate()
        {
            // TODO: escape doublequotes and newlines
            let comment_id_i = comment_i.id();
            let cstring_comment_i = escape_as_cstr(comment_i.text());
            println!("{} \"{}\"", comment_id_i, cstring_comment_i);
        }
    }

    if matches.opt_present("d") {
        matched = true;
        let commented = Docx::open_commented(&input_path).expect("Cannot open file");
        for (i,comment_i) in commented.iter().enumerate()
        {
            // TODO: escape doublequotes and newlines
            let commented_id_i = comment_i.id();
            let cstring_comment_i = escape_as_cstr(comment_i.text());
            println!("{} \"{}\"", commented_id_i, cstring_comment_i);
        }
    }

    if matches.opt_present("h") {
        matched = true;
        let (stringtable,highlighted) = Docx::open_highlighted(&input_path).expect("Cannot open file");
        for highlighted_i in highlighted.iter()
        {
            let h_id = highlighted_i.id();
            let h_value = match stringtable.get(&h_id) { Some(color) => color, None => "??" };
            let cstring_highlighted_i = escape_as_cstr(highlighted_i.text());
            println!("{} \"{}\"", h_value, cstring_highlighted_i);
        }
    }

    if ! matched {
        print_usage(&program_name, opts);
        return;
    }
}


fn print_usage(program_name: &str, opts: Options) {
    let brief = format!("Usage: {} OPTIONS FILE", program_name);
    print!("{}", opts.usage(&brief));
}


// TODO: use a real escaping function instead of this dummy
fn escape_as_cstr(s: &str) -> String {
    s.replace("\"","\\\"").replace("\n","\\n").to_string()
}

// TODO: program to grep commented areas
// TODO: program to grep sections
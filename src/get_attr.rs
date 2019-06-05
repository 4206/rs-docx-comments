use xml::events::BytesStart;

/// get an xml attribute by key and return is string value
pub trait GetAttr {
    fn get_attr(&self, key:&[u8]) -> String;
}

impl<'a> GetAttr for BytesStart<'a> {
    fn get_attr(&self, key:&[u8]) -> String {
        let cstr = self.attributes().find(|a| a.as_ref().unwrap().key==key ).map(|a| a.unwrap().value ).expect(&format!("malformed word/comments.xml: missing attribute '{:?}' on tag 'comment'",key));
        String::from_utf8(cstr.to_vec()).unwrap()
    }
}
#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    let cursor = Cursor::new(data);
    // Try reading the first sheet by index.
    let _ = opensheet_core::reader::xlsx::read_single_sheet(cursor, None, Some(0));
});

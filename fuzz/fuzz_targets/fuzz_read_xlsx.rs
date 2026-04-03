#![no_main]
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

fuzz_target!(|data: &[u8]| {
    let cursor = Cursor::new(data);
    // We don't care about the result — only that it doesn't panic or hang.
    let _ = opensheet_core::reader::xlsx::read_xlsx(cursor);
});

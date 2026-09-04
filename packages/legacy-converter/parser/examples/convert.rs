//! Local diagnostic tool; never invoke Office or resolve linked resources.
use legacy_office_converter::{convert_native, LegacyFormat};
use std::{env, fs, path::Path};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = env::args().skip(1).collect();
    if args.len() != 2 {
        return Err("usage: convert <input.doc|xls|ppt> <new-output-path>".into());
    }
    let format = match Path::new(&args[0]).extension().and_then(|s| s.to_str()) {
        Some("doc") => LegacyFormat::Doc,
        Some("xls") => LegacyFormat::Xls,
        Some("ppt") => LegacyFormat::Ppt,
        _ => return Err("expected a lowercase doc, xls or ppt extension".into()),
    };
    let input = fs::read(&args[0])?;
    let output = convert_native(&input, format, 256 * 1024 * 1024)?;
    use std::io::Write;
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&args[1])?;
    file.write_all(&output.bytes)?;
    for warning in output.warnings {
        eprintln!("{warning}");
    }
    Ok(())
}

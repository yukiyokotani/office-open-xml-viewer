//! Native-only evidence for XLS drawing integration; not a worksheet renderer.
use std::io::{Read, Write};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let source = args
        .next()
        .ok_or("expected XLS input and optional fresh output directory")?;
    let directory = args.next();
    if args.next().is_some() {
        return Err("too many arguments".into());
    }
    let mut bytes = Vec::new();
    std::fs::File::open(source)?
        .take(256 * 1024 * 1024 + 1)
        .read_to_end(&mut bytes)?;
    let images = legacy_office_converter::inspect_xls_images(&bytes)?;
    // Never overwrite an existing folder or source/sample artifact. Passing no
    // directory prints metadata only; catalog membership is not visibility.
    if let Some(directory) = &directory {
        std::fs::create_dir(directory)?;
    }
    println!("index,extension,bytes");
    for (index, extension, data) in images {
        if let Some(directory) = &directory {
            let path = std::path::Path::new(directory).join(format!("image{index}.{extension}"));
            std::fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(path)?
                .write_all(&data)?;
        }
        println!("{index},{extension},{}", data.len());
    }
    Ok(())
}

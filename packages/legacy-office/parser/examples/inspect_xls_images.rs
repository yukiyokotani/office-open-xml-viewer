//! Native-only evidence for XLS drawing integration; not a worksheet renderer.
//!
//! Extracts the supported passive PNG, JPEG, EMF and WMF entries of a BIFF8
//! global image store (MS-XLS 2.4.58/171, MS-ODRAW 2.2.12/20/22) without font
//! metrics:
//!
//! ```sh
//! cargo run -p legacy-office-parser --features inspection \
//!   --example inspect_xls_images -- [--used] sample.xls [fresh-output-directory]
//! ```
//!
//! Without a directory it prints catalog indices, formats and byte counts.
//! The directory must not exist yet. `--used` keeps only images referenced by
//! explicitly owned plain picture objects (`inspect_xls_pictures`). A catalog
//! entry is not proof that an image is displayed. Extracted images can hold
//! private document content and must stay local. The inspection APIs are
//! available only on native targets.
mod support;
use std::io::Write;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let first = args
        .next()
        .ok_or("expected XLS input and optional fresh output directory")?;
    let used_only = first == "--used";
    let source = if used_only {
        args.next().ok_or("expected XLS input after --used")?
    } else {
        first
    };
    let directory = args.next();
    if args.next().is_some() {
        return Err("too many arguments".into());
    }
    let bytes = support::read_source(std::fs::File::open(source)?)?;
    let images = if used_only {
        let result = legacy_office_parser::inspect_xls_pictures(&bytes)?;
        eprintln!("passive-picture-anchors:{}", result.anchors.len());
        result.images
    } else {
        legacy_office_parser::inspect_xls_images(&bytes)?
    };
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

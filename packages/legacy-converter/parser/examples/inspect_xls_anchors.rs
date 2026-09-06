//! Passive native development metadata, not a list of renderable images.
use std::io::Read;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let source = args.next().ok_or("expected one XLS input")?;
    if args.next().is_some() {
        return Err("too many arguments".into());
    }
    let mut bytes = Vec::new();
    std::fs::File::open(source)?
        .take(256 * 1024 * 1024 + 1)
        .read_to_end(&mut bytes)?;
    let anchors = legacy_office_converter::inspect_xls_anchors(&bytes)?;
    println!("sheet,shape,shapeFlags,object,objectType,objectFlags,groupDepth,behavior,colL,rowT,dxL,dyT,colR,rowB,dxR,dyB,pictureIndex,cropTop,cropBottom,cropLeft,cropRight,rotation,clipboardFormat,autoPicture");
    for a in anchors {
        print!(
            "{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}",
            a.sheet,
            a.shape_id,
            a.shape_flags,
            a.object_id,
            a.object_type,
            a.object_flags,
            a.group_depth,
            a.behavior,
            a.from.column,
            a.from.row,
            a.from.dx,
            a.from.dy,
            a.to.column,
            a.to.row,
            a.to.dx,
            a.to.dy
        );
        let p = a
            .picture
            .unwrap_or(legacy_office_converter::PictureReference {
                store_index: 0,
                crop: [0; 4],
                rotation: 0,
                clipboard_format: 0,
                auto_picture: false,
            });
        println!(
            ",{},{},{},{},{},{},{},{}",
            p.store_index,
            p.crop[0],
            p.crop[1],
            p.crop[2],
            p.crop[3],
            p.rotation,
            p.clipboard_format,
            u8::from(p.auto_picture)
        );
    }
    Ok(())
}

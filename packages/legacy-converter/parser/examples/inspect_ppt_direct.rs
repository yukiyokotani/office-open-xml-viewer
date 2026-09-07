//! Export the native PPT renderer model and only its admitted passive resources.
use pptx_model::{Fill, Slide, SlideElement};
use serde_json::{json, Value};
use std::collections::BTreeSet;
use std::io::{Read, Write};

fn fill_key(fill: &Fill, output: &mut BTreeSet<String>) {
    if let Fill::Image { image_path, .. } = fill {
        output.insert(image_path.clone());
    }
}

fn resource_keys(slides: &[Slide]) -> BTreeSet<String> {
    let mut output = BTreeSet::new();
    for slide in slides {
        if let Some(fill) = &slide.background {
            fill_key(fill, &mut output);
        }
        for element in &slide.elements {
            match element {
                SlideElement::Shape(shape) => {
                    if let Some(fill) = &shape.fill {
                        fill_key(fill, &mut output);
                    }
                }
                SlideElement::Picture(picture) => {
                    output.insert(picture.image_path.clone());
                }
                _ => {}
            }
        }
    }
    output
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let source = args
        .next()
        .ok_or("expected PPT input and fresh output directory")?;
    let output = args
        .next()
        .ok_or("expected PPT input and fresh output directory")?;
    if args.next().is_some() {
        return Err("too many arguments".into());
    }
    let mut bytes = Vec::new();
    std::fs::File::open(source)?
        .take(256 * 1024 * 1024 + 1)
        .read_to_end(&mut bytes)?;
    if bytes.len() > 256 * 1024 * 1024 {
        return Err("PPT inspection source byte budget exceeded".into());
    }
    std::fs::create_dir(&output)?;

    let mut session = legacy_office_converter::inspect_ppt_direct(&bytes)?;
    let (width, height) = session.size();
    let mut slides = Vec::new();
    for index in 0..session.slide_count() {
        slides.push(session.slide(index)?);
    }
    let slide_values = serde_json::to_value(&slides)?;
    let keys = resource_keys(&slides);
    let mut manifest = serde_json::Map::new();
    for (ordinal, key) in keys.iter().enumerate() {
        let (extension, data) = session.resource(key)?;
        let filename = format!("resource-{}.{}", ordinal + 1, extension);
        std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(std::path::Path::new(&output).join(&filename))?
            .write_all(data)?;
        manifest.insert(key.clone(), Value::String(filename));
    }
    let model = json!({
        "slideWidth": width,
        "slideHeight": height,
        "slides": slide_values,
        "resources": manifest,
    });
    std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(std::path::Path::new(&output).join("model.json"))?
        .write_all(serde_json::to_vec_pretty(&model)?.as_slice())?;
    println!("slides:{} resources:{}", slides.len(), keys.len());
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pptx_model::PptxComment;

    fn image(path: &str) -> Fill {
        Fill::Image {
            image_path: path.to_owned(),
            mime_type: "image/png".to_owned(),
            svg_image_path: None,
            dpi: None,
            rot_with_shape: None,
            src_rect: None,
            fill_rect: None,
            stretch: true,
            tile: None,
            alpha: None,
            duotone: None,
        }
    }

    #[test]
    fn typed_collection_ignores_resource_looking_text_and_keeps_image_fill() {
        let slide = Slide {
            index: 0,
            slide_number: 1,
            part_name: None,
            background: Some(image("legacy-ppt/image/7")),
            elements: Vec::new(),
            element_sources: Vec::new(),
            notes: Some("legacy-ppt/image/8".to_owned()),
            comments: vec![PptxComment {
                author_id: None,
                modern_author_id: None,
                id: None,
                index: None,
                author: None,
                date: None,
                x: None,
                y: None,
                anchors: Vec::new(),
                status: None,
                text: "legacy-ppt/image/9".to_owned(),
                replies: Vec::new(),
            }],
            hidden: false,
            parse_error: None,
        };
        assert_eq!(
            resource_keys(&[slide]),
            BTreeSet::from(["legacy-ppt/image/7".to_owned()])
        );
    }
}

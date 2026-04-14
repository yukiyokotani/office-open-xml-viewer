fn main() {
    let path = std::env::args().nth(1)
        .unwrap_or_else(|| "/Users/yokotani/開発/office-open-xml-viewer/public/sample.pptx".into());
    let data = std::fs::read(&path).expect("read pptx");
    
    // Call the library's internal logic via the public WASM function
    // (We can't call parse_pptx directly since it's wasm_bindgen, but we can 
    //  replicate the core logic inline here for testing)
    
    use std::io::{Cursor, Read};
    let cursor = Cursor::new(data.as_slice());
    let mut zip = zip::ZipArchive::new(cursor).unwrap();
    
    let pres_xml = {
        let mut f = zip.by_name("ppt/presentation.xml").unwrap();
        let mut s = String::new(); f.read_to_string(&mut s).unwrap(); s
    };
    let pres_doc = roxmltree::Document::parse(&pres_xml).unwrap();
    let root = pres_doc.root_element();
    
    // Find sldIdLst
    let sld_id_lst = root.descendants().find(|n| n.is_element() && n.tag_name().name() == "sldIdLst");
    println!("sldIdLst found: {}", sld_id_lst.is_some());
    
    if let Some(lst) = sld_id_lst {
        for sld in lst.children().filter(|n| n.is_element()) {
            println!("  sldId attrs:");
            for a in sld.attributes() {
                println!("    name={} ns={:?} val={}", a.name(), a.namespace(), a.value());
            }
        }
    }
    
    let pres_rels = {
        let mut f = zip.by_name("ppt/_rels/presentation.xml.rels").unwrap();
        let mut s = String::new(); f.read_to_string(&mut s).unwrap(); s
    };
    let rels_doc = roxmltree::Document::parse(&pres_rels).unwrap();
    println!("\nPresentation rels:");
    for rel in rels_doc.root_element().children().filter(|n| n.is_element()) {
        let id = rel.attributes().find(|a| a.name() == "Id" && a.namespace().is_none()).map(|a| a.value());
        let target = rel.attributes().find(|a| a.name() == "Target" && a.namespace().is_none()).map(|a| a.value());
        let typ = rel.attributes().find(|a| a.name() == "Type" && a.namespace().is_none()).map(|a| a.value());
        if typ.map(|t| t.contains("/slide\"") || t.ends_with("/slide")).unwrap_or(false) {
            println!("  id={:?} target={:?}", id, target);
        }
    }
}

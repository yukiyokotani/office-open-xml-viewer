use std::collections::HashMap;
use roxmltree::Document as XmlDoc;
use crate::xml_util::*;

#[derive(Debug, Clone)]
pub struct LevelDef {
    pub format: String,  // "decimal" | "bullet" | etc.
    pub text: String,    // lvlText val, e.g. "%1." or "•"
    pub indent_left: f64,   // pt
    pub tab: f64,            // pt
    pub start: u32,
}

impl Default for LevelDef {
    fn default() -> Self {
        LevelDef {
            format: "decimal".to_string(),
            text: "%1.".to_string(),
            indent_left: 36.0,
            tab: 36.0,
            start: 1,
        }
    }
}

#[derive(Default)]
pub struct NumberingMap {
    /// abstractNumId → [level0..level8]
    abstract_nums: HashMap<u32, Vec<LevelDef>>,
    /// numId → abstractNumId
    num_to_abstract: HashMap<u32, u32>,
    /// numId → level override starts
    num_overrides: HashMap<u32, HashMap<u32, u32>>,
    /// per-numId per-level counter
    pub counters: HashMap<u32, HashMap<u32, u32>>,
}

impl NumberingMap {
    pub fn parse(xml: &str) -> Self {
        let mut map = NumberingMap::default();
        let doc = match XmlDoc::parse(xml) {
            Ok(d) => d,
            Err(_) => return map,
        };
        let root = doc.root_element();

        // Parse abstractNum definitions
        for abs_node in children_w(root, "abstractNum") {
            let Some(abs_id_s) = attr_w(abs_node, "abstractNumId") else { continue };
            let abs_id: u32 = abs_id_s.parse().unwrap_or(0);
            let mut levels = vec![];
            for lvl_node in children_w(abs_node, "lvl") {
                let start = child_w(lvl_node, "start")
                    .and_then(|n| attr_w(n, "val"))
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(1);
                let format = child_w(lvl_node, "numFmt")
                    .and_then(|n| attr_w(n, "val"))
                    .unwrap_or_else(|| "decimal".to_string());
                let text = child_w(lvl_node, "lvlText")
                    .and_then(|n| attr_w(n, "val"))
                    .unwrap_or_else(|| "%1.".to_string());
                let indent_left = child_w(lvl_node, "pPr")
                    .and_then(|p| child_w(p, "ind"))
                    .and_then(|i| attr_w(i, "left"))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(720.0 / 20.0 * (levels.len() as f64 + 1.0));
                let tab = child_w(lvl_node, "pPr")
                    .and_then(|p| child_w(p, "ind"))
                    .and_then(|i| attr_w(i, "hanging").or_else(|| attr_w(i, "firstLine")))
                    .map(|v| twips_to_pt(&v))
                    .unwrap_or(36.0);
                levels.push(LevelDef { format, text, indent_left, tab, start });
            }
            map.abstract_nums.insert(abs_id, levels);
        }

        // Parse num → abstractNum
        for num_node in children_w(root, "num") {
            let Some(num_id_s) = attr_w(num_node, "numId") else { continue };
            let num_id: u32 = num_id_s.parse().unwrap_or(0);
            if let Some(abs_ref) = child_w(num_node, "abstractNumId").and_then(|n| attr_w(n, "val")) {
                let abs_id: u32 = abs_ref.parse().unwrap_or(0);
                map.num_to_abstract.insert(num_id, abs_id);
            }
            // Level overrides
            let mut overrides = HashMap::new();
            for lvl_ov in children_w(num_node, "lvlOverride") {
                let ilvl: u32 = attr_w(lvl_ov, "ilvl").and_then(|v| v.parse().ok()).unwrap_or(0);
                if let Some(start_ov) = child_w(lvl_ov, "startOverride").and_then(|n| attr_w(n, "val")) {
                    overrides.insert(ilvl, start_ov.parse().unwrap_or(1));
                }
            }
            if !overrides.is_empty() {
                map.num_overrides.insert(num_id, overrides);
            }
        }

        map
    }

    pub fn get_level(&self, num_id: u32, level: u32) -> Option<&LevelDef> {
        let abs_id = self.num_to_abstract.get(&num_id)?;
        let levels = self.abstract_nums.get(abs_id)?;
        levels.get(level as usize)
    }

    pub fn get_start(&self, num_id: u32, level: u32) -> u32 {
        if let Some(ov) = self.num_overrides.get(&num_id).and_then(|m| m.get(&level)) {
            return *ov;
        }
        self.get_level(num_id, level).map(|l| l.start).unwrap_or(1)
    }

    /// Advance counter for (numId, level), resetting deeper levels.
    pub fn advance(&mut self, num_id: u32, level: u32) -> u32 {
        // Pre-compute start values to avoid borrow conflicts
        let starts: Vec<u32> = (0..=level).map(|l| self.get_start(num_id, l)).collect();

        let entry = self.counters.entry(num_id).or_default();

        // Reset deeper levels
        let keys: Vec<u32> = entry.keys().copied().filter(|&l| l > level).collect();
        for k in keys { entry.remove(&k); }

        // Ensure levels from 0 to level-1 are initialized
        for (lvl, &start) in starts.iter().enumerate().take(level as usize) {
            entry.entry(lvl as u32).or_insert(start);
        }

        let current = entry.entry(level).or_insert(starts[level as usize]);
        let val = *current;
        *current = val + 1;
        val
    }

    /// Resolve the display text for a counter value in the given level.
    pub fn resolve_text(&self, num_id: u32, level: u32, counter: u32) -> String {
        let Some(lvl) = self.get_level(num_id, level) else { return format!("{}.", counter) };

        let formatted = format_counter(counter, &lvl.format);

        // Replace %N placeholders in lvlText
        let mut text = lvl.text.clone();
        // Only replace the placeholder for current level (simplified)
        let placeholder = format!("%{}", level + 1);
        text = text.replace(&placeholder, &formatted);
        text
    }
}

fn format_counter(n: u32, format: &str) -> String {
    match format {
        "decimal" => n.to_string(),
        "bullet" => "•".to_string(),
        "lowerLetter" => {
            let c = (b'a' + ((n - 1) % 26) as u8) as char;
            c.to_string()
        }
        "upperLetter" => {
            let c = (b'A' + ((n - 1) % 26) as u8) as char;
            c.to_string()
        }
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        _ => n.to_string(),
    }
}

fn to_roman(n: u32) -> String {
    let vals = [(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),(100,"C"),(90,"XC"),
                (50,"L"),(40,"XL"),(10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")];
    let mut n = n;
    let mut s = String::new();
    for (v, r) in &vals {
        while n >= *v { s.push_str(r); n -= v; }
    }
    s
}

//! Relative IRI resolution inside a package (ODF Part 2 §3.7).

/// Where a reference resolves.
#[derive(Debug, PartialEq, Eq)]
pub enum Target {
    /// A file or directory name inside the package ("" is the package root),
    /// plus an optional fragment.
    Package {
        path: String,
        fragment: Option<String>,
    },
    /// The same document (§3.7 step 4: an empty relative path).
    SameDocument { fragment: Option<String> },
    /// Resolved against the package base IRI (§3.7 steps 3, 9, 10): outside
    /// the package. Readers never fetch it.
    External,
}

/// Resolve `reference`, found in the package file `file_entry_path`.
/// Structure (`./`, `.`, `..`) is processed on the undecoded IRI text exactly
/// as §3.7 steps 5–11 describe; only afterwards are the IRI-origin segments
/// percent-decoded once (step 12), while the copied file-entry prefix keeps
/// ZIP path syntax. A query component, an encoded `/`, `\` or NUL, and
/// invalid UTF-8 are rejected (`None`): such names cannot address a package
/// file unambiguously.
pub fn resolve(file_entry_path: &str, reference: &str) -> Option<Target> {
    let (without_fragment, fragment) = match reference.split_once('#') {
        Some((head, fragment)) => (head, Some(fragment.to_string())),
        None => (reference, None),
    };
    if without_fragment.contains('?') {
        return None;
    }
    // §3.7 step 3: an IRI with a scheme (not a relative-ref) or an absolute
    // path resolves against the package base IRI.
    if without_fragment.starts_with('/') || has_scheme(without_fragment) {
        return Some(Target::External);
    }
    if without_fragment.starts_with("//") {
        return Some(Target::External);
    }
    // §3.7 step 4.
    if without_fragment.is_empty() {
        return Some(Target::SameDocument { fragment });
    }
    // Step 5/6: the IRI buffer and the file-entry directory prefix.
    let mut buffer = without_fragment.to_string();
    let mut prefix: Vec<String> = match file_entry_path.rfind('/') {
        Some(index) => file_entry_path[..index]
            .split('/')
            .map(str::to_string)
            .collect(),
        None => Vec::new(),
    };
    loop {
        // Step 7.
        if let Some(rest) = buffer.strip_prefix("./") {
            buffer = rest.to_string();
            continue;
        }
        // Step 8.
        if buffer == "." {
            buffer.clear();
            break;
        }
        // Step 9: "..".
        if buffer == ".." {
            if prefix.pop().is_none() {
                return Some(Target::External);
            }
            buffer.clear();
            break;
        }
        // Step 10: "../".
        if let Some(rest) = buffer.strip_prefix("../") {
            if prefix.pop().is_none() {
                return Some(Target::External);
            }
            buffer = rest.to_string();
            continue;
        }
        break;
    }
    // Step 11/12: prefix segments (ZIP syntax, literal) + IRI segments
    // (decoded once).
    let mut segments = prefix;
    if !buffer.is_empty() {
        for segment in buffer.split('/') {
            segments.push(percent_decode(segment)?);
        }
    }
    Some(Target::Package {
        path: segments.join("/"),
        fragment,
    })
}

fn has_scheme(text: &str) -> bool {
    let Some(colon) = text.find(':') else {
        return false;
    };
    let scheme = &text[..colon];
    let first_segment_end = text.find('/').unwrap_or(text.len());
    colon < first_segment_end
        && scheme.starts_with(|c: char| c.is_ascii_alphabetic())
        && scheme
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '.'))
}

fn percent_decode(segment: &str) -> Option<String> {
    let bytes = segment.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let hex = segment.get(index + 1..index + 3)?;
            let value = u8::from_str_radix(hex, 16).ok()?;
            if matches!(value, b'/' | b'\\' | 0) {
                return None;
            }
            out.push(value);
            index += 3;
        } else {
            out.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(out).ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn package(path: &str) -> Option<Target> {
        Some(Target::Package {
            path: path.into(),
            fragment: None,
        })
    }

    #[test]
    fn resolves_relative_references_from_the_containing_file() {
        assert_eq!(
            resolve("content.xml", "Pictures/a.png"),
            package("Pictures/a.png")
        );
        assert_eq!(resolve("content.xml", "./Object 1"), package("Object 1"));
        assert_eq!(
            resolve("Object 1/content.xml", "../Pictures/a.png"),
            package("Pictures/a.png")
        );
        assert_eq!(resolve("Object 1/content.xml", "."), package("Object 1"));
        assert_eq!(resolve("Object 1/content.xml", ".."), package(""));
        assert_eq!(resolve("content.xml", "a%20b.png"), package("a b.png"));
    }

    #[test]
    fn decodes_after_structure_so_encoded_dots_are_names() {
        assert_eq!(resolve("content.xml", "%2E%2E/x"), package("../x"));
        assert_eq!(resolve("content.xml", "a%2Fb"), None);
        assert_eq!(resolve("content.xml", "a%5Cb"), None);
        assert_eq!(resolve("content.xml", "a%00b"), None);
        assert_eq!(resolve("content.xml", "a%FF"), None);
    }

    #[test]
    fn escaping_the_root_absolute_and_schemed_references_are_external() {
        assert_eq!(resolve("content.xml", "../x.png"), Some(Target::External));
        assert_eq!(resolve("content.xml", ".."), Some(Target::External));
        assert_eq!(resolve("content.xml", "/x.png"), Some(Target::External));
        assert_eq!(
            resolve("content.xml", "https://example.test/x.png"),
            Some(Target::External)
        );
        assert_eq!(resolve("content.xml", "//host/x"), Some(Target::External));
        assert_eq!(resolve("content.xml", "x.png?q=1"), None);
        assert_eq!(
            resolve("content.xml", "#Table1"),
            Some(Target::SameDocument {
                fragment: Some("Table1".into())
            })
        );
    }
}

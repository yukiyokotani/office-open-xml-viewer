pub use pptx_model::*;

#[cfg(test)]
mod group_transform_tests {
    use super::*;

    #[test]
    fn quarter_turned_child_uses_shared_non_uniform_group_mapping() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 127_000,
            cy: 254_000,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 127_000,
            ch_cy: 127_000,
            ..Default::default()
        };
        let mapped = group.apply_to_transform(Transform {
            x: 0,
            y: 50_800,
            cx: 127_000,
            cy: 25_400,
            rot: 90.0,
            ..Default::default()
        });

        assert_eq!((mapped.x, mapped.y), (0, 101_600));
        assert_eq!((mapped.cx, mapped.cy), (127_000, 50_800));
        assert_eq!(mapped.rot, 90.0);
    }

    #[test]
    fn rotated_flipped_group_uses_shared_annex_l_mapping() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 200,
            cy: 100,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 100,
            ch_cy: 100,
            rot: 90.0,
            flip_h: true,
            flip_v: false,
        };
        let mapped = group.apply_to_transform(Transform {
            x: 10,
            y: 20,
            cx: 20,
            cy: 10,
            rot: 15.0,
            flip_h: false,
            flip_v: true,
        });

        assert_eq!(
            (mapped.x, mapped.y, mapped.cx, mapped.cy),
            (105, 105, 40, 10)
        );
        assert_eq!(mapped.rot, 75.0);
        assert!(mapped.flip_h);
        assert!(mapped.flip_v);
    }

    #[test]
    fn group_transform_preserves_table_chart_and_media_frame_orientation() {
        let group = GroupTransform {
            x: 0,
            y: 0,
            cx: 200,
            cy: 100,
            ch_x: 0,
            ch_y: 0,
            ch_cx: 100,
            ch_cy: 100,
            rot: 90.0,
            flip_h: true,
            flip_v: false,
        };

        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let chart = crate::chart::parse_legacy_chart(xml, &std::collections::HashMap::new())
            .expect("minimal chart");
        let mut elements = vec![
            SlideElement::Table(TableElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                cols: vec![],
                rows: vec![],
                rtl: false,
            }),
            SlideElement::Chart(ChartElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                ..chart
            }),
            SlideElement::Media(MediaElement {
                id: None,
                x: 10,
                y: 20,
                width: 20,
                height: 10,
                rotation: 15.0,
                flip_h: false,
                flip_v: true,
                media_kind: "video".into(),
                poster_path: String::new(),
                poster_mime_type: String::new(),
                media_path: String::new(),
                mime_type: "video/mp4".into(),
            }),
        ];

        for element in &mut elements {
            apply_group_transform_to_element(element, &group);
            match element {
                SlideElement::Table(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                SlideElement::Chart(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                SlideElement::Media(frame) => {
                    assert_eq!(frame.rotation, 75.0);
                    assert!(frame.flip_h);
                    assert!(frame.flip_v);
                }
                _ => unreachable!(),
            }
        }
    }
}

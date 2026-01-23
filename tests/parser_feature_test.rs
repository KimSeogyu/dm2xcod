use rs_docx::document::{Inline, Pict, Run, RunContent};
use rs_docx::formatting::{GridSpan, VMerge, VMergeType};
use hard_xml::XmlRead;

#[test]
fn test_vml_pict_parsing() {
    let xml = r#"<w:pict><v:shape id="Shape1" style="width:100;height:100"><v:imagedata r:id="rId1" o:title="Test Image"/></v:shape></w:pict>"#;
    let pict = Pict::from_str(xml).expect("Failed to parse Pict");

    let shape = pict.shape.expect("Shape not found");
    assert_eq!(shape.id.as_deref(), Some("Shape1"));

    let image_data = shape.image_data.expect("ImageData not found");
    assert_eq!(image_data.id.as_deref(), Some("rId1"));
    assert_eq!(image_data.title.as_deref(), Some("Test Image"));
}

#[test]
fn test_table_merge_parsing() {
    let xml_vmerge = r#"<w:vMerge w:val="restart"/>"#;
    let vmerge = VMerge::from_str(xml_vmerge).expect("Failed to parse VMerge");
    assert_eq!(vmerge.val, Some(VMergeType::Restart));

    let xml_gridspan = r#"<w:gridSpan w:val="3"/>"#;
    let gridspan = GridSpan::from_str(xml_gridspan).expect("Failed to parse GridSpan");
    assert_eq!(gridspan.val, 3);
}

#[test]
fn test_run_with_pict() {
    let xml = r#"<w:r><w:pict><v:shape id="S1"/></w:pict></w:r>"#;
    let run = Run::from_str(xml).expect("Failed to parse Run with Pict");

    let has_pict = run.content.iter().any(|c| matches!(c, RunContent::Pict(_)));
    assert!(has_pict, "Run should contain Pict");
}

#[test]
fn test_drawing_inline_extension() {
    // wp14:anchorId and editId
    let xml = r#"<wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="12345678" wp14:editId="87654321"><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="X"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="uri"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name=""/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill><pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>"#;

    // Note: XmlRead usually ignores unknown attributes unless strictly defined?
    // hard-xml (our parser) typically ignores unless defined.
    // We defined anchorId/editId, so it should parse them (or at least not fail, and hopefully capture them).

    let inline = Inline::from_str(xml).expect("Failed to parse Inline");

    // Note: namespaces in attributes might be tricky in hard-xml if not fully handled,
    // but our struct definition is #[xml(attr = "wp14:anchorId")].
    assert_eq!(inline.anchor_id.as_deref(), Some("12345678"));
    assert_eq!(inline.edit_id.as_deref(), Some("87654321"));
}

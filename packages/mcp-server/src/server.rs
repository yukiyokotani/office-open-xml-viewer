use rmcp::{
    ServerHandler,
    handler::server::router::tool::ToolRouter,
    handler::server::wrapper::Parameters,
    model::{ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};

use crate::tools::{
    docx::DocxTools,
    pptx::PptxTools,
    xlsx::XlsxTools,
};
use crate::tools::docx::DocxPathParam;
use crate::tools::pptx::{PptxPathParam, PptxSlideParam, PptxTextParam};
use crate::tools::xlsx::{XlsxCellRangeParam, XlsxPathParam, XlsxSheetParam};

#[derive(Clone)]
pub struct OoxmlServer {
    #[allow(dead_code)]
    tool_router: ToolRouter<OoxmlServer>,
}

#[tool_router]
impl OoxmlServer {
    pub fn new() -> Self {
        Self {
            tool_router: Self::tool_router(),
        }
    }

    // ── xlsx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Parse an XLSX file and return workbook overview including sheet names and IDs")]
    fn xlsx_parse(&self, Parameters(p): Parameters<XlsxPathParam>) -> String {
        XlsxTools::xlsx_parse(Parameters(p))
    }

    #[tool(description = "Return only the list of sheet names from an XLSX file")]
    fn xlsx_get_sheet_names(&self, Parameters(p): Parameters<XlsxPathParam>) -> String {
        XlsxTools::xlsx_get_sheet_names(Parameters(p))
    }

    #[tool(description = "Return the dimensions (max row and column) of a worksheet")]
    fn xlsx_get_sheet_dimensions(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_sheet_dimensions(Parameters(p))
    }

    #[tool(description = "Return cell values and formulas for a given range (e.g. \"A1:C10\") in a worksheet")]
    fn xlsx_get_cell_range(&self, Parameters(p): Parameters<XlsxCellRangeParam>) -> String {
        XlsxTools::xlsx_get_cell_range(Parameters(p))
    }

    #[tool(description = "Return all cells that contain formulas in a worksheet")]
    fn xlsx_get_formulas(&self, Parameters(p): Parameters<XlsxSheetParam>) -> String {
        XlsxTools::xlsx_get_formulas(Parameters(p))
    }

    // ── docx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Extract all plain text from a DOCX file")]
    fn docx_extract_text(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_extract_text(Parameters(p))
    }

    #[tool(description = "Return the document structure (paragraphs and tables) of a DOCX file")]
    fn docx_get_structure(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_structure(Parameters(p))
    }

    #[tool(description = "Return all tables from a DOCX file with their cell contents")]
    fn docx_get_tables(&self, Parameters(p): Parameters<DocxPathParam>) -> String {
        DocxTools::docx_get_tables(Parameters(p))
    }

    // ── pptx tools ────────────────────────────────────────────────────────────

    #[tool(description = "Return the number of slides and each slide's title from a PPTX file")]
    fn pptx_get_slides(&self, Parameters(p): Parameters<PptxPathParam>) -> String {
        PptxTools::pptx_get_slides(Parameters(p))
    }

    #[tool(description = "Extract plain text from a PPTX file; optionally filter to a single slide by 0-based index")]
    fn pptx_extract_text(&self, Parameters(p): Parameters<PptxTextParam>) -> String {
        PptxTools::pptx_extract_text(Parameters(p))
    }

    #[tool(description = "Return the structure (elements with position, size, text) of a single slide")]
    fn pptx_get_slide_structure(&self, Parameters(p): Parameters<PptxSlideParam>) -> String {
        PptxTools::pptx_get_slide_structure(Parameters(p))
    }
}

#[tool_handler]
impl ServerHandler for OoxmlServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(
            ServerCapabilities::builder()
                .enable_tools()
                .build(),
        )
    }
}

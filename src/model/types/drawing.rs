//! DrawingML types — images, pictures, shapes, anchoring, and preset geometry.

use crate::model::dimension::{Dimension, Emu, SixtieThousandthDeg, ThousandthPercent};
use crate::model::geometry::{EdgeInsets, Offset, Size};

use super::content::Block;
use super::identifiers::RelId;

/// Format of an embedded image, detected from the OOXML relationship target path
/// (§M.1.1) with magic-byte fallback.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    /// Windows Enhanced Metafile (MS-EMF).
    Emf,
    /// Windows Metafile (MS-WMF).
    Wmf,
    Svg,
    WebP,
    Unknown,
}

impl ImageFormat {
    /// Detect format from the relationship target file extension (OOXML §M.1.1)
    /// with magic-byte fallback for unrecognised or missing extensions.
    pub fn detect(target_path: &str, data: &[u8]) -> Self {
        let ext = target_path
            .rsplit('.')
            .next()
            .map(|e| e.to_ascii_lowercase());
        match ext.as_deref() {
            Some("png") => Self::Png,
            Some("jpg" | "jpeg") => Self::Jpeg,
            Some("gif") => Self::Gif,
            Some("bmp") => Self::Bmp,
            Some("tif" | "tiff") => Self::Tiff,
            Some("emf") => Self::Emf,
            Some("wmf") => Self::Wmf,
            Some("svg" | "svgz") => Self::Svg,
            Some("webp") => Self::WebP,
            _ => Self::detect_by_magic(data),
        }
    }

    /// Magic-byte detection for the most common raster formats.
    fn detect_by_magic(data: &[u8]) -> Self {
        match data {
            [0x89, b'P', b'N', b'G', ..] => Self::Png,
            [0xFF, 0xD8, 0xFF, ..] => Self::Jpeg,
            [b'G', b'I', b'F', b'8', ..] => Self::Gif,
            [b'B', b'M', ..] => Self::Bmp,
            // EMF header: RecordType=0x00000001 followed by size, then EMF signature 0x464D4520
            [0x01, 0x00, 0x00, 0x00, _, _, _, _, 0x20, 0x45, 0x4D, 0x46, ..] => Self::Emf,
            [b'R', b'I', b'F', b'F', _, _, _, _, b'W', b'E', b'B', b'P', ..] => Self::WebP,
            [b'<', ..] => Self::Svg,
            _ => Self::Unknown,
        }
    }
}

#[derive(Clone, Debug)]
pub struct Image {
    /// §20.4.2.7: drawing extent.
    pub extent: Size<Emu>,
    /// §20.4.2.6: additional extent for effects.
    pub effect_extent: Option<EdgeInsets<Emu>>,
    /// §20.1.2.2.8: non-visual drawing properties (wp:docPr).
    pub doc_properties: DocProperties,
    /// §20.4.2.4: graphic frame locking properties.
    pub graphic_frame_locks: Option<GraphicFrameLocks>,
    /// Graphic content from a:graphic > a:graphicData.
    pub graphic: Option<GraphicContent>,
    /// Inline or anchor placement.
    pub placement: ImagePlacement,
}

/// How the image is placed in the document flow.
#[derive(Clone, Debug)]
pub enum ImagePlacement {
    /// §20.4.2.8: inline with text — no wrapping.
    Inline {
        /// Distance from surrounding text.
        distance: EdgeInsets<Emu>,
    },
    /// §20.4.2.3: floating/anchored with text wrapping.
    Anchor(AnchorProperties),
}

/// §20.1.2.2.8 CT_NonVisualDrawingProps — shared by wp:docPr and pic:cNvPr.
#[derive(Clone, Debug)]
pub struct DocProperties {
    /// Unique identifier.
    pub id: u32,
    /// Element name.
    pub name: String,
    /// Alternative text description.
    pub description: Option<String>,
    /// Whether the element is hidden.
    pub hidden: Option<bool>,
    /// Title (Office 2010+).
    pub title: Option<String>,
}

/// §20.1.2.2.19 CT_GraphicalObjectFrameLocking.
#[derive(Clone, Copy, Debug)]
pub struct GraphicFrameLocks {
    pub no_change_aspect: Option<bool>,
    pub no_drilldown: Option<bool>,
    pub no_grp: Option<bool>,
    pub no_move: Option<bool>,
    pub no_resize: Option<bool>,
    pub no_select: Option<bool>,
}

/// Content type inside a:graphicData.
#[derive(Clone, Debug)]
pub enum GraphicContent {
    /// §19.3.1.37: picture.
    Picture(Picture),
    /// §14.5 wps:wsp: Word Processing Shape.
    WordProcessingShape(WordProcessingShape),
}

/// §14.5 wps:wsp — a Word Processing Shape.
/// Contains shape properties and optional text body.
#[derive(Clone, Debug)]
pub struct WordProcessingShape {
    /// §20.1.2.2.8: non-visual drawing properties.
    pub cnv_pr: Option<DocProperties>,
    /// §20.1.2.2.35: shape properties.
    pub shape_properties: Option<ShapeProperties>,
    /// §20.1.2.1.1: body properties (text layout within the shape).
    pub body_pr: Option<BodyProperties>,
    /// §17.17.1: text content inside the shape.
    pub txbx_content: Vec<Block>,
}

/// §20.1.2.1.1 CT_TextBodyProperties — text body properties inside a shape.
#[derive(Clone, Debug)]
pub struct BodyProperties {
    /// Rotation of the text body in 60,000ths of a degree.
    pub rotation: Option<Dimension<SixtieThousandthDeg>>,
    /// §20.1.10.82 ST_TextVerticalType: vertical text mode.
    pub vert: Option<TextVerticalType>,
    /// §20.1.10.85 ST_TextWrappingType: text wrapping within the shape.
    pub wrap: Option<TextWrappingType>,
    /// Left inset in EMU.
    pub left_inset: Option<Dimension<Emu>>,
    /// Top inset in EMU.
    pub top_inset: Option<Dimension<Emu>>,
    /// Right inset in EMU.
    pub right_inset: Option<Dimension<Emu>>,
    /// Bottom inset in EMU.
    pub bottom_inset: Option<Dimension<Emu>>,
    /// §20.1.10.59 ST_TextAnchoringType: vertical anchor.
    pub anchor: Option<TextAnchoringType>,
    /// Auto-fit mode.
    pub auto_fit: Option<TextAutoFit>,
}

/// §20.1.10.82 ST_TextVerticalType.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextVerticalType {
    Horz,
    Vert,
    Vert270,
    WordArtVert,
    EaVert,
    MongolianVert,
    WordArtVertRtl,
}

/// §20.1.10.85 ST_TextWrappingType.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextWrappingType {
    None,
    Square,
}

/// §20.1.10.59 ST_TextAnchoringType.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAnchoringType {
    Top,
    Center,
    Bottom,
    Justified,
    Distributed,
}

/// Text auto-fit mode for shape text bodies.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextAutoFit {
    /// §20.1.2.1.16: no auto-fit.
    NoAutoFit,
    /// §20.1.2.1.18: shrink text to fit.
    NormalAutoFit,
    /// §20.1.2.1.20: resize shape to fit text.
    SpAutoFit,
}

/// §19.3.1.37 pic:pic — a picture element.
#[derive(Clone, Debug)]
pub struct Picture {
    /// §19.3.1.32: non-visual picture properties.
    pub nv_pic_pr: NvPicProperties,
    /// §20.1.8.14: blip fill (picture data + crop + fill mode).
    pub blip_fill: BlipFill,
    /// §20.1.2.2.35: shape properties (transform, geometry, outline).
    pub shape_properties: Option<ShapeProperties>,
}

/// §19.3.1.32 pic:nvPicPr — non-visual picture properties.
#[derive(Clone, Debug)]
pub struct NvPicProperties {
    /// §20.1.2.2.8 pic:cNvPr.
    pub cnv_pr: DocProperties,
    /// §19.3.1.4 pic:cNvPicPr.
    pub cnv_pic_pr: Option<CnvPicProperties>,
}

/// §19.3.1.4 pic:cNvPicPr — non-visual picture drawing properties.
#[derive(Clone, Debug)]
pub struct CnvPicProperties {
    pub prefer_relative_resize: Option<bool>,
    /// §20.1.2.2.31: picture locking.
    pub pic_locks: Option<PicLocks>,
}

/// §20.1.2.2.31 a:picLocks — picture locking constraints.
#[derive(Clone, Copy, Debug)]
pub struct PicLocks {
    pub no_change_aspect: Option<bool>,
    pub no_crop: Option<bool>,
    pub no_resize: Option<bool>,
    pub no_move: Option<bool>,
    pub no_rot: Option<bool>,
    pub no_select: Option<bool>,
    pub no_edit_points: Option<bool>,
    pub no_adjust_handles: Option<bool>,
    pub no_change_arrowheads: Option<bool>,
    pub no_change_shape_type: Option<bool>,
    pub no_grp: Option<bool>,
}

/// §20.1.8.14 pic:blipFill — picture fill properties.
#[derive(Clone, Debug)]
pub struct BlipFill {
    pub rotate_with_shape: Option<bool>,
    pub dpi: Option<u32>,
    /// §20.1.8.13: blip reference.
    pub blip: Option<Blip>,
    /// §20.1.10.48: source rectangle (crop).
    pub src_rect: Option<RelativeRect>,
    /// §20.1.8.56: stretch fill mode.
    pub stretch: Option<StretchFill>,
}

/// §20.1.8.13 a:blip — reference to image data.
#[derive(Clone, Debug)]
pub struct Blip {
    /// r:embed — relationship ID for embedded image.
    pub embed: Option<RelId>,
    /// r:link — relationship ID for linked (external) image.
    pub link: Option<RelId>,
    /// §20.1.10.7: compression state.
    pub compression: Option<BlipCompression>,
}

/// §20.1.10.7 ST_BlipCompression.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BlipCompression {
    Email,
    Hqprint,
    None,
    Print,
    Screen,
}

/// §20.1.10.48 CT_RelativeRect — relative rectangle (thousandths of percent).
/// Used for a:srcRect and a:fillRect.
#[derive(Clone, Copy, Debug)]
pub struct RelativeRect {
    pub left: Option<Dimension<ThousandthPercent>>,
    pub top: Option<Dimension<ThousandthPercent>>,
    pub right: Option<Dimension<ThousandthPercent>>,
    pub bottom: Option<Dimension<ThousandthPercent>>,
}

/// §20.1.8.56 a:stretch — stretch fill mode.
#[derive(Clone, Copy, Debug)]
pub struct StretchFill {
    /// §20.1.10.48: fill rectangle.
    pub fill_rect: Option<RelativeRect>,
}

/// §20.1.2.2.35 CT_ShapeProperties — shape visual properties.
#[derive(Clone, Debug)]
pub struct ShapeProperties {
    /// §20.1.10.10: black-and-white mode.
    pub bw_mode: Option<BlackWhiteMode>,
    /// §20.1.7.6: 2D transform.
    pub transform: Option<Transform2D>,
    /// §20.1.9.18: preset geometry.
    pub preset_geometry: Option<PresetGeometryDef>,
    /// Fill type (noFill, solidFill, etc.).
    pub fill: Option<DrawingFill>,
    /// §20.1.2.2.24: outline/line properties.
    pub outline: Option<Outline>,
}

/// §20.1.7.6 CT_Transform2D — 2D transform.
#[derive(Clone, Copy, Debug)]
pub struct Transform2D {
    /// Rotation in 60,000ths of a degree (§20.1.10.3).
    pub rotation: Option<Dimension<SixtieThousandthDeg>>,
    pub flip_h: Option<bool>,
    pub flip_v: Option<bool>,
    /// §20.1.7.4: offset (x, y).
    pub offset: Option<Offset<Emu>>,
    /// §20.1.7.3: extent (cx, cy).
    pub extent: Option<Size<Emu>>,
}

/// §20.1.9.18 CT_PresetGeometry2D — preset shape geometry.
#[derive(Clone, Debug)]
pub struct PresetGeometryDef {
    /// §20.1.10.56: preset shape type.
    pub preset: PresetShapeType,
    /// §20.1.9.5: adjustment values.
    pub adjust_values: Vec<GeomGuide>,
}

/// §20.1.9.11 CT_GeomGuide — geometry guide (named formula).
#[derive(Clone, Debug)]
pub struct GeomGuide {
    pub name: String,
    pub formula: String,
}

/// §20.1.10.56 ST_ShapeType — preset shape types (subset).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PresetShapeType {
    Rect,
    RoundRect,
    Ellipse,
    Triangle,
    RtTriangle,
    Diamond,
    Parallelogram,
    Trapezoid,
    Pentagon,
    Hexagon,
    Octagon,
    Star4,
    Star5,
    Star6,
    Star8,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Line,
    Plus,
    Can,
    Cube,
    Donut,
    NoSmoking,
    BlockArc,
    Heart,
    Sun,
    Moon,
    SmileyFace,
    LightningBolt,
    Cloud,
    Arc,
    Plaque,
    Frame,
    Bevel,
    FoldedCorner,
    Chevron,
    HomePlate,
    Ribbon,
    Ribbon2,
    Pie,
    PieWedge,
    Chord,
    Teardrop,
    Arrow,
    LeftArrow,
    RightArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,
    QuadArrow,
    BentArrow,
    UturnArrow,
    CircularArrow,
    CurvedRightArrow,
    CurvedLeftArrow,
    CurvedUpArrow,
    CurvedDownArrow,
    StripedRightArrow,
    NotchedRightArrow,
    BentUpArrow,
    LeftUpArrow,
    LeftRightUpArrow,
    LeftArrowCallout,
    RightArrowCallout,
    UpArrowCallout,
    DownArrowCallout,
    LeftRightArrowCallout,
    UpDownArrowCallout,
    QuadArrowCallout,
    SwooshArrow,
    LeftCircularArrow,
    LeftRightCircularArrow,
    Callout1,
    Callout2,
    Callout3,
    AccentCallout1,
    AccentCallout2,
    AccentCallout3,
    BorderCallout1,
    BorderCallout2,
    BorderCallout3,
    AccentBorderCallout1,
    AccentBorderCallout2,
    AccentBorderCallout3,
    WedgeRectCallout,
    WedgeRoundRectCallout,
    WedgeEllipseCallout,
    CloudCallout,
    LeftBracket,
    RightBracket,
    LeftBrace,
    RightBrace,
    BracketPair,
    BracePair,
    StraightConnector1,
    BentConnector2,
    BentConnector3,
    BentConnector4,
    BentConnector5,
    CurvedConnector2,
    CurvedConnector3,
    CurvedConnector4,
    CurvedConnector5,
    FlowChartProcess,
    FlowChartDecision,
    FlowChartInputOutput,
    FlowChartPredefinedProcess,
    FlowChartInternalStorage,
    FlowChartDocument,
    FlowChartMultidocument,
    FlowChartTerminator,
    FlowChartPreparation,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartConnector,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSummingJunction,
    FlowChartOr,
    FlowChartCollate,
    FlowChartSort,
    FlowChartExtract,
    FlowChartMerge,
    FlowChartOfflineStorage,
    FlowChartOnlineStorage,
    FlowChartMagneticTape,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartDisplay,
    FlowChartDelay,
    FlowChartAlternateProcess,
    FlowChartOffpageConnector,
    ActionButtonBlank,
    ActionButtonHome,
    ActionButtonHelp,
    ActionButtonInformation,
    ActionButtonForwardNext,
    ActionButtonBackPrevious,
    ActionButtonEnd,
    ActionButtonBeginning,
    ActionButtonReturn,
    ActionButtonDocument,
    ActionButtonSound,
    ActionButtonMovie,
    IrregularSeal1,
    IrregularSeal2,
    Wave,
    DoubleWave,
    EllipseRibbon,
    EllipseRibbon2,
    VerticalScroll,
    HorizontalScroll,
    LeftRightRibbon,
    Gear6,
    Gear9,
    Funnel,
    MathPlus,
    MathMinus,
    MathMultiply,
    MathDivide,
    MathEqual,
    MathNotEqual,
    CornerTabs,
    SquareTabs,
    PlaqueTabs,
    ChartX,
    ChartStar,
    ChartPlus,
    HalfFrame,
    Corner,
    DiagStripe,
    NonIsoscelesTrapezoid,
    Heptagon,
    Decagon,
    Dodecagon,
    Round1Rect,
    Round2SameRect,
    Round2DiagRect,
    SnipRoundRect,
    Snip1Rect,
    Snip2SameRect,
    Snip2DiagRect,
    /// Unrecognized shape type — preserved as raw string.
    Other(String),
}

/// §20.1.10.10 ST_BlackWhiteMode.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BlackWhiteMode {
    Auto,
    Black,
    BlackGray,
    BlackWhite,
    Clr,
    Gray,
    GrayWhite,
    Hidden,
    InvGray,
    LtGray,
    White,
}

/// Drawing fill type.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DrawingFill {
    /// §20.1.8.44: no fill.
    NoFill,
    // SolidFill, GradFill, etc. — extend as needed.
}

/// §20.1.2.2.24 CT_LineProperties — outline/line properties.
#[derive(Clone, Debug)]
pub struct Outline {
    /// Line width in EMUs.
    pub width: Option<Dimension<Emu>>,
    /// §20.1.10.31: line cap style.
    pub cap: Option<LineCap>,
    /// §20.1.10.15: compound line type.
    pub compound: Option<CompoundLine>,
    /// §20.1.10.39: pen alignment.
    pub alignment: Option<PenAlignment>,
    /// Line fill.
    pub fill: Option<DrawingFill>,
}

/// §20.1.10.31 ST_LineCap.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LineCap {
    Flat,
    Round,
    Square,
}

/// §20.1.10.15 ST_CompoundLine.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CompoundLine {
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

/// §20.1.10.39 ST_PenAlignment.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PenAlignment {
    Center,
    Inset,
}

/// §20.4.2.3 CT_Anchor — anchor/floating drawing properties.
#[derive(Clone, Copy, Debug)]
pub struct AnchorProperties {
    /// §20.4.2.3: distance from surrounding text.
    pub distance: EdgeInsets<Emu>,
    /// §20.4.2.13: simple positioning point.
    pub simple_pos: Option<Offset<Emu>>,
    /// §20.4.2.3 @simplePos: whether to use simplePos coordinates.
    pub use_simple_pos: Option<bool>,
    /// §20.4.2.10: horizontal position.
    pub horizontal_position: AnchorPosition,
    /// §20.4.2.11: vertical position.
    pub vertical_position: AnchorPosition,
    /// Text wrapping mode.
    pub wrap: TextWrap,
    /// §20.4.2.3 @behindDoc: behind document text.
    pub behind_text: bool,
    /// §20.4.2.3 @locked: anchor is locked to position.
    pub lock_anchor: bool,
    /// §20.4.2.3 @allowOverlap: can overlap other anchored objects.
    pub allow_overlap: bool,
    /// §20.4.2.3 @relativeHeight: z-ordering value.
    pub relative_height: u32,
    /// §20.4.2.3 @layoutInCell: allow layout inside table cell.
    pub layout_in_cell: Option<bool>,
    /// §20.4.2.3 @hidden: whether the anchor is hidden.
    pub hidden: Option<bool>,
}

/// §20.4.2.10 / §20.4.2.11: anchor position (offset or alignment).
#[derive(Clone, Copy, Debug)]
pub enum AnchorPosition {
    /// §20.4.2.12: position by EMU offset from relativeFrom.
    Offset {
        relative_from: AnchorRelativeFrom,
        offset: Dimension<Emu>,
    },
    /// §20.4.2.1: position by alignment within relativeFrom.
    Align {
        relative_from: AnchorRelativeFrom,
        alignment: AnchorAlignment,
    },
}

/// §20.4.3.4 ST_RelFromH / §20.4.3.5 ST_RelFromV — relative-from values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AnchorRelativeFrom {
    Page,
    Margin,
    Column,
    Character,
    Paragraph,
    Line,
    InsideMargin,
    OutsideMargin,
    TopMargin,
    BottomMargin,
    LeftMargin,
    RightMargin,
}

/// §20.4.3.1 ST_AlignH / §20.4.3.2 ST_AlignV — alignment values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AnchorAlignment {
    Left,
    Center,
    Right,
    Inside,
    Outside,
    Top,
    Bottom,
}

/// Text wrapping mode for anchored drawings.
#[derive(Clone, Copy, Debug)]
pub enum TextWrap {
    /// §20.4.2.15: no wrapping.
    None,
    /// §20.4.2.17: square wrapping.
    Square {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
    /// §20.4.2.16: tight wrapping.
    Tight {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
    /// §20.4.2.18: text above and below only.
    TopAndBottom {
        distance_top: Dimension<Emu>,
        distance_bottom: Dimension<Emu>,
    },
    /// §20.4.2.14: through wrapping.
    Through {
        distance: EdgeInsets<Emu>,
        wrap_text: WrapText,
    },
}

/// §20.4.3.7 ST_WrapText — which sides text wraps on.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum WrapText {
    BothSides,
    Left,
    Right,
    Largest,
}

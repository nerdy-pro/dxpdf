//! VML (Vector Markup Language) types — legacy shapes, paths, formulas, styling.

use super::content::Block;
use super::identifiers::{RelId, VmlShapeId};

/// §17.3.3.19: VML picture container. Wraps legacy VML shape content.
#[derive(Clone, Debug)]
pub struct Pict {
    /// VML §14.1.2.20: optional reusable shape type definition.
    pub shape_type: Option<VmlShapeType>,
    /// VML §14.1.2.19: shape instances.
    pub shapes: Vec<VmlShape>,
}

// ── Path Commands ────────────────────────────────────────────────────────────

/// VML §14.2.1.6: path coordinate — literal integer or `@n` formula reference.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlPathCoord {
    /// Literal integer coordinate value.
    Literal(i64),
    /// `@n` — reference to formula result n.
    FormulaRef(u32),
}

/// VML §14.2.1.6: a single path command in the shape path language.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlPathCommand {
    /// `m x,y` — move to absolute position.
    MoveTo { x: VmlPathCoord, y: VmlPathCoord },
    /// `l x,y` — line to absolute position.
    LineTo { x: VmlPathCoord, y: VmlPathCoord },
    /// `c x1,y1,x2,y2,x,y` — cubic bezier to absolute position.
    CurveTo {
        x1: VmlPathCoord,
        y1: VmlPathCoord,
        x2: VmlPathCoord,
        y2: VmlPathCoord,
        x: VmlPathCoord,
        y: VmlPathCoord,
    },
    /// `r dx,dy` — relative line to.
    RLineTo { dx: VmlPathCoord, dy: VmlPathCoord },
    /// `v dx1,dy1,dx2,dy2,dx,dy` — relative cubic bezier.
    RCurveTo {
        dx1: VmlPathCoord,
        dy1: VmlPathCoord,
        dx2: VmlPathCoord,
        dy2: VmlPathCoord,
        dx: VmlPathCoord,
        dy: VmlPathCoord,
    },
    /// `t dx,dy` — relative move to.
    RMoveTo { dx: VmlPathCoord, dy: VmlPathCoord },
    /// `x` — close subpath.
    Close,
    /// `e` — end of path.
    End,
    /// `qx x,y` — elliptical quadrant, x-axis first.
    QuadrantX { x: VmlPathCoord, y: VmlPathCoord },
    /// `qy x,y` — elliptical quadrant, y-axis first.
    QuadrantY { x: VmlPathCoord, y: VmlPathCoord },
    /// `nf` — no fill for following subpath.
    NoFill,
    /// `ns` — no stroke for following subpath.
    NoStroke,
    /// `wa/wr/at/ar x1,y1,x2,y2,x3,y3,x4,y4` — arc commands.
    Arc {
        /// `wa` (angle clockwise), `wr` (angle counter-clockwise),
        /// `at` (to clockwise), `ar` (to counter-clockwise).
        kind: VmlArcKind,
        bounding_x1: VmlPathCoord,
        bounding_y1: VmlPathCoord,
        bounding_x2: VmlPathCoord,
        bounding_y2: VmlPathCoord,
        start_x: VmlPathCoord,
        start_y: VmlPathCoord,
        end_x: VmlPathCoord,
        end_y: VmlPathCoord,
    },
}

/// VML §14.2.1.6: arc sub-type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlArcKind {
    /// `wa` — clockwise arc (angle).
    WA,
    /// `wr` — counter-clockwise arc (angle).
    WR,
    /// `at` — clockwise arc (to point).
    AT,
    /// `ar` — counter-clockwise arc (to point).
    AR,
}

// ── Formulas ─────────────────────────────────────────────────────────────────

/// VML §14.1.2.6: a single formula equation (`v:f eqn="..."`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VmlFormula {
    pub operation: VmlFormulaOp,
    pub args: [VmlFormulaArg; 3],
}

/// VML §14.1.2.6: formula operations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFormulaOp {
    /// `val` — returns arg1.
    Val,
    /// `sum` — arg1 + arg2 - arg3.
    Sum,
    /// `product` — arg1 * arg2 / arg3.
    Product,
    /// `mid` — (arg1 + arg2) / 2.
    Mid,
    /// `abs` — |arg1|.
    Abs,
    /// `min` — min(arg1, arg2).
    Min,
    /// `max` — max(arg1, arg2).
    Max,
    /// `if` — if arg1 > 0 then arg2 else arg3.
    If,
    /// `sqrt` — sqrt(arg1).
    Sqrt,
    /// `mod` — sqrt(arg1² + arg2² + arg3²).
    Mod,
    /// `sin` — arg1 * sin(arg2).
    Sin,
    /// `cos` — arg1 * cos(arg2).
    Cos,
    /// `tan` — arg1 * tan(arg2).
    Tan,
    /// `atan2` — atan2(arg1, arg2) in fd units.
    Atan2,
    /// `sinatan2` — arg1 * sin(atan2(arg2, arg3)).
    SinAtan2,
    /// `cosatan2` — arg1 * cos(atan2(arg2, arg3)).
    CosAtan2,
    /// `sumangle` — arg1 + arg2 * 2^16 - arg3 * 2^16 (angle arithmetic).
    SumAngle,
    /// `ellipse` — arg3 * sqrt(1 - (arg1/arg2)²).
    Ellipse,
}

/// VML §14.1.2.6: formula argument — a reference or literal.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFormulaArg {
    /// Integer literal.
    Literal(i64),
    /// `#n` — adjustment value reference.
    AdjRef(u32),
    /// `@n` — formula result reference.
    FormulaRef(u32),
    /// Named guide value (width, height, xcenter, ycenter, etc.).
    Guide(VmlGuide),
}

/// VML §14.1.2.6: named guide constants available in formulas.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlGuide {
    Width,
    Height,
    XCenter,
    YCenter,
    XRange,
    YRange,
    PixelWidth,
    PixelHeight,
    PixelLineWidth,
    EmuWidth,
    EmuHeight,
    EmuWidth2,
    EmuHeight2,
}

// ── Shape Types and Shapes ───────────────────────────────────────────────────

/// VML Vector2D — unitless coordinate pair (e.g., coordsize="21600,21600").
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct VmlVector2D {
    pub x: i64,
    pub y: i64,
}

/// VML §14.1.2.20: shape type definition (reusable template for shapes).
#[derive(Clone, Debug)]
pub struct VmlShapeType {
    /// Shape type identifier (e.g., "_x0000_t202").
    pub id: Option<VmlShapeId>,
    /// Coordinate space for the shape (VML Vector2D, e.g., 21600,21600).
    pub coord_size: Option<VmlVector2D>,
    /// o:spt — shape type number (Office extension, xsd:float).
    pub spt: Option<f32>,
    /// Adjustment values for parameterized shapes (comma-separated integers).
    pub adj: Vec<i64>,
    /// VML §14.2.1.6: parsed shape path commands.
    pub path: Vec<VmlPathCommand>,
    /// Whether the shape is filled by default.
    pub filled: Option<bool>,
    /// Whether the shape is stroked by default.
    pub stroked: Option<bool>,
    /// VML §14.1.2.21: stroke child element.
    pub stroke: Option<VmlStroke>,
    /// VML §14.1.2.14: path child element.
    pub vml_path: Option<VmlPath>,
    /// VML §14.1.2.6: formula definitions.
    pub formulas: Vec<VmlFormula>,
    /// Office VML extension: editing locks.
    pub lock: Option<VmlLock>,
}

/// VML §14.1.2.19: a shape instance.
#[derive(Clone, Debug)]
pub struct VmlShape {
    /// Shape identifier.
    pub id: Option<VmlShapeId>,
    /// Reference to a shape type (e.g., "#_x0000_t202").
    pub shape_type_ref: Option<VmlShapeId>,
    /// Parsed CSS2 style properties.
    pub style: VmlStyle,
    /// Fill color.
    pub fill_color: Option<VmlColor>,
    /// Whether the shape has a stroke.
    pub stroked: Option<bool>,
    /// VML §14.1.2.21: stroke child element.
    pub stroke: Option<VmlStroke>,
    /// VML §14.1.2.14: path child element.
    pub vml_path: Option<VmlPath>,
    /// VML §14.1.2.22: text box child element.
    pub text_box: Option<VmlTextBox>,
    /// VML §14.1.2.23: text wrapping around shape.
    pub wrap: Option<VmlWrap>,
    /// VML §14.1.2.11: image data reference.
    pub image_data: Option<VmlImageData>,
}

// ── Wrapping ─────────────────────────────────────────────────────────────────

/// VML §14.1.2.23: text wrapping element.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct VmlWrap {
    /// Wrapping type.
    pub wrap_type: Option<VmlWrapType>,
    /// Which side(s) text wraps on.
    pub side: Option<VmlWrapSide>,
}

/// VML §14.1.2.23: wrap type values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlWrapType {
    TopAndBottom,
    Square,
    None,
    Tight,
    Through,
}

/// VML §14.1.2.23: wrap side values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlWrapSide {
    Both,
    Left,
    Right,
    Largest,
}

// ── Lock / Extension ─────────────────────────────────────────────────────────

/// Office VML extension: editing locks on a shape type or shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct VmlLock {
    /// Lock aspect ratio.
    pub aspect_ratio: Option<bool>,
    /// v:ext — extension handling mode.
    pub ext: Option<VmlExtHandling>,
}

/// VML v:ext — extension handling mode.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlExtHandling {
    Edit,
    View,
    BackwardCompatible,
}

// ── Image Data ───────────────────────────────────────────────────────────────

/// VML §14.1.2.11: image data reference.
#[derive(Clone, Debug)]
pub struct VmlImageData {
    /// r:id — relationship ID to the image part.
    pub rel_id: Option<RelId>,
    /// o:title — image title/alt text.
    pub title: Option<String>,
}

// ── VML Style ────────────────────────────────────────────────────────────────

/// VML style — parsed CSS2 properties from the `style` attribute (§14.1.2.19).
#[derive(Clone, Debug, Default)]
pub struct VmlStyle {
    /// CSS `position`.
    pub position: Option<CssPosition>,
    /// CSS `left`.
    pub left: Option<VmlLength>,
    /// CSS `top`.
    pub top: Option<VmlLength>,
    /// CSS `width`.
    pub width: Option<VmlLength>,
    /// CSS `height`.
    pub height: Option<VmlLength>,
    /// CSS `margin-left`.
    pub margin_left: Option<VmlLength>,
    /// CSS `margin-top`.
    pub margin_top: Option<VmlLength>,
    /// CSS `margin-right`.
    pub margin_right: Option<VmlLength>,
    /// CSS `margin-bottom`.
    pub margin_bottom: Option<VmlLength>,
    /// CSS `z-index`.
    pub z_index: Option<i64>,
    /// CSS `rotation` (degrees).
    pub rotation: Option<f64>,
    /// VML `flip`.
    pub flip: Option<VmlFlip>,
    /// CSS `visibility`.
    pub visibility: Option<CssVisibility>,
    /// Office `mso-position-horizontal`.
    pub mso_position_horizontal: Option<MsoPositionH>,
    /// Office `mso-position-horizontal-relative`.
    pub mso_position_horizontal_relative: Option<MsoPositionHRelative>,
    /// Office `mso-position-vertical`.
    pub mso_position_vertical: Option<MsoPositionV>,
    /// Office `mso-position-vertical-relative`.
    pub mso_position_vertical_relative: Option<MsoPositionVRelative>,
    /// Office `mso-wrap-distance-left`.
    pub mso_wrap_distance_left: Option<VmlLength>,
    /// Office `mso-wrap-distance-right`.
    pub mso_wrap_distance_right: Option<VmlLength>,
    /// Office `mso-wrap-distance-top`.
    pub mso_wrap_distance_top: Option<VmlLength>,
    /// Office `mso-wrap-distance-bottom`.
    pub mso_wrap_distance_bottom: Option<VmlLength>,
    /// Office `mso-wrap-style`.
    pub mso_wrap_style: Option<MsoWrapStyle>,
}

// ── VML Color ────────────────────────────────────────────────────────────────

/// VML color value (§14.1.2.1 ST_ColorType).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum VmlColor {
    /// Hex-specified sRGB color (r, g, b).
    Rgb(u8, u8, u8),
    /// Predefined named color.
    Named(VmlNamedColor),
}

/// VML/CSS named colors (§14.1.2.1, CSS2.1 §4.3.6, and SVG/CSS3 extended colors).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum VmlNamedColor {
    // CSS2.1 §4.3.6 — the 17 standard colors.
    Black,
    Silver,
    Gray,
    White,
    Maroon,
    Red,
    Purple,
    Fuchsia,
    Green,
    Lime,
    Olive,
    Yellow,
    Navy,
    Blue,
    Teal,
    Aqua,
    Orange,
    // SVG/CSS3 extended named colors used in Office documents.
    AliceBlue,
    AntiqueWhite,
    Beige,
    Bisque,
    BlanchedAlmond,
    BlueViolet,
    Brown,
    BurlyWood,
    CadetBlue,
    Chartreuse,
    Chocolate,
    Coral,
    CornflowerBlue,
    Cornsilk,
    Crimson,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGoldenrod,
    DarkGray,
    DarkGreen,
    DarkKhaki,
    DarkMagenta,
    DarkOliveGreen,
    DarkOrange,
    DarkOrchid,
    DarkRed,
    DarkSalmon,
    DarkSeaGreen,
    DarkSlateBlue,
    DarkSlateGray,
    DarkTurquoise,
    DarkViolet,
    DeepPink,
    DeepSkyBlue,
    DimGray,
    DodgerBlue,
    Firebrick,
    FloralWhite,
    ForestGreen,
    Gainsboro,
    GhostWhite,
    Gold,
    Goldenrod,
    GreenYellow,
    Honeydew,
    HotPink,
    IndianRed,
    Indigo,
    Ivory,
    Khaki,
    Lavender,
    LavenderBlush,
    LawnGreen,
    LemonChiffon,
    LightBlue,
    LightCoral,
    LightCyan,
    LightGoldenrodYellow,
    LightGray,
    LightGreen,
    LightPink,
    LightSalmon,
    LightSeaGreen,
    LightSkyBlue,
    LightSlateGray,
    LightSteelBlue,
    LightYellow,
    LimeGreen,
    Linen,
    Magenta,
    MediumAquamarine,
    MediumBlue,
    MediumOrchid,
    MediumPurple,
    MediumSeaGreen,
    MediumSlateBlue,
    MediumSpringGreen,
    MediumTurquoise,
    MediumVioletRed,
    MidnightBlue,
    MintCream,
    MistyRose,
    Moccasin,
    NavajoWhite,
    OldLace,
    OliveDrab,
    OrangeRed,
    Orchid,
    PaleGoldenrod,
    PaleGreen,
    PaleTurquoise,
    PaleVioletRed,
    PapayaWhip,
    PeachPuff,
    Peru,
    Pink,
    Plum,
    PowderBlue,
    RosyBrown,
    RoyalBlue,
    SaddleBrown,
    Salmon,
    SandyBrown,
    SeaGreen,
    Seashell,
    Sienna,
    SkyBlue,
    SlateBlue,
    SlateGray,
    Snow,
    SpringGreen,
    SteelBlue,
    Tan,
    Thistle,
    Tomato,
    Turquoise,
    Violet,
    Wheat,
    WhiteSmoke,
    YellowGreen,
    // VML §14.1.2.1 system colors.
    ButtonFace,
    ButtonHighlight,
    ButtonShadow,
    ButtonText,
    CaptionText,
    GrayText,
    Highlight,
    HighlightText,
    InactiveBorder,
    InactiveCaption,
    InactiveCaptionText,
    InfoBackground,
    InfoText,
    Menu,
    MenuText,
    Scrollbar,
    ThreeDDarkShadow,
    ThreeDFace,
    ThreeDHighlight,
    ThreeDLightShadow,
    ThreeDShadow,
    Window,
    WindowFrame,
    WindowText,
}

// ── CSS Enums ────────────────────────────────────────────────────────────────

/// CSS2 `position` property values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CssPosition {
    Static,
    Relative,
    Absolute,
}

/// CSS2 `visibility` property values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CssVisibility {
    Visible,
    Hidden,
    Inherit,
}

/// VML `flip` attribute values (§14.1.2.19).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlFlip {
    /// Flip along the x-axis.
    X,
    /// Flip along the y-axis.
    Y,
    /// Flip along both axes.
    XY,
}

// ── MSO Position / Wrap ──────────────────────────────────────────────────────

/// Office `mso-position-horizontal` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionH {
    Absolute,
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

/// Office `mso-position-horizontal-relative` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionHRelative {
    Margin,
    Page,
    Text,
    Char,
    LeftMarginArea,
    RightMarginArea,
    InnerMarginArea,
    OuterMarginArea,
}

/// Office `mso-position-vertical` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionV {
    Absolute,
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
}

/// Office `mso-position-vertical-relative` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoPositionVRelative {
    Margin,
    Page,
    Text,
    Line,
    TopMarginArea,
    BottomMarginArea,
    InnerMarginArea,
    OuterMarginArea,
}

/// Office `mso-wrap-style` values.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MsoWrapStyle {
    Square,
    None,
    Tight,
    Through,
}

// ── VML Length ────────────────────────────────────────────────────────────────

/// A CSS length value with unit.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct VmlLength {
    pub value: f64,
    pub unit: VmlLengthUnit,
}

/// CSS length unit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlLengthUnit {
    /// Points (pt).
    Pt,
    /// Inches (in).
    In,
    /// Centimeters (cm).
    Cm,
    /// Millimeters (mm).
    Mm,
    /// Pixels (px).
    Px,
    /// Em units (em).
    Em,
    /// Percentage (%).
    Percent,
    /// No unit (bare number, treated as EMU in VML context).
    None,
}

// ── Stroke ───────────────────────────────────────────────────────────────────

/// VML §14.1.2.21: stroke styling.
#[derive(Clone, Debug)]
pub struct VmlStroke {
    /// Dash pattern.
    pub dash_style: Option<VmlDashStyle>,
    /// Line join style.
    pub join_style: Option<VmlJoinStyle>,
}

/// VML §14.1.2.21: stroke dash patterns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlDashStyle {
    Solid,
    ShortDash,
    ShortDot,
    ShortDashDot,
    ShortDashDotDot,
    Dot,
    Dash,
    LongDash,
    DashDot,
    LongDashDot,
    LongDashDotDot,
}

/// VML §14.1.2.21: line join styles.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlJoinStyle {
    Round,
    Bevel,
    Miter,
}

// ── Path Properties ──────────────────────────────────────────────────────────

/// VML §14.1.2.14: path properties.
#[derive(Clone, Debug)]
pub struct VmlPath {
    /// o:gradientshapeok — whether the path supports gradient fill.
    pub gradient_shape_ok: Option<bool>,
    /// o:connecttype — connection point type.
    pub connect_type: Option<VmlConnectType>,
    /// o:extrusionok — whether the path supports 3D extrusion.
    pub extrusion_ok: Option<bool>,
}

/// VML o:connecttype — how connection points are defined on the shape.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VmlConnectType {
    /// No connection points.
    None,
    /// Four connection points at the rectangle midpoints.
    Rect,
    /// Connection points derived from path segments.
    Segments,
    /// Custom connection points.
    Custom,
}

// ── Text Box ─────────────────────────────────────────────────────────────────

/// VML §14.1.2.22: text box inset margins (comma-separated CSS lengths).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct VmlTextBoxInset {
    pub left: Option<VmlLength>,
    pub top: Option<VmlLength>,
    pub right: Option<VmlLength>,
    pub bottom: Option<VmlLength>,
}

/// VML §14.1.2.22: text box within a shape.
#[derive(Clone, Debug)]
pub struct VmlTextBox {
    /// Parsed CSS2 style properties.
    pub style: VmlStyle,
    /// VML §14.1.2.22: inset margins (top, left, bottom, right).
    pub inset: Option<VmlTextBoxInset>,
    /// §17.17.1: block-level content from w:txbxContent.
    pub content: Vec<Block>,
}

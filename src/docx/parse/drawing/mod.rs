//! DrawingML parsing — fully serde-based.
//!
//! See [`schema`] for the type-by-type schemas (color, fill, stroke, effect,
//! geometry, picture, shape, anchor, inline). Consumers (body, numbering,
//! notes) call the top-level `InlineXml::into_image(ctx)` /
//! `AnchorXml::into_image(ctx)` / `PictureXml` / `WspXml::into_model(ctx)`
//! entry points directly.

pub mod schema;

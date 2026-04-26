//! Emoji rendering pipeline.
//!
//! This module implements the staged plan in `docs/emoji-rendering.md`:
//! cluster classification (UAX #29 + UTS #51), host-OS color emoji typeface
//! resolution, and Skia raster-backend rasterization with a per-render cache.
//!
//! The rest of the renderer interacts with this module through typed ADTs;
//! no string-name allowlists, no font bundling.

pub mod cluster;
pub mod raster;
pub mod resolve;
pub mod shape;

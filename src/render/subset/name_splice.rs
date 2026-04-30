//! Splice the original SFNT `name` table back into a subsetted font.
//!
//! `fontcull` (via `klippa`) drops every record from the `name` table because
//! its public API hard-codes an empty `name_ids` set. The resulting subset has
//! no family/full/PostScript name, so Skia falls back to a synthetic
//! `font<hex>` identifier and the PDF embeds that synthetic name. Splicing the
//! original `name` table back in restores the real PostScript name (with the
//! standard `ABCDEF+` subset prefix that Skia adds at PDF write time).
//!
//! The `name` table is independent of the glyph order — it stores font
//! metadata, not per-glyph data — so substituting it cannot affect text
//! shaping or rendering.

const HEAD_TAG: &[u8; 4] = b"head";
const NAME_TAG: &[u8; 4] = b"name";
const SFNT_HEADER_SIZE: usize = 12;
const TABLE_RECORD_SIZE: usize = 16;
const HEAD_CHECKSUM_ADJUSTMENT_OFFSET: usize = 8;
const HEAD_MAGIC: u32 = 0xB1B0_AFBA;

/// Replace the `name` table in `subsetted` with the one from `original`.
/// Returns the rebuilt SFNT bytes. If either input lacks a `name` table the
/// subsetted bytes are returned unchanged.
pub fn splice_original_name(subsetted: &[u8], original: &[u8]) -> Result<Vec<u8>, String> {
    let Some(name_bytes) = find_table(original, NAME_TAG)? else {
        return Ok(subsetted.to_vec());
    };
    replace_table(subsetted, NAME_TAG, &name_bytes)
}

/// Locate a table by tag and return its raw bytes (unpadded).
fn find_table(sfnt: &[u8], tag: &[u8; 4]) -> Result<Option<Vec<u8>>, String> {
    let num_tables = read_num_tables(sfnt)?;
    for i in 0..num_tables {
        let rec = SFNT_HEADER_SIZE + i * TABLE_RECORD_SIZE;
        if rec + TABLE_RECORD_SIZE > sfnt.len() {
            return Err("table directory overflow".into());
        }
        if &sfnt[rec..rec + 4] == tag.as_slice() {
            let off = read_u32(sfnt, rec + 8) as usize;
            let len = read_u32(sfnt, rec + 12) as usize;
            if off.checked_add(len).is_none_or(|end| end > sfnt.len()) {
                return Err("table data overflow".into());
            }
            return Ok(Some(sfnt[off..off + len].to_vec()));
        }
    }
    Ok(None)
}

/// Build a new SFNT identical to `sfnt` except the table at `tag` is replaced
/// with `new_data`. Recomputes table directory, table checksums, and the
/// `head` table's `checksumAdjustment`.
fn replace_table(sfnt: &[u8], tag: &[u8; 4], new_data: &[u8]) -> Result<Vec<u8>, String> {
    if sfnt.len() < SFNT_HEADER_SIZE {
        return Err("sfnt too short".into());
    }
    let num_tables = read_num_tables(sfnt)?;

    // Collect (tag, data) pairs for every table, substituting `new_data` for
    // the target tag. Tables that don't appear are added.
    let mut tables: Vec<([u8; 4], Vec<u8>)> = Vec::with_capacity(num_tables);
    let mut found = false;
    for i in 0..num_tables {
        let rec = SFNT_HEADER_SIZE + i * TABLE_RECORD_SIZE;
        if rec + TABLE_RECORD_SIZE > sfnt.len() {
            return Err("table directory overflow".into());
        }
        let mut t = [0u8; 4];
        t.copy_from_slice(&sfnt[rec..rec + 4]);
        let data = if &t == tag {
            found = true;
            new_data.to_vec()
        } else {
            let off = read_u32(sfnt, rec + 8) as usize;
            let len = read_u32(sfnt, rec + 12) as usize;
            if off.checked_add(len).is_none_or(|end| end > sfnt.len()) {
                return Err("table data overflow".into());
            }
            sfnt[off..off + len].to_vec()
        };
        tables.push((t, data));
    }
    if !found {
        tables.push((*tag, new_data.to_vec()));
    }
    // SFNT directory is sorted by tag.
    tables.sort_by_key(|t| t.0);

    rebuild_sfnt(&sfnt[0..4], &tables)
}

fn rebuild_sfnt(version: &[u8], tables: &[([u8; 4], Vec<u8>)]) -> Result<Vec<u8>, String> {
    let n = u16::try_from(tables.len()).map_err(|_| "too many tables".to_string())?;
    // Per OpenType spec: entrySelector = floor(log2(numTables));
    // searchRange = 2^entrySelector * 16; rangeShift = numTables*16 - searchRange.
    let entry_selector = if n == 0 {
        0
    } else {
        15 - n.leading_zeros() as u16
    };
    let search_range = (1u16 << entry_selector) * 16;
    let range_shift = n * 16 - search_range;

    let dir_size = SFNT_HEADER_SIZE + tables.len() * TABLE_RECORD_SIZE;
    let mut offsets = Vec::with_capacity(tables.len());
    let mut cur = dir_size;
    for (_, data) in tables {
        offsets.push(cur);
        cur += data.len();
        cur = (cur + 3) & !3; // 4-byte alignment between tables
    }
    let total_size = cur;

    let mut out = vec![0u8; total_size];
    out[0..4].copy_from_slice(version);
    out[4..6].copy_from_slice(&n.to_be_bytes());
    out[6..8].copy_from_slice(&search_range.to_be_bytes());
    out[8..10].copy_from_slice(&entry_selector.to_be_bytes());
    out[10..12].copy_from_slice(&range_shift.to_be_bytes());

    for (i, ((t, data), &off)) in tables.iter().zip(offsets.iter()).enumerate() {
        let rec = SFNT_HEADER_SIZE + i * TABLE_RECORD_SIZE;
        out[rec..rec + 4].copy_from_slice(t);
        // checksum filled below
        out[rec + 8..rec + 12].copy_from_slice(&(off as u32).to_be_bytes());
        out[rec + 12..rec + 16].copy_from_slice(&(data.len() as u32).to_be_bytes());
        out[off..off + data.len()].copy_from_slice(data);
    }

    // Per-table checksums are computed over the padded table region.
    for (i, ((_, data), &off)) in tables.iter().zip(offsets.iter()).enumerate() {
        let rec = SFNT_HEADER_SIZE + i * TABLE_RECORD_SIZE;
        let padded_len = (data.len() + 3) & !3;
        let checksum = sfnt_checksum(&out[off..off + padded_len]);
        out[rec + 4..rec + 8].copy_from_slice(&checksum.to_be_bytes());
    }

    // §5.2: head.checksumAdjustment = 0xB1B0AFBA - whole-file-checksum,
    // computed with checksumAdjustment itself zeroed.
    if let Some((i, _)) = tables.iter().enumerate().find(|(_, (t, _))| t == HEAD_TAG) {
        let head_off = offsets[i];
        if head_off + HEAD_CHECKSUM_ADJUSTMENT_OFFSET + 4 > out.len() {
            return Err("head table truncated".into());
        }
        let adj_pos = head_off + HEAD_CHECKSUM_ADJUSTMENT_OFFSET;
        out[adj_pos..adj_pos + 4].copy_from_slice(&0u32.to_be_bytes());
        let total = sfnt_checksum(&out);
        let adjustment = HEAD_MAGIC.wrapping_sub(total);
        out[adj_pos..adj_pos + 4].copy_from_slice(&adjustment.to_be_bytes());
    }

    Ok(out)
}

fn read_num_tables(sfnt: &[u8]) -> Result<usize, String> {
    if sfnt.len() < SFNT_HEADER_SIZE {
        return Err("sfnt too short".into());
    }
    Ok(u16::from_be_bytes([sfnt[4], sfnt[5]]) as usize)
}

fn read_u32(buf: &[u8], offset: usize) -> u32 {
    u32::from_be_bytes([
        buf[offset],
        buf[offset + 1],
        buf[offset + 2],
        buf[offset + 3],
    ])
}

/// Sum the buffer as big-endian u32 values, with zero-padding at the tail
/// if the length isn't 4-aligned.
fn sfnt_checksum(data: &[u8]) -> u32 {
    let mut sum: u32 = 0;
    let chunks = data.chunks_exact(4);
    let rem = chunks.remainder();
    for c in chunks {
        sum = sum.wrapping_add(u32::from_be_bytes([c[0], c[1], c[2], c[3]]));
    }
    if !rem.is_empty() {
        let mut padded = [0u8; 4];
        padded[..rem.len()].copy_from_slice(rem);
        sum = sum.wrapping_add(u32::from_be_bytes(padded));
    }
    sum
}

#[cfg(test)]
mod tests {
    use super::*;
    use skia_safe::{Data, FontMgr, FontStyle};

    #[test]
    fn splice_preserves_family_name() {
        let mgr = FontMgr::new();
        let Some(tf) = mgr.match_family_style("Carlito", FontStyle::normal()) else {
            // Carlito isn't installed on every CI image — skip rather than fail.
            return;
        };
        let Some((bytes, _)) = tf.to_font_data() else {
            return;
        };
        let unicodes: Vec<u32> = (0x20u32..0x7Fu32).collect();
        let subsetted = match fontcull::subset_font_data_unicode(&bytes, &unicodes, &[]) {
            Ok(s) => s,
            Err(_) => return,
        };

        let pre = mgr
            .new_from_data(&Data::new_copy(&subsetted), 0)
            .expect("subsetted parses");
        assert_eq!(
            pre.family_name(),
            "",
            "fontcull is expected to wipe the name table — if this changes, the splice is unnecessary"
        );

        let spliced = splice_original_name(&subsetted, &bytes).expect("splice");
        let post = mgr
            .new_from_data(&Data::new_copy(&spliced), 0)
            .expect("spliced parses");
        assert_eq!(post.family_name(), "Carlito");
    }
}

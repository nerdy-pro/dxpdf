//! C ABI for the Go bindings (`go/`), enabled by the `capi` feature.
//!
//! This is the only non-Rust, non-Python entry point into the crate: Go
//! cannot call Rust directly, so `go/dxpdf.go` reaches this module through
//! cgo against the `staticlib` artifact `crate-type` produces. The surface
//! is deliberately narrow — one conversion entry point, two matching
//! `free` calls, and two constant getters — because every function here is
//! `unsafe` at the boundary in a way the Python bindings (via PyO3) are not:
//! PyO3 catches panics and manages memory for us, cgo does neither.
//!
//! Unlike `src/lib.rs`'s `python` module, there is no `dxpdf_convert_file`:
//! file I/O is done in pure Go on top of [`dxpdf_convert`], the same way the
//! Python module's own `convert_file` wraps `std::fs::read`/`write` around
//! `convert_with_options` rather than pushing file paths across the FFI
//! boundary. That keeps the `unsafe` surface to a single function.
//!
//! Concurrency: `FontRegistry` (`src/render/fonts/mod.rs`) is owned per
//! render, not a `thread_local!`/process-global cache, so calling
//! [`dxpdf_convert`] concurrently from multiple Go goroutines/OS threads
//! needs no additional locking on dxpdf's side.

use std::ffi::CString;
use std::os::raw::c_char;
use std::panic::{self, AssertUnwindSafe};
use std::ptr;

/// A byte buffer handed back across the C ABI.
///
/// `cap` travels alongside `len` because [`dxpdf_free_buffer`] reconstructs
/// the exact `Vec<u8>` that produced this buffer via [`Vec::from_raw_parts`],
/// which requires the real allocated capacity — not the buffer's logical
/// length — to hand back to Rust's allocator.
#[repr(C)]
pub struct DxpdfBuffer {
    pub data: *mut u8,
    pub len: usize,
    pub cap: usize,
}

impl DxpdfBuffer {
    fn empty() -> Self {
        DxpdfBuffer {
            data: ptr::null_mut(),
            len: 0,
            cap: 0,
        }
    }

    fn from_vec(mut bytes: Vec<u8>) -> Self {
        let buf = DxpdfBuffer {
            data: bytes.as_mut_ptr(),
            len: bytes.len(),
            cap: bytes.capacity(),
        };
        std::mem::forget(bytes);
        buf
    }
}

/// Writes `message` into `*out_error` as an owned, NUL-terminated C string.
/// A no-op if `out_error` is null. Interior NUL bytes are stripped rather
/// than rejected: `message` can carry arbitrary content from a malformed
/// DOCX (e.g. a path or XML value quoted into an error), and this is the
/// FFI boundary — the one place that content has to be made safe for a
/// C string, not an internal invariant we get to assume.
fn set_error(out_error: *mut *mut c_char, message: &str) {
    if out_error.is_null() {
        return;
    }
    let sanitized: String = message.chars().filter(|&c| c != '\0').collect();
    let c_string = CString::new(sanitized).expect("NUL bytes were just filtered out");
    unsafe {
        *out_error = c_string.into_raw();
    }
}

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    if let Some(s) = payload.downcast_ref::<&str>() {
        (*s).to_string()
    } else if let Some(s) = payload.downcast_ref::<String>() {
        s.clone()
    } else {
        "unknown panic".to_string()
    }
}

/// Converts DOCX bytes to PDF bytes.
///
/// Reads `docx_len` bytes from `docx_ptr` and converts at `image_dpi` (pass
/// [`dxpdf_default_image_dpi`]'s result for the library default). On success
/// writes the PDF bytes into `*out_pdf` and returns `0`. On failure — a
/// conversion error or a caught panic — writes a message into `*out_error`
/// instead, leaves `*out_pdf` zeroed, and returns a nonzero code.
///
/// # Safety
///
/// `docx_ptr` must be null or point to `docx_len` readable, initialized
/// bytes. `out_pdf` and `out_error`, if non-null, must each point to valid,
/// writable storage of the matching type. The buffer written to `*out_pdf`
/// must be released with [`dxpdf_free_buffer`], and the string written to
/// `*out_error` with [`dxpdf_free_error`] — both were allocated by Rust's
/// allocator and must not be passed to a C `free()`.
#[no_mangle]
pub unsafe extern "C" fn dxpdf_convert(
    docx_ptr: *const u8,
    docx_len: usize,
    image_dpi: f32,
    out_pdf: *mut DxpdfBuffer,
    out_error: *mut *mut c_char,
) -> i32 {
    if !out_pdf.is_null() {
        unsafe {
            *out_pdf = DxpdfBuffer::empty();
        }
    }
    if !out_error.is_null() {
        unsafe {
            *out_error = ptr::null_mut();
        }
    }
    if docx_ptr.is_null() {
        set_error(out_error, "dxpdf_convert: docx_ptr is null");
        return -1;
    }

    let docx_bytes = unsafe { std::slice::from_raw_parts(docx_ptr, docx_len) };
    let options = crate::RenderOptions::default().with_image_dpi(image_dpi);

    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        crate::convert_with_options(docx_bytes, &options)
    }));

    match result {
        Ok(Ok(pdf_bytes)) => {
            if !out_pdf.is_null() {
                unsafe {
                    *out_pdf = DxpdfBuffer::from_vec(pdf_bytes);
                }
            }
            0
        }
        Ok(Err(e)) => {
            set_error(out_error, &e.to_string());
            1
        }
        Err(panic_payload) => {
            let message = panic_message(panic_payload.as_ref());
            set_error(out_error, &format!("dxpdf: internal error: {message}"));
            2
        }
    }
}

/// Frees a buffer returned by [`dxpdf_convert`]'s `out_pdf`. Safe to call on
/// a zeroed buffer (e.g. one left behind by a failed conversion).
///
/// # Safety
///
/// `buf` must be a [`DxpdfBuffer`] previously returned by `dxpdf_convert`,
/// unmodified, and must not be freed more than once.
#[no_mangle]
pub unsafe extern "C" fn dxpdf_free_buffer(buf: DxpdfBuffer) {
    if buf.data.is_null() {
        return;
    }
    unsafe {
        drop(Vec::from_raw_parts(buf.data, buf.len, buf.cap));
    }
}

/// Frees an error string returned by [`dxpdf_convert`]'s `out_error`. A
/// no-op on null.
///
/// # Safety
///
/// `err` must be a pointer previously returned by `dxpdf_convert`'s
/// `out_error` (or null), unmodified, and must not be freed more than once.
#[no_mangle]
pub unsafe extern "C" fn dxpdf_free_error(err: *mut c_char) {
    if err.is_null() {
        return;
    }
    unsafe {
        drop(CString::from_raw(err));
    }
}

/// The library's default embedded-image DPI ([`crate::DEFAULT_IMAGE_DPI`]).
#[no_mangle]
pub extern "C" fn dxpdf_default_image_dpi() -> f32 {
    crate::DEFAULT_IMAGE_DPI
}

/// The floor `image_dpi` is clamped to ([`crate::MIN_IMAGE_DPI`]).
#[no_mangle]
pub extern "C" fn dxpdf_min_image_dpi() -> f32 {
    crate::MIN_IMAGE_DPI
}

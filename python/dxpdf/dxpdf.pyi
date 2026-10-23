from typing import Optional

class ResourceLimits:
    """Inclusive per-conversion budgets. Zero disables the resource, not its limit.

    These are not total-memory or CPU limits. Untrusted documents still require
    a resource-limited worker process.
    """

    def __init__(
        self,
        *,
        max_input_bytes: int = 33554432,
        max_zip_entries: int = 4096,
        max_zip_entry_bytes: int = 67108864,
        max_zip_total_bytes: int = 268435456,
        max_image_pixels: int = 25000000,
    ) -> None: ...

    @property
    def max_input_bytes(self) -> int: ...
    @property
    def max_zip_entries(self) -> int: ...
    @property
    def max_zip_entry_bytes(self) -> int: ...
    @property
    def max_zip_total_bytes(self) -> int: ...
    @property
    def max_image_pixels(self) -> int: ...

def convert(
    docx_bytes: bytes,
    image_dpi: float = 220.0,
    *,
    limits: Optional[ResourceLimits] = None,
) -> bytes:
    """Convert DOCX bytes to PDF bytes.

    Args:
        docx_bytes: Raw bytes of a .docx file.
        image_dpi: Target embedded-image resolution, default 220.
        limits: Resource budgets; None uses finite defaults.

    Returns:
        PDF file contents as bytes.

    Raises:
        RuntimeError: If the DOCX is invalid, a budget is exceeded, or conversion fails.
    """
    ...

def convert_file(
    input: str,
    output: str,
    image_dpi: float = 220.0,
    *,
    limits: Optional[ResourceLimits] = None,
) -> None:
    """Convert a DOCX file to a PDF file.

    Args:
        input: Path to the input .docx file.
        output: Path to the output .pdf file.
        image_dpi: Target embedded-image resolution, default 220.
        limits: Resource budgets; None uses finite defaults.

    Raises:
        RuntimeError: If a budget is exceeded or reading, conversion, or writing fails.
    """
    ...

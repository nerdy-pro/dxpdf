def convert(docx_bytes: bytes) -> bytes:
    """Convert DOCX bytes to PDF bytes.

    Args:
        docx_bytes: Raw bytes of a .docx file.

    Returns:
        PDF file contents as bytes.

    Raises:
        RuntimeError: If the DOCX file is invalid or conversion fails.
    """
    ...

def convert_file(input: str, output: str) -> None:
    """Convert a DOCX file to a PDF file.

    Args:
        input: Path to the input .docx file.
        output: Path to the output .pdf file.

    Raises:
        RuntimeError: If reading, conversion, or writing fails.
    """
    ...

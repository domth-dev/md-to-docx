# md_to_docx/errors.py

class MdToDocxError(Exception):
    """Base exception for all package errors."""


class InputNotFoundError(MdToDocxError):
    """Input file is missing or unreadable."""


class TemplateError(MdToDocxError):
    """Template issues detected (e.g., missing or corrupted parts)."""


class NumberingError(MdToDocxError):
    """No suitable numbering found or applied in the template."""


class ImageDownloadError(MdToDocxError):
    """Image download failed (timeout, content-type, size, etc.)."""


class MathConversionError(MdToDocxError):
    """LaTeX to OMML conversion failed unexpectedly (no fallback available)."""
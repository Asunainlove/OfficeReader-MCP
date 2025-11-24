"""
Core converter module for Office documents to Markdown conversion.
Supports: docx, doc, xlsx, xls, pptx, ppt
"""

import base64
import hashlib
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET
from io import BytesIO

import mammoth
from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image

from .image_optimizer import ImageOptimizer


# Supported file extensions
SUPPORTED_FORMATS = {
    "word": [".docx", ".doc"],
    "excel": [".xlsx", ".xls"],
    "powerpoint": [".pptx", ".ppt"],
}


def get_file_type(file_path: Path) -> Optional[str]:
    """Determine file type from extension."""
    suffix = file_path.suffix.lower()
    for file_type, extensions in SUPPORTED_FORMATS.items():
        if suffix in extensions:
            return file_type
    return None


class DocxConverter:
    """Convert docx/doc files to Markdown with image extraction."""

    def __init__(self, cache_dir: Optional[str] = None, image_optimizer: Optional[ImageOptimizer] = None):
        """
        Initialize the converter.

        Args:
            cache_dir: Directory to store converted files and images.
                      Defaults to system temp directory.
            image_optimizer: Optional image optimizer instance.
        """
        if cache_dir:
            self.cache_dir = Path(cache_dir)
        else:
            self.cache_dir = Path(tempfile.gettempdir()) / "officereader_mcp_cache"

        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.images_dir = self.cache_dir / "images"
        self.images_dir.mkdir(exist_ok=True)
        self.output_dir = self.cache_dir / "output"
        self.output_dir.mkdir(exist_ok=True)

        # Image optimizer for compression
        self.image_optimizer = image_optimizer or ImageOptimizer()

    def convert_file(
        self,
        file_path: str,
        extract_images: bool = True,
        image_format: str = "file",  # "file", "base64", or "both"
        output_name: Optional[str] = None,
    ) -> dict:
        """
        Convert a docx/doc file to Markdown.

        Args:
            file_path: Path to the docx/doc file
            extract_images: Whether to extract images
            image_format: How to handle images - "file" (save to disk),
                         "base64" (embed in markdown), or "both"
            output_name: Custom name for output file (without extension)

        Returns:
            dict with keys:
                - markdown: The converted markdown content
                - output_path: Path to the saved markdown file
                - images: List of extracted image paths
                - metadata: Document metadata
        """
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        suffix = file_path.suffix.lower()
        if suffix not in [".docx", ".doc"]:
            raise ValueError(f"Unsupported file format: {suffix}. Only .docx and .doc are supported.")

        # Generate unique output name based on file hash
        if not output_name:
            file_hash = self._get_file_hash(file_path)
            output_name = f"{file_path.stem}_{file_hash[:8]}"

        # Create output subdirectory for this conversion
        conversion_dir = self.output_dir / output_name
        conversion_dir.mkdir(exist_ok=True)
        conversion_images_dir = conversion_dir / "images"
        conversion_images_dir.mkdir(exist_ok=True)

        # Convert based on file type
        if suffix == ".docx":
            result = self._convert_docx(
                file_path,
                extract_images,
                image_format,
                conversion_dir,
                conversion_images_dir
            )
        else:
            # For .doc files, use mammoth which can handle older formats
            result = self._convert_doc_with_mammoth(
                file_path,
                extract_images,
                image_format,
                conversion_dir,
                conversion_images_dir
            )

        # Save markdown to file
        md_path = conversion_dir / f"{output_name}.md"
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(result["markdown"])

        result["output_path"] = str(md_path)
        result["conversion_dir"] = str(conversion_dir)

        return result

    def _get_file_hash(self, file_path: Path) -> str:
        """Generate MD5 hash of file for unique identification."""
        hasher = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def _convert_docx(
        self,
        file_path: Path,
        extract_images: bool,
        image_format: str,
        conversion_dir: Path,
        images_dir: Path,
    ) -> dict:
        """Convert docx file using python-docx for better control."""
        doc = Document(file_path)

        markdown_parts = []
        images = []
        image_counter = 0

        # Extract metadata
        metadata = self._extract_metadata(doc)

        # Process document body
        for element in doc.element.body:
            if element.tag.endswith("p"):
                # Paragraph
                para = Paragraph(element, doc)
                md_text, new_images = self._process_paragraph(
                    para, doc, extract_images, image_format,
                    images_dir, image_counter
                )
                if md_text:
                    markdown_parts.append(md_text)
                images.extend(new_images)
                image_counter += len(new_images)

            elif element.tag.endswith("tbl"):
                # Table
                table = Table(element, doc)
                md_table = self._process_table(table)
                if md_table:
                    markdown_parts.append(md_table)

        markdown = "\n\n".join(markdown_parts)

        # Clean up markdown
        markdown = self._clean_markdown(markdown)

        return {
            "markdown": markdown,
            "images": images,
            "metadata": metadata,
        }

    def _convert_doc_with_mammoth(
        self,
        file_path: Path,
        extract_images: bool,
        image_format: str,
        conversion_dir: Path,
        images_dir: Path,
    ) -> dict:
        """Convert doc/docx file using mammoth for broader compatibility."""
        images = []
        image_counter = [0]  # Use list to allow modification in closure

        def handle_image(image):
            """Handle image extraction during mammoth conversion."""
            nonlocal images
            image_counter[0] += 1

            with image.open() as img_stream:
                img_data = img_stream.read()

            # Determine image extension
            content_type = image.content_type
            ext_map = {
                "image/png": ".png",
                "image/jpeg": ".jpg",
                "image/gif": ".gif",
                "image/bmp": ".bmp",
                "image/webp": ".webp",
            }
            ext = ext_map.get(content_type, ".png")

            img_name = f"image_{image_counter[0]:03d}{ext}"
            img_path = images_dir / img_name

            result_parts = []

            if image_format in ["file", "both"]:
                with open(img_path, "wb") as f:
                    f.write(img_data)
                images.append(str(img_path))
                result_parts.append(f"![image](images/{img_name})")

            if image_format in ["base64", "both"]:
                b64_data = base64.b64encode(img_data).decode("utf-8")
                if image_format == "base64":
                    result_parts.append(f"![image](data:{content_type};base64,{b64_data})")

            return {"html": f'<img src="images/{img_name}" alt="image" />'}

        # Convert with mammoth
        with open(file_path, "rb") as f:
            result = mammoth.convert_to_html(
                f,
                convert_image=mammoth.images.img_element(handle_image) if extract_images else None
            )

        html_content = result.value

        # Convert HTML to Markdown
        markdown = self._html_to_markdown(html_content)

        # Extract metadata (limited for mammoth conversion)
        metadata = {
            "source": str(file_path),
            "converted_at": datetime.now().isoformat(),
            "warnings": result.messages,
        }

        return {
            "markdown": markdown,
            "images": images,
            "metadata": metadata,
        }

    def _extract_metadata(self, doc: DocumentType) -> dict:
        """Extract document metadata."""
        core_props = doc.core_properties
        return {
            "title": core_props.title or "",
            "author": core_props.author or "",
            "created": str(core_props.created) if core_props.created else "",
            "modified": str(core_props.modified) if core_props.modified else "",
            "subject": core_props.subject or "",
            "keywords": core_props.keywords or "",
        }

    def _process_paragraph(
        self,
        para: Paragraph,
        doc: DocumentType,
        extract_images: bool,
        image_format: str,
        images_dir: Path,
        image_counter: int,
    ) -> tuple[str, list]:
        """Process a paragraph and convert to Markdown."""
        images = []

        # Check for heading style
        style_name = para.style.name if para.style else ""
        heading_level = 0

        if style_name.startswith("Heading"):
            try:
                heading_level = int(style_name.replace("Heading", "").strip())
            except ValueError:
                heading_level = 0

        # Process runs (text segments with formatting)
        text_parts = []
        for run in para.runs:
            text = run.text
            if not text:
                continue

            # Apply formatting
            if run.bold:
                text = f"**{text}**"
            if run.italic:
                text = f"*{text}*"
            if run.underline:
                text = f"<u>{text}</u>"
            if run.font.strike:
                text = f"~~{text}~~"

            text_parts.append(text)

        # Check for images in paragraph
        if extract_images:
            for run in para.runs:
                for drawing in run.element.findall(".//a:blip",
                    {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}):
                    # Get image relationship
                    embed_id = drawing.get(qn("r:embed"))
                    if embed_id:
                        try:
                            image_part = doc.part.related_parts.get(embed_id)
                            if image_part:
                                image_counter += 1
                                img_data = image_part.blob
                                content_type = image_part.content_type

                                ext_map = {
                                    "image/png": ".png",
                                    "image/jpeg": ".jpg",
                                    "image/gif": ".gif",
                                    "image/bmp": ".bmp",
                                }
                                ext = ext_map.get(content_type, ".png")

                                img_name = f"image_{image_counter:03d}{ext}"
                                img_path = images_dir / img_name

                                if image_format in ["file", "both"]:
                                    with open(img_path, "wb") as f:
                                        f.write(img_data)
                                    images.append(str(img_path))

                                if image_format == "base64":
                                    b64_data = base64.b64encode(img_data).decode("utf-8")
                                    text_parts.append(f"\n![image](data:{content_type};base64,{b64_data})\n")
                                else:
                                    text_parts.append(f"\n![image](images/{img_name})\n")
                        except Exception:
                            pass  # Skip problematic images

        full_text = "".join(text_parts).strip()

        if not full_text:
            return "", images

        # Apply heading formatting
        if heading_level > 0:
            full_text = f"{'#' * heading_level} {full_text}"

        # Check for list items
        if para._element.pPr is not None:
            numPr = para._element.pPr.find(qn("w:numPr"))
            if numPr is not None:
                ilvl = numPr.find(qn("w:ilvl"))
                level = int(ilvl.get(qn("w:val"))) if ilvl is not None else 0
                indent = "  " * level
                full_text = f"{indent}- {full_text}"

        return full_text, images

    def _process_table(self, table: Table) -> str:
        """Convert a table to Markdown format."""
        if not table.rows:
            return ""

        md_rows = []

        for i, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                # Get cell text, handling merged cells
                cell_text = cell.text.strip().replace("\n", " ").replace("|", "\\|")
                cells.append(cell_text)

            md_rows.append("| " + " | ".join(cells) + " |")

            # Add header separator after first row
            if i == 0:
                separator = "| " + " | ".join(["---"] * len(cells)) + " |"
                md_rows.append(separator)

        return "\n".join(md_rows)

    def _html_to_markdown(self, html: str) -> str:
        """Convert HTML to Markdown."""
        from markdownify import markdownify as md

        markdown = md(html, heading_style="ATX", bullets="-")
        return self._clean_markdown(markdown)

    def _clean_markdown(self, markdown: str) -> str:
        """Clean up and normalize markdown output."""
        # Remove excessive blank lines
        markdown = re.sub(r"\n{3,}", "\n\n", markdown)

        # Fix spacing around headers
        markdown = re.sub(r"(#{1,6})\s*\n+", r"\1 ", markdown)

        # Clean up list formatting
        markdown = re.sub(r"\n\s*-\s*\n", "\n", markdown)

        # Remove trailing whitespace on lines
        lines = [line.rstrip() for line in markdown.split("\n")]
        markdown = "\n".join(lines)

        return markdown.strip()

    def get_cache_info(self) -> dict:
        """Get information about cached conversions."""
        conversions = []
        total_size = 0

        if not self.output_dir.exists():
            return {
                "cache_dir": str(self.cache_dir),
                "output_dir": str(self.output_dir),
                "conversions": [],
                "total_conversions": 0,
                "total_size_bytes": 0,
                "total_size_human": "0 B",
            }

        for item in self.output_dir.iterdir():
            if item.is_dir():
                md_files = list(item.glob("*.md"))
                images_dir = item / "images"
                images = list(images_dir.glob("*")) if images_dir.exists() else []

                # Calculate sizes
                md_size = sum(f.stat().st_size for f in md_files if f.exists())
                img_size = sum(f.stat().st_size for f in images if f.exists())
                dir_size = md_size + img_size
                total_size += dir_size

                # Get modification time
                mod_time = ""
                if md_files:
                    mod_time = datetime.fromtimestamp(md_files[0].stat().st_mtime).isoformat()

                conversions.append({
                    "name": item.name,
                    "path": str(item),
                    "markdown_files": [str(f) for f in md_files],
                    "image_count": len(images),
                    "image_paths": [str(f) for f in images],
                    "size_bytes": dir_size,
                    "size_human": self._human_readable_size(dir_size),
                    "modified": mod_time,
                })

        return {
            "cache_dir": str(self.cache_dir),
            "output_dir": str(self.output_dir),
            "conversions": conversions,
            "total_conversions": len(conversions),
            "total_size_bytes": total_size,
            "total_size_human": self._human_readable_size(total_size),
        }

    def _human_readable_size(self, size_bytes: int) -> str:
        """Convert bytes to human readable format."""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"

    def clear_cache(self) -> dict:
        """Clear all cached conversions."""
        import shutil

        cleared = []
        for item in self.output_dir.iterdir():
            if item.is_dir():
                shutil.rmtree(item)
                cleared.append(str(item))

        return {
            "cleared": cleared,
            "count": len(cleared),
        }


class OfficeConverter:
    """
    Unified converter for all Office document types.
    Automatically detects file type and uses appropriate converter.
    """

    def __init__(self, cache_dir: Optional[str] = None):
        """
        Initialize the unified Office converter.

        Args:
            cache_dir: Directory for caching outputs
        """
        if cache_dir:
            self.cache_dir = Path(cache_dir)
        else:
            self.cache_dir = Path(tempfile.gettempdir()) / "officereader_mcp_cache"

        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir = self.cache_dir / "output"
        self.output_dir.mkdir(exist_ok=True)

        # Initialize image optimizer
        self.image_optimizer = ImageOptimizer(
            max_width=1920,
            max_height=1080,
            jpeg_quality=85,
            webp_quality=80,
            prefer_webp=True,
        )

        # Initialize individual converters
        self._docx_converter = DocxConverter(
            cache_dir=str(self.cache_dir),
            image_optimizer=self.image_optimizer
        )

        # Lazy-load Excel and PowerPoint converters
        self._excel_converter = None
        self._pptx_converter = None

    @property
    def excel_converter(self):
        """Lazy-load Excel converter."""
        if self._excel_converter is None:
            from .excel_converter import ExcelConverter
            self._excel_converter = ExcelConverter(
                self.cache_dir,
                self.image_optimizer
            )
        return self._excel_converter

    @property
    def pptx_converter(self):
        """Lazy-load PowerPoint converter."""
        if self._pptx_converter is None:
            from .pptx_converter import PowerPointConverter
            self._pptx_converter = PowerPointConverter(
                self.cache_dir,
                self.image_optimizer
            )
        return self._pptx_converter

    def convert_file(
        self,
        file_path: str,
        extract_images: bool = True,
        image_format: str = "file",
        output_name: Optional[str] = None,
    ) -> dict:
        """
        Convert any supported Office document to Markdown.

        Args:
            file_path: Path to the document
            extract_images: Whether to extract images
            image_format: "file", "base64", or "both"
            output_name: Custom output name

        Returns:
            dict with markdown, images, metadata, and paths
        """
        file_path = Path(file_path)

        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        file_type = get_file_type(file_path)
        if file_type is None:
            supported = []
            for exts in SUPPORTED_FORMATS.values():
                supported.extend(exts)
            raise ValueError(
                f"Unsupported file format: {file_path.suffix}. "
                f"Supported formats: {', '.join(supported)}"
            )

        # Generate output name
        if not output_name:
            file_hash = self._get_file_hash(file_path)
            output_name = f"{file_path.stem}_{file_hash[:8]}"

        # Create output directory
        conversion_dir = self.output_dir / output_name
        conversion_dir.mkdir(exist_ok=True)
        images_dir = conversion_dir / "images"
        images_dir.mkdir(exist_ok=True)

        # Convert based on file type
        if file_type == "word":
            result = self._docx_converter.convert_file(
                str(file_path),
                extract_images=extract_images,
                image_format=image_format,
                output_name=output_name,
            )
        elif file_type == "excel":
            result = self.excel_converter.convert_file(
                str(file_path),
                extract_images=extract_images,
                image_format=image_format,
                output_name=output_name,
                conversion_dir=conversion_dir,
                images_dir=images_dir,
            )
        elif file_type == "powerpoint":
            result = self.pptx_converter.convert_file(
                str(file_path),
                extract_images=extract_images,
                image_format=image_format,
                output_name=output_name,
                conversion_dir=conversion_dir,
                images_dir=images_dir,
            )
        else:
            raise ValueError(f"Unknown file type: {file_type}")

        # Save markdown if not already saved (for Excel/PPT)
        md_path = conversion_dir / f"{output_name}.md"
        if "output_path" not in result:
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(result["markdown"])
            result["output_path"] = str(md_path)
            result["conversion_dir"] = str(conversion_dir)

        # Add file type info
        result["file_type"] = file_type

        return result

    def _get_file_hash(self, file_path: Path) -> str:
        """Generate MD5 hash of file for unique identification."""
        hasher = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def get_cache_info(self) -> dict:
        """Get information about cached conversions."""
        return self._docx_converter.get_cache_info()

    def clear_cache(self) -> dict:
        """Clear all cached conversions."""
        return self._docx_converter.clear_cache()

    @staticmethod
    def get_supported_formats() -> dict:
        """Return supported file formats."""
        return SUPPORTED_FORMATS.copy()

"""
Image optimization utility for reducing file size while maintaining quality.
"""

from io import BytesIO
from typing import Tuple, Optional

from PIL import Image


class ImageOptimizer:
    """Optimize images for reduced file size while maintaining readability."""

    def __init__(
        self,
        max_width: int = 1920,
        max_height: int = 1080,
        jpeg_quality: int = 85,
        webp_quality: int = 80,
        prefer_webp: bool = True,
        size_threshold_kb: int = 100,
    ):
        """
        Initialize the image optimizer.

        Args:
            max_width: Maximum image width (larger images will be resized)
            max_height: Maximum image height (larger images will be resized)
            jpeg_quality: JPEG compression quality (0-100)
            webp_quality: WebP compression quality (0-100)
            prefer_webp: Use WebP format when possible (better compression)
            size_threshold_kb: Compress images larger than this size (in KB)
        """
        self.max_width = max_width
        self.max_height = max_height
        self.jpeg_quality = jpeg_quality
        self.webp_quality = webp_quality
        self.prefer_webp = prefer_webp
        self.size_threshold_kb = size_threshold_kb

    def optimize_image(
        self,
        image: Image.Image,
        name_hint: str = "",
        force_format: Optional[str] = None,
    ) -> Tuple[bytes, str]:
        """
        Optimize an image for reduced size.

        Args:
            image: PIL Image object
            name_hint: Hint for the image name (for logging)
            force_format: Force specific output format ("png", "jpg", "webp")

        Returns:
            Tuple of (optimized_bytes, extension)
        """
        original_mode = image.mode
        original_size = image.size

        # Convert RGBA to RGB for JPEG (JPEG doesn't support transparency)
        has_transparency = original_mode in ("RGBA", "LA", "P")

        # Resize if needed
        image = self._resize_if_needed(image)

        # Determine best format
        if force_format:
            output_format = force_format.lower()
        else:
            output_format = self._choose_format(image, has_transparency)

        # Optimize based on format
        if output_format == "webp":
            return self._save_as_webp(image, has_transparency)
        elif output_format == "png":
            return self._save_as_png(image)
        else:
            return self._save_as_jpeg(image, has_transparency)

    def _resize_if_needed(self, image: Image.Image) -> Image.Image:
        """Resize image if it exceeds maximum dimensions."""
        width, height = image.size

        if width <= self.max_width and height <= self.max_height:
            return image

        # Calculate new size maintaining aspect ratio
        ratio = min(self.max_width / width, self.max_height / height)
        new_width = int(width * ratio)
        new_height = int(height * ratio)

        # Use high-quality resampling
        return image.resize((new_width, new_height), Image.Resampling.LANCZOS)

    def _choose_format(self, image: Image.Image, has_transparency: bool) -> str:
        """
        Choose optimal format based on image characteristics.

        Strategy:
        - Screenshots/graphics with transparency -> PNG or WebP
        - Photos -> JPEG or WebP
        - Small images with few colors -> PNG
        """
        # Check if image is likely a photo or graphics
        is_photo = self._is_photo_like(image)

        if has_transparency:
            # Need format that supports transparency
            return "webp" if self.prefer_webp else "png"

        if is_photo:
            # Photos compress better with lossy formats
            return "webp" if self.prefer_webp else "jpg"

        # Graphics/screenshots
        if self.prefer_webp:
            return "webp"

        # For simple graphics, PNG might be better
        return "png"

    def _is_photo_like(self, image: Image.Image) -> bool:
        """
        Determine if image is photo-like (continuous tones) or graphics-like.

        Photos have many unique colors, graphics tend to have fewer.
        """
        # Convert to RGB for color counting
        if image.mode != "RGB":
            test_image = image.convert("RGB")
        else:
            test_image = image

        # Sample a portion of the image for speed
        width, height = test_image.size
        sample_size = min(100, width) * min(100, height)

        # Get colors (with limit)
        colors = test_image.getcolors(maxcolors=1000)

        if colors is None:
            # More than 1000 colors - likely a photo
            return True

        # Many unique colors relative to size suggests photo
        unique_ratio = len(colors) / sample_size if sample_size > 0 else 0
        return unique_ratio > 0.3

    def _save_as_webp(
        self, image: Image.Image, has_transparency: bool
    ) -> Tuple[bytes, str]:
        """Save image as WebP format."""
        buffer = BytesIO()

        if has_transparency:
            # Ensure RGBA mode for transparency
            if image.mode != "RGBA":
                image = image.convert("RGBA")
            image.save(
                buffer,
                format="WEBP",
                quality=self.webp_quality,
                lossless=False,
                method=4,  # Balanced compression
            )
        else:
            if image.mode == "RGBA":
                image = image.convert("RGB")
            image.save(
                buffer,
                format="WEBP",
                quality=self.webp_quality,
                lossless=False,
                method=4,
            )

        return buffer.getvalue(), ".webp"

    def _save_as_png(self, image: Image.Image) -> Tuple[bytes, str]:
        """Save image as PNG format with optimization."""
        buffer = BytesIO()

        # PNG is lossless, use compression
        image.save(
            buffer,
            format="PNG",
            optimize=True,
            compress_level=6,  # Balanced speed/size
        )

        return buffer.getvalue(), ".png"

    def _save_as_jpeg(
        self, image: Image.Image, has_transparency: bool
    ) -> Tuple[bytes, str]:
        """Save image as JPEG format."""
        buffer = BytesIO()

        # JPEG doesn't support transparency - convert to RGB
        if image.mode in ("RGBA", "LA"):
            # Create white background and paste image
            background = Image.new("RGB", image.size, (255, 255, 255))
            if image.mode == "RGBA":
                background.paste(image, mask=image.split()[3])
            else:
                background.paste(image)
            image = background
        elif image.mode != "RGB":
            image = image.convert("RGB")

        image.save(
            buffer,
            format="JPEG",
            quality=self.jpeg_quality,
            optimize=True,
            progressive=True,  # Better for web
        )

        return buffer.getvalue(), ".jpg"

    def optimize_bytes(
        self,
        image_bytes: bytes,
        name_hint: str = "",
        force_format: Optional[str] = None,
    ) -> Tuple[bytes, str]:
        """
        Optimize image from bytes.

        Args:
            image_bytes: Raw image bytes
            name_hint: Hint for logging
            force_format: Force specific format

        Returns:
            Tuple of (optimized_bytes, extension)
        """
        image = Image.open(BytesIO(image_bytes))
        return self.optimize_image(image, name_hint, force_format)

    def get_compression_ratio(
        self, original_bytes: bytes, optimized_bytes: bytes
    ) -> float:
        """Calculate compression ratio."""
        if len(original_bytes) == 0:
            return 1.0
        return len(optimized_bytes) / len(original_bytes)

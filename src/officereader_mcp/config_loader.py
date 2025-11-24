"""
Configuration loader for OfficeReader-MCP.
Loads settings from config.json file or environment variables.
"""

import json
import os
from pathlib import Path
from typing import Optional, Dict, Any


class Config:
    """Configuration manager for OfficeReader-MCP."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration.

        Args:
            config_path: Path to config.json file. If not provided, searches in:
                        1. Current directory
                        2. Script directory
                        3. User home directory
        """
        self.config_data: Dict[str, Any] = {}
        self.config_path: Optional[Path] = None

        # Try to find and load config file
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = self._find_config_file()

        if self.config_path and self.config_path.exists():
            self._load_config()

    def _find_config_file(self) -> Optional[Path]:
        """Find config.json in standard locations."""
        search_paths = [
            Path.cwd() / "config.json",  # Current directory
            Path(__file__).parent.parent.parent / "config.json",  # Project root
            Path.home() / ".officereader-mcp" / "config.json",  # User home
        ]

        for path in search_paths:
            if path.exists():
                return path

        return None

    def _load_config(self):
        """Load configuration from JSON file."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config_data = json.load(f)
        except Exception as e:
            print(f"Warning: Failed to load config file: {e}")
            self.config_data = {}

    def get_cache_dir(self) -> Optional[str]:
        """
        Get cache directory from config or environment.

        Priority:
        1. OFFICEREADER_CACHE_DIR environment variable
        2. cache_dir from config.json
        3. None (will use default temp directory)
        """
        # Environment variable takes highest priority
        env_cache = os.environ.get("OFFICEREADER_CACHE_DIR")
        if env_cache:
            return env_cache

        # Then check config file
        cache_dir = self.config_data.get("cache_dir")
        if cache_dir:
            # Convert relative path to absolute
            cache_path = Path(cache_dir)
            if not cache_path.is_absolute() and self.config_path:
                # Relative to config file location
                cache_path = self.config_path.parent / cache_dir
            return str(cache_path)

        return None

    def get_output_dir(self) -> Optional[str]:
        """Get custom output directory if specified."""
        output_dir = self.config_data.get("output_dir")
        if output_dir:
            output_path = Path(output_dir)
            if not output_path.is_absolute() and self.config_path:
                output_path = self.config_path.parent / output_dir
            return str(output_path)
        return None

    def get_image_optimization_settings(self) -> Dict[str, Any]:
        """Get image optimization settings."""
        defaults = {
            "enabled": True,
            "max_dimension": 1920,
            "quality": 80,
        }

        settings = self.config_data.get("image_optimization", {})
        return {**defaults, **settings}

    def get_default_settings(self) -> Dict[str, Any]:
        """Get default conversion settings."""
        defaults = {
            "extract_images": True,
            "image_format": "file",
        }

        settings = self.config_data.get("default_settings", {})
        return {**defaults, **settings}

    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value by key."""
        return self.config_data.get(key, default)

    def __repr__(self) -> str:
        """String representation of config."""
        return f"Config(path={self.config_path}, data={self.config_data})"


# Global config instance
_config: Optional[Config] = None


def get_config(config_path: Optional[str] = None) -> Config:
    """
    Get or create the global configuration instance.

    Args:
        config_path: Optional path to config file

    Returns:
        Config instance
    """
    global _config
    if _config is None:
        _config = Config(config_path)
    return _config


def load_config(config_path: Optional[str] = None) -> Config:
    """
    Force reload configuration.

    Args:
        config_path: Optional path to config file

    Returns:
        New Config instance
    """
    global _config
    _config = Config(config_path)
    return _config

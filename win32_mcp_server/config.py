"""
Server configuration with sensible defaults.

All settings can be overridden at runtime via the configure tool
or by modifying the config singleton.
"""

from dataclasses import dataclass, field
from typing import Literal


#: Valid OCR preprocessing modes — canonical definition, re-exported by utils.imaging
PreprocessMode = Literal["auto", "light_bg", "dark_bg", "high_contrast", "none"]

VALID_PREPROCESS_MODES = {"auto", "light_bg", "dark_bg", "high_contrast", "none"}


@dataclass
class OCRConfig:
    """OCR engine configuration."""
    lang: str = "eng"
    confidence_threshold: int = 60
    preprocess_mode: PreprocessMode = "auto"
    scale_small_images: bool = True
    min_dimension_for_scaling: int = 500
    upscale_factor: int = 2


@dataclass
class CaptureConfig:
    """Screenshot capture configuration."""
    default_format: str = "png"  # png | jpeg | webp
    default_quality: int = 85  # 1-100, for jpeg/webp
    default_scale: float = 1.0  # 0.1-1.0


@dataclass
class AutomationConfig:
    """Mouse/keyboard automation configuration."""
    click_delay: float = 0.05  # seconds after click
    type_interval: float = 0.01  # seconds between keystrokes
    drag_duration: float = 0.5  # seconds for drag operations
    move_duration: float = 0.25  # seconds for mouse move


@dataclass
class ServerConfig:
    """Top-level server configuration."""
    ocr: OCRConfig = field(default_factory=OCRConfig)
    capture: CaptureConfig = field(default_factory=CaptureConfig)
    automation: AutomationConfig = field(default_factory=AutomationConfig)
    validate_coordinates: bool = True
    default_timeout: float = 10.0
    window_retry_attempts: int = 3
    window_retry_delay: float = 0.3
    min_operation_interval: float = 0.05
    fuzzy_match_threshold: int = 65
    debug: bool = False


# Module-level singleton — importable from anywhere
config = ServerConfig()

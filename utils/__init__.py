"""
PowerPoint utilities package.
Organized utility functions for PowerPoint manipulation.
"""

from .core_utils import *
from .presentation_utils import *
from .content_utils import *
from .design_utils import *
from .validation_utils import *

__all__ = [
    # Core utilities
    "safe_operation",
    "try_multiple_approaches",
    
    # Presentation utilities
    "create_presentation",
    "open_presentation",
    "save_presentation",
    "create_presentation_from_template",
    "get_presentation_info",
    "get_template_info",
    "set_core_properties",
    "get_core_properties",
    "move_slide",
    "swap_slides",
    "reorder_slides",
    "delete_slide",
    "delete_slides",
    "duplicate_slide",
    
    # Content utilities
    "add_slide",
    "get_slide_info",
    "set_title",
    "populate_placeholder",
    "add_bullet_points",
    "add_textbox",
    "format_text",
    "format_text_advanced",
    "add_image",
    "add_table",
    "format_table_cell",
    "add_chart",
    "format_chart",
    "get_shape_info",
    "find_shapes_by_type",
    "get_all_textboxes",
    "format_keywords_in_text",
    "extract_slide_text_content",
    
    # Design utilities
    "get_professional_color",
    "get_professional_font",
    "get_color_schemes",
    "add_professional_slide",
    "apply_professional_theme",
    "enhance_existing_slide",
    "apply_professional_image_enhancement",
    "enhance_image_with_pillow",
    "set_slide_gradient_background",
    "create_professional_gradient_background",
    "format_shape",
    "apply_picture_shadow",
    "apply_picture_reflection",
    "apply_picture_glow",
    "apply_picture_soft_edges",
    "apply_picture_rotation",
    "apply_picture_transparency",
    "apply_picture_bevel",
    "apply_picture_filter",
    "analyze_font_file",
    "optimize_font_for_presentation",
    "get_font_recommendations",
    "copy_slide_format",
    "apply_text_style_to_all",
    
    # Validation utilities
    "validate_text_fit",
    "validate_and_fix_slide"
]
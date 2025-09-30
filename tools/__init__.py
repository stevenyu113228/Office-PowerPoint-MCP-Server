"""
Tools package for PowerPoint MCP Server.
Organizes tools into logical modules for better maintainability.
"""

from .presentation_tools import register_presentation_tools
from .content_tools import register_content_tools
from .structural_tools import register_structural_tools
from .professional_tools import register_professional_tools
from .template_tools import register_template_tools
from .hyperlink_tools import register_hyperlink_tools
from .chart_tools import register_chart_tools
from .connector_tools import register_connector_tools
from .master_tools import register_master_tools
from .transition_tools import register_transition_tools
from .notes_tools import register_notes_tools
from .shape_positioning_tools import register_shape_positioning_tools
from .shape_alignment_tools import register_shape_alignment_tools
from .slide_management_tools import register_slide_management_tools

__all__ = [
    "register_presentation_tools",
    "register_content_tools",
    "register_structural_tools",
    "register_professional_tools",
    "register_template_tools",
    "register_hyperlink_tools",
    "register_chart_tools",
    "register_connector_tools",
    "register_master_tools",
    "register_transition_tools",
    "register_notes_tools",
    "register_shape_positioning_tools",
    "register_shape_alignment_tools",
    "register_slide_management_tools"
]
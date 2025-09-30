"""
Slide management tools for PowerPoint MCP Server.
Handles slide reordering, duplication, search, and batch operations.
"""
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp import FastMCP
import utils as ppt_utils
import re


def register_slide_management_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register slide management tools with the FastMCP app"""

    # ===== Slide Deletion Tools =====

    @app.tool()
    def delete_slide(
        slide_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Delete a slide from the presentation.

        Args:
            slide_index: Index of the slide to delete (0-based)
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with deletion results

        Example:
            # Delete slide 3
            delete_slide(slide_index=3)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.delete_slide(pres, slide_index)
            return result
        except Exception as e:
            return {
                "error": f"Failed to delete slide: {str(e)}"
            }

    @app.tool()
    def delete_slides(
        slide_indices: List[int],
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Delete multiple slides from the presentation.

        Args:
            slide_indices: List of slide indices to delete (0-based)
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with deletion results

        Example:
            # Delete slides 2, 5, and 7
            delete_slides(slide_indices=[2, 5, 7])
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.delete_slides(pres, slide_indices)
            return result
        except Exception as e:
            return {
                "error": f"Failed to delete slides: {str(e)}"
            }

    # ===== Slide Reordering Tools =====

    @app.tool()
    def move_slide(
        from_index: int,
        to_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Move a slide from one position to another.

        Args:
            from_index: Current index of the slide (0-based)
            to_index: Target index for the slide (0-based)
            presentation_id: Optional presentation ID

        Example:
            # Move slide 5 to position 2
            move_slide(from_index=5, to_index=2)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.move_slide(pres, from_index, to_index)
            return result
        except Exception as e:
            return {
                "error": f"Failed to move slide: {str(e)}"
            }

    @app.tool()
    def swap_slides(
        index_a: int,
        index_b: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Swap two slides.

        Args:
            index_a: Index of first slide
            index_b: Index of second slide
            presentation_id: Optional presentation ID

        Example:
            # Swap slides at positions 2 and 5
            swap_slides(index_a=2, index_b=5)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.swap_slides(pres, index_a, index_b)
            return result
        except Exception as e:
            return {
                "error": f"Failed to swap slides: {str(e)}"
            }

    @app.tool()
    def reorder_slides(
        new_order: List[int],
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Reorder slides according to a new arrangement.

        Args:
            new_order: List of slide indices in desired order
            presentation_id: Optional presentation ID

        Example:
            # Rearrange: slide 0 stays, slide 3 becomes position 1, etc.
            reorder_slides(new_order=[0, 3, 1, 2, 4])
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.reorder_slides(pres, new_order)
            return result
        except Exception as e:
            return {
                "error": f"Failed to reorder slides: {str(e)}"
            }

    # ===== Slide Duplication Tools =====

    @app.tool()
    def duplicate_slide(
        slide_index: int,
        insert_position: Optional[int] = None,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Duplicate a slide within the presentation.

        Args:
            slide_index: Index of the slide to duplicate
            insert_position: Where to insert duplicate (None = after original)
            presentation_id: Optional presentation ID

        Example:
            # Duplicate slide 5 and insert it at position 6
            duplicate_slide(slide_index=5, insert_position=6)
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.duplicate_slide(pres, slide_index, insert_position)
            return result
        except Exception as e:
            return {
                "error": f"Failed to duplicate slide: {str(e)}"
            }

    # ===== Text Search and Replace Tools =====

    @app.tool()
    def find_slides_by_text(
        search_text: str,
        match_type: str = "contains",  # "contains", "exact", "regex"
        search_in: str = "all",  # "title", "body", "all"
        case_sensitive: bool = False,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Find slides containing specific text.

        Args:
            search_text: Text to search for
            match_type: Type of match ("contains", "exact", "regex")
            search_in: Where to search ("title", "body", "all")
            case_sensitive: Whether search is case-sensitive
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with matching slides and their details

        Example:
            # Find all slides containing "Terraform"
            find_slides_by_text(search_text="Terraform", match_type="contains")
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            matches = []
            flags = 0 if case_sensitive else re.IGNORECASE

            for slide_index, slide in enumerate(pres.slides):
                slide_title = ""
                slide_body_text = []
                match_count = 0

                # Extract title
                if hasattr(slide.shapes, 'title') and slide.shapes.title:
                    try:
                        slide_title = slide.shapes.title.text_frame.text
                    except:
                        pass

                # Extract body text
                for shape in slide.shapes:
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        try:
                            text = shape.text_frame.text
                            if text and text != slide_title:
                                slide_body_text.append(text)
                        except:
                            pass

                # Search based on criteria
                search_targets = []
                if search_in in ["title", "all"] and slide_title:
                    search_targets.append(("title", slide_title))
                if search_in in ["body", "all"]:
                    for text in slide_body_text:
                        search_targets.append(("body", text))

                # Perform search
                for location, text in search_targets:
                    if match_type == "contains":
                        if case_sensitive:
                            if search_text in text:
                                match_count += text.count(search_text)
                        else:
                            if search_text.lower() in text.lower():
                                match_count += text.lower().count(search_text.lower())
                    elif match_type == "exact":
                        if case_sensitive:
                            if text == search_text:
                                match_count += 1
                        else:
                            if text.lower() == search_text.lower():
                                match_count += 1
                    elif match_type == "regex":
                        pattern = re.compile(search_text, flags)
                        found = pattern.findall(text)
                        match_count += len(found)

                if match_count > 0:
                    matches.append({
                        "slide_index": slide_index,
                        "title": slide_title if slide_title else "(No title)",
                        "layout_name": slide.slide_layout.name,
                        "match_count": match_count
                    })

            return {
                "success": True,
                "search_text": search_text,
                "match_type": match_type,
                "total_matches": len(matches),
                "matches": matches
            }

        except Exception as e:
            return {
                "error": f"Failed to search slides: {str(e)}"
            }

    @app.tool()
    def find_slides_by_layout(
        layout_name: str,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Find all slides using a specific layout.

        Args:
            layout_name: Name of the layout to search for
            presentation_id: Optional presentation ID

        Example:
            # Find all slides using "Title Slide" layout
            find_slides_by_layout(layout_name="Title Slide")
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            matches = []

            for slide_index, slide in enumerate(pres.slides):
                if slide.slide_layout.name == layout_name:
                    slide_title = ""
                    if hasattr(slide.shapes, 'title') and slide.shapes.title:
                        try:
                            slide_title = slide.shapes.title.text_frame.text
                        except:
                            pass

                    matches.append({
                        "slide_index": slide_index,
                        "title": slide_title if slide_title else "(No title)"
                    })

            return {
                "success": True,
                "layout_name": layout_name,
                "total_matches": len(matches),
                "matches": matches
            }

        except Exception as e:
            return {
                "error": f"Failed to find slides by layout: {str(e)}"
            }

    @app.tool()
    def count_slides_by_type(
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Count slides by layout type.

        Args:
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with slide count statistics by layout
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            layout_counts = {}

            for slide in pres.slides:
                layout_name = slide.slide_layout.name
                layout_counts[layout_name] = layout_counts.get(layout_name, 0) + 1

            return {
                "success": True,
                "total_slides": len(pres.slides),
                "by_layout": layout_counts
            }

        except Exception as e:
            return {
                "error": f"Failed to count slides: {str(e)}"
            }

    @app.tool()
    def replace_text_in_presentation(
        find_text: str,
        replace_text: str,
        case_sensitive: bool = False,
        whole_word_only: bool = False,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Replace text throughout the entire presentation.

        Args:
            find_text: Text to find
            replace_text: Text to replace with
            case_sensitive: Whether search is case-sensitive
            whole_word_only: Whether to match whole words only
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with replacement statistics

        Example:
            # Replace all "AWS" with "Amazon Web Services"
            replace_text_in_presentation(
                find_text="AWS",
                replace_text="Amazon Web Services"
            )
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            total_replacements = 0
            slides_affected = []

            for slide_index, slide in enumerate(pres.slides):
                slide_replacements = 0

                for shape in slide.shapes:
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                original_text = run.text

                                # Perform replacement
                                if whole_word_only:
                                    # Use regex for whole word matching
                                    pattern = r'\b' + re.escape(find_text) + r'\b'
                                    flags = 0 if case_sensitive else re.IGNORECASE
                                    new_text = re.sub(pattern, replace_text, original_text, flags=flags)
                                else:
                                    if case_sensitive:
                                        new_text = original_text.replace(find_text, replace_text)
                                    else:
                                        # Case-insensitive replacement
                                        pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                                        new_text = pattern.sub(replace_text, original_text)

                                if new_text != original_text:
                                    run.text = new_text
                                    replacements_in_run = original_text.count(find_text) if case_sensitive else original_text.lower().count(find_text.lower())
                                    slide_replacements += replacements_in_run

                if slide_replacements > 0:
                    slides_affected.append({
                        "slide_index": slide_index,
                        "replacements": slide_replacements
                    })
                    total_replacements += slide_replacements

            return {
                "success": True,
                "find_text": find_text,
                "replace_text": replace_text,
                "total_replacements": total_replacements,
                "slides_affected": len(slides_affected),
                "details": slides_affected
            }

        except Exception as e:
            return {
                "error": f"Failed to replace text: {str(e)}"
            }

    @app.tool()
    def batch_replace_text(
        replacements: Dict[str, str],
        case_sensitive: bool = False,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Perform multiple text replacements in one operation.

        Args:
            replacements: Dictionary of {find_text: replace_text} pairs
            case_sensitive: Whether search is case-sensitive
            presentation_id: Optional presentation ID

        Example:
            batch_replace_text({
                "[Repository]": "https://github.com/user/repo",
                "<env-name>": "production",
                "PLACEHOLDER": "Actual Content"
            })
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            all_results = []
            total_replacements = 0

            for find_text, replace_text in replacements.items():
                # Perform individual replacement
                result = replace_text_in_presentation(
                    find_text=find_text,
                    replace_text=replace_text,
                    case_sensitive=case_sensitive,
                    presentation_id=pres_id
                )

                if result.get("success"):
                    all_results.append({
                        "find_text": find_text,
                        "replace_text": replace_text,
                        "replacements": result["total_replacements"]
                    })
                    total_replacements += result["total_replacements"]

            return {
                "success": True,
                "total_replacements": total_replacements,
                "replacement_count": len(replacements),
                "details": all_results
            }

        except Exception as e:
            return {
                "error": f"Failed to batch replace text: {str(e)}"
            }

    # ===== Placeholder Query Tools =====

    @app.tool()
    def list_placeholders(
        slide_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        List all placeholders in a slide with their details.

        Args:
            slide_index: Index of the slide
            presentation_id: Optional presentation ID

        Returns:
            Detailed list of placeholders with indices, types, and names
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }

        slide = pres.slides[slide_index]

        try:
            placeholders = []

            for placeholder in slide.placeholders:
                placeholder_info = {
                    "idx": placeholder.placeholder_format.idx,
                    "type": str(placeholder.placeholder_format.type),
                    "name": placeholder.name
                }
                placeholders.append(placeholder_info)

            return {
                "success": True,
                "slide_index": slide_index,
                "layout_name": slide.slide_layout.name,
                "placeholder_count": len(placeholders),
                "placeholders": placeholders
            }

        except Exception as e:
            return {
                "error": f"Failed to list placeholders: {str(e)}"
            }

    @app.tool()
    def get_placeholder_by_name(
        slide_index: int,
        placeholder_name: str,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Get placeholder information by name.

        Args:
            slide_index: Index of the slide
            placeholder_name: Name of the placeholder
            presentation_id: Optional presentation ID

        Returns:
            Placeholder details including its index
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }

        slide = pres.slides[slide_index]

        try:
            for placeholder in slide.placeholders:
                if placeholder.name == placeholder_name:
                    return {
                        "success": True,
                        "slide_index": slide_index,
                        "found": True,
                        "placeholder": {
                            "idx": placeholder.placeholder_format.idx,
                            "type": str(placeholder.placeholder_format.type),
                            "name": placeholder.name
                        }
                    }

            return {
                "success": True,
                "slide_index": slide_index,
                "found": False,
                "message": f"No placeholder named '{placeholder_name}' found on slide {slide_index}"
            }

        except Exception as e:
            return {
                "error": f"Failed to get placeholder by name: {str(e)}"
            }

    # ===== Format Copying Tools =====

    @app.tool()
    def copy_slide_format(
        source_slide_index: int,
        target_slide_indices: List[int],
        copy_background: bool = True,
        copy_font_styles: bool = True,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Copy formatting from source slide to multiple target slides.

        Args:
            source_slide_index: Index of the source slide
            target_slide_indices: List of target slide indices
            copy_background: Whether to copy background formatting
            copy_font_styles: Whether to copy font styles
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with formatting results

        Example:
            # Copy formatting from slide 0 to slides 1, 2, and 3
            copy_slide_format(
                source_slide_index=0,
                target_slide_indices=[1, 2, 3],
                copy_background=True,
                copy_font_styles=True
            )
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        try:
            result = ppt_utils.copy_slide_format(
                pres,
                source_slide_index,
                target_slide_indices,
                copy_background,
                copy_font_styles
            )
            return result
        except Exception as e:
            return {
                "error": f"Failed to copy slide format: {str(e)}"
            }

    @app.tool()
    def apply_text_style_to_all(
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        font_color: Optional[List[int]] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        apply_to: str = "body",
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Apply consistent text styling to all slides in the presentation.

        Args:
            font_name: Font name to apply (e.g., "Arial", "Calibri")
            font_size: Font size in points
            font_color: RGB color as [R, G, B] list
            bold: Whether to make text bold
            italic: Whether to make text italic
            apply_to: Where to apply ("title", "body", "all")
            presentation_id: Optional presentation ID

        Returns:
            Dictionary with styling results

        Example:
            # Make all body text Arial, 14pt
            apply_text_style_to_all(
                font_name="Arial",
                font_size=14,
                apply_to="body"
            )

            # Make all titles bold and blue
            apply_text_style_to_all(
                bold=True,
                font_color=[0, 120, 215],
                apply_to="title"
            )
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[pres_id]

        # Convert font_color list to tuple if provided
        font_color_tuple = tuple(font_color) if font_color else None

        try:
            result = ppt_utils.apply_text_style_to_all(
                pres,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color_tuple,
                bold=bold,
                italic=italic,
                apply_to=apply_to
            )
            return result
        except Exception as e:
            return {
                "error": f"Failed to apply text style: {str(e)}"
            }

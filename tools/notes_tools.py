"""
Notes management tools for PowerPoint MCP Server.
Handles slide notes reading, writing, and management.
"""
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp import FastMCP


def register_notes_tools(app: FastMCP, presentations: Dict, get_current_presentation_id, validate_parameters):
    """Register notes management tools with the FastMCP app"""
    
    @app.tool()
    def get_slide_notes(
        slide_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Get notes text from a specific slide.
        
        Args:
            slide_index: Index of the slide (0-based)
            presentation_id: Optional presentation ID (uses current if not provided)
            
        Returns:
            Dictionary with notes information
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
        
        try:
            slide = pres.slides[slide_index]
            
            if slide.has_notes_slide:
                notes_text = slide.notes_slide.notes_text_frame.text
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "has_notes": True,
                    "notes_text": notes_text
                }
            else:
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "has_notes": False,
                    "notes_text": ""
                }
                
        except Exception as e:
            return {
                "error": f"Failed to get slide notes: {str(e)}"
            }
    
    @app.tool()
    def set_slide_notes(
        slide_index: int,
        notes_text: str,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Set or update notes text for a specific slide.
        
        Args:
            slide_index: Index of the slide (0-based)
            notes_text: Text content for the notes
            presentation_id: Optional presentation ID (uses current if not provided)
            
        Returns:
            Dictionary with operation results
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
        
        try:
            slide = pres.slides[slide_index]
            
            # Access or create notes slide
            if not slide.has_notes_slide:
                # This will create a notes slide if it doesn't exist
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide
            
            # Set the notes text
            notes_slide.notes_text_frame.clear()
            notes_slide.notes_text_frame.text = notes_text
            
            return {
                "success": True,
                "slide_index": slide_index,
                "notes_text": notes_text,
                "message": f"Notes updated for slide {slide_index}"
            }
                
        except Exception as e:
            return {
                "error": f"Failed to set slide notes: {str(e)}"
            }
    
    @app.tool()
    def get_all_slide_notes(
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Get notes text from all slides in the presentation.
        
        Args:
            presentation_id: Optional presentation ID (uses current if not provided)
            
        Returns:
            Dictionary with all slides' notes information
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        try:
            slides_notes = []
            
            for i, slide in enumerate(pres.slides):
                # Try to get slide title
                slide_title = ""
                try:
                    for shape in slide.shapes:
                        if hasattr(shape, 'text') and shape.text.strip():
                            slide_title = shape.text.strip().split('\n')[0]
                            break
                except:
                    pass
                
                # Get notes
                if slide.has_notes_slide:
                    notes_text = slide.notes_slide.notes_text_frame.text
                    has_notes = bool(notes_text.strip())
                else:
                    notes_text = ""
                    has_notes = False
                
                slides_notes.append({
                    "slide_index": i,
                    "slide_title": slide_title,
                    "has_notes": has_notes,
                    "notes_text": notes_text
                })
            
            return {
                "success": True,
                "presentation_id": pres_id,
                "total_slides": len(pres.slides),
                "slides_with_notes": sum(1 for slide_info in slides_notes if slide_info["has_notes"]),
                "slides_notes": slides_notes
            }
                
        except Exception as e:
            return {
                "error": f"Failed to get all slide notes: {str(e)}"
            }
    
    @app.tool()
    def clear_slide_notes(
        slide_index: int,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Clear notes text from a specific slide.
        
        Args:
            slide_index: Index of the slide (0-based)
            presentation_id: Optional presentation ID (uses current if not provided)
            
        Returns:
            Dictionary with operation results
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
        
        try:
            slide = pres.slides[slide_index]
            
            if slide.has_notes_slide:
                slide.notes_slide.notes_text_frame.clear()
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "message": f"Notes cleared for slide {slide_index}"
                }
            else:
                return {
                    "success": True,
                    "slide_index": slide_index,
                    "message": f"Slide {slide_index} had no notes to clear"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to clear slide notes: {str(e)}"
            }
    
    @app.tool()
    def append_slide_notes(
        slide_index: int,
        additional_notes: str,
        separator: str = "\n\n",
        presentation_id: Optional[str] = None
    ) -> Dict:
        """
        Append text to existing notes of a specific slide.
        
        Args:
            slide_index: Index of the slide (0-based)
            additional_notes: Text to append to existing notes
            separator: Separator between existing and new notes (default: double newline)
            presentation_id: Optional presentation ID (uses current if not provided)
            
        Returns:
            Dictionary with operation results
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
        
        try:
            slide = pres.slides[slide_index]
            
            # Get existing notes
            if slide.has_notes_slide:
                existing_notes = slide.notes_slide.notes_text_frame.text
            else:
                existing_notes = ""
            
            # Combine existing and new notes
            if existing_notes.strip():
                new_notes = existing_notes + separator + additional_notes
            else:
                new_notes = additional_notes
            
            # Set the combined notes
            if not slide.has_notes_slide:
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide
            
            notes_slide.notes_text_frame.clear()
            notes_slide.notes_text_frame.text = new_notes
            
            return {
                "success": True,
                "slide_index": slide_index,
                "previous_notes": existing_notes,
                "appended_notes": additional_notes,
                "final_notes": new_notes,
                "message": f"Notes appended to slide {slide_index}"
            }
                
        except Exception as e:
            return {
                "error": f"Failed to append slide notes: {str(e)}"
            }
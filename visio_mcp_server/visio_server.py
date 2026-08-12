"""
Visio MCP Server

This MCP server provides tools for creating and editing Visio files.
It uses the Microsoft.Office.Interop.Visio API via Python's win32com interface.
"""

import os
import sys
import json
import glob
import tempfile
import atexit
import time
import winreg
from typing import Dict, Any, List, Optional
import win32com.client
from mcp.server.fastmcp import FastMCP

# Initialize FastMCP server
mcp = FastMCP("visio-server")

# Constants
DEFAULT_SAVE_PATH = os.path.expandvars(r"%USERPROFILE%\Documents")

# Global variables
visio_app = None
open_documents = {}

def check_visio_installed():
    """Check if Visio is installed without launching it."""
    try:
        # Just check if the COM object is registered
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        return True
    except:
        return False

def get_visio_app():
    """Initialize or get the Visio Application object.
    
    Tries multiple methods to ensure Visio launches properly.
    """
    global visio_app
    
    if visio_app is None:
        try:
            # First try regular Dispatch (simplest method)
            visio_app = win32com.client.Dispatch("Visio.Application")
            # Wait for Visio to initialize
            time.sleep(1)
            visio_app.Visible = True
        except Exception as e1:
            try:
                # If Dispatch fails, try dynamic dispatch
                visio_app = win32com.client.dynamic.Dispatch("Visio.Application")
                time.sleep(1)
                visio_app.Visible = True
            except Exception as e2:
                try:
                    # Last resort: try DispatchEx
                    visio_app = win32com.client.DispatchEx("Visio.Application")
                    time.sleep(1)
                    visio_app.Visible = True
                except Exception as e3:
                    error_msgs = [str(e1), str(e2), str(e3)]
                    raise Exception(f"Failed to initialize Visio application after multiple attempts. Errors: {error_msgs}")
    
    return visio_app

def close_visio_app():
    """Properly close the Visio Application."""
    global visio_app, open_documents
    
    # Close all open documents first
    for path, doc in list(open_documents.items()):
        try:
            doc.Close()
        except:
            pass
    
    open_documents = {}
    
    # Then quit Visio
    if visio_app:
        try:
            visio_app.Quit()
        except:
            pass
        visio_app = None

@mcp.tool()
async def create_visio_file(template_path: Optional[str] = None, save_path: Optional[str] = None) -> str:
    """Create a new Visio file.
    
    Args:
        template_path: Path to the Visio template file (.vstx, .vst, etc.) to use.
                      If not provided, a default template will be used.
        save_path: Path where the file should be saved. If not provided,
                  it will be saved in the user's Documents folder with a default name.
    
    Returns:
        The path to the created Visio file.
    """
    try:
        app = get_visio_app()
        
        # Use default save path if not provided
        if not save_path:
            filename = f"New_Diagram_{int(time.time())}.vsdx"
            save_path = os.path.join(DEFAULT_SAVE_PATH, filename)
        else:
            # Check if save_path is just a filename (no directory)
            if os.path.dirname(save_path) == '':
                # If it's just a filename, put it in the default save directory
                save_path = os.path.join(DEFAULT_SAVE_PATH, save_path)
        
        # Make sure the directory exists
        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        
        # Create a new document
        try:
            if template_path and os.path.exists(template_path):
                doc = app.Documents.Add(template_path)
            else:
                # Create a blank document
                doc = app.Documents.Add("")
            
            # Wait for document to initialize
            time.sleep(1)

            # Default page size is too small for multi-shape flowcharts (shapes
            # end up drawn outside the page frame). Give it generous room.
            try:
                page_sheet = doc.Pages.Item(1).PageSheet
                page_sheet.CellsU("PageWidth").FormulaU = "60 in"
                page_sheet.CellsU("PageHeight").FormulaU = "90 in"
            except Exception:
                pass  # non-fatal - proceed with default page size if this fails

            # Save the document
            doc.SaveAs(save_path)
            
            # Add to open documents
            open_documents[save_path] = doc
            
            return f"Visio file created successfully at: {save_path}"
        except Exception as e:
            return f"Error creating Visio file: {str(e)}\nTried to create file at: {save_path}"
    except Exception as e:
        return f"Error initializing Visio: {str(e)}"

@mcp.tool()
async def open_visio_file(file_path: str) -> str:
    """Open an existing Visio file.
    
    Args:
        file_path: Path to the Visio file to open.
    
    Returns:
        Result message indicating success or failure.
    """
    try:
        app = get_visio_app()
        
        # Check if file exists
        if not os.path.exists(file_path):
            return f"Error: File does not exist at path: {file_path}"
        
        # Check if file is already open
        if file_path in open_documents:
            try:
                # Test if document is valid
                _ = open_documents[file_path].Name
                return f"Visio file is already open: {file_path}"
            except:
                # Document reference is invalid, remove it
                del open_documents[file_path]
        
        # Open the document
        try:
            doc = app.Documents.Open(file_path)
            open_documents[file_path] = doc
            return f"Visio file opened successfully: {file_path}"
        except Exception as e:
            return f"Error opening Visio file: {str(e)}"
    except Exception as e:
        return f"Error initializing Visio: {str(e)}"

@mcp.tool()
async def add_shape(file_path: str, shape_type: str, x: float, y: float, 
                    width: Optional[float] = 1.0, height: Optional[float] = 1.0) -> str:
    """Add a shape to an existing Visio document.
    
    Args:
        file_path: Path to the Visio file.
        shape_type: Type of shape to add (e.g., "Rectangle", "Circle", "Line", etc.).
        x: X-coordinate for the shape.
        y: Y-coordinate for the shape.
        width: Width of the shape (default: 1.0).
        height: Height of the shape (default: 1.0).
    
    Returns:
        Result message indicating success or failure.
    """
    try:
        # Ensure Visio is running
        app = get_visio_app()
        
        # If file doesn't exist, try to create it
        if not os.path.exists(file_path):
            creation_result = await create_visio_file(save_path=file_path)
            if "Error" in creation_result:
                return f"Cannot add shape - file does not exist and could not be created: {creation_result}"
        
        # If file is not open, try to open it
        if file_path not in open_documents:
            open_result = await open_visio_file(file_path)
            if "Error" in open_result:
                return f"Cannot add shape - file could not be opened: {open_result}"
        
        # Get the document
        doc = open_documents[file_path]
        
        # Get the active page
        page = app.ActivePage
        
        # Create shape based on type
        shape = None
        shape_type_lower = shape_type.lower()
        
        if shape_type_lower == "rectangle":
            shape = page.DrawRectangle(x, y, x + width, y + height)
        elif shape_type_lower in ["circle", "ellipse"]:
            shape = page.DrawOval(x, y, x + width, y + height)
        elif shape_type_lower == "line":
            shape = page.DrawLine(x, y, x + width, y + height)
        else:
            # Default to rectangle if shape type not recognized
            shape = page.DrawRectangle(x, y, x + width, y + height)
        
        # Set shape text
        if shape:
            shape.Text = shape_type
        
        # Save document
        doc.Save()
        
        return f"Shape '{shape_type}' added to the Visio file at ({x}, {y}) with ID {shape.ID}"
    except Exception as e:
        return f"Error adding shape to Visio file: {str(e)}"

@mcp.tool()
async def connect_shapes(file_path: str, shape1_id: int, shape2_id: int, 
                        connector_type: Optional[str] = "Dynamic") -> str:
    """Connect two shapes in a Visio document.
    
    Args:
        file_path: Path to the Visio file.
        shape1_id: ID of the first shape.
        shape2_id: ID of the second shape.
        connector_type: Type of connector (options: "Dynamic", "Straight", "Curved").
    
    Returns:
        Result message indicating success or failure.
    """
    try:
        # Ensure Visio is running
        app = get_visio_app()
        
        # If file is not open, try to open it
        if file_path not in open_documents:
            open_result = await open_visio_file(file_path)
            if "Error" in open_result:
                return f"Cannot connect shapes - file could not be opened: {open_result}"
        
        # Get the document
        doc = open_documents[file_path]
        
        # Get the active page
        page = app.ActivePage
        
        # Find shapes by ID
        shape1 = None
        shape2 = None
        
        for shape in page.Shapes:
            if shape.ID == shape1_id:
                shape1 = shape
            elif shape.ID == shape2_id:
                shape2 = shape
        
        if not shape1 or not shape2:
            return f"Error: Could not find shapes with IDs {shape1_id} and {shape2_id}"

        # Page.AutoConnectMany(FromShapeIDs, ToShapeIDs, PlacementDirs, [Connector])
        # is the documented, code-driven way to connect shapes by ID:
        # https://learn.microsoft.com/en-us/office/vba/api/visio.page.autoconnectmany
        # The previous implementation used app.ConnectorToolDataObject + page.Drop(),
        # which simulates an interactive drag-and-drop of the Connector Tool - this
        # hangs in unattended automation (observed: consistent ~4min timeout, no
        # visible dialog). AutoConnectMany avoids UI-drag simulation and ActiveWindow
        # selection state entirely.
        before_ids = {s.ID for s in page.Shapes}

        # visAutoConnectDirNone = 0 - connect without relocating the existing shapes
        # (they're already positioned by add_shape calls).
        connected_count = page.AutoConnectMany([shape1_id], [shape2_id], [0])

        if connected_count == 0:
            return f"Error: AutoConnectMany could not connect shapes {shape1_id} and {shape2_id}"

        after_ids = {s.ID for s in page.Shapes}
        new_ids = after_ids - before_ids
        connector = None
        if new_ids:
            new_id = next(iter(new_ids))
            for shape in page.Shapes:
                if shape.ID == new_id:
                    connector = shape
                    break

        connector_type_lower = (connector_type or "dynamic").lower()
        if connector is not None and connector_type_lower in ("straight", "curved"):
            connector.CellsU("LinePattern").FormulaU = "0"
            connector.CellsU("Rounding").FormulaU = "0 mm" if connector_type_lower == "straight" else "5 mm"

        # Save document
        doc.Save()

        return f"Shapes {shape1_id} and {shape2_id} connected successfully with {connector_type} connector"
    except Exception as e:
        return f"Error connecting shapes: {str(e)}"

@mcp.tool()
async def add_text(file_path: str, shape_id: int, text: str) -> str:
    """Add text to a shape in a Visio document.
    
    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to add text to.
        text: Text to add to the shape.
    
    Returns:
        Result message indicating success or failure.
    """
    try:
        # Ensure Visio is running
        app = get_visio_app()
        
        # If file is not open, try to open it
        if file_path not in open_documents:
            open_result = await open_visio_file(file_path)
            if "Error" in open_result:
                return f"Cannot add text - file could not be opened: {open_result}"
        
        # Get the document
        doc = open_documents[file_path]
        
        # Get the active page
        page = app.ActivePage
        
        # Find shape by ID
        target_shape = None
        for shape in page.Shapes:
            if shape.ID == shape_id:
                target_shape = shape
                break
        
        if not target_shape:
            return f"Error: Could not find shape with ID {shape_id}"
        
        # Add text to shape
        target_shape.Text = text
        
        # Save document
        doc.Save()
        
        return f"Text added to shape {shape_id} successfully"
    except Exception as e:
        return f"Error adding text to shape: {str(e)}"

@mcp.tool()
async def set_shape_color(file_path: str, shape_id: int, fill_color: Optional[str] = None,
                          line_color: Optional[str] = None) -> str:
    """Set the fill and/or line color of a shape in a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to color.
        fill_color: Fill color as a hex string, e.g. "#DBEAFE" or "DBEAFE" (optional).
        line_color: Outline/stroke color as a hex string, e.g. "#2563EB" (optional).

    Returns:
        Result message indicating success or failure.
    """
    def hex_to_rgb_formula(hex_color: str) -> str:
        h = hex_color.strip().lstrip("#")
        if len(h) != 6:
            raise ValueError(f"Invalid hex color: {hex_color!r} (expected 6 hex digits)")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f"RGB({r},{g},{b})"

    try:
        # Ensure Visio is running
        app = get_visio_app()

        # If file is not open, try to open it
        if file_path not in open_documents:
            open_result = await open_visio_file(file_path)
            if "Error" in open_result:
                return f"Cannot set shape color - file could not be opened: {open_result}"

        # Get the document
        doc = open_documents[file_path]

        # Get the active page
        page = app.ActivePage

        # Find shape by ID
        target_shape = None
        for shape in page.Shapes:
            if shape.ID == shape_id:
                target_shape = shape
                break

        if not target_shape:
            return f"Error: Could not find shape with ID {shape_id}"

        if fill_color is None and line_color is None:
            return "Error: Provide at least one of fill_color or line_color"

        applied = []
        if fill_color is not None:
            target_shape.CellsU("FillForegnd").FormulaU = hex_to_rgb_formula(fill_color)
            target_shape.CellsU("FillPattern").FormulaU = "1"  # solid fill
            applied.append(f"fill={fill_color}")
        if line_color is not None:
            target_shape.CellsU("LineColor").FormulaU = hex_to_rgb_formula(line_color)
            applied.append(f"line={line_color}")

        # Save document
        doc.Save()

        return f"Shape {shape_id} colored successfully ({', '.join(applied)})"
    except Exception as e:
        return f"Error setting shape color: {str(e)}"

@mcp.tool()
async def list_shapes(file_path: str) -> str:
    """List all shapes in a Visio document.
    
    Args:
        file_path: Path to the Visio file.
    
    Returns:
        JSON string containing information about all shapes in the document.
    """
    try:
        # Ensure Visio is running
        app = get_visio_app()
        
        # If file is not open, try to open it
        if file_path not in open_documents:
            open_result = await open_visio_file(file_path)
            if "Error" in open_result:
                return f"Cannot list shapes - file could not be opened: {open_result}"
        
        # Get the document
        doc = open_documents[file_path]
        
        # Get the active page
        page = app.ActivePage
        
        # Collect shape information
        shapes_info = []
        for shape in page.Shapes:
            shape_info = {
                "ID": shape.ID,
                "Name": shape.Name,
                "Text": shape.Text,
                "Type": shape.Type,
                "Position": {
                    "X": shape.Cells("PinX").Result(""),
                    "Y": shape.Cells("PinY").Result("")
                },
                "Size": {
                    "Width": shape.Cells("Width").Result(""),
                    "Height": shape.Cells("Height").Result("")
                }
            }
            shapes_info.append(shape_info)
        
        return json.dumps(shapes_info, indent=2)
    except Exception as e:
        return f"Error listing shapes: {str(e)}"

@mcp.tool()
async def close_document(file_path: str, save_changes: Optional[bool] = True) -> str:
    """Close a Visio document.
    
    Args:
        file_path: Path to the Visio file.
        save_changes: Whether to save changes before closing (default: True).
    
    Returns:
        Result message indicating success or failure.
    """
    global open_documents
    
    try:
        if file_path in open_documents:
            doc = open_documents[file_path]
            
            # Save changes if requested
            if save_changes:
                doc.Save()
            
            # Close document
            doc.Close()
            
            # Remove from open documents
            del open_documents[file_path]
            
            return f"Document {file_path} closed successfully"
        else:
            return f"Document {file_path} is not currently open"
    except Exception as e:
        return f"Error closing document: {str(e)}"

# Register the cleanup function with atexit
atexit.register(close_visio_app)
    
    
def main():
    """Entry point for the MCP server."""
    # Check if Visio is installed before starting
    if not check_visio_installed():
        sys.stderr.write("Microsoft Visio is not installed. This MCP server requires Visio to function.\n")
        sys.exit(1)
    
    # Ensure Visio is initialized before accepting requests
    try:
        # Pre-initialize Visio app to ensure it's ready
        _ = get_visio_app()
        # Run the server with proper initialization
        mcp.run(transport='stdio')
    except Exception as e:
        sys.stderr.write(f"Error initializing Visio MCP Server: {str(e)}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
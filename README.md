# Visio MCP Server (maintained fork)

A MCP server that provides tools for creating and editing Microsoft Visio diagrams programmatically via a standardized API.

![](https://badge.mcpx.dev?type=server "MCP Server")

> **Fork notice.** This is a fork of [GongRzhe/Office-Visio-MCP-Server](https://github.com/GongRzhe/Office-Visio-MCP-Server), which was archived (read-only) on 2026-03-03. This fork fixes a hang in `connect_shapes` and adds shape coloring — see [Changes in this fork](#changes-in-this-fork) below. All credit for the original design and implementation goes to GongRzhe.

## Overview

Visio MCP Server allows you to automate Visio diagram creation and editing using Python. It leverages Microsoft's COM interface to control Visio, enabling you to programmatically create diagrams, add shapes, connect them, add text, color them, and more.

## Example

![demo](./public/demo.gif)

## Requirements

- Windows operating system
- Microsoft Visio (Professional or Standard) installed
- Python 3.12+
- Python packages:
  - `mcp==1.29.0` (pinned — see [Changes in this fork](#changes-in-this-fork))
  - `win32com.client` (pywin32)

## Installation

1. Ensure Microsoft Visio is installed on your system
2. Install required Python packages:

```bash
pip install -r requirements.txt
```

3. Clone or download this repository
4. Run the server:

```bash
python visio_mcp_server/visio_server.py
```

## Features

The server provides the following functionality:

### Creating and Opening Files
- Create new Visio diagrams (with a page sized for multi-shape diagrams — see below)
- Open existing Visio diagrams

### Shape Management
- Add various shapes (Rectangle, Circle, Line, etc.)
- Connect shapes with different connector types
- Add text to shapes
- Set shape fill and line color
- List all shapes in a document

### File Operations
- Save documents to specified locations
- Export diagrams as images
- Close documents safely

## MCP Configuration

### Local Python Server

Add the server to your MCP settings configuration file:

```json
{
  "mcpServers": {
    "visio": {
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\visio_mcp_server\\visio_server.py"],
      "env": {}
    }
  }
}
```

## API Reference

### Create a Visio File
Creates a new Visio diagram.

```json
{
  "template_path": "[optional] Path to Visio template (.vstx, .vst)",
  "save_path": "[optional] Where to save the file"
}
```

### Open a Visio File
Opens an existing Visio diagram.

```json
{
  "file_path": "Path to the Visio file to open"
}
```

### Add Shape
Adds a shape to a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "shape_type": "Type of shape (Rectangle, Circle, Line, etc.)",
  "x": 1.0,
  "y": 1.0,
  "width": 1.0,
  "height": 1.0
}
```

### Connect Shapes
Connects two shapes in a Visio diagram, by ID.

```json
{
  "file_path": "Path to the Visio file",
  "shape1_id": 1,
  "shape2_id": 2,
  "connector_type": "Dynamic, Straight, or Curved"
}
```

### Add Text
Adds text to a shape in a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "shape_id": 1,
  "text": "Text to add to the shape"
}
```

### Set Shape Color *(new in this fork)*
Sets the fill and/or line color of a shape.

```json
{
  "file_path": "Path to the Visio file",
  "shape_id": 1,
  "fill_color": "#DBEAFE",
  "line_color": "#2563EB"
}
```

Both `fill_color` and `line_color` are optional hex strings (with or without the leading `#`); pass at least one.

### List Shapes
Lists all shapes in a Visio diagram.

```json
{
  "file_path": "Path to the Visio file"
}
```

## Changes in this fork

### Fixed: `connect_shapes` hung indefinitely (~4 minute timeout, every call)

The original implementation built a connector via `app.ConnectorToolDataObject` + `page.Drop(...)` — this simulates an interactive drag-and-drop of Visio's Connector Tool. In unattended MCP automation (no real mouse driving the drag), this call blocks and eventually times out, with no visible dialog to explain why.

**Fix:** use [`Page.AutoConnectMany`](https://learn.microsoft.com/en-us/office/vba/api/visio.page.autoconnectmany), the documented, code-driven API for connecting shapes by ID. It doesn't depend on `ActiveWindow.Selection` or any drag simulation, and returns immediately.

### Fixed: new diagrams used the default page size, shapes ended up outside the page frame

A default new Visio document uses a small default page (Letter/A4-sized). A flowchart with more than a handful of shapes quickly exceeds that, and shapes get placed outside the visible page frame. `create_visio_file` now sets a generous page size (60in × 90in) on creation.

### Added: `set_shape_color`

The original README listed "Color and fill pattern customization" under Future Features. This fork adds a `set_shape_color` tool (fill + line color, hex input) using `Shape.CellsU("FillForegnd")` / `Shape.CellsU("LineColor")`.

### Pinned `mcp==1.29.0`

`requirements.txt` originally specified `mcp>=1.2.0` with no upper bound. `mcp` 2.0.0 renamed/restructured `mcp.server.fastmcp`, which this server imports directly — installing latest breaks the import (`ModuleNotFoundError`). Pinned to the last 1.x release.

## Troubleshooting

### Common Issues:

1. **Visio Not Launching**:
   - Ensure Visio is correctly installed and can be opened manually
   - Check that you have sufficient permissions to launch COM applications

2. **Template Not Found**:
   - The server will create a blank diagram if templates aren't found
   - Specify an absolute path to a template if needed

3. **Invalid Shape Type**:
   - If a shape type isn't recognized, the server will default to a rectangle
   - Check spelling and case of shape names

4. **COM Errors**:
   - Restarting Visio manually may help resolve COM interface issues
   - Ensure no existing Visio processes are hanging in Task Manager

5. **Automating over SSH / a non-interactive session**:
   - Visio's COM server needs a real interactive Windows session to initialize. Launching the MCP server (or any script that touches `win32com.client.Dispatch("Visio.Application")`) from a non-interactive SSH/service session fails (`RPC failed`) even though the same code works fine when the MCP client (e.g. Claude Desktop) runs in the interactive session itself.

## License

This project is licensed under the MIT License - see the LICENSE file for details. Original work Copyright (c) 2025 GongRzhe.

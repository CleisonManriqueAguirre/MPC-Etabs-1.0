# -*- coding: utf-8 -*-
"""
ETABS MCP Server - Shared ETABS connection

Exposes module-level `etabs` / `model` globals that create_elements,
change_elements, etabs_queries, table_extraction_tools, and lateral_loads
all read directly as `etabs_connection.model`.
"""

import sys

import comtypes.client
import psutil

# Global variables
etabs = None
model = None


def connect_to_existing() -> str:
    """Connect to existing ETABS instance - SIMPLE VERSION"""
    global etabs, model

    try:
        helper = comtypes.client.CreateObject('ETABSv1.Helper')
        import comtypes.gen.ETABSv1 as ETABSv1

        helper = helper.QueryInterface(ETABSv1.cHelper)

        etabs_raw = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        etabs = etabs_raw.QueryInterface(ETABSv1.cOAPI)
        model = etabs.SapModel

        # Get basic model info
        model_name = model.GetModelFilename()
        return f"Connected to: {model_name}"

    except Exception as e:
        etabs = None
        model = None
        return f"Failed to connect: {e}"


def get_model_info() -> str:
    """Get basic model information"""
    if model is None:
        return "Not connected to ETABS"

    try:
        model_name = model.GetModelFilename()
        units = model.GetPresentUnits()
        num_points = model.PointObj.Count()
        num_frames = model.FrameObj.Count()

        return f"""Model Info:
File: {model_name}
Units: {units}
Points: {num_points}
Frames: {num_frames}"""

    except Exception as e:
        return f"Error getting model info: {e}"


def get_etabs_connection():
    """Return this module (which exposes `.model`) if connected, else None.

    tools.lateral_loads calls etabs_connection.get_etabs_connection() and
    reads `.model` off the result.
    """
    if model is None:
        return None
    return sys.modules[__name__]

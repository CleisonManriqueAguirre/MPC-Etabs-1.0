"""Claude-compatible MCP server for ETABS."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from tools.etabs_client import EtabsClient
from tools.etabs_connection import connect_to_existing
from tools.extraccion_tablas import get_etabs_data_api, get_etabs_table


mcp = FastMCP(
    name="ETABS MCP",
    instructions=(
        "Use this server to inspect and automate CSI ETABS through the local Windows installation. "
        "Connect via connect_etabs(), then use get_etabs_table()/get_etabs_data_api() to pull model data."
    ),
)

_client = EtabsClient()


@mcp.tool()
def ping() -> str:
    """Confirm the server is reachable from Claude."""
    return "ETABS MCP server is running."


@mcp.tool()
def etabs_status() -> dict[str, str]:
    """Return the local ETABS connection status."""
    return _client.status()


@mcp.tool()
def connect_etabs() -> str:
    """Connect to an already-running ETABS instance via comtypes."""
    return connect_to_existing()

mcp.tool()(get_etabs_table)
mcp.tool()(get_etabs_data_api)


def main() -> None:
    """Run the MCP server over stdio for Claude Desktop."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

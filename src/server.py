from fastmcp import FastMCP
import os

mcp = FastMCP("M365 Graph MCP")

@mcp.tool
def ping() -> str:
    return "pong"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))

    os.environ["FASTMCP_STATELESS_HTTP"] = "true"

    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=port
    )

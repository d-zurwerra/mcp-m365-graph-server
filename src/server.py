import os
from fastmcp import FastMCP

# Wichtig: stateless_http + json_response für Copilot Studio (streamable HTTP, kein SSE)
mcp = FastMCP(
    "M365 Graph MCP",
    stateless_http=True,
    json_response=True
)

@mcp.tool(description="Health check tool.")
def ping() -> str:
    return "pong"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=port
    )

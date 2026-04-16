from fastmcp import FastMCP

mcp = FastMCP(
    "M365 Graph MCP",
    stateless_http=True,
    json_response=True   # wichtig für Copilot Studio
)

@mcp.tool
def ping() -> str:
    return "pong"

if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8000
    )

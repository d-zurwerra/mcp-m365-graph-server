import os
from fastmcp import FastMCP

mcp = FastMCP("m365-graph-mcp-dev")

@mcp.tool(description="Simple health tool.")
def ping() -> str:
    return "OK"

app = mcp.http_app()

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port)

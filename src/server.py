import os

# ✅ MUSS vor Import gesetzt werden: Copilot Studio braucht stateless MCP
os.environ["FASTMCP_STATELESS_HTTP"] = "true"

from fastmcp import FastMCP

mcp = FastMCP("M365 Graph MCP")

@mcp.tool
def ping() -> str:
    return "pong"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))

    # ✅ WICHTIG: HTTP-Runner statt run()
    # → akzeptiert GET / OPTIONS / POST
    mcp.run_http_async(
        host="0.0.0.0",
        port=port,
        stateless_http=True
    )
``

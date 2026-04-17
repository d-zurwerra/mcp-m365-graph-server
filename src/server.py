import os
import asyncio

# MUSS vor Import
os.environ["FASTMCP_STATELESS_HTTP"] = "true"

from fastmcp import FastMCP

mcp = FastMCP("M365 Graph MCP")

@mcp.tool
def ping() -> str:
    return "pong"

async def main():
    port = int(os.environ.get("PORT", "8080"))

    # Server starten
    server_task = asyncio.create_task(
        mcp.run_http_async(
            host="0.0.0.0",
            port=port,
            stateless_http=True
        )
    )

    # ✅ Prozess am Leben halten
    await server_task

if __name__ == "__main__":
    asyncio.run(main())

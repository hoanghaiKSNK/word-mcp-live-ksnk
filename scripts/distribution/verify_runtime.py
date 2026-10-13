"""Check a fresh runtime over stdio without entering Word or reading a document."""
import argparse
import asyncio
import json
from pathlib import Path
from fastmcp import Client
from fastmcp.client.transports import StdioTransport


async def verify(runtime):
    async with Client(StdioTransport(command=str(runtime/'.venv/Scripts/word_mcp_server.exe'),
                                     args=[], cwd=str(runtime), env={'MCP_TRANSPORT': 'stdio'},
                                     log_file=runtime/'runtime-protocol.log'), timeout=45) as client:
        tools = await client.list_tools()
        result = await client.call_tool('word_mcp_dispatch_status', {})
        status = json.loads(next(block.text for block in result.content if hasattr(block, 'text')))
        assert not result.is_error and status['version'] == '1.6.21+ksnk.2'
        assert len(tools) == 121
        return {'passed': True, 'version': status['version'], 'tools': len(tools)}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--runtime', type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(asyncio.run(verify(args.runtime))))

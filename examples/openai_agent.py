"""OpenAI agent example that attaches the Excel MCP server as a tool.

Prerequisites
------------
- pip install --upgrade openai
- dotnet publish src/ExcelMcp.Server/ExcelMcp.Server.csproj -c Release
- Set EXCEL_MCP_SERVER_PATH and EXCEL_MCP_WORKBOOK
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

from openai import OpenAI


def resolve_server_path() -> str:
    """Return the MCP server executable path, validating existence."""
    root = Path(__file__).resolve().parents[1]
    default_path = root / "src" / "ExcelMcp.Server" / "bin" / "Release" / "net9.0" / "ExcelMcp.Server.exe"
    server_path = Path(os.environ.get("EXCEL_MCP_SERVER_PATH", str(default_path)))
    if not server_path.exists():
        raise FileNotFoundError(
            "Excel MCP server executable not found. Set EXCEL_MCP_SERVER_PATH to the published .exe."
        )
    return str(server_path)


def resolve_workbook_path() -> str:
    """Return the workbook path provided by the caller."""
    workbook_path = Path(os.environ.get("EXCEL_MCP_WORKBOOK", ""))
    if not workbook_path.exists():
        raise FileNotFoundError(
            "Workbook not found. Set EXCEL_MCP_WORKBOOK to the .xlsx file you want to inspect."
        )
    return str(workbook_path)


def print_thread_messages(client: OpenAI, thread_id: str) -> None:
    """Print the conversation so far in role-prefixed form."""
    messages = client.threads.messages.list(thread_id=thread_id, order="asc")
    for message in messages.data:
        print(f"\n{message.role.upper()}:")
        for item in message.content:
            if getattr(item, "type", None) == "text":
                print(item.text.value)


def main() -> int:
    server_path = resolve_server_path()
    workbook_path = resolve_workbook_path()
    model = os.environ.get("OPENAI_AGENT_MODEL", "gpt-4.1-mini")

    client = OpenAI()

    agent = client.agents.create(
        name="Excel Workbook Assistant",
        model=model,
        instructions=(
            "You may call the Excel MCP tools to explore the workbook. "
            "Summarize worksheets, run searches, or preview tables when helpful."
        ),
        tools=[
            {
                "type": "mcp_server",
                "server": {
                    "name": "excel-workbook-mcp",
                    "transport": {
                        "type": "stdio",
                        "command": server_path,
                        "args": ["--workbook", workbook_path],
                    },
                },
            }
        ],
    )

    thread = client.threads.create()
    client.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content="List the worksheets and tables that appear in this workbook.",
    )

    run = client.threads.runs.create(thread_id=thread.id, agent_id=agent.id)

    while run.status in {"queued", "in_progress", "requires_action"}:
        time.sleep(1)
        run = client.threads.runs.retrieve(thread_id=thread.id, run_id=run.id)

    print_thread_messages(client, thread.id)

    if run.status != "completed":
        raise RuntimeError(f"Agent run did not complete successfully: {run.status}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - diagnostic print only
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)

"""MCP Server implementation for Microsoft Graph API."""

import asyncio
import json
import logging
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.lowlevel.server import NotificationOptions
import mcp.server.stdio
import mcp.types as types

from .handlers import (
    AuthHandler,
    UserHandler,
    EmailHandler,
    CalendarHandler,
    # FileHandler,  # Disabled: OneDrive not needed
    # TeamsHandler,  # Disabled: Teams not needed
)
from .tools import ToolRegistry
from .config import settings
from .utils import read_bcc_from_csv

logger = logging.getLogger(__name__)


class MicrosoftGraphMCPServer:
    """MCP Server for Microsoft Graph API integration."""

    def __init__(self):
        logger.info("=" * 70)
        logger.info("MicrosoftGraphMCPServer.__init__() called")
        logger.info(f"Server name: {settings.server_name}")
        logger.info(f"Server version: {settings.server_version}")
        logger.info("=" * 70)

        self.server = Server("microsoft-graph-mcp-server")
        self.auth_handler = AuthHandler()
        self.user_handler = UserHandler()
        self.email_handler = EmailHandler()
        self.calendar_handler = CalendarHandler()
        # self.file_handler = FileHandler()  # Disabled: OneDrive not needed
        # self.teams_handler = TeamsHandler()  # Disabled: Teams not needed

        self._build_dispatch_table()
        self._register_handlers()

        logger.info("Handlers registered successfully")

    def _build_dispatch_table(self):
        """Build tool dispatch table for O(1) lookup."""
        self.tool_dispatch = {
            "auth": (self.auth_handler, "handle_auth"),
            "user_settings": (self.user_handler, "handle_user_settings"),
            "search_contacts": (self.user_handler, "handle_search_contacts"),
            "manage_mail_folder": (self.email_handler, "handle_manage_mail_folder"),
            "manage_emails": (self.email_handler, "handle_manage_emails"),
            "browse_email_cache": (self.email_handler, "handle_browse_email_cache"),
            "search_emails": (self.email_handler, "handle_search_emails"),
            "get_email_content": (self.email_handler, "handle_get_email_content"),
            "send_email": (self.email_handler, "handle_send_email"),
            "create_draft": (self.email_handler, "handle_create_draft"),
            "browse_events": (self.calendar_handler, "handle_browse_events"),
            "get_event_detail": (self.calendar_handler, "handle_get_event"),
            "search_events": (self.calendar_handler, "handle_search_events"),
            "check_attendee_availability": (
                self.calendar_handler,
                "handle_check_attendee_availability",
            ),
            "manage_event_as_attendee": (self.calendar_handler, "handle_respond_to_event"),
            "manage_event_as_organizer": (self.calendar_handler, "handle_manage_my_event"),
            # "list_files": (self.file_handler, "handle_list_files"),  # Disabled: OneDrive not needed
            # "get_teams": (self.teams_handler, "handle_get_teams"),  # Disabled: Teams not needed
            # "get_team_channels": (self.teams_handler, "handle_get_team_channels"),  # Disabled: Teams not needed
        }

    def _register_handlers(self):
        """Register MCP tool handlers."""

        @self.server.list_tools()
        async def handle_list_tools() -> list[types.Tool]:
            """List all available tools."""
            return ToolRegistry.get_all_tools()

        @self.server.call_tool()
        async def handle_call_tool(
            name: str, arguments: dict | None
        ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
            """Handle tool execution requests."""

            if arguments is None:
                arguments = {}

            # Normalize date_range and time_range to lowercase for case-insensitive handling
            if "date_range" in arguments and arguments["date_range"]:
                arguments["date_range"] = arguments["date_range"].lower()
            if "time_range" in arguments and arguments["time_range"]:
                arguments["time_range"] = arguments["time_range"].lower()

            # Normalize boolean parameters (return_html) from string to bool
            if "return_html" in arguments and isinstance(arguments["return_html"], str):
                arguments["return_html"] = arguments["return_html"].lower() == "true"

            logger.info(f"Server: Tool called - {name} with args: {arguments}")

            try:
                if name == "auth":
                    logger.info(
                        f"Server: Routing to auth_handler with action: {arguments.get('action')}"
                    )

                handler_info = self.tool_dispatch.get(name)
                if not handler_info:
                    raise ValueError(f"Unknown tool: {name}")

                handler, method_name = handler_info
                method = getattr(handler, method_name)
                return await method(arguments)

            except Exception as e:
                return [
                    types.TextContent(
                        type="text", text=f"Error executing tool {name}: {str(e)}"
                    )
                ]

    async def run(self):
        """Run the MCP server."""
        logger.info("=" * 70)
        logger.info("MicrosoftGraphMCPServer.run() called - Starting MCP server")
        logger.info("=" * 70)

        try:
            # Wrap sys.stdin to strip BOM characters before MCP library reads it
            import sys

            original_stdin = sys.stdin
            bom_stripped = False

            class BOMStrippingTextIOWrapper:
                """Wrapper that strips BOM from stdin."""

                def __init__(self, original):
                    self._original = original

                def read(self, size=-1):
                    data = self._original.read(size)
                    nonlocal bom_stripped
                    if not bom_stripped and data:
                        # Strip UTF-8 BOM (0xEF 0xBB 0xBF) or Unicode BOM (\ufeff)
                        if data.startswith('\ufeff'):
                            data = data[1:]
                        elif data.startswith('\xef\xbb\xbf'):
                            data = data[3:]
                        bom_stripped = True
                    return data

                def readline(self, size=-1):
                    data = self._original.readline(size)
                    nonlocal bom_stripped
                    if not bom_stripped and data:
                        if data.startswith('\ufeff'):
                            data = data[1:]
                        elif data.startswith('\xef\xbb\xbf'):
                            data = data[3:]
                        bom_stripped = True
                    return data

                def __getattr__(self, name):
                    return getattr(self._original, name)

                @property
                def encoding(self):
                    return self._original.encoding

            sys.stdin = BOMStrippingTextIOWrapper(original_stdin)

            try:
                async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
                    logger.info("stdio_server acquired, starting server.run()...")
                    await self.server.run(
                        read_stream,
                        write_stream,
                        InitializationOptions(
                            server_name=settings.server_name,
                            server_version=settings.server_version,
                            capabilities=self.server.get_capabilities(
                                notification_options=NotificationOptions(),
                                experimental_capabilities={},
                            ),
                        ),
                    )
            finally:
                sys.stdin = original_stdin

        except Exception as e:
            logger.error(f"Error in server.run(): {e}", exc_info=True)
            raise


def run_stdio_server():
    """Entry point for stdio MCP server without click."""
    # CRITICAL: Disable PYTHONSTARTUP to prevent it from corrupting stdio MCP protocol
    # This must be done BEFORE any imports or stdio operations
    import os
    if "PYTHONSTARTUP" in os.environ:
        del os.environ["PYTHONSTARTUP"]

    server = MicrosoftGraphMCPServer()
    asyncio.run(server.run())

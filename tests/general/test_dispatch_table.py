"""Test tool dispatch table in server.py."""

import pytest
from unittest.mock import AsyncMock, patch

from microsoft_graph_mcp_server.server import MicrosoftGraphMCPServer


class TestToolDispatchTable:
    """Tests for tool dispatch table implementation."""

    @pytest.mark.asyncio
    async def test_dispatch_table_has_all_tools(self):
        """Verify dispatch table has entries for all 18 tools."""
        server = MicrosoftGraphMCPServer()

        expected_tools = [
            "auth",
            "user_settings",
            "search_contacts",
            "manage_mail_folder",
            "manage_emails",
            "browse_email_cache",
            "search_emails",
            "get_email_content",
            "send_email",
            "browse_events",
            "get_event_detail",
            "search_events",
            "check_attendee_availability",
            "manage_event_as_attendee",
            "manage_event_as_organizer",
            "teams_chats",
        ]

        for tool_name in expected_tools:
            assert tool_name in server.tool_dispatch, (
                f"Tool '{tool_name}' not found in dispatch table"
            )

        print(f"   ✓ All {len(expected_tools)} tools registered in dispatch table")

    @pytest.mark.asyncio
    async def test_dispatch_table_handlers_are_callable(self):
        """Verify all handlers in dispatch table have callable methods."""
        server = MicrosoftGraphMCPServer()

        for tool_name, (handler, method_name) in server.tool_dispatch.items():
            assert hasattr(handler, method_name), (
                f"Handler for '{tool_name}' missing method '{method_name}'"
            )
            method = getattr(handler, method_name)
            assert callable(method), (
                f"Method '{method_name}' for '{tool_name}' is not callable"
            )

        print(f"   ✓ All dispatch table handlers are callable")

    @pytest.mark.asyncio
    async def test_auth_tool_dispatch(self):
        """Test auth tool routing via dispatch table."""
        server = MicrosoftGraphMCPServer()

        with patch.object(
            server.auth_handler, "handle_auth", new_callable=AsyncMock
        ) as mock_handler:
            mock_handler.return_value = [{"type": "text", "text": "test result"}]

            handler, method_name = server.tool_dispatch["auth"]
            method = getattr(handler, method_name)
            result = await method({"action": "complete"})

            mock_handler.assert_called_once_with({"action": "complete"})
            print("   ✓ Auth tool dispatch works")

    @pytest.mark.asyncio
    async def test_search_emails_tool_dispatch(self):
        """Test search_emails tool routing via dispatch table."""
        server = MicrosoftGraphMCPServer()

        with patch.object(
            server.email_handler, "handle_search_emails", new_callable=AsyncMock
        ) as mock_handler:
            mock_handler.return_value = [{"type": "text", "text": "test result"}]

            handler, method_name = server.tool_dispatch["search_emails"]
            method = getattr(handler, method_name)
            result = await method({"query": "test"})

            mock_handler.assert_called_once_with({"query": "test"})
            print("   ✓ Search emails tool dispatch works")

    @pytest.mark.asyncio
    async def test_browse_events_tool_dispatch(self):
        """Test browse_events tool routing via dispatch table."""
        server = MicrosoftGraphMCPServer()

        with patch.object(
            server.calendar_handler, "handle_browse_events", new_callable=AsyncMock
        ) as mock_handler:
            mock_handler.return_value = [{"type": "text", "text": "test result"}]

            handler, method_name = server.tool_dispatch["browse_events"]
            method = getattr(handler, method_name)
            result = await method({"page_number": 1})

            mock_handler.assert_called_once_with({"page_number": 1})
            print("   ✓ Browse events tool dispatch works")


async def main():
    """Run all dispatch table tests."""
    print("=" * 70)
    print("Testing tool dispatch table implementation")
    print("=" * 70)

    tests = [
        test_dispatch_table_has_all_tools,
        test_dispatch_table_handlers_are_callable,
        test_auth_tool_dispatch,
        test_search_emails_tool_dispatch,
        test_browse_events_tool_dispatch,
    ]

    failed_tests = []

    for test in tests:
        try:
            await test()
        except Exception as e:
            print(f"\n✗ Test failed: {test.__name__}")
            print(f"  Error: {e}")
            failed_tests.append(test.__name__)

    print("\n" + "=" * 70)
    if failed_tests:
        print(f"✗ {len(failed_tests)} test(s) failed:")
        for test_name in failed_tests:
            print(f"  - {test_name}")
    else:
        print("✓ All dispatch table tests passed!")
    print("=" * 70)


if __name__ == "__main__":
    import asyncio

    asyncio.run(main())

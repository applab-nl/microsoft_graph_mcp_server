"""Teams handlers for MCP tools."""

import json
from typing import Any, Dict

from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client
from ..utils.html_text import html_to_text


def map_chat(c: Dict[str, Any]) -> Dict[str, Any]:
    """Graph chat -> compact row for the teams_chats tool."""
    return {
        "id": c.get("id"),
        "chat_type": c.get("chatType"),
        "topic": c.get("topic"),
        "last_updated": c.get("lastUpdatedDateTime"),
        "last_message_at": (c.get("lastMessagePreview") or {}).get("createdDateTime"),
        "join_url": (c.get("onlineMeetingInfo") or {}).get("joinWebUrl"),
        "members": [
            {"user_id": m.get("userId"), "display_name": m.get("displayName"), "email": m.get("email")}
            for m in c.get("members") or []
        ],
    }


def _reply_to_user_id(m: Dict[str, Any]) -> str | None:
    """Author of the quoted message when this message is a reply (messageReference attachment)."""
    for a in m.get("attachments") or []:
        if a.get("contentType") != "messageReference":
            continue
        try:
            ref = json.loads(a.get("content") or "{}")
        except ValueError:
            continue
        uid = ((ref.get("messageSender") or {}).get("user") or {}).get("id")
        if uid:
            return uid
    return None


def _mentioned_user_id(x: Dict[str, Any]) -> str | None:
    return ((x.get("mentioned") or {}).get("user") or {}).get("id")


def map_message(m: Dict[str, Any]) -> Dict[str, Any]:
    """Graph chatMessage -> compact row with plain-text body."""
    sender = (m.get("from") or {}).get("user") or {}
    body = m.get("body") or {}
    content = body.get("content") or ""
    text = html_to_text(content) if body.get("contentType") == "html" else content.strip()
    return {
        "id": m.get("id"),
        "created": m.get("createdDateTime"),
        "sender_id": sender.get("id"),
        "sender_name": sender.get("displayName"),
        "text": text,
        "mention_user_ids": [uid for uid in (_mentioned_user_id(x) for x in m.get("mentions") or []) if uid],
        "reply_to_user_id": _reply_to_user_id(m),
        "web_url": m.get("webUrl"),
    }


class TeamsHandler(BaseHandler):
    """Handler for Teams-related tools."""

    async def handle_get_teams(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_teams tool."""
        teams = await graph_client.get_teams()
        return self._format_response(teams)

    async def handle_get_team_channels(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle get_team_channels tool."""
        team_id = arguments["team_id"]
        channels = await graph_client.get_team_channels(team_id)
        return self._format_response(channels)

    async def handle_teams_chats(self, arguments: dict) -> list[types.TextContent]:
        """Handle teams_chats tool (me / list_chats / list_messages)."""
        tc = graph_client.teams_client
        action = arguments.get("action")
        if action == "me":
            me = await tc.get_me_identity()
            return self._format_response({
                "id": me.get("id"),
                "user_principal_name": me.get("userPrincipalName"),
                "display_name": me.get("displayName"),
            })
        if action == "list_chats":
            return self._format_response({"chats": [map_chat(c) for c in await tc.list_chats()]})
        if action == "list_messages":
            chat_id, since = arguments.get("chat_id"), arguments.get("since")
            if not chat_id or not since:
                return self._format_response({"error": "chat_id and since are required"})
            msgs = await tc.list_chat_messages(chat_id, since)
            return self._format_response({"messages": [map_message(m) for m in msgs]})
        return self._format_response({"error": f"unknown action: {action}"})

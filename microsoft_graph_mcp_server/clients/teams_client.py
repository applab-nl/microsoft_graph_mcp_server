"""Teams client for Microsoft Graph API."""

from typing import List, Dict, Any

from .base_client import BaseGraphClient


class TeamsClient(BaseGraphClient):
    """Client for Teams-related operations."""

    async def get_teams(self) -> List[Dict[str, Any]]:
        """Get list of Teams."""
        result = await self.get("/me/joinedTeams")
        return result.get("value", [])

    async def get_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """Get channels for a specific Team."""
        result = await self.get(f"/teams/{team_id}/channels")
        return result.get("value", [])

    MAX_PAGES = 20

    async def _follow(self, next_link: str) -> Dict[str, Any]:
        relative = next_link[len(self.base_url):] if next_link.startswith(self.base_url) else next_link
        return await self.get(relative)

    async def get_me_identity(self) -> Dict[str, Any]:
        """The signed-in user's Graph id, UPN and display name."""
        return await self.get("/me", params={"$select": "id,userPrincipalName,displayName"})

    async def list_chats(self) -> List[Dict[str, Any]]:
        """All chats (1:1, group, meeting) with members and the last-message preview expanded.

        lastUpdatedDateTime only moves on renames/member changes; the preview's
        createdDateTime is what tells a caller a chat has new messages.
        """
        result = await self.get(
            "/me/chats", params={"$expand": "members,lastMessagePreview", "$top": "50"}
        )
        chats = list(result.get("value", []))
        next_link, pages = result.get("@odata.nextLink"), 1
        while next_link and pages < self.MAX_PAGES:
            result = await self._follow(next_link)
            chats.extend(result.get("value", []))
            next_link, pages = result.get("@odata.nextLink"), pages + 1
        return chats

    async def list_chat_messages(self, chat_id: str, since: str) -> List[Dict[str, Any]]:
        """Messages newer than `since` (ISO), oldest first; user messages only."""
        result = await self.get(
            f"/chats/{chat_id}/messages",
            params={"$top": "50", "$orderby": "createdDateTime desc"},
        )
        out: List[Dict[str, Any]] = []
        pages = 1
        while True:
            reached_since = False
            for m in result.get("value", []):
                if m.get("createdDateTime", "") <= since:
                    reached_since = True
                    break
                if m.get("messageType") == "message" and not m.get("deletedDateTime"):
                    out.append(m)
            next_link = result.get("@odata.nextLink")
            if reached_since or not next_link or pages >= self.MAX_PAGES:
                break
            result = await self._follow(next_link)
            pages += 1
        out.reverse()
        return out

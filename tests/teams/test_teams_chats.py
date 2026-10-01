import json

from microsoft_graph_mcp_server.clients.teams_client import TeamsClient
from microsoft_graph_mcp_server.handlers.teams_handlers import map_chat, map_message
from microsoft_graph_mcp_server.utils.html_text import html_to_text


def test_html_to_text_keeps_line_breaks_and_decodes_entities():
    html = '<p>Hi <at id="0">Dylan</at>,</p><p>can you review&nbsp;this?<br>Thanks &amp; bye</p>'
    assert html_to_text(html) == "Hi Dylan,\ncan you review this?\nThanks & bye"


def test_html_to_text_image_only_is_empty():
    assert html_to_text('<p><img src="x.png"></p>') == ""


def _msg(i, created, **kw):
    base = {
        "id": str(i), "messageType": "message", "deletedDateTime": None,
        "createdDateTime": created,
        "from": {"user": {"id": "u-other", "displayName": "Ann"}},
        "body": {"contentType": "html", "content": f"<p>msg {i}</p>"},
        "mentions": [], "attachments": [], "webUrl": f"https://teams.microsoft.com/l/message/c/{i}",
    }
    base.update(kw)
    return base


class FakeTeamsClient(TeamsClient):
    """Pages of messages newest-first, like Graph with $orderby=createdDateTime desc."""

    def __init__(self, pages):
        super().__init__()
        self.pages = pages
        self.calls = []

    async def get(self, endpoint, params=None, headers=None):
        self.calls.append((endpoint, params))
        idx = len(self.calls) - 1
        page = {"value": self.pages[idx]}
        if idx + 1 < len(self.pages):
            page["@odata.nextLink"] = f"{self.base_url}/chats/c/messages?page={idx + 1}"
        return page


async def test_list_chat_messages_pages_until_since_and_returns_oldest_first():
    page1 = [_msg(60 - i, f"2026-10-01T10:{59 - i:02d}:00Z") for i in range(50)]
    page2 = [_msg(10, "2026-10-01T10:09:00Z"), _msg(9, "2026-09-30T08:00:00Z")]
    client = FakeTeamsClient([page1, page2])
    msgs = await client.list_chat_messages("c", "2026-10-01T00:00:00Z")
    assert len(client.calls) == 2
    assert len(msgs) == 51  # 50 from page 1 + 1 newer than `since` from page 2
    assert msgs[0]["id"] == "10"  # oldest first
    assert all(m["createdDateTime"] > "2026-10-01T00:00:00Z" for m in msgs)


async def test_list_chat_messages_drops_system_and_deleted():
    page = [
        _msg(3, "2026-10-01T10:03:00Z"),
        _msg(2, "2026-10-01T10:02:00Z", messageType="systemEventMessage"),
        _msg(1, "2026-10-01T10:01:00Z", deletedDateTime="2026-10-01T10:05:00Z"),
    ]
    msgs = await FakeTeamsClient([page]).list_chat_messages("c", "2026-10-01T00:00:00Z")
    assert [m["id"] for m in msgs] == ["3"]


def test_map_message_extracts_mentions_and_quoted_reply_author():
    raw = _msg(
        7, "2026-10-01T10:00:00Z",
        mentions=[{"mentioned": {"user": {"id": "U-OWNER", "displayName": "Dylan"}}}],
        attachments=[{"contentType": "messageReference",
                      "content": json.dumps({"messageSender": {"user": {"id": "U-OWNER"}}})}],
    )
    m = map_message(raw)
    assert m == {
        "id": "7", "created": "2026-10-01T10:00:00Z", "sender_id": "u-other", "sender_name": "Ann",
        "text": "msg 7", "mention_user_ids": ["U-OWNER"], "reply_to_user_id": "U-OWNER",
        "web_url": "https://teams.microsoft.com/l/message/c/7",
    }


def test_map_chat_reads_join_url_and_members():
    raw = {
        "id": "19:meeting_x@thread.v2", "chatType": "meeting", "topic": "Weekly",
        "lastUpdatedDateTime": "2026-10-01T09:00:00Z",
        "onlineMeetingInfo": {"joinWebUrl": "https://teams.microsoft.com/l/meetup-join/abc"},
        "members": [{"userId": "u1", "displayName": "Ann", "email": "ann@x.test"}],
    }
    assert map_chat(raw) == {
        "id": "19:meeting_x@thread.v2", "chat_type": "meeting", "topic": "Weekly",
        "last_updated": "2026-10-01T09:00:00Z",
        "join_url": "https://teams.microsoft.com/l/meetup-join/abc",
        "members": [{"user_id": "u1", "display_name": "Ann", "email": "ann@x.test"}],
    }

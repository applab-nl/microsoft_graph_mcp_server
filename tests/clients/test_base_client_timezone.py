"""Tests for BaseGraphClient.get_user_timezone caching.

A tenant that denies /me/mailboxSettings does so persistently (the app
registration is missing MailboxSettings.Read). Re-probing Graph every cache-TTL
window buys nothing and floods the log, so a denial must be cached for the whole
session while a *successful* lookup keeps its TTL.
"""

import logging

import pytest

from microsoft_graph_mcp_server.clients import base_client as bc
from microsoft_graph_mcp_server.clients.base_client import BaseGraphClient


DENIED = (
    'Graph API request failed: 403 - {"error":{"code":"ErrorAccessDenied",'
    '"message":"Access is denied. Check credentials and try again."}}'
)


@pytest.fixture(autouse=True)
def clear_timezone_cache():
    """The cache is class-level, so it leaks between tests unless cleared."""
    BaseGraphClient._user_timezone_cache = None
    BaseGraphClient._user_timezone_cache_time = None
    yield
    BaseGraphClient._user_timezone_cache = None
    BaseGraphClient._user_timezone_cache_time = None


@pytest.fixture
def fake_clock(monkeypatch):
    """Controllable clock so TTL expiry is testable without sleeping."""

    class Clock:
        now = 1_000_000.0

        def advance(self, seconds):
            self.now += seconds

    clock = Clock()
    monkeypatch.setattr(bc.time, "time", lambda: clock.now)
    return clock


class DeniedClient(BaseGraphClient):
    """Client whose every Graph call is refused, counting the attempts."""

    def __init__(self):
        super().__init__()
        self.calls = 0

    async def get(self, endpoint, params=None):
        self.calls += 1
        raise Exception(DENIED)


class WorkingClient(BaseGraphClient):
    """Client that returns a mailbox timezone, counting the attempts."""

    def __init__(self, timezone="W. Europe Standard Time"):
        super().__init__()
        self.calls = 0
        self.timezone = timezone

    async def get(self, endpoint, params=None):
        self.calls += 1
        return {"mailboxSettings": {"timeZone": self.timezone}}


class TestPermissionDenied:
    """A persistent 403 must be resolved once per process, not once per TTL."""

    async def test_denied_tenant_is_probed_once_and_warns_once(
        self, fake_clock, caplog, monkeypatch
    ):
        monkeypatch.setattr(bc.settings, "user_timezone", "Europe/Amsterdam")
        client = DeniedClient()

        with caplog.at_level(logging.WARNING, logger=bc.__name__):
            for _ in range(5):
                assert await client.get_user_timezone() == "Europe/Amsterdam"

        assert client.calls == 1, "a denied tenant must not be re-probed"
        warnings = [r for r in caplog.records if r.levelno == logging.WARNING]
        assert len(warnings) == 1, f"expected one warning, got {len(warnings)}"

    async def test_denial_is_not_retried_after_the_cache_ttl(
        self, fake_clock, caplog, monkeypatch
    ):
        """The regression: the TTL used to expire the failure and re-warn hourly."""
        monkeypatch.setattr(bc.settings, "user_timezone", "Europe/Amsterdam")
        client = DeniedClient()

        with caplog.at_level(logging.WARNING, logger=bc.__name__):
            await client.get_user_timezone()
            # Simulate a long-running session: several TTL windows go by.
            for _ in range(5):
                fake_clock.advance(BaseGraphClient._TIMEZONE_CACHE_TTL + 1)
                assert await client.get_user_timezone() == "Europe/Amsterdam"

        assert client.calls == 1, "denial was re-probed after the TTL expired"
        warnings = [r for r in caplog.records if r.levelno == logging.WARNING]
        assert len(warnings) == 1, f"warning repeated {len(warnings)} times"

    async def test_denial_is_shared_across_client_instances(
        self, fake_clock, caplog, monkeypatch
    ):
        """The cache is class-level, so a second client must not re-probe."""
        monkeypatch.setattr(bc.settings, "user_timezone", "Europe/Amsterdam")
        first, second = DeniedClient(), DeniedClient()

        with caplog.at_level(logging.WARNING, logger=bc.__name__):
            await first.get_user_timezone()
            await second.get_user_timezone()

        assert first.calls == 1
        assert second.calls == 0
        assert len([r for r in caplog.records if r.levelno == logging.WARNING]) == 1

    async def test_warning_is_actionable(self, fake_clock, caplog, monkeypatch):
        """It must name the missing permission AND the timezone we fell back to."""
        monkeypatch.setattr(bc.settings, "user_timezone", "Europe/Amsterdam")
        client = DeniedClient()

        with caplog.at_level(logging.WARNING, logger=bc.__name__):
            await client.get_user_timezone()

        message = caplog.records[0].getMessage()
        assert "MailboxSettings.Read" in message
        assert "Europe/Amsterdam" in message, "must say which timezone is in use"
        assert "403" in message, "must keep the underlying cause"


class TestFallbackResolution:
    """Which timezone a denied tenant actually lands on."""

    async def test_falls_back_to_user_timezone_setting(self, fake_clock, monkeypatch):
        monkeypatch.setattr(bc.settings, "user_timezone", "Asia/Shanghai")
        assert await DeniedClient().get_user_timezone() == "Asia/Shanghai"

    async def test_converts_windows_timezone_names(self, fake_clock, monkeypatch):
        monkeypatch.setattr(bc.settings, "user_timezone", "China Standard Time")
        assert await DeniedClient().get_user_timezone() == "Asia/Shanghai"

    async def test_falls_back_to_system_timezone_when_setting_is_utc(
        self, fake_clock, monkeypatch
    ):
        monkeypatch.setattr(bc.settings, "user_timezone", "UTC")
        result = await DeniedClient().get_user_timezone()
        assert result, "must always resolve to something"


class TestSuccessfulLookup:
    """A working tenant keeps the existing TTL semantics."""

    async def test_success_is_cached(self, fake_clock):
        client = WorkingClient()

        assert await client.get_user_timezone() == "Europe/Amsterdam"
        assert await client.get_user_timezone() == "Europe/Amsterdam"
        assert client.calls == 1

    async def test_success_is_refreshed_after_the_ttl(self, fake_clock):
        """A timezone the user changes in Outlook must still be picked up."""
        client = WorkingClient()
        await client.get_user_timezone()

        client.timezone = "China Standard Time"
        fake_clock.advance(BaseGraphClient._TIMEZONE_CACHE_TTL + 1)

        assert await client.get_user_timezone() == "Asia/Shanghai"
        assert client.calls == 2, "success must expire after the TTL"

    async def test_success_does_not_warn(self, fake_clock, caplog):
        with caplog.at_level(logging.WARNING, logger=bc.__name__):
            await WorkingClient().get_user_timezone()

        assert [r for r in caplog.records if r.levelno == logging.WARNING] == []

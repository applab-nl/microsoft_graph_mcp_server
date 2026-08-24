"""Tests for token persistence in TokenManager.

The token file holds the long-lived refresh_token: losing it forces a manual
interactive re-login, so these tests pin down when the file may be removed and
that it is never written non-atomically.
"""

import builtins
import json
import os
import stat
from pathlib import Path

import pytest

from microsoft_graph_mcp_server.auth_modules import token_manager as tm


VALID_TOKENS = {
    "access_token": "fake-access-token",
    "refresh_token": "fake-refresh-token",
    "token_expiry": 9999999999.0,
    "authenticated": True,
}


@pytest.fixture
def token_file(tmp_path, monkeypatch):
    """Point TOKEN_FILE at an isolated temp file."""
    path = tmp_path / ".microsoft_graph_mcp_tokens.json"
    monkeypatch.setattr(tm, "TOKEN_FILE", path)
    return path


class TestLoadTokensFromDisk:
    """Which failures may delete the token file, and which may not."""

    def test_transient_os_error_preserves_token_file(self, token_file, monkeypatch):
        """An OS/IO error while reading must NOT destroy the refresh_token."""
        token_file.write_text(json.dumps(VALID_TOKENS))

        real_open = builtins.open

        def failing_open(file, *args, **kwargs):
            if str(file) == str(token_file):
                raise OSError("Input/output error")
            return real_open(file, *args, **kwargs)

        monkeypatch.setattr(builtins, "open", failing_open)

        manager = tm.TokenManager()

        monkeypatch.setattr(builtins, "open", real_open)
        assert token_file.exists(), "token file was deleted on a transient read error"
        assert json.loads(token_file.read_text()) == VALID_TOKENS
        assert manager is not None

    def test_permission_error_preserves_token_file(self, token_file, monkeypatch):
        """A PermissionError while reading must NOT destroy the refresh_token."""
        token_file.write_text(json.dumps(VALID_TOKENS))

        real_open = builtins.open

        def failing_open(file, *args, **kwargs):
            if str(file) == str(token_file):
                raise PermissionError("Permission denied")
            return real_open(file, *args, **kwargs)

        monkeypatch.setattr(builtins, "open", failing_open)

        tm.TokenManager()

        monkeypatch.setattr(builtins, "open", real_open)
        assert token_file.exists(), "token file was deleted on a permission error"

    def test_empty_file_is_moved_aside(self, token_file):
        """An empty/truncated file is unusable: quarantine it, don't leave it."""
        token_file.write_text("")

        tm.TokenManager()

        assert not token_file.exists()
        corrupt = token_file.with_name(token_file.name + ".corrupt")
        assert corrupt.exists(), "corrupt token file should be kept for diagnosis"
        assert corrupt.read_text() == ""

    def test_wrong_shape_is_moved_aside(self, token_file):
        """Valid JSON of the wrong shape is corrupt too."""
        token_file.write_text(json.dumps(["not", "a", "token", "object"]))

        tm.TokenManager()

        assert not token_file.exists()
        assert token_file.with_name(token_file.name + ".corrupt").exists()

    def test_valid_file_is_loaded(self, token_file):
        """Sanity check: a healthy file still loads."""
        token_file.write_text(json.dumps(VALID_TOKENS))

        manager = tm.TokenManager()

        assert manager.authenticated is True
        assert manager.refresh_token == VALID_TOKENS["refresh_token"]
        assert token_file.exists()

    def test_expired_access_token_keeps_refresh_token(self, token_file):
        """Expiry must not clear the refresh_token or the file."""
        token_file.write_text(json.dumps({**VALID_TOKENS, "token_expiry": 1.0}))

        manager = tm.TokenManager()

        assert manager.authenticated is False
        assert manager.access_token is None
        assert manager.refresh_token == VALID_TOKENS["refresh_token"]
        assert token_file.exists()


class TestSaveTokensToDisk:
    """Writes must be atomic and the file must stay private."""

    def test_failed_write_leaves_previous_file_intact(self, token_file, monkeypatch):
        """A crash mid-write must not truncate the existing token file."""
        token_file.write_text(json.dumps(VALID_TOKENS))

        def exploding_dump(*args, **kwargs):
            raise RuntimeError("disk full")

        monkeypatch.setattr(tm.json, "dump", exploding_dump)

        manager = tm.TokenManager()
        manager.update_token("new-access-token", expires_in=3600)

        assert json.loads(token_file.read_text()) == VALID_TOKENS

    def test_failed_write_leaves_no_temp_files(self, token_file, monkeypatch):
        """A failed atomic write cleans up after itself."""
        token_file.write_text(json.dumps(VALID_TOKENS))
        manager = tm.TokenManager()

        monkeypatch.setattr(
            tm.json, "dump", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom"))
        )
        manager.update_token("new-access-token")

        leftovers = [p.name for p in token_file.parent.iterdir() if p != token_file]
        assert leftovers == [], f"temp files left behind: {leftovers}"

    def test_token_file_is_not_world_readable(self, token_file):
        """The file holds OAuth refresh tokens: 0600, never group/world readable."""
        manager = tm.TokenManager()
        manager.update_token("an-access-token", refresh_token="a-refresh-token")

        mode = stat.S_IMODE(os.stat(token_file).st_mode)
        assert mode & 0o077 == 0, f"token file mode is {oct(mode)}, expected 0o600"

    def test_write_replaces_file_atomically(self, token_file, monkeypatch):
        """The token path is only ever swapped in via os.replace, never opened 'w'."""
        replaced: list = []
        real_replace = os.replace

        def tracking_replace(src, dst, **kwargs):
            replaced.append((str(src), str(dst)))
            return real_replace(src, dst, **kwargs)

        monkeypatch.setattr(tm.os, "replace", tracking_replace)

        manager = tm.TokenManager()
        manager.update_token("an-access-token", refresh_token="a-refresh-token")

        assert [dst for _, dst in replaced] == [str(token_file)]
        assert json.loads(token_file.read_text())["refresh_token"] == "a-refresh-token"

    def test_round_trip(self, token_file):
        """What we save is what we load back."""
        manager = tm.TokenManager()
        manager.update_token("an-access-token", refresh_token="a-refresh-token")

        reloaded = tm.TokenManager()
        assert reloaded.access_token == "an-access-token"
        assert reloaded.refresh_token == "a-refresh-token"
        assert reloaded.authenticated is True

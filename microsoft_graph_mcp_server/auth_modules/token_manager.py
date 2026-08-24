"""Token management for Microsoft Graph API authentication."""

import json
import logging
import os
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, Optional

from ..config import settings

logger = logging.getLogger(__name__)


TOKEN_FILE = Path.home() / ".microsoft_graph_mcp_tokens.json"
DEVICE_FLOW_FILE = Path.home() / ".microsoft_graph_mcp_device_flows.json"
LATEST_DEVICE_CODE_FILE = Path.home() / ".microsoft_graph_mcp_latest_device_code.json"


class TokenManager:
    """Manages authentication tokens for Microsoft Graph API."""

    def __init__(self):
        self.access_token: Optional[str] = None
        self.token_expiry: float = 0
        self.refresh_token: Optional[str] = None
        self.authenticated: bool = False

        self.load_tokens_from_disk()

    def save_tokens_to_disk(self) -> None:
        """Save authentication tokens to disk atomically."""
        try:
            token_data = {
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "token_expiry": self.token_expiry,
                "authenticated": self.authenticated,
            }
            # Write to a temp file in the same directory and swap it in, so a
            # crash or a concurrent read can never observe a truncated (or
            # zero-length) token file. mkstemp creates it mode 0600, which is
            # what we want for a file holding a long-lived refresh_token.
            fd, tmp_path = tempfile.mkstemp(
                dir=str(TOKEN_FILE.parent),
                prefix=f".{TOKEN_FILE.name}.",
                suffix=".tmp",
            )
            try:
                with os.fdopen(fd, "w") as f:
                    json.dump(token_data, f, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp_path, TOKEN_FILE)
            except Exception:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
                raise
        except Exception as e:
            logger.warning(f"Failed to save tokens to disk: {e}")

    def load_tokens_from_disk(self) -> None:
        """Load authentication tokens from disk."""
        if not TOKEN_FILE.exists():
            self.access_token = None
            self.token_expiry = 0
            self.refresh_token = None
            self.authenticated = False
            return

        try:
            with open(TOKEN_FILE, "r") as f:
                token_data = json.load(f)

            if not isinstance(token_data, dict):
                raise ValueError(
                    f"expected a JSON object, got {type(token_data).__name__}"
                )
        except (ValueError, UnicodeDecodeError) as e:
            # Genuinely corrupt content (empty/truncated file, or valid JSON of
            # the wrong shape). json.JSONDecodeError is a ValueError. The file
            # is unusable, so move it aside - it stays diagnosable, and the next
            # save writes a clean one.
            logger.warning(f"Token file is corrupt, moving it aside: {e}")
            self.quarantine_corrupt_token_file()
            return
        except Exception as e:
            # Transient failure (OS/IO/permission error, interrupted read).
            # NEVER delete the file here: it holds the long-lived refresh_token,
            # and destroying it turns one bad read into a dead session that only
            # an interactive re-login can fix.
            logger.warning(f"Failed to load tokens from disk: {e}")
            return

        self.access_token = token_data.get("access_token")
        self.refresh_token = token_data.get("refresh_token")
        self.token_expiry = token_data.get("token_expiry", 0)
        self.authenticated = token_data.get("authenticated", False)

        # Check if access token is expired
        # IMPORTANT: We keep refresh_token even when access_token expires
        # This allows automatic token refresh without requiring user to login again
        if self.authenticated and self.access_token:
            current_time = time.time()
            if current_time >= self.token_expiry - 60:
                # Access token expired, but keep refresh_token for auto-refresh
                logger.info("Access token expired on load, but refresh_token is preserved for auto-refresh")
                self.authenticated = False
                self.access_token = None
                # DO NOT delete tokens from disk - keep refresh_token

    def quarantine_corrupt_token_file(self) -> None:
        """Move an unreadable token file aside so the failure stays diagnosable."""
        corrupt_file = TOKEN_FILE.with_name(TOKEN_FILE.name + ".corrupt")
        try:
            os.replace(TOKEN_FILE, corrupt_file)
        except Exception as e:
            logger.warning(f"Failed to move corrupt token file aside: {e}")
            self.delete_tokens_from_disk()

    def delete_tokens_from_disk(self) -> None:
        """Delete authentication tokens from disk."""
        try:
            if TOKEN_FILE.exists():
                os.remove(TOKEN_FILE)
        except Exception as e:
            logger.warning(f"Failed to delete tokens from disk: {e}")

    def update_token(
        self,
        access_token: str,
        expires_in: int = 3600,
        refresh_token: Optional[str] = None,
    ) -> None:
        """Update the access token and related information."""
        self.access_token = access_token
        self.token_expiry = time.time() + expires_in
        self.refresh_token = refresh_token or self.refresh_token
        self.authenticated = True
        self.save_tokens_to_disk()

    def clear_tokens(self) -> None:
        """Clear all authentication tokens."""
        self.access_token = None
        self.token_expiry = 0
        self.refresh_token = None
        self.authenticated = False
        self.delete_tokens_from_disk()

    def is_token_valid(self, buffer_seconds: int = 60) -> bool:
        """Check if the current token is valid and not expired.
        
        Args:
            buffer_seconds: Buffer time in seconds before actual expiry (default: 60)
                           Token is considered invalid if it expires within this buffer.
        """
        if not self.authenticated or not self.access_token:
            return False
        return time.time() < self.token_expiry - buffer_seconds

    def get_token_expiry_info(self) -> dict:
        """Get token expiry information."""
        remaining_seconds = int(self.token_expiry - time.time())
        remaining_minutes = remaining_seconds // 60
        remaining_hours = remaining_minutes // 60
        remaining_minutes_display = remaining_minutes % 60

        # Simple display string
        if remaining_hours > 0:
            display = f"{remaining_hours}h {remaining_minutes_display}m"
        else:
            display = f"{remaining_minutes_display}m"

        return {
            "seconds": remaining_seconds,
            "display": display,
        }

    def save_device_flow(self, device_code: str, device_flow: Dict[str, Any]) -> None:
        """Save device flow to disk using device_code as key."""
        try:
            flows = {}
            if DEVICE_FLOW_FILE.exists():
                with open(DEVICE_FLOW_FILE, "r") as f:
                    flows = json.load(f)

            flows[device_code] = device_flow
            with open(DEVICE_FLOW_FILE, "w") as f:
                json.dump(flows, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save device flow to disk: {e}")

    def load_device_flow(self, device_code: str) -> Optional[Dict[str, Any]]:
        """Load device flow from disk using device_code as key."""
        if not DEVICE_FLOW_FILE.exists():
            return None

        try:
            with open(DEVICE_FLOW_FILE, "r") as f:
                flows = json.load(f)

            return flows.get(device_code)
        except Exception as e:
            logger.warning(f"Failed to load device flow from disk: {e}")
            return None

    def delete_device_flow(self, device_code: str) -> None:
        """Delete device flow from disk using device_code as key."""
        try:
            if DEVICE_FLOW_FILE.exists():
                with open(DEVICE_FLOW_FILE, "r") as f:
                    flows = json.load(f)

                if device_code in flows:
                    del flows[device_code]
                    with open(DEVICE_FLOW_FILE, "w") as f:
                        json.dump(flows, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to delete device flow from disk: {e}")

    def cleanup_expired_device_flows(self) -> None:
        """Clean up expired device flows from disk."""
        if not DEVICE_FLOW_FILE.exists():
            return

        try:
            with open(DEVICE_FLOW_FILE, "r") as f:
                flows = json.load(f)

            current_time = time.time()
            flows_to_delete = []

            for device_code, flow in flows.items():
                expires_at = flow.get("expires_at")
                if expires_at and current_time >= expires_at:
                    flows_to_delete.append(device_code)

            if flows_to_delete:
                for device_code in flows_to_delete:
                    del flows[device_code]

                with open(DEVICE_FLOW_FILE, "w") as f:
                    json.dump(flows, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to cleanup expired device flows from disk: {e}")

    def save_latest_device_code(self, device_code: str) -> None:
        """Save the latest device_code to disk."""
        try:
            data = {"device_code": device_code, "timestamp": time.time()}
            with open(LATEST_DEVICE_CODE_FILE, "w") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save latest device_code to disk: {e}")

    def get_latest_device_code(self) -> Optional[str]:
        """Get the latest device_code from disk."""
        if not LATEST_DEVICE_CODE_FILE.exists():
            return None

        try:
            with open(LATEST_DEVICE_CODE_FILE, "r") as f:
                data = json.load(f)

            device_code = data.get("device_code")

            if device_code:
                return device_code
            return None
        except Exception as e:
            logger.warning(f"Failed to load latest device_code from disk: {e}")
            return None

    def clear_latest_device_code(self) -> None:
        """Clear the latest device_code from disk."""
        try:
            if LATEST_DEVICE_CODE_FILE.exists():
                os.remove(LATEST_DEVICE_CODE_FILE)
        except Exception as e:
            logger.warning(f"Failed to clear latest device_code from disk: {e}")

    def clear_all_device_flows(self) -> None:
        """Clear all device flows from disk."""
        try:
            if DEVICE_FLOW_FILE.exists():
                os.remove(DEVICE_FLOW_FILE)
        except Exception as e:
            logger.warning(f"Failed to clear device flows from disk: {e}")

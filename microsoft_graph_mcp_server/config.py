"""Configuration module for Microsoft Graph MCP Server."""

import os
from pathlib import Path
from typing import Optional

try:
    from pydantic_settings import BaseSettings, SettingsConfigDict
except ImportError:
    from pydantic import BaseSettings, ConfigDict as SettingsConfigDict


# Microsoft Graph API service limits
MAX_RECIPIENTS_LIMIT = 500  # Maximum total recipients (TO + CC + BCC) per email


class Settings(BaseSettings):
    """Application settings."""

    # Microsoft Graph API configuration
    client_id: str = os.getenv("CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")
    client_secret: str = os.getenv("CLIENT_SECRET", "")
    tenant_id: str = os.getenv("TENANT_ID", "organizations")

    # User timezone configuration (fallback if API fails)
    user_timezone: str = os.getenv("USER_TIMEZONE", "UTC")

    # MCP Server configuration
    server_name: str = "microsoft-graph-mcp-server"
    server_version: str = "0.1.0"

    # API endpoints
    graph_api_base_url: str = "https://graph.microsoft.com/v1.0"
    auth_url: str = "https://login.microsoftonline.com"

    # Cache settings
    cache_ttl: int = 300  # 5 minutes

    # Search settings
    default_search_days: int = int(
        os.getenv("DEFAULT_SEARCH_DAYS", "7")
    )  # Default search range in days when not specified
    max_search_days: int = int(
        os.getenv("MAX_SEARCH_DAYS", "90")
    )  # Maximum allowed search range in days

    # Pagination settings
    page_size: int = int(
        os.getenv("PAGE_SIZE", "5")
    )  # Number of items per page for user browsing
    llm_page_size: int = int(
        os.getenv("LLM_PAGE_SIZE", "20")
    )  # Number of items per page for LLM browsing
    # How many messages search_emails fetches into the cache in one call.
    # Distinct from the two page sizes above, which paginate BROWSING of an
    # already-populated cache; this bounds what the search puts there at all.
    # Iris injects it (see buildMcpEnv in src/lib/graph/client.ts) so it always
    # equals Iris's own per-run cap — its watermark logic treats
    # "fewer results than the cap" as proof a search was not truncated, so the
    # two numbers must not drift. Ceiling is MAX_EMAIL_SEARCH_LIMIT (1000).
    search_page_size: int = int(os.getenv("SEARCH_PAGE_SIZE", "100"))

    # Contact search settings
    contact_search_limit: int = int(
        os.getenv("CONTACT_SEARCH_LIMIT", "10")
    )  # Maximum contacts to return in search results

    # Calendar search settings
    calendar_search_past_days: int = int(
        os.getenv("CALENDAR_SEARCH_PAST_DAYS", "30")
    )  # Default past days to search for calendar events
    calendar_search_future_days: int = int(
        os.getenv("CALENDAR_SEARCH_FUTURE_DAYS", "90")
    )  # Default future days to search for calendar events

    # LLM capability settings
    multimodal_supported: bool = False  # Whether LLM supports multimodal (image processing), read from .env by pydantic-settings

    # Image compression settings (for multimodal support)
    # Note: IMAGE_MAX_SIZE_KB in .env is in KB, converted to bytes internally
    image_max_size_kb: int = 150  # Max image size in KB
    image_max_dimension: int = 1024  # Max width/height in pixels
    image_quality: int = 75  # JPEG quality (1-100)

    @property
    def image_max_size(self) -> int:
        """Get max image size in bytes."""
        return self.image_max_size_kb * 1024

    model_config = SettingsConfigDict(
        env_file=str(Path(__file__).parent.parent / ".env"),
        env_file_encoding="utf-8",
        extra="ignore"
    )


# Global settings instance
settings = Settings()

# Debug log for multimodal setting
import logging
logger = logging.getLogger(__name__)
logger.info(f"[CONFIG DEBUG] Settings loaded from .env: multimodal_supported={settings.multimodal_supported}")

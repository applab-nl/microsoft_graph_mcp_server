"""Base client class for Microsoft Graph API clients."""

import asyncio
import logging
import time
from datetime import datetime
from typing import Any, Dict, Optional

import httpx

from ..auth import auth_manager
from ..config import settings
from ..utils import DateHandler as date_handler

logger = logging.getLogger(__name__)


class RateLimitError(Exception):
    """Exception raised when rate limit is exceeded."""
    
    def __init__(self, message: str, retry_after: Optional[int] = None):
        super().__init__(message)
        self.retry_after = retry_after


class BaseGraphClient:
    """Base client for Microsoft Graph API operations."""

    # Class-level cache for user timezone (shared across all instances)
    _user_timezone_cache: Optional[str] = None
    _user_timezone_cache_time: Optional[float] = None
    _TIMEZONE_CACHE_TTL = 3600  # 1 hour cache TTL

    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
        self._semaphore = asyncio.Semaphore(20)
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_client(self) -> httpx.AsyncClient:
        """Get or create the HTTP client."""
        if self._client is None or self._client.is_closed:
            self._client = httpx.AsyncClient(timeout=self.timeout)
        return self._client

    async def close(self):
        """Close the HTTP client."""
        if self._client and not self._client.is_closed:
            await self._client.aclose()

    async def _make_request(
        self,
        method: str,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
        max_retries: int = 3,
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API with concurrency control and rate limiting handling."""

        async with self._semaphore:
            access_token = await auth_manager.get_access_token()

            default_headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            }

            if headers:
                default_headers.update(headers)

            url = f"{self.base_url}{endpoint}"

            client = await self._get_client()
            
            for attempt in range(max_retries + 1):
                response = await client.request(
                    method=method,
                    url=url,
                    params=params,
                    json=data,
                    headers=default_headers,
                )

                if response.status_code in (200, 201):
                    return response.json()
                elif response.status_code == 202:
                    return {"status": "accepted"}
                elif response.status_code == 204:
                    return {"status": "success"}
                elif response.status_code == 429:
                    retry_after = self._extract_retry_after(response)
                    if attempt < max_retries:
                        wait_time = retry_after if retry_after else 5
                        wait_time = min(wait_time * (2 ** attempt), 60)
                        logger.warning(
                            f"Rate limited (429). Waiting {wait_time} seconds before retry {attempt + 1}/{max_retries}"
                        )
                        await asyncio.sleep(wait_time)
                        continue
                    else:
                        error_msg = response.text
                        raise RateLimitError(
                            f"Rate limit exceeded after {max_retries} retries. {error_msg}",
                            retry_after=retry_after
                        )
                else:
                    raise Exception(
                        f"Graph API request failed: {response.status_code} - {response.text}"
                    )

    def _extract_retry_after(self, response: httpx.Response) -> Optional[int]:
        """Extract Retry-After header from response.
        
        Returns:
            Number of seconds to wait, or None if not specified
        """
        retry_after = response.headers.get("Retry-After")
        if retry_after:
            try:
                return int(retry_after)
            except ValueError:
                pass
        return None

    async def get(
        self,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make GET request to Graph API."""
        return await self._make_request("GET", endpoint, params=params, headers=headers)

    async def post(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make POST request to Graph API."""
        return await self._make_request("POST", endpoint, data=data)

    async def patch(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make PATCH request to Graph API."""
        return await self._make_request("PATCH", endpoint, data=data)

    async def put(
        self, endpoint: str, data: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Make PUT request to Graph API."""
        return await self._make_request("PUT", endpoint, data=data)

    async def delete(self, endpoint: str) -> Dict[str, Any]:
        """Make DELETE request to Graph API."""
        return await self._make_request("DELETE", endpoint)

    @classmethod
    def _cache_user_timezone(cls, timezone: str, sticky: bool) -> None:
        """Store the resolved timezone in the class-level cache.

        A ``sticky`` entry never expires: it records that Graph refused to tell
        us, which is a tenant/app-registration state that only changes when an
        admin grants the permission (and the server is restarted). Everything
        else expires after _TIMEZONE_CACHE_TTL so a timezone the user changes
        in Outlook is still picked up.
        """
        BaseGraphClient._user_timezone_cache = timezone
        BaseGraphClient._user_timezone_cache_time = None if sticky else time.time()

    @staticmethod
    def _resolve_fallback_timezone() -> str:
        """Timezone to use when Graph can't tell us: USER_TIMEZONE, then system, then UTC."""
        user_tz = date_handler.convert_to_iana_timezone(settings.user_timezone)
        if user_tz != "UTC":
            return user_tz

        try:
            local_tz = datetime.now().astimezone().tzinfo
            if local_tz:
                tz_str = str(local_tz)
                if tz_str and tz_str != "UTC":
                    return date_handler.convert_to_iana_timezone(tz_str)
        except Exception:
            pass

        return "UTC"

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier from Microsoft Graph mailbox settings.

        First attempts to get timezone from Graph API mailbox settings.
        Falls back to config setting or system timezone if unavailable.
        Uses a class-level cache to avoid repeated API calls. A refusal is
        cached for the whole session, so a tenant that denies mailboxSettings
        is probed (and warned about) ONCE per process rather than every time
        the cache TTL lapses.
        """
        # Check cache first (cache_time None = sticky entry, never expires)
        current_time = time.time()
        if BaseGraphClient._user_timezone_cache is not None and (
            BaseGraphClient._user_timezone_cache_time is None
            or current_time - BaseGraphClient._user_timezone_cache_time
            < BaseGraphClient._TIMEZONE_CACHE_TTL
        ):
            return BaseGraphClient._user_timezone_cache

        # Try to get timezone from Graph API mailbox settings
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            mailbox_settings = result.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
        except Exception as e:
            fallback_tz = self._resolve_fallback_timezone()
            # Log ONCE - the sticky cache below keeps subsequent calls silent.
            # Name the timezone actually in use: this drives meeting times, so a
            # wrong fallback shows the wrong times.
            logger.warning(
                f"Failed to get timezone from Graph API - falling back to "
                f"{fallback_tz} for the rest of this session (all times are "
                f"shown in {fallback_tz}). A 403 ErrorAccessDenied here means "
                f"the app registration is missing the MailboxSettings.Read "
                f"permission; grant it and restart the server, or set "
                f"USER_TIMEZONE to the correct timezone. Cause: {e}"
            )
            self._cache_user_timezone(fallback_tz, sticky=True)
            return fallback_tz

        if timezone:
            iana_tz = date_handler.convert_to_iana_timezone(timezone)
            logger.info(f"Retrieved timezone from Graph API: {timezone} -> {iana_tz}")
            self._cache_user_timezone(iana_tz, sticky=False)
            return iana_tz

        # Graph answered but has no timezone set - keep the TTL so a value the
        # user sets later is picked up.
        fallback_tz = self._resolve_fallback_timezone()
        logger.info(f"No timezone in Graph API mailbox settings, using {fallback_tz}")
        self._cache_user_timezone(fallback_tz, sticky=False)
        return fallback_tz

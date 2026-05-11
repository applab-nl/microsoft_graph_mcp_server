"""User client for Microsoft Graph API."""

import logging
from datetime import datetime
from typing import Optional, List, Dict, Any

from .base_client import BaseGraphClient, RateLimitError
from ..utils import DateHandler as date_handler
from ..config import settings

logger = logging.getLogger(__name__)


class UserClient(BaseGraphClient):
    """Client for user-related operations."""

    def __init__(self):
        """Initialize the user client."""
        super().__init__()
        self._user_email_cache: Optional[str] = None
        # Cached timezone: None = not yet attempted; str = resolved (either
        # from Graph or a fallback). Once populated — including via a failure
        # path — subsequent calls return immediately. This stops corporate
        # tenants that deny /me/mailboxSettings from spamming a 403 warning
        # on every get_email_content() call.
        self._user_timezone_cache: Optional[str] = None

    def _get_system_timezone(self) -> str:
        """Get the system's local timezone as a fallback.

        Returns:
            IANA timezone name for the system's local timezone
        """
        try:
            local_tz = datetime.now().astimezone().tzinfo
            logger.debug(f"System timezone info: {local_tz}, type: {type(local_tz)}")

            if hasattr(local_tz, "key"):
                tz_key = str(local_tz.key)
                logger.debug(f"System timezone key: {tz_key}")
                if tz_key and tz_key != "UTC":
                    converted = date_handler.convert_to_iana_timezone(tz_key)
                    logger.debug(f"Converted system timezone: {converted}")
                    return converted

            tz_name = str(local_tz)
            logger.debug(f"System timezone name: {tz_name}")
            if tz_name and tz_name != "UTC":
                converted = date_handler.convert_to_iana_timezone(tz_name)
                if converted != "UTC":
                    logger.debug(f"Converted system timezone from name: {converted}")
                    return converted
        except Exception as e:
            logger.warning(f"Failed to get system timezone: {e}")

        logger.debug("Defaulting to UTC")
        return "UTC"

    async def get_user_email(self) -> Optional[str]:
        """Get the current user's email address with caching."""
        # Return cached email if available
        if self._user_email_cache is not None:
            return self._user_email_cache

        try:
            params = {"$select": "mail,userPrincipalName"}
            result = await self.get("/me", params=params)
            user_email = result.get("mail") or result.get("userPrincipalName")
            if user_email:
                self._user_email_cache = user_email
            return user_email
        except Exception as e:
            logger.warning(f"Failed to get user email: {e}")
            return None

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier from Microsoft Graph mailbox settings.

        Cached per-process. First call attempts /me/mailboxSettings; on
        success or failure the resolved value (or fallback) is cached and
        all subsequent calls return immediately without hitting Graph.
        Restart the MCP server to retry — useful after a tenant admin
        grants the missing permission.
        """
        if self._user_timezone_cache is not None:
            return self._user_timezone_cache

        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            mailbox_settings = result.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                iana_tz = date_handler.convert_to_iana_timezone(timezone)
                logger.info(
                    f"Retrieved timezone from Graph API: {timezone} -> {iana_tz}"
                )
                self._user_timezone_cache = iana_tz
                return iana_tz
            logger.info("No timezone in Graph API mailbox settings")
        except Exception as e:
            # Log ONCE — subsequent calls hit the cache and stay silent.
            logger.warning(
                f"Failed to get timezone from Graph API (using fallback for "
                f"the rest of this session): {e}"
            )

        user_tz = date_handler.convert_to_iana_timezone(settings.user_timezone)
        if user_tz != "UTC":
            logger.info(
                f"Using USER_TIMEZONE setting: {settings.user_timezone} -> {user_tz}"
            )
            self._user_timezone_cache = user_tz
            return user_tz

        system_tz = self._get_system_timezone()
        logger.info(f"Using system timezone as fallback: {system_tz}")
        self._user_timezone_cache = system_tz
        return system_tz

    async def get_users(
        self, filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get("/users", params=params)
        return result.get("value", [])

    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.get(f"/users/{user_id}")

    async def get_user_by_email(self, email: str) -> Optional[Dict[str, Any]]:
        """Get user by email address with mailbox settings."""
        try:
            params = {
                "$filter": f"mail eq '{email}'",
                "$select": "id,displayName,mail,mailboxSettings",
            }
            result = await self.get("/users", params=params)
            users = result.get("value", [])
            if users:
                return users[0]
        except Exception:
            pass
        return None

    async def get_user_timezone_by_email(self, email: str) -> Optional[str]:
        """Get user's timezone by email address."""
        user = await self.get_user_by_email(email)
        if user:
            mailbox_settings = user.get("mailboxSettings", {})
            timezone = mailbox_settings.get("timeZone")
            if timezone:
                return timezone
        return None

    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search for people by name or email address in organization directory.

        This searches the organization directory (/users) to find people.
        Use this when you need to find information about a person,
        such as "who is Joyson Barrago" or "find contact with email joyson@ibm.com".

        This is NOT for searching email messages - use search_emails for that.

        Uses smart detection to automatically choose the optimal search method:
        - Email addresses: Fast $filter with exact match
        - Names: $search with tokenization for contains-like behavior
        """
        try:
            # Strategy: Detect query type and use optimal method
            
            # 1. If query looks like an email address (contains @ and . after @)
            if "@" in query and "." in query.split("@")[-1]:
                # Use $filter for exact match - FASTEST
                params = {
                    "$filter": f"mail eq '{query}' OR userPrincipalName eq '{query}'",
                    "$top": top
                }
                result = await self.get("/users", params=params)
            # 2. For name searches (use $search for contains-like behavior)
            else:
                # Use $search which provides tokenized search (contains-like)
                params = {"$search": f'"displayName:{query}"', "$top": top}
                headers = {"ConsistencyLevel": "eventual"}
                result = await self.get("/users", params=params, headers=headers)
            
            contacts = result.get("value", [])
            return contacts[:top]
        except RateLimitError as e:
            wait_time = e.retry_after if e.retry_after else "a few minutes"
            logger.warning(f"Rate limit exceeded. Please wait {wait_time} seconds before retrying.")
            raise RateLimitError(
                f"Rate limit exceeded. Please wait {wait_time} seconds before retrying.",
                retry_after=e.retry_after
            )
        except Exception as e:
            logger.error(f"Error searching contacts: {e}")
            raise

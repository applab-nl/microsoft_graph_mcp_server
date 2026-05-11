"""Calendar handlers for MCP tools."""

import json
import logging
import re
import urllib.parse
from datetime import datetime
from mcp import types

from .base import BaseHandler
from ..graph_client import graph_client
from ..config import settings
from ..cache import event_cache
from ..utils import DateHandler, normalize_email_html
from ..clients.calendar_client import MAX_EVENT_SEARCH_LIMIT

logger = logging.getLogger(__name__)


class CalendarHandler(BaseHandler):
    """Handler for calendar-related tools."""

    def _convert_microsoft_timezone_to_iana(self, ms_timezone: str) -> str:
        """Convert Microsoft timezone name to IANA timezone name.

        Args:
            ms_timezone: Microsoft timezone name (e.g., "India Standard Time")

        Returns:
            IANA timezone name (e.g., "Asia/Kolkata")
        """
        timezone_mapping = {
            "India Standard Time": "Asia/Kolkata",
            "China Standard Time": "Asia/Shanghai",
            "Tokyo Standard Time": "Asia/Tokyo",
            "Singapore Standard Time": "Asia/Singapore",
            "Eastern Standard Time": "America/New_York",
            "Pacific Standard Time": "America/Los_Angeles",
            "Central Standard Time": "America/Chicago",
            "GMT Standard Time": "Europe/London",
            "W. Europe Standard Time": "Europe/Paris",
            "Romance Standard Time": "Europe/Paris",
            "FLE Standard Time": "Europe/Kiev",  # Finland, Russia (Kaliningrad), Estonia, Latvia, Lithuania
            "E. Europe Standard Time": "Europe/Chisinau",
            "Central Europe Standard Time": "Europe/Budapest",
            "Central European Standard Time": "Europe/Warsaw",
            "AUS Eastern Standard Time": "Australia/Sydney",
            "UTC": "UTC",
        }

        return timezone_mapping.get(ms_timezone, ms_timezone)

    def _normalize_teams_url(self, url: str) -> str:
        """Normalize Teams URL to ensure it opens in new Teams experience.

        Args:
            url: The Teams URL to normalize

        Returns:
            Normalized URL that works with new Teams
        """
        url = url.strip()

        if "teams.microsoft.com" in url:
            parsed_url = urllib.parse.urlparse(url)
            query_params = urllib.parse.parse_qs(parsed_url.query)

            if "context" in query_params:
                return url

            new_params = {}
            for key, values in query_params.items():
                if key in ["anon", "p", "subject", "startTime", "duration"]:
                    new_params[key] = values[0]

            new_params["allowNewTeams"] = "true"

            new_query = urllib.parse.urlencode(new_params)

            if new_query:
                normalized_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}?{new_query}"
            else:
                normalized_url = (
                    f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"
                )

            return normalized_url

        return url

    def _extract_meeting_info(self, event: dict) -> dict:
        """Extract meeting information from event, including Teams and other meeting links.

        Args:
            event: Full event object from Graph API

        Returns:
            Dictionary with meeting information including:
            - type: Meeting type (teams, zoom, google_meet, other, none)
            - joinUrl: Join URL if available
            - onlineMeeting: Full onlineMeeting object if available
            - locationMeetingUrl: Meeting URL extracted from location field
            - bodyMeetingUrl: Meeting URL extracted from body content
        """
        meeting_info = {
            "type": "none",
            "joinUrl": None,
            "onlineMeeting": None,
            "locationMeetingUrl": None,
            "bodyMeetingUrl": None,
        }

        online_meeting = event.get("onlineMeeting", {})
        if online_meeting:
            meeting_info["onlineMeeting"] = online_meeting
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                join_url = join_url.strip().strip("`").strip()
                join_url = self._normalize_teams_url(join_url)
                meeting_info["type"] = "teams"
                meeting_info["joinUrl"] = join_url

        location = event.get("location", {})
        location_display = location.get("displayName", "")

        if location_display and not meeting_info["joinUrl"]:
            url_patterns = {
                "zoom": r"https?://(?:[\w-]+\.)?zoom\.us/\S+",
                "google_meet": r"https?://meet\.google\.com/[a-z0-9-]+",
                "teams": r"https?://teams\.microsoft\.com/\S+",
                "webex": r"https?://[\w-]*webex\.com/\S+",
                "skype": r"https?://join\.skype\.com/\S+",
                "generic": r"https?://\S+",
            }

            for meet_type, pattern in url_patterns.items():
                match = re.search(pattern, location_display, re.IGNORECASE)
                if match:
                    url = match.group(0).strip().strip("`").strip()
                    if meet_type == "teams":
                        url = self._normalize_teams_url(url)
                    meeting_info["locationMeetingUrl"] = url
                    if meet_type == "generic":
                        meeting_info["type"] = "other"
                    else:
                        meeting_info["type"] = meet_type
                        meeting_info["joinUrl"] = url
                    break

        if not meeting_info["joinUrl"]:
            body = event.get("body", {})
            body_content = body.get("content", "")
            content_type = body.get("contentType", "")

            if body_content and content_type == "html":
                body_url_patterns = {
                    "teams": r'https?://teams\.microsoft\.com/l/meetup-join/[^\s"<>]+',
                    "zoom": r'https?://(?:[\w-]+\.)?zoom\.us/[j|s|w][^\s"<>]*',
                    "google_meet": r"https?://meet\.google\.com/[a-z0-9-]+",
                    "webex": r'https?://[\w-]*webex\.com/[^\s"<>]+',
                    "skype": r'https?://join\.skype\.com/[^\s"<>]+',
                    "generic": r'https?://[^\s"<>]+',
                }

                for meet_type, pattern in body_url_patterns.items():
                    match = re.search(pattern, body_content, re.IGNORECASE)
                    if match:
                        url = match.group(0).strip().strip("`").strip()
                        if meet_type == "teams":
                            url = self._normalize_teams_url(url)
                        meeting_info["bodyMeetingUrl"] = url
                        meeting_info["type"] = (
                            meet_type if meet_type != "generic" else "other"
                        )
                        meeting_info["joinUrl"] = url
                        break

        return meeting_info

    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content with aggressive cleanup (optimized)."""
        text = html_content

        text = re.sub(
            r"<(head|style|script).*?>.*?</\1>",
            "",
            text,
            flags=re.DOTALL | re.IGNORECASE,
        )
        text = re.sub(
            r'\s*(style|class|id|data-outlook-trace)="[^"]*"',
            "",
            text,
            flags=re.IGNORECASE,
        )
        text = re.sub(r"<(img|hr)[^>]*>", "", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(
            r"&(nbsp|amp|lt|gt|quot|#39);",
            lambda m: {
                "nbsp": " ",
                "amp": "&",
                "lt": "<",
                "gt": ">",
                "quot": '"',
                "#39": "'",
            }[m.group(1)],
            text,
        )
        text = re.sub(r"\s+", " ", text)
        text = re.sub(r"(\n\s*){3,}", "\n\n", text)

        return text.strip()

    async def handle_browse_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle browse_events tool."""
        page_number = arguments["page_number"]
        mode = arguments.get("mode", "llm")

        page_size = settings.llm_page_size if mode == "llm" else settings.page_size
        skip = (page_number - 1) * page_size

        cached_events = event_cache.get_cached_events()
        total_count = len(cached_events)

        if total_count == 0:
            return self._format_response(
                {
                    "message": "No events in cache. Use search_events to load events first.",
                    "events": [],
                    "count": 0,
                }
            )

        start_idx = skip
        end_idx = start_idx + page_size
        page_events = cached_events[start_idx:end_idx]

        filtered_events = []
        for event in page_events:
            attendees_list = event.get("attendees_list", [])
            attendees_display = []
            for attendee in attendees_list:
                email_info = attendee.get("emailAddress", {})
                name = email_info.get("name", "")
                email = email_info.get("address", "")
                if name:
                    attendees_display.append(f"{name} ({email})")
                elif email:
                    attendees_display.append(email)

            filtered_event = {
                "number": event.get("number", 0),
                "subject": event.get("subject", ""),
                "start": event.get("start", ""),
                "end": event.get("end", ""),
                "location": event.get("location", ""),
                "organizer": event.get("organizer", {}),
                "attendees": len(attendees_list),
                "attendees_list": attendees_display,
                "isAllDay": event.get("isAllDay", False),
                "isCancelled": event.get("isCancelled", False),
                "status": event.get("status", {}),
                "showAs": event.get("showAs", ""),
                "importance": event.get("importance", "normal"),
                "type": event.get("type", "singleInstance"),
                "recurrence": event.get("recurrence", False),
                "recurrenceInfo": event.get("recurrenceInfo"),
                "seriesMasterId": event.get("seriesMasterId"),
                "responseStatus": event.get("responseStatus", {}),
                "sensitivity": event.get("sensitivity", "normal"),
                "meetingType": event.get("meetingType", "none"),
                "hasMeeting": event.get("hasMeeting", False),
            }
            filtered_events.append(filtered_event)

        success, user_timezone, error = await self._handle_auth_error(
            lambda: graph_client.get_user_timezone(), "getting user timezone"
        )
        if not success:
            return self._format_error(error)

        return self._format_response(
            {
                "events": filtered_events,
                "count": len(filtered_events),
                "total_count": total_count,
                "current_page": page_number,
                "total_pages": (total_count + page_size - 1) // page_size,
                "page_size": page_size,
                "mode": mode,
                "timezone": user_timezone,
            }
        )

    async def handle_get_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle get_event tool."""
        event_number = int(arguments["cache_number"])

        cached_events = event_cache.get_cached_events()

        for event in cached_events:
            if event.get("number") == event_number:
                event_id = event.get("id")
                if event_id:
                    success, full_event, error = await self._handle_auth_error(
                        lambda: graph_client.get_event(event_id),
                        "getting event details",
                    )
                    if not success:
                        return self._format_error(error)

                    meeting_info = self._extract_meeting_info(full_event)

                    body = full_event.get("body", {})
                    body_content = body.get("content", "")
                    body_type = body.get("contentType", "Text")

                    if body_type.lower() == "html" and body_content:
                        body_content = self._extract_text_from_html(body_content)

                    # Get user timezone for time conversion
                    user_timezone = await graph_client.get_user_timezone()

                    # Format start time
                    start_obj = full_event.get("start", {})
                    start_utc = start_obj.get("dateTime", "")
                    start_formatted = {
                        "time": "",
                        "timeZone": user_timezone,
                        # Raw Microsoft Graph UTC ISO 8601 string. Surfaced so
                        # MCP clients can do their own TZ math without trusting
                        # `time` (which depends on whatever the user's M365
                        # mailbox timezone happens to be set to).
                        "dateTime": start_utc,
                    }
                    if start_utc:
                        start_formatted["time"] = DateHandler.convert_utc_to_user_timezone(
                            start_utc, user_timezone, "%a %m/%d/%Y %I:%M %p"
                        )

                    # Format end time
                    end_obj = full_event.get("end", {})
                    end_utc = end_obj.get("dateTime", "")
                    end_formatted = {
                        "time": "",
                        "timeZone": user_timezone,
                        "dateTime": end_utc,
                    }
                    if end_utc:
                        end_formatted["time"] = DateHandler.convert_utc_to_user_timezone(
                            end_utc, user_timezone, "%a %m/%d/%Y %I:%M %p"
                        )

                    readable_event = {
                        "subject": full_event.get("subject", ""),
                        "start": start_formatted,
                        "end": end_formatted,
                        "location": full_event.get("location", {}),
                        "organizer": full_event.get("organizer", {}),
                        "attendees": full_event.get("attendees", []),
                        "isAllDay": full_event.get("isAllDay", False),
                        "showAs": full_event.get("showAs", ""),
                        "importance": full_event.get("importance", "normal"),
                        "sensitivity": full_event.get("sensitivity", "normal"),
                        "body": {"contentType": body_type, "content": body_content},
                        "responseStatus": full_event.get("responseStatus", {}),
                        "type": full_event.get("type", "singleInstance"),
                        "createdDateTime": full_event.get("createdDateTime", ""),
                        "lastModifiedDateTime": full_event.get(
                            "lastModifiedDateTime", ""
                        ),
                        "meetingInfo": meeting_info,
                    }

                    return self._format_response(readable_event)
                else:
                    return self._format_error(
                        f"Event #{event_number} does not have a valid ID"
                    )

        available_numbers = [e.get("number") for e in cached_events]
        return self._format_error(
            f"Event #{event_number} not found in cache. Available numbers: {available_numbers}"
        )

    async def handle_search_events(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_events tool. This operation clears and reloads the cache."""
        query = arguments.get("query")
        search_type = arguments.get("search_type", "organizer")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        date_range = arguments.get("time_range")

        user_timezone = await graph_client.get_user_timezone()
        today_date = DateHandler.get_today_date(user_timezone)

        if date_range:
            display_range, start_date_local, end_date_local = DateHandler.parse_date_range(
                date_range, user_timezone
            )
            start_date_display = DateHandler.format_date_with_weekday(
                start_date_local, user_timezone
            )
            end_date_display = DateHandler.format_date_with_weekday(
                end_date_local, user_timezone
            )
            # Convert local timezone dates to UTC for API call
            # parse_date_range returns timezone-aware ISO strings, parse and convert to UTC
            from zoneinfo import ZoneInfo

            # Normalize ISO strings for Python compatibility (handles 7-digit fractional seconds)
            start_dt = datetime.fromisoformat(DateHandler.normalize_iso_datetime(start_date_local))
            end_dt = datetime.fromisoformat(DateHandler.normalize_iso_datetime(end_date_local))
            start_date = start_dt.astimezone(ZoneInfo("UTC")).isoformat().replace("+00:00", "Z")
            end_date = end_dt.astimezone(ZoneInfo("UTC")).isoformat().replace("+00:00", "Z")
        else:
            if start_date:
                start_date = DateHandler.parse_local_date_to_utc(
                    start_date, user_timezone
                )
                start_date_display = DateHandler.format_date_with_weekday(
                    start_date, user_timezone
                )
            if end_date:
                end_date = DateHandler.parse_local_date_to_utc(end_date, user_timezone)
                end_date_display = DateHandler.format_date_with_weekday(
                    end_date, user_timezone
                )

            if start_date and end_date and start_date == end_date:
                from datetime import timedelta

                end_dt = datetime.fromisoformat(DateHandler.normalize_iso_datetime(end_date))
                end_dt = end_dt + timedelta(days=1)
                end_date = end_dt.isoformat().replace("+00:00", "Z")

        result = await graph_client.search_events(
            query, search_type, start_date, end_date, MAX_EVENT_SEARCH_LIMIT
        )

        # Clear cache and reload with search results - this is intentional for search operations
        await event_cache.set_mode("search")
        await event_cache.update_search_state(
            query=query,
            start_date=start_date,
            end_date=end_date,
            top=MAX_EVENT_SEARCH_LIMIT,
            total_count=result["count"],
            metadata=result.get("events", []),
        )

        response_data = {
            "query": query,
            "date_range": display_range if date_range else None,
            "count": result["count"],
            "timezone": user_timezone,
            "today": today_date,
            "hint": f"Found {result['count']} events. Use browse_events to view the results.",
        }

        if date_range or start_date:
            response_data["start_date"] = start_date_display
        if date_range or end_date:
            response_data["end_date"] = end_date_display

        return self._format_response(response_data)

    async def handle_respond_to_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_event_as_attendee tool for managing events as attendee: accept, decline, tentatively_accept, propose_new_time, delete_cancelled, email_attendees."""
        action = arguments["action"]

        if action == "accept":
            return await self._handle_accept_event_action(arguments)
        elif action == "decline":
            return await self._handle_decline_event_action(arguments)
        elif action == "tentatively_accept":
            return await self._handle_tentatively_accept_event_action(arguments)
        elif action == "propose_new_time":
            return await self._handle_propose_new_time_action(arguments)
        elif action == "delete_cancelled":
            return await self._handle_delete_cancelled_event_action(arguments)
        elif action == "email_attendees":
            return await self._handle_email_attendees_as_attendee_action(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'accept', 'decline', 'tentatively_accept', 'propose_new_time', 'delete_cancelled', or 'email_attendees'."
            )

    async def handle_manage_my_event(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_event_as_organizer tool for managing user's own events: create, update, cancel, forward, email_attendees."""
        action = arguments["action"]

        if action == "create":
            return await self._handle_create_event_action(arguments)
        elif action == "update":
            return await self._handle_update_event_action(arguments)
        elif action == "cancel":
            return await self._handle_cancel_event_action(arguments)
        elif action == "forward":
            return await self._handle_forward_event_action(arguments)
        elif action == "email_attendees":
            return await self._handle_reply_event_action(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'create', 'update', 'cancel', 'forward', or 'email_attendees'."
            )

    async def _handle_create_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle create event action with calendar conflict detection."""
        user_timezone = arguments.get("timezone")
        if not user_timezone:
            user_timezone = await graph_client.get_user_timezone()

        start_local = arguments["start"]
        end_local = arguments["end"]

        # Convert local times to UTC for conflict checking
        start_utc = DateHandler.parse_local_date_to_utc(start_local, user_timezone)
        end_utc = DateHandler.parse_local_date_to_utc(end_local, user_timezone)

        # Check for calendar conflicts before creating the event
        conflict_result = await graph_client.check_calendar_conflict(
            start_date=start_utc,
            end_date=end_utc,
        )
        conflict_warning = None
        if conflict_result.get("has_conflict"):
            conflicting = conflict_result.get("conflicting_events", [])
            conflict_list = ", ".join([e.get("subject", "Unknown") for e in conflicting[:3]])
            if len(conflicting) > 3:
                conflict_list += f" (+{len(conflicting) - 3} more)"
            conflict_warning = f"⚠️ Calendar Conflict Warning: {conflict_result.get('count')} conflicting event(s) found: {conflict_list}"

        event_data = {
            "subject": arguments["subject"],
            "start": {"dateTime": start_local, "timeZone": user_timezone},
            "end": {"dateTime": end_local, "timeZone": user_timezone},
        }

        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": normalize_email_html(arguments["body"]),
            }

        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}

        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        is_online_meeting = arguments.get("isOnlineMeeting")
        online_meeting_provider = arguments.get("onlineMeetingProvider")

        teams_integration_warning = None
        custom_link_note = None

        if is_online_meeting and online_meeting_provider:
            if online_meeting_provider == "teamsForBusiness":
                has_teams = await graph_client.calendar_client.check_teams_integration()
                if not has_teams:
                    teams_integration_warning = "Note: Current user does not have Teams integration. Teams meeting link may not be created if the organizer also lacks Teams access."
            elif online_meeting_provider == "unknown":
                custom_link_note = "Note: For 'unknown' provider, make sure to include the custom join link in the event body."

            event_data["isOnlineMeeting"] = is_online_meeting
            event_data["onlineMeetingProvider"] = online_meeting_provider

        if "recurrence" in arguments:
            event_data["recurrence"] = arguments["recurrence"]

        result = await graph_client.create_event(event_data)

        start_utc = result.get("start", {}).get("dateTime", "")
        end_utc = result.get("end", {}).get("dateTime", "")

        start_local = ""
        end_local = ""

        if start_utc:
            start_local = DateHandler.convert_utc_to_user_timezone(
                start_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        if end_utc:
            end_local = DateHandler.convert_utc_to_user_timezone(
                end_utc, user_timezone, "%Y-%m-%d %H:%M"
            )

        key_info = {
            "subject": result.get("subject", ""),
            "start": start_local,
            "end": end_local,
            "timezone": user_timezone,
            "location": result.get("location", {}).get("displayName", ""),
            "isOnlineMeeting": result.get("isOnlineMeeting", False),
            "webLink": result.get("webLink", ""),
        }

        attendees = result.get("attendees", [])
        if attendees:
            key_info["attendees_count"] = len(attendees)
            attendee_emails = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                if email:
                    attendee_emails.append(email)
            if attendee_emails:
                key_info["attendees"] = attendee_emails

        online_meeting = result.get("onlineMeeting", {})
        if online_meeting:
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                key_info["joinUrl"] = join_url

        if "recurrence" in arguments:
            key_info["recurrence"] = result.get("recurrence", {})

        await event_cache.add_event_to_cache(
            {
                "id": result.get("id", ""),
                "subject": result.get("subject", ""),
                "start_datetime": start_utc,
                "end_datetime": end_utc,
                "start": start_utc,
                "end": end_utc,
                "location": result.get("location", {}).get("displayName", ""),
                "isOnlineMeeting": result.get("isOnlineMeeting", False),
                "webLink": result.get("webLink", ""),
                "recurrence": "recurrence" in arguments,
            }
        )

        cached_events = event_cache.get_cached_events()
        cache_number = None
        for event in cached_events:
            if event.get("id") == result.get("id", ""):
                cache_number = event.get("number")
                break

        if cache_number:
            key_info["cache_number"] = cache_number

        response_message = f"Event created successfully and stored in cache as number {cache_number}.\n\n{json.dumps(key_info, indent=2, ensure_ascii=False)}"
        if conflict_warning:
            response_message = f"{conflict_warning}\n\n{response_message}"
        if teams_integration_warning:
            response_message = f"{teams_integration_warning}\n\n{response_message}"
        if custom_link_note:
            response_message = f"{custom_link_note}\n\n{response_message}"

        return self._format_response(response_message)

    async def _handle_update_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle update event action using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        cache_number = int(cache_number_str)
        event_info = await self._resolve_event_id(cache_number)

        user_timezone = arguments.get("timezone")
        if not user_timezone:
            user_timezone = await graph_client.get_user_timezone()

        event_data = {}
        conflict_warning = None
        
        # Check for conflicts if time is being changed
        if "start" in arguments or "end" in arguments:
            # Get current event to have complete time info
            current_event = await graph_client.get_event(event_info["event_id"])
            current_start = current_event.get("start", {}).get("dateTime", "")
            current_end = current_event.get("end", {}).get("dateTime", "")
            
            # Use new times if provided, otherwise use current times
            check_start = arguments.get("start")
            check_end = arguments.get("end")
            
            if check_start or check_end:
                start_utc = DateHandler.parse_local_date_to_utc(
                    check_start, user_timezone
                ) if check_start else current_start
                end_utc = DateHandler.parse_local_date_to_utc(
                    check_end, user_timezone
                ) if check_end else current_end
                
                # Check for conflicts, excluding the current event
                conflict_result = await graph_client.check_calendar_conflict(
                    start_date=start_utc,
                    end_date=end_utc,
                    exclude_event_id=event_info["event_id"],
                )
                if conflict_result.get("has_conflict"):
                    conflicting = conflict_result.get("conflicting_events", [])
                    conflict_list = ", ".join([e.get("subject", "Unknown") for e in conflicting[:3]])
                    if len(conflicting) > 3:
                        conflict_list += f" (+{len(conflicting) - 3} more)"
                    conflict_warning = f"⚠️ Calendar Conflict Warning: {conflict_result.get('count')} conflicting event(s) found: {conflict_list}"
        
        if "subject" in arguments:
            event_data["subject"] = arguments["subject"]
        if "start" in arguments:
            start_utc = DateHandler.parse_local_date_to_utc(
                arguments["start"], user_timezone
            )
            event_data["start"] = {"dateTime": start_utc, "timeZone": "UTC"}
        if "end" in arguments:
            end_utc = DateHandler.parse_local_date_to_utc(
                arguments["end"], user_timezone
            )
            event_data["end"] = {"dateTime": end_utc, "timeZone": "UTC"}
        if "body" in arguments:
            event_data["body"] = {
                "contentType": arguments.get("body_content_type", "HTML"),
                "content": normalize_email_html(arguments["body"]),
            }
        if "location" in arguments:
            event_data["location"] = {"displayName": arguments["location"]}
        if "attendees" in arguments:
            event_data["attendees"] = [
                {"emailAddress": {"address": email}, "type": "required"}
                for email in arguments["attendees"]
            ]

        if "optional_attendees" in arguments:
            if "attendees" not in event_data:
                event_data["attendees"] = []
            event_data["attendees"].extend(
                [
                    {"emailAddress": {"address": email}, "type": "optional"}
                    for email in arguments["optional_attendees"]
                ]
            )

        is_online_meeting = arguments.get("isOnlineMeeting")
        online_meeting_provider = arguments.get("onlineMeetingProvider")

        teams_integration_warning = None
        custom_link_note = None

        if is_online_meeting is not None or online_meeting_provider is not None:
            if online_meeting_provider == "teamsForBusiness":
                has_teams = await graph_client.calendar_client.check_teams_integration()
                if not has_teams:
                    teams_integration_warning = "Note: Current user does not have Teams integration. Teams meeting link may not be created if the organizer also lacks Teams access."
            elif online_meeting_provider == "unknown":
                custom_link_note = "Note: For 'unknown' provider, make sure to include the custom join link in the event body."

            if is_online_meeting is not None:
                event_data["isOnlineMeeting"] = is_online_meeting
            if online_meeting_provider is not None:
                event_data["onlineMeetingProvider"] = online_meeting_provider

        if "recurrence" in arguments:
            event_data["recurrence"] = arguments["recurrence"]

        result = await graph_client.update_event(event_info["event_id"], event_data)

        start_utc = result.get("start", {}).get("dateTime", "")
        end_utc = result.get("end", {}).get("dateTime", "")

        start_local = ""
        end_local = ""

        if start_utc:
            start_local = DateHandler.convert_utc_to_user_timezone(
                start_utc, user_timezone, "%Y-%m-%d %H:%M"
            )
        if end_utc:
            end_local = DateHandler.convert_utc_to_user_timezone(
                end_utc, user_timezone, "%Y-%m-%d %H:%M"
            )

        key_info = {
            "subject": result.get("subject", ""),
            "start": start_local,
            "end": end_local,
            "timezone": user_timezone,
            "location": result.get("location", {}).get("displayName", ""),
            "isOnlineMeeting": result.get("isOnlineMeeting", False),
            "webLink": result.get("webLink", ""),
        }

        attendees = result.get("attendees", [])
        if attendees:
            key_info["attendees_count"] = len(attendees)
            attendee_emails = []
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                if email:
                    attendee_emails.append(email)
            if attendee_emails:
                key_info["attendees"] = attendee_emails

        online_meeting = result.get("onlineMeeting", {})
        if online_meeting:
            join_url = online_meeting.get("joinWebUrl") or online_meeting.get("joinUrl")
            if join_url:
                key_info["joinUrl"] = join_url

        if "recurrence" in result:
            key_info["recurrence"] = result.get("recurrence", {})

        response_message = f"Event updated successfully: {json.dumps(key_info, indent=2, ensure_ascii=False)}"
        if conflict_warning:
            response_message = f"{conflict_warning}\n\n{response_message}"
        if teams_integration_warning:
            response_message = f"{teams_integration_warning}\n\n{response_message}"
        if custom_link_note:
            response_message = f"{custom_link_note}\n\n{response_message}"

        await event_cache.update_event_in_cache(
            event_info["event_id"],
            {
                "id": result.get("id", ""),
                "subject": result.get("subject", ""),
                "start_datetime": start_utc,
                "end_datetime": end_utc,
                "start": start_local,
                "end": end_local,
                "location": result.get("location", {}).get("displayName", ""),
                "isOnlineMeeting": result.get("isOnlineMeeting", False),
                "webLink": result.get("webLink", ""),
                "recurrence": "recurrence" in result,
            },
        )

        return self._format_response(response_message)

    async def _handle_cancel_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle cancel event action using cache number."""
        cache_number_param = arguments["cache_number"]
        comment = arguments.get("comment")

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        cache_number = int(cache_number_str)
        event_info = await self._resolve_event_id(cache_number)

        await graph_client.cancel_event(event_info["event_id"], comment)
        await event_cache.remove_event_from_cache(event_info["event_id"])
        return self._format_response(
            f"Event cancelled successfully. Cancellation notifications sent to attendees."
        )

    async def _resolve_event_id(self, event_number) -> dict:
        """Resolve cache event number to actual event ID and series information.

        Args:
            event_number: The cache number from browse_events (can be int or string)

        Returns:
            Dictionary with event_id, series_master_id, and is_recurring

        Raises:
            ValueError: If event number not found in cache
        """
        # Convert to int if it's a string
        if isinstance(event_number, str):
            if not event_number.isdigit():
                raise ValueError(f"Invalid event number format: '{event_number}'")
            event_number = int(event_number)

        cached_events = event_cache.get_cached_events()
        for event in cached_events:
            if event.get("number") == event_number:
                return {
                    "event_id": event.get("id"),
                    "series_master_id": event.get("seriesMasterId"),
                    "is_recurring": event.get("recurrence", False),
                }
        available_numbers = [e.get("number") for e in cached_events]
        raise ValueError(
            f"Event number {event_number} not found in cache. "
            f"Available numbers: {available_numbers}. "
            f"Please search for events first using search_events to find the correct cache number."
        )

    async def _handle_forward_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle forward event action (add optional attendees) using cache number."""
        cache_number_param = arguments["cache_number"]
        attendees = arguments["attendees"]
        comment = arguments.get("comment")

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        cache_number = int(cache_number_str)
        event_info = await self._resolve_event_id(cache_number)

        attendee_list = []
        for attendee in attendees:
            if isinstance(attendee, str):
                attendee_list.append({"address": attendee})
            elif isinstance(attendee, dict):
                attendee_list.append(attendee)

        await graph_client.forward_event(event_info["event_id"], attendee_list, comment)
        return self._format_response(
            f"Event forwarded successfully to {len(attendee_list)} attendee(s)."
        )

    async def _handle_reply_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle email attendees action for organizer (send email to attendees using event body as email content)."""
        return await self._send_email_to_event_attendees(
            arguments=arguments,
            include_organizer=False,
            success_message="Email sent successfully to event attendees"
        )

    async def _handle_email_attendees_as_attendee_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle email attendees action for attendee (send email to other attendees and organizer)."""
        return await self._send_email_to_event_attendees(
            arguments=arguments,
            include_organizer=True,
            success_message="Email sent successfully to event attendees and organizer"
        )

    async def _send_email_to_event_attendees(
        self,
        arguments: dict,
        include_organizer: bool,
        success_message: str
    ) -> list[types.TextContent]:
        """Shared method to send email to event attendees.

        Args:
            arguments: Tool arguments containing cache_number, email_subject, body, to, cc
            include_organizer: Whether to include the event organizer in recipients
            success_message: Message to return on success
        """
        cache_number_param = arguments["cache_number"]
        subject = arguments.get("email_subject", "Re: Event")
        body = normalize_email_html(arguments.get("body"))
        to_recipients = arguments.get("to")
        cc_recipients = arguments.get("cc")

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or search_events. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or search_events. "
                f"Use search_events and browse_events to find the cache number."
            )

        cache_number = int(cache_number_str)
        event_info = await self._resolve_event_id(cache_number)

        full_event = await graph_client.get_event(event_info["event_id"])

        if not body:
            event_body = full_event.get("body", {})
            body = event_body.get("content", "")
            body_type = event_body.get("contentType", "Text")
            if body_type.lower() == "html" and body:
                body = self._extract_text_from_html(body)

        if not to_recipients:
            attendees = full_event.get("attendees", [])
            to_recipients = []
            cc_recipients = []
            # Get user's email to filter out from recipients
            user_email = await graph_client.email_client.get_user_email()
            for attendee in attendees:
                email = attendee.get("emailAddress", {}).get("address", "")
                # Skip user's own email
                if email == user_email:
                    continue
                attendee_type = attendee.get("type", "required")
                if attendee_type == "required":
                    to_recipients.append(email)
                else:
                    cc_recipients.append(email)

            # For attendees: include the organizer in TO (they are not in attendees list)
            if include_organizer:
                organizer = full_event.get("organizer", {})
                organizer_email = organizer.get("emailAddress", {}).get("address", "")
                if organizer_email and organizer_email != user_email and organizer_email not in to_recipients:
                    to_recipients.insert(0, organizer_email)

        if not to_recipients:
            return self._format_error(
                "No recipients found. Please provide 'to' recipients or ensure the event has other attendees."
            )

        result = await graph_client.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            body_content_type="Text",
        )
        return self._format_response(f"{success_message}: {result}")

    async def _handle_accept_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle accept event action using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        event_number = int(cache_number_str)
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        try:
            if series and event_info["series_master_id"]:
                await graph_client.accept_event(
                    event_info["series_master_id"], comment, send_response, series
                )
                return self._format_response(
                    f"Entire recurring series accepted successfully."
                )

            await graph_client.accept_event(
                event_info["event_id"], comment, send_response, series
            )
            return self._format_response(f"Event accepted successfully.")
        except Exception as e:
            error_msg = str(e)
            if "organizer hasn't requested a response" in error_msg:
                # Organizer didn't request responses, so we update the event to mark it as busy
                # This uses the PATCH endpoint to set showAs property (same as Outlook behavior)
                try:
                    # Update the event to mark as busy with reminders
                    event_data = {
                        "showAs": "busy",
                        "isReminderOn": True
                    }
                    
                    await graph_client.update_event(event_info["event_id"], event_data)
                    logger.info(f"Updated event {event_info['event_id']} to show as busy")
                    
                    return self._format_response(
                        f"Event marked as busy on your calendar. The organizer didn't request a response, so the event was updated to show as busy with reminders."
                    )
                except Exception as update_error:
                    logger.error(f"Failed to update event to busy: {update_error}")
                    return self._format_response(
                        f"Event is on your calendar. The organizer didn't request a response, so no acceptance was sent. Note: Unable to update event to show as busy due to: {str(update_error)}"
                    )
            raise

    async def _handle_decline_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle decline event action using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        event_number = int(cache_number_str)
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        try:
            if series and event_info["series_master_id"]:
                await graph_client.decline_event(
                    event_info["series_master_id"], comment, send_response, series
                )
                return self._format_response(
                    f"Entire recurring series declined successfully."
                )

            await graph_client.decline_event(
                event_info["event_id"], comment, send_response, series
            )
            return self._format_response(f"Event declined successfully.")
        except Exception as e:
            error_msg = str(e)
            if "organizer hasn't requested a response" in error_msg:
                # Cannot formally decline when responses aren't requested
                # Delete the event from the calendar since the user declined it
                try:
                    await graph_client.delete_event(event_info["event_id"])
                    logger.info(f"Deleted event {event_info['event_id']} from calendar (declined without response requested)")
                    
                    return self._format_response(
                        f"Event deleted from your calendar. The organizer didn't request a response, so the event was deleted since you declined it."
                    )
                except Exception as delete_error:
                    logger.error(f"Failed to delete event after decline: {delete_error}")
                    return self._format_response(
                        f"Event is on your calendar. The organizer didn't request a response, so it cannot be formally declined. Note: Unable to delete event due to: {str(delete_error)}"
                    )
            raise

    async def _handle_tentatively_accept_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle tentatively accept event action using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        event_number = int(cache_number_str)
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        series = arguments.get("series", False)
        event_info = await self._resolve_event_id(event_number)

        try:
            if series and event_info["series_master_id"]:
                await graph_client.tentatively_accept_event(
                    event_info["series_master_id"], comment, send_response, series
                )
                return self._format_response(
                    f"Entire recurring series tentatively accepted successfully."
                )

            await graph_client.tentatively_accept_event(
                event_info["event_id"], comment, send_response, series
            )
            return self._format_response(f"Event tentatively accepted successfully.")
        except Exception as e:
            error_msg = str(e)
            if "organizer hasn't requested a response" in error_msg:
                # Organizer didn't request responses, so we update the event to mark it as tentative
                # This uses the PATCH endpoint to set showAs property (same as Outlook behavior)
                try:
                    # Update the event to mark as tentative with reminders
                    event_data = {
                        "showAs": "tentative",
                        "isReminderOn": True
                    }
                    
                    await graph_client.update_event(event_info["event_id"], event_data)
                    logger.info(f"Updated event {event_info['event_id']} to show as tentative")
                    
                    return self._format_response(
                        f"Event marked as tentative on your calendar. The organizer didn't request a response, so the event was updated to show as tentative with reminders."
                    )
                except Exception as update_error:
                    logger.error(f"Failed to update event to tentative: {update_error}")
                    return self._format_response(
                        f"Event is on your calendar. The organizer didn't request a response, so no tentative acceptance was sent. Note: Unable to update event to show as tentative due to: {str(update_error)}"
                    )
            raise

    async def _handle_propose_new_time_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle propose new time action (decline event and propose new time) using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        event_number = int(cache_number_str)
        propose_new_time = arguments["propose_new_time"]
        comment = arguments.get("comment")
        send_response = arguments.get("send_response", True)
        event_info = await self._resolve_event_id(event_number)

        user_timezone = await graph_client.get_user_timezone()

        proposed_time_local = propose_new_time.get("dateTime")
        proposed_time_utc = DateHandler.parse_local_date_to_utc(
            proposed_time_local, user_timezone
        )

        from datetime import timedelta

        proposed_end_time_utc = (
            datetime.fromisoformat(DateHandler.normalize_iso_datetime(proposed_time_utc))
            + timedelta(hours=1)
        ).strftime("%Y-%m-%dT%H:%M:%S")

        propose_new_time_data = {
            "start": {"dateTime": proposed_time_utc, "timeZone": "UTC"},
            "end": {"dateTime": proposed_end_time_utc, "timeZone": "UTC"},
        }

        try:
            await graph_client.propose_new_time(
                event_info["event_id"], propose_new_time_data, comment, send_response
            )
            return self._format_response(
                f"Event declined successfully with proposed new time: {proposed_time_local} ({user_timezone})."
            )
        except Exception as e:
            error_msg = str(e)
            if "organizer hasn't requested a response" in error_msg:
                # Cannot propose new time when responses aren't requested
                # Delete the event and tell user to contact organizer directly
                try:
                    await graph_client.delete_event(event_info["event_id"])
                    logger.info(f"Deleted event {event_info['event_id']} from calendar (propose new time without response requested)")
                    
                    return self._format_response(
                        f"Event deleted from your calendar. The organizer didn't request a response, so you cannot propose a new time through the system. Please contact the organizer directly to suggest: {proposed_time_local} ({user_timezone})."
                    )
                except Exception as delete_error:
                    logger.error(f"Failed to delete event after propose new time: {delete_error}")
                    return self._format_response(
                        f"Event is on your calendar. The organizer didn't request a response, so you cannot propose a new time. Please contact the organizer directly to suggest: {proposed_time_local} ({user_timezone}). Note: Unable to delete event due to: {str(delete_error)}"
                    )
            raise

    async def _handle_delete_cancelled_event_action(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete cancelled event action using cache number."""
        cache_number_param = arguments["cache_number"]

        # Handle both string and integer cache_number inputs
        if cache_number_param is None:
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        # Convert to string to check if it's numeric
        cache_number_str = str(cache_number_param)
        if not cache_number_str.isdigit():
            return self._format_error(
                f"Invalid cache number format: '{cache_number_param}'. "
                f"Please use the cache number (e.g., '1', '2', '3') from browse_events or returned when creating an event. "
                f"Use search_events and browse_events to find the cache number."
            )

        event_number = int(cache_number_str)
        event_info = await self._resolve_event_id(event_number)

        await graph_client.delete_event(event_info["event_id"])
        await event_cache.remove_event_from_cache(event_info["event_id"])
        return self._format_response(
            f"Cancelled event deleted from your calendar."
        )

    async def handle_check_attendee_availability(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle check_attendee_availability tool."""
        mandatory_attendees = arguments["attendees"]
        optional_attendees = arguments.get("optional_attendees", [])
        date = arguments["date"]
        availability_view_interval = arguments.get("availability_view_interval", 30)
        time_zone = arguments.get("time_zone")
        top_slots = arguments.get("top_slots", 5)
        meeting_duration = arguments.get("meeting_duration", 30)

        if time_zone:
            timezone_str = time_zone
        else:
            timezone_str = await graph_client.get_user_timezone()

        # Get user email with caching to avoid 429 rate limit errors
        user_email = await graph_client.get_user_email()
        logger.debug(f"[DEBUG] Organizer email retrieved: {user_email}")

        schedules = mandatory_attendees + optional_attendees
        if user_email and user_email not in schedules:
            schedules = [user_email] + schedules
        logger.debug(f"[DEBUG] Final schedules list: {schedules}")

        from datetime import datetime, timedelta
        from zoneinfo import ZoneInfo

        date_obj = datetime.strptime(date, "%Y-%m-%d").date()

        result = await graph_client.check_availability(
            schedules, None, None, availability_view_interval, date
        )

        availability_data = result.get("value", [])
        logger.debug(f"[DEBUG] availability_data count: {len(availability_data)}, scheduleIds: {[a.get('scheduleId') for a in availability_data]}")

        # Initialize JSON response structure
        json_response = {
            "date": date,
            "interval_minutes": availability_view_interval,
            "timezone": timezone_str,
            "attendees": [],
            "summary": {"top_time_slots": [], "total_attendees": 0},
        }

        all_attendee_availability = []
        attendee_names = []

        for attendee_info in availability_data:
            schedule_id = attendee_info.get("scheduleId", "Unknown")
            availability_view = attendee_info.get("availabilityView", "")
            schedule_items = attendee_info.get("scheduleItems", [])
            working_hours = attendee_info.get("workingHours", {})

            # Debug: Log raw working_hours from getSchedule API
            logger.debug(
                f"[DEBUG] {schedule_id} - getSchedule working_hours: {working_hours}"
            )

            all_attendee_availability.append(
                {
                    "schedule_id": schedule_id,
                    "availability_view": availability_view,
                    "working_hours": working_hours,
                }
            )
            attendee_names.append(schedule_id)

            attendee_timezone = None
            attendee_timezone_found = False

            if working_hours:
                working_hours_tz = working_hours.get("timeZone", {})
                if working_hours_tz:
                    attendee_timezone = working_hours_tz.get("name")
                    if attendee_timezone:
                        attendee_timezone_found = True

            if not attendee_timezone_found and schedule_items:
                for item in schedule_items:
                    item_start = item.get("start", {})
                    item_tz = item_start.get("timeZone")
                    if item_tz and item_tz != "UTC":
                        attendee_timezone = item_tz
                        attendee_timezone_found = True
                        break

            if not attendee_timezone_found:
                try:
                    attendee_timezone = await graph_client.get_user_timezone_by_email(
                        schedule_id
                    )
                    if attendee_timezone:
                        attendee_timezone_found = True
                except Exception as e:
                    pass

            # Get working hours from mailbox settings
            # This may override the working_hours from getSchedule API
            # Try multiple methods to get mailbox settings
            mailbox_settings = None
            
            # Method 1: Direct mailboxSettings endpoint
            try:
                mailbox_settings = await graph_client.calendar_client.get_mailbox_settings(schedule_id)
                if "workingHours" in mailbox_settings:
                    working_hours = mailbox_settings["workingHours"]
                    logger.debug(
                        f"[DEBUG] {schedule_id} - mailboxSettings (direct) working_hours: {working_hours}"
                    )
            except Exception as e:
                logger.debug(
                    f"[DEBUG] {schedule_id} - Direct mailboxSettings failed: {e}"
                )
            
            # Method 2: Try via user search (alternative permission path)
            if not mailbox_settings or "workingHours" not in mailbox_settings:
                try:
                    user_info = await graph_client.user_client.get_user_by_email(schedule_id)
                    if user_info and "mailboxSettings" in user_info:
                        mailbox_settings = user_info["mailboxSettings"]
                        if "workingHours" in mailbox_settings:
                            working_hours = mailbox_settings["workingHours"]
                            logger.debug(
                                f"[DEBUG] {schedule_id} - mailboxSettings (via user search) working_hours: {working_hours}"
                            )
                except Exception as e:
                    logger.debug(
                        f"[DEBUG] {schedule_id} - User search mailboxSettings failed: {e}"
                    )

            attendee_type = (
                "Organizer"
                if schedule_id == user_email
                else ("Optional" if schedule_id in optional_attendees else "Mandatory")
            )

            # Create attendee JSON object
            attendee_data = {
                "email": schedule_id,
                "type": attendee_type,
                "working_hours": None,
                "free_time_slots": [],
                "scheduled_items": [],
                "timezone": None,
            }

            if availability_view:
                from datetime import datetime, timedelta
                from zoneinfo import ZoneInfo

                if attendee_timezone_found and attendee_timezone and working_hours:
                    try:
                        attendee_tz = ZoneInfo(
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                        )
                        user_tz = ZoneInfo(timezone_str)

                        same_timezone = (
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                            == timezone_str
                        )

                        today = date_obj

                        # Check if today is a working day
                        days_of_week = working_hours.get("daysOfWeek", [])
                        weekday_map = {
                            "monday": 0,
                            "tuesday": 1,
                            "wednesday": 2,
                            "thursday": 3,
                            "friday": 4,
                            "saturday": 5,
                            "sunday": 6,
                        }
                        today_weekday = today.weekday()
                        is_working_day = True  # Default to True if daysOfWeek not specified
                        if days_of_week:
                            is_working_day = any(
                                weekday_map.get(day.lower()) == today_weekday
                                for day in days_of_week
                            )

                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")

                        # Debug: Log raw working hours data
                        logger.info(
                            f"[DEBUG] {schedule_id} - working_hours raw: startTime={working_start}, "
                            f"endTime={working_end}, daysOfWeek={days_of_week}, "
                            f"timeZone={attendee_timezone}, is_working_day={is_working_day}"
                        )

                        if working_start and working_end and is_working_day:
                            try:
                                working_start_clean = working_start.split(".")[0]
                                working_end_clean = working_end.split(".")[0]

                                working_start_time = datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time()
                                working_end_time = datetime.strptime(
                                    working_end_clean, "%H:%M:%S"
                                ).time()

                                working_start_dt = datetime.combine(
                                    today, working_start_time, tzinfo=attendee_tz
                                )
                                working_end_dt = datetime.combine(
                                    today, working_end_time, tzinfo=attendee_tz
                                )

                                working_start_user = working_start_dt.astimezone(
                                    user_tz
                                )
                                working_end_user = working_end_dt.astimezone(user_tz)

                                if same_timezone:
                                    attendee_data["working_hours"] = {
                                        "start": working_start_dt.strftime("%H:%M"),
                                        "end": working_end_dt.strftime("%H:%M"),
                                        "timezone": timezone_str,
                                        "days_of_week": days_of_week if days_of_week else None,
                                        "raw": {
                                            "startTime": working_start,
                                            "endTime": working_end,
                                            "timeZone": attendee_timezone,
                                        },
                                    }
                                else:
                                    attendee_data["working_hours"] = {
                                        "start": working_start_dt.strftime("%H:%M"),
                                        "end": working_end_dt.strftime("%H:%M"),
                                        "timezone": attendee_timezone,
                                        "days_of_week": days_of_week if days_of_week else None,
                                        "organizer_timezone": {
                                            "start": working_start_user.strftime(
                                                "%H:%M"
                                            ),
                                            "end": working_end_user.strftime("%H:%M"),
                                            "timezone": timezone_str,
                                        },
                                        "raw": {
                                            "startTime": working_start,
                                            "endTime": working_end,
                                            "timeZone": attendee_timezone,
                                        },
                                    }
                            except Exception as e:
                                attendee_data["working_hours"] = None
                        else:
                            attendee_data["working_hours"] = None
                    except Exception as e:
                        attendee_data["working_hours"] = None
                else:
                    attendee_data["working_hours"] = None
                attendee_data["timezone"] = (
                    attendee_timezone if attendee_timezone_found else None
                )

                # Process free time slots
                if attendee_timezone_found and attendee_timezone and working_hours:
                    try:
                        attendee_tz = ZoneInfo(
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                        )
                        user_tz = ZoneInfo(timezone_str)

                        same_timezone = (
                            self._convert_microsoft_timezone_to_iana(attendee_timezone)
                            == timezone_str
                        )

                        # Check if today is a working day
                        days_of_week = working_hours.get("daysOfWeek", [])
                        weekday_map = {
                            "monday": 0,
                            "tuesday": 1,
                            "wednesday": 2,
                            "thursday": 3,
                            "friday": 4,
                            "saturday": 5,
                            "sunday": 6,
                        }
                        today_weekday = today.weekday()
                        is_working_day = True
                        if days_of_week:
                            is_working_day = any(
                                weekday_map.get(day.lower()) == today_weekday
                                for day in days_of_week
                            )

                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")

                        if working_start and working_end and is_working_day:
                            working_start_clean = working_start.split(".")[0]
                            working_end_clean = working_end.split(".")[0]

                            working_start_dt = datetime.combine(
                                today,
                                datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time(),
                                tzinfo=attendee_tz,
                            )
                            working_end_dt = datetime.combine(
                                today,
                                datetime.strptime(working_end_clean, "%H:%M:%S").time(),
                                tzinfo=attendee_tz,
                            )

                            utc_midnight = datetime.combine(
                                today,
                                datetime.strptime("00:00:00", "%H:%M:%S").time(),
                                tzinfo=ZoneInfo("UTC"),
                            )

                            # Convert working start time to UTC to get correct slot index
                            working_start_utc = working_start_dt.astimezone(ZoneInfo("UTC"))

                            # Calculate slot index based on UTC time
                            minutes_from_utc_midnight = int(
                                (working_start_utc - utc_midnight).total_seconds() / 60
                            )
                            start_slot_index = (
                                minutes_from_utc_midnight // availability_view_interval
                            )

                            if start_slot_index < 0:
                                start_slot_index = 0

                            in_free_slot = False
                            free_slot_start = 0

                            for i in range(start_slot_index, len(availability_view)):
                                status_code = (
                                    availability_view[i]
                                    if i < len(availability_view)
                                    else "?"
                                )
                                slot_start_utc = utc_midnight + timedelta(
                                    minutes=i * availability_view_interval
                                )
                                slot_end_utc = utc_midnight + timedelta(
                                    minutes=(i + 1) * availability_view_interval
                                )

                                slot_start_attendee = slot_start_utc.astimezone(
                                    attendee_tz
                                )

                                # Check if slot is past working hours (in attendee's timezone)
                                if slot_start_attendee >= working_end_dt:
                                    break

                                # Check if slot is within working hours
                                if slot_start_attendee < working_start_dt:
                                    continue

                                if status_code == "0":
                                    if not in_free_slot:
                                        free_slot_start = i
                                        in_free_slot = True
                                else:
                                    if in_free_slot:
                                        slot_start_utc_free = utc_midnight + timedelta(
                                            minutes=free_slot_start
                                            * availability_view_interval
                                        )
                                        slot_end_utc_free = utc_midnight + timedelta(
                                            minutes=i * availability_view_interval
                                        )
                                        slot_start_attendee_free = (
                                            slot_start_utc_free.astimezone(attendee_tz)
                                        )
                                        slot_end_attendee_free = (
                                            slot_end_utc_free.astimezone(attendee_tz)
                                        )
                                        slot_start_user_free = (
                                            slot_start_attendee_free.astimezone(user_tz)
                                        )
                                        slot_end_user_free = (
                                            slot_end_attendee_free.astimezone(user_tz)
                                        )
                                        if same_timezone:
                                            attendee_data["free_time_slots"].append(
                                                {
                                                    "start": slot_start_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "end": slot_end_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "timezone": timezone_str,
                                                }
                                            )
                                        else:
                                            attendee_data["free_time_slots"].append(
                                                {
                                                    "start": slot_start_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "end": slot_end_attendee_free.strftime(
                                                        "%Y-%m-%d %H:%M"
                                                    ),
                                                    "timezone": attendee_timezone,
                                                    "organizer_timezone": {
                                                        "start": slot_start_user_free.strftime(
                                                            "%Y-%m-%d %H:%M"
                                                        ),
                                                        "end": slot_end_user_free.strftime(
                                                            "%Y-%m-%d %H:%M"
                                                        ),
                                                        "timezone": timezone_str,
                                                    },
                                                }
                                            )
                                        in_free_slot = False

                            if in_free_slot:
                                slot_start_utc_free = utc_midnight + timedelta(
                                    minutes=free_slot_start * availability_view_interval
                                )
                                slot_end_utc_free = utc_midnight + timedelta(
                                    minutes=len(availability_view)
                                    * availability_view_interval
                                )
                                slot_start_attendee_free = (
                                    slot_start_utc_free.astimezone(attendee_tz)
                                )
                                slot_end_attendee_free = slot_end_utc_free.astimezone(
                                    attendee_tz
                                )
                                if slot_end_attendee_free > working_end_dt:
                                    slot_end_attendee_free = working_end_dt
                                slot_start_user_free = (
                                    slot_start_attendee_free.astimezone(user_tz)
                                )
                                slot_end_user_free = slot_end_attendee_free.astimezone(
                                    user_tz
                                )
                                if same_timezone:
                                    attendee_data["free_time_slots"].append(
                                        {
                                            "start": slot_start_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "end": slot_end_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "timezone": timezone_str,
                                        }
                                    )
                                else:
                                    attendee_data["free_time_slots"].append(
                                        {
                                            "start": slot_start_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "end": slot_end_attendee_free.strftime(
                                                "%Y-%m-%d %H:%M"
                                            ),
                                            "timezone": attendee_timezone,
                                            "organizer_timezone": {
                                                "start": slot_start_user_free.strftime(
                                                    "%Y-%m-%d %H:%M"
                                                ),
                                                "end": slot_end_user_free.strftime(
                                                    "%Y-%m-%d %H:%M"
                                                ),
                                                "timezone": timezone_str,
                                            },
                                        }
                                    )
                    except Exception as e:
                        pass

                # Add attendee to JSON response
                json_response["attendees"].append(attendee_data)

                non_free_items = [
                    item
                    for item in schedule_items
                    if item.get("status", "Unknown").lower() != "free"
                ]

                if non_free_items:
                    for item in non_free_items:
                        status = item.get("status", "Unknown")
                        item_start = item.get("start", {}).get("dateTime", "")
                        item_end = item.get("end", {}).get("dateTime", "")

                        try:
                            # Normalize ISO datetime for Python compatibility
                            item_start_dt = datetime.fromisoformat(
                                DateHandler.normalize_iso_datetime(item_start)
                            )
                            item_end_dt = datetime.fromisoformat(
                                DateHandler.normalize_iso_datetime(item_end)
                            )
                            # Ensure UTC timezone
                            if item_start_dt.tzinfo is None:
                                item_start_dt = item_start_dt.replace(tzinfo=ZoneInfo("UTC"))
                            if item_end_dt.tzinfo is None:
                                item_end_dt = item_end_dt.replace(tzinfo=ZoneInfo("UTC"))

                            if attendee_timezone_found and attendee_timezone:
                                attendee_tz = ZoneInfo(
                                    self._convert_microsoft_timezone_to_iana(
                                        attendee_timezone
                                    )
                                )
                                user_tz = ZoneInfo(timezone_str)

                                same_timezone = (
                                    self._convert_microsoft_timezone_to_iana(
                                        attendee_timezone
                                    )
                                    == timezone_str
                                )

                                item_start_attendee = item_start_dt.astimezone(
                                    attendee_tz
                                )
                                item_end_attendee = item_end_dt.astimezone(attendee_tz)
                                item_start_user = item_start_dt.astimezone(user_tz)
                                item_end_user = item_end_dt.astimezone(user_tz)

                                item_start_attendee_str = item_start_attendee.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_attendee_str = item_end_attendee.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_start_user_str = item_start_user.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_user_str = item_end_user.strftime(
                                    "%Y-%m-%d %H:%M"
                                )

                                if same_timezone:
                                    attendee_data["scheduled_items"].append(
                                        {
                                            "status": status,
                                            "start": item_start_attendee_str,
                                            "end": item_end_attendee_str,
                                            "timezone": timezone_str,
                                        }
                                    )
                                else:
                                    attendee_data["scheduled_items"].append(
                                        {
                                            "status": status,
                                            "start": item_start_attendee_str,
                                            "end": item_end_attendee_str,
                                            "timezone": attendee_timezone,
                                            "organizer_timezone": {
                                                "start": item_start_user_str,
                                                "end": item_end_user_str,
                                                "timezone": timezone_str,
                                            },
                                        }
                                    )
                            else:
                                item_start_dt = item_start_dt.astimezone(
                                    ZoneInfo(timezone_str)
                                )
                                item_end_dt = item_end_dt.astimezone(
                                    ZoneInfo(timezone_str)
                                )
                                item_start_str = item_start_dt.strftime(
                                    "%Y-%m-%d %H:%M"
                                )
                                item_end_str = item_end_dt.strftime("%Y-%m-%d %H:%M")
                                attendee_data["scheduled_items"].append(
                                    {
                                        "status": status,
                                        "start": item_start_str,
                                        "end": item_end_str,
                                        "timezone": timezone_str,
                                    }
                                )
                        except Exception as e:
                            pass

        # Update total attendees count
        json_response["summary"]["total_attendees"] = len(all_attendee_availability)

        if all_attendee_availability:
            try:
                from datetime import datetime, timedelta
                from zoneinfo import ZoneInfo
                from collections import defaultdict

                user_tz = ZoneInfo(timezone_str)
                utc_midnight = datetime.combine(
                    date_obj,
                    datetime.strptime("00:00:00", "%H:%M:%S").time(),
                    tzinfo=ZoneInfo("UTC"),
                )

                # Track free status for each slot by attendee type
                # Key: (slot_start, slot_end), Value: {"mandatory": set(), "optional": set(), "organizer": set()}
                slot_free_status = defaultdict(
                    lambda: {"mandatory": set(), "optional": set(), "organizer": set()}
                )
                # Track attendees outside working hours for each slot
                slot_outside_hours = defaultdict(
                    lambda: {"mandatory": set(), "optional": set(), "organizer": set()}
                )
                # Track tentative (not yet accepted) events for each slot
                # These are events that might conflict if the attendee accepts
                slot_tentative = defaultdict(
                    lambda: {"mandatory": set(), "optional": set(), "organizer": set()}
                )
                slot_unavailable = {}

                for attendee_data in all_attendee_availability:
                    schedule_id = attendee_data["schedule_id"]
                    availability_view = attendee_data["availability_view"]
                    working_hours = attendee_data["working_hours"]

                    # Determine attendee type
                    if schedule_id == user_email:
                        attendee_type = "organizer"
                    elif schedule_id in mandatory_attendees:
                        attendee_type = "mandatory"
                    else:
                        attendee_type = "optional"

                    attendee_timezone = None
                    attendee_timezone_found = False

                    if working_hours:
                        working_hours_tz = working_hours.get("timeZone", {})
                        if working_hours_tz:
                            attendee_timezone = working_hours_tz.get("name")
                            if attendee_timezone:
                                attendee_timezone_found = True

                    if not attendee_timezone_found:
                        attendee_timezone = timezone_str
                        attendee_timezone_found = True

                    attendee_tz = ZoneInfo(
                        self._convert_microsoft_timezone_to_iana(attendee_timezone)
                    )

                    working_start_dt = None
                    working_end_dt = None
                    is_working_day = True  # Default to True

                    if working_hours:
                        working_start = working_hours.get("startTime")
                        working_end = working_hours.get("endTime")
                        days_of_week = working_hours.get("daysOfWeek", [])

                        # Check if today is a working day
                        weekday_map = {
                            "monday": 0,
                            "tuesday": 1,
                            "wednesday": 2,
                            "thursday": 3,
                            "friday": 4,
                            "saturday": 5,
                            "sunday": 6,
                        }
                        today_weekday = date_obj.weekday()
                        if days_of_week:
                            is_working_day = any(
                                weekday_map.get(day.lower()) == today_weekday
                                for day in days_of_week
                            )

                        if working_start and working_end and is_working_day:
                            try:
                                working_start_clean = working_start.split(".")[0]
                                working_end_clean = working_end.split(".")[0]

                                working_start_time = datetime.strptime(
                                    working_start_clean, "%H:%M:%S"
                                ).time()
                                working_end_time = datetime.strptime(
                                    working_end_clean, "%H:%M:%S"
                                ).time()

                                working_start_dt = datetime.combine(
                                    date_obj, working_start_time, tzinfo=attendee_tz
                                )
                                working_end_dt = datetime.combine(
                                    date_obj, working_end_time, tzinfo=attendee_tz
                                )
                            except Exception as e:
                                pass

                    for i, status_code in enumerate(availability_view):
                        slot_start_utc = utc_midnight + timedelta(
                            minutes=i * availability_view_interval
                        )
                        slot_end_utc = utc_midnight + timedelta(
                            minutes=(i + 1) * availability_view_interval
                        )

                        slot_start_user = slot_start_utc.astimezone(user_tz)
                        slot_end_user = slot_end_utc.astimezone(user_tz)

                        slot_key = (slot_start_user, slot_end_user)

                        is_free = status_code == "0"
                        is_tentative = status_code == "1"  # Tentative = not yet accepted
                        is_outside_working_hours = False

                        # If not a working day, all slots are outside working hours
                        if not is_working_day:
                            is_free = False
                            is_tentative = False
                            is_outside_working_hours = True
                        elif working_start_dt and working_end_dt:
                            slot_start_attendee = slot_start_utc.astimezone(attendee_tz)
                            if (
                                slot_start_attendee < working_start_dt
                                or slot_start_attendee >= working_end_dt
                            ):
                                is_free = False
                                is_tentative = False
                                is_outside_working_hours = True
                                status_code = "4"

                        if is_free:
                            slot_free_status[slot_key][attendee_type].add(schedule_id)
                        elif is_tentative:
                            # Track tentative (not yet accepted) events
                            # These are "potentially busy" slots
                            slot_tentative[slot_key][attendee_type].add(schedule_id)
                            if slot_key not in slot_unavailable:
                                slot_unavailable[slot_key] = []
                            slot_unavailable[slot_key].append(
                                {
                                    "schedule_id": schedule_id,
                                    "status": "Tentative (not responded)",
                                    "type": attendee_type.capitalize(),
                                }
                            )
                        elif is_outside_working_hours:
                            # Track attendees outside working hours
                            slot_outside_hours[slot_key][attendee_type].add(schedule_id)
                            if slot_key not in slot_unavailable:
                                slot_unavailable[slot_key] = []
                            slot_unavailable[slot_key].append(
                                {
                                    "schedule_id": schedule_id,
                                    "status": "Outside working hours",
                                    "type": attendee_type.capitalize(),
                                }
                            )
                        else:
                            status_map = {
                                "2": "Busy",
                                "3": "Out of office",
                                "4": "Working elsewhere / Outside hours",
                                "?": "Unknown",
                            }
                            status_text = status_map.get(status_code, "Unknown")

                            if slot_key not in slot_unavailable:
                                slot_unavailable[slot_key] = []
                            slot_unavailable[slot_key].append(
                                {
                                    "schedule_id": schedule_id,
                                    "status": status_text,
                                    "type": attendee_type.capitalize(),
                                }
                            )

                # Find meeting slots that fit the duration and are within all attendees' working hours
                def find_meeting_slots(
                    slot_status,
                    slot_outside,
                    slot_tentative_info,
                    mandatory_attendee_list,
                    meeting_duration_min,
                    interval_min,
                ):
                    """Find all possible meeting slots that fit the duration.

                    Only returns slots where ALL mandatory attendees are within their working hours.
                    For optional attendees, being outside working hours just means they're unavailable.
                    Also tracks tentative (not yet accepted) events for scoring.
                    """
                    if not slot_status and not slot_outside:
                        return []

                    # Collect all unique slot keys from free, outside_hours, and tentative
                    all_slot_keys = set(slot_status.keys()) | set(slot_outside.keys()) | set(slot_tentative_info.keys())
                    sorted_slots = sorted(all_slot_keys, key=lambda x: x[0])

                    if not sorted_slots:
                        return []

                    # Build a unified status dict for all slots
                    # Each slot has free attendees, outside_hours attendees, and tentative attendees
                    all_slots_data = []
                    for slot_key in sorted_slots:
                        free_data = slot_status.get(slot_key, {
                            "mandatory": set(),
                            "optional": set(),
                            "organizer": set(),
                        })
                        outside_data = slot_outside.get(slot_key, {
                            "mandatory": set(),
                            "optional": set(),
                            "organizer": set(),
                        })
                        tentative_data = slot_tentative_info.get(slot_key, {
                            "mandatory": set(),
                            "optional": set(),
                            "organizer": set(),
                        })
                        all_slots_data.append({
                            "slot_key": slot_key,
                            "free": free_data,
                            "outside_hours": outside_data,
                            "tentative": tentative_data,
                        })

                    required_slots = max(1, meeting_duration_min // interval_min)
                    meeting_slots = []

                    # Slide through slots to find continuous windows that fit the duration
                    for i in range(len(all_slots_data)):
                        # Check if we have enough slots starting from position i
                        if i + required_slots > len(all_slots_data):
                            break

                        # Get the starting slot info
                        start_data = all_slots_data[i]
                        start_slot_key = start_data["slot_key"]

                        # Check continuity: all slots must be consecutive
                        is_continuous = True
                        current_end = start_slot_key[1]

                        for j in range(1, required_slots):
                            next_slot_key = all_slots_data[i + j]["slot_key"]
                            next_start = next_slot_key[0]

                            if next_start != current_end:
                                is_continuous = False
                                break
                            current_end = next_slot_key[1]

                        if not is_continuous:
                            continue

                        # Check if any mandatory attendee or organizer is outside working hours in ANY slot
                        has_mandatory_outside = False
                        for j in range(required_slots):
                            outside_mandatory = all_slots_data[i + j]["outside_hours"]["mandatory"]
                            outside_organizer = all_slots_data[i + j]["outside_hours"]["organizer"]
                            # Check if any mandatory attendee is outside working hours
                            if outside_mandatory & set(mandatory_attendee_list):
                                has_mandatory_outside = True
                                break
                            # Check if organizer is outside working hours
                            if outside_organizer:
                                has_mandatory_outside = True
                                break

                        if has_mandatory_outside:
                            continue

                        # Calculate the meeting slot boundaries
                        meeting_start = start_slot_key[0]
                        meeting_end = all_slots_data[i + required_slots - 1]["slot_key"][1]

                        # Calculate intersection of free attendees across all slots in the window
                        mandatory_intersection = set(start_data["free"]["mandatory"])
                        optional_intersection = set(start_data["free"]["optional"])
                        organizer_intersection = set(start_data["free"]["organizer"])

                        # Collect all tentative attendees (union across slots)
                        mandatory_tentative = set(start_data["tentative"]["mandatory"])
                        optional_tentative = set(start_data["tentative"]["optional"])
                        organizer_tentative = set(start_data["tentative"]["organizer"])

                        for j in range(1, required_slots):
                            next_free = all_slots_data[i + j]["free"]
                            mandatory_intersection &= next_free["mandatory"]
                            optional_intersection &= next_free["optional"]
                            organizer_intersection &= next_free["organizer"]

                            # Union for tentative (any tentative event in the window)
                            mandatory_tentative |= all_slots_data[i + j]["tentative"]["mandatory"]
                            optional_tentative |= all_slots_data[i + j]["tentative"]["optional"]
                            organizer_tentative |= all_slots_data[i + j]["tentative"]["organizer"]

                        meeting_slots.append({
                            "start": meeting_start,
                            "end": meeting_end,
                            "mandatory_free": mandatory_intersection,
                            "optional_free": optional_intersection,
                            "organizer_free": organizer_intersection,
                            "mandatory_tentative": mandatory_tentative,
                            "optional_tentative": optional_tentative,
                            "organizer_tentative": organizer_tentative,
                        })

                    return meeting_slots

                # Calculate meeting slots with scoring
                meeting_slots = find_meeting_slots(
                    slot_free_status,
                    slot_outside_hours,
                    slot_tentative,
                    mandatory_attendees,
                    meeting_duration,
                    availability_view_interval,
                )

                def score_slot(slot_info, mandatory_list, optional_list):
                    """Score a slot based on attendee availability.

                    Also considers tentative (not yet accepted) events:
                    - Tentative events create uncertainty - the attendee might accept them
                    - Slots with tentative events are ranked lower to reduce scheduling conflicts
                    """
                    mandatory_count = len(mandatory_list)
                    optional_count = len(optional_list)

                    mandatory_free_count = len(slot_info["mandatory_free"])
                    optional_free_count = len(slot_info["optional_free"])
                    organizer_free = len(slot_info["organizer_free"]) > 0

                    # Get tentative counts
                    mandatory_tentative_count = len(slot_info.get("mandatory_tentative", set()))
                    optional_tentative_count = len(slot_info.get("optional_tentative", set()))
                    organizer_tentative = len(slot_info.get("organizer_tentative", set())) > 0

                    # Check if all mandatory attendees are free
                    all_mandatory_free = mandatory_free_count == mandatory_count

                    # Score components (total max: 200):
                    # 1. Mandatory attendees score (0-100):
                    #    - 100 if all mandatory are free
                    #    - Otherwise: proportional score based on how many are free
                    #    - This gives partial credit when some mandatory are busy
                    if mandatory_count > 0:
                        if all_mandatory_free:
                            mandatory_score = 100
                        else:
                            # Partial score: e.g., 2/3 mandatory free = 60 points
                            mandatory_score = (mandatory_free_count / mandatory_count) * 80
                    else:
                        mandatory_score = 80  # No mandatory attendees, give good score

                    # 2. Organizer free bonus (0-40)
                    organizer_score = 40 if organizer_free else 0

                    # 3. Optional attendees score (0-30)
                    if optional_count > 0:
                        optional_score = (optional_free_count / optional_count) * 30
                    else:
                        optional_score = 30  # No optional attendees, give max score

                    # 4. All mandatory free bonus (0-30)
                    #    Extra bonus when ALL mandatory attendees are available
                    all_mandatory_bonus = 30 if all_mandatory_free else 0

                    # 5. Tentative events penalty (0 to -50)
                    #    Tentative events are "pending" meetings that might be accepted
                    #    This creates uncertainty for scheduling
                    tentative_penalty = 0

                    # Penalty for mandatory attendees with tentative events
                    if mandatory_count > 0:
                        tentative_penalty -= (mandatory_tentative_count / mandatory_count) * 20

                    # Penalty for optional attendees with tentative events (less severe)
                    if optional_count > 0:
                        tentative_penalty -= (optional_tentative_count / optional_count) * 10

                    # Penalty for organizer with tentative events
                    if organizer_tentative:
                        tentative_penalty -= 20

                    # Ensure penalty doesn't go below -50
                    tentative_penalty = max(tentative_penalty, -50)

                    total_score = mandatory_score + organizer_score + optional_score + all_mandatory_bonus + tentative_penalty

                    return {
                        "total_score": total_score,
                        "mandatory_score": mandatory_score,
                        "organizer_score": organizer_score,
                        "optional_score": optional_score,
                        "tentative_penalty": tentative_penalty,
                        "all_mandatory_free": all_mandatory_free,
                        "organizer_free": organizer_free,
                        "mandatory_free_count": mandatory_free_count,
                        "optional_free_count": optional_free_count,
                        "mandatory_tentative_count": mandatory_tentative_count,
                        "optional_tentative_count": optional_tentative_count,
                        "organizer_tentative": organizer_tentative,
                        "mandatory_total": mandatory_count,
                        "optional_total": optional_count,
                    }

                # Score and rank all slots
                scored_slots = []

                for slot_info in meeting_slots:
                    score_info = score_slot(
                        slot_info, mandatory_attendees, optional_attendees
                    )

                    scored_slots.append(
                        {
                            "slot_info": slot_info,
                            "score_info": score_info,
                        }
                    )

                # Sort by score (descending), then by start time (ascending for same score)
                scored_slots.sort(
                    key=lambda x: (-x["score_info"]["total_score"], x["slot_info"]["start"])
                )

                # Build final top slots list
                rank = 0
                seen_slots = set()

                for item in scored_slots:
                    if rank >= top_slots:
                        break

                    slot_info = item["slot_info"]
                    score_info = item["score_info"]

                    slot_start = slot_info["start"]
                    slot_end = slot_info["end"]

                    # Deduplicate same time slots
                    slot_key = (slot_start, slot_end)
                    if slot_key in seen_slots:
                        continue
                    seen_slots.add(slot_key)

                    rank += 1

                    # Calculate total free attendees
                    total_free = (
                        len(slot_info["mandatory_free"])
                        + len(slot_info["optional_free"])
                        + len(slot_info["organizer_free"])
                    )
                    total_attendees = len(all_attendee_availability)
                    percentage = (total_free / total_attendees) * 100 if total_attendees > 0 else 0

                    time_slot = {
                        "rank": rank,
                        "start_time": slot_start.strftime("%Y-%m-%d %H:%M"),
                        "end_time": slot_end.strftime("%Y-%m-%d %H:%M"),
                        "duration_minutes": meeting_duration,
                        "timezone": timezone_str,
                        "score": score_info["total_score"],
                        "all_mandatory_free": score_info["all_mandatory_free"],
                        "organizer_free": score_info["organizer_free"],
                        "free_attendees": total_free,
                        "mandatory_free_count": score_info["mandatory_free_count"],
                        "optional_free_count": score_info["optional_free_count"],
                        "tentative_attendees": {
                            "mandatory_count": score_info["mandatory_tentative_count"],
                            "optional_count": score_info["optional_tentative_count"],
                            "organizer_has_tentative": score_info["organizer_tentative"],
                        },
                        "tentative_penalty": score_info["tentative_penalty"],
                        "total_attendees": total_attendees,
                        "percentage_free": round(percentage, 1),
                        "unavailable_attendees": [],
                    }

                    # Add unavailable attendees info
                    # For each slot in this meeting duration, collect unavailable info
                    current_time = slot_start
                    while current_time < slot_end:
                        single_slot_key = (current_time, current_time + timedelta(minutes=availability_view_interval))
                        if single_slot_key in slot_unavailable:
                            for unavailable_info in slot_unavailable[single_slot_key]:
                                # Avoid duplicates
                                email = unavailable_info["schedule_id"]
                                if not any(
                                    u["email"] == email
                                    for u in time_slot["unavailable_attendees"]
                                ):
                                    time_slot["unavailable_attendees"].append(
                                        {
                                            "email": email,
                                            "status": unavailable_info["status"],
                                            "type": unavailable_info["type"],
                                        }
                                    )
                        current_time += timedelta(minutes=availability_view_interval)

                    json_response["summary"]["top_time_slots"].append(time_slot)

                # Add meeting duration to response
                json_response["summary"]["meeting_duration_minutes"] = meeting_duration

            except Exception as e:
                json_response["summary"]["error"] = f"Error generating summary: {e}"

        return self._format_response(json.dumps(json_response, indent=2))

"""Microsoft Graph API client module.

This module provides backward compatibility by re-exporting the graph_client
from the new modular clients package.
"""

import asyncio
from typing import Any, Dict, List, Optional

import httpx

from .auth import auth_manager
from .config import settings
from .clients import UserClient, EmailClient, CalendarClient, FileClient, TeamsClient


class GraphClient:
    """Client for Microsoft Graph API operations.

    This class provides backward compatibility by delegating to specialized clients.
    """

    def __init__(self):
        self.base_url = settings.graph_api_base_url
        self.timeout = 30.0
        self._semaphore = asyncio.Semaphore(10)

        self.user_client = UserClient()
        self.email_client = EmailClient()
        self.calendar_client = CalendarClient()
        self.file_client = FileClient()
        self.teams_client = TeamsClient()

    async def _make_request(
        self,
        method: str,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        data: Optional[Dict[str, Any]] = None,
        headers: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API with concurrency control."""

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

            async with httpx.AsyncClient(timeout=self.timeout) as client:
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
                else:
                    raise Exception(
                        f"Graph API request failed: {response.status_code} - {response.text}"
                    )

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

    # User management methods - delegated to UserClient
    async def get_me(self) -> Dict[str, Any]:
        """Get current user information."""
        return await self.get("/me")

    async def get_mailbox_settings(self) -> Dict[str, Any]:
        """Get user's mailbox settings including timezone."""
        try:
            params = {"$select": "mailboxSettings"}
            result = await self.get("/me", params=params)
            return result.get("mailboxSettings", {})
        except Exception:
            return {}

    async def get_user_email(self) -> Optional[str]:
        """Get the current user's email address."""
        return await self.user_client.get_user_email()

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier. Uses server local timezone."""
        return await self.user_client.get_user_timezone()

    async def get_users(
        self, filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get list of users in organization."""
        return await self.user_client.get_users(filter_query)

    async def get_user(self, user_id: str) -> Dict[str, Any]:
        """Get specific user by ID."""
        return await self.user_client.get_user(user_id)

    async def get_user_timezone_by_email(self, email: str) -> Optional[str]:
        """Get user's timezone by email address."""
        return await self.user_client.get_user_timezone_by_email(email)

    async def search_contacts(self, query: str, top: int = 10) -> List[Dict[str, Any]]:
        """Search contacts and people relevant to the user."""
        return await self.user_client.search_contacts(query, top)

    # Mail management methods - delegated to EmailClient
    async def get_messages(
        self, folder: str = "Inbox", top: int = 10, filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get messages from specified folder."""
        return await self.email_client.get_messages(folder, top, filter_query)

    async def list_mail_folders(self) -> List[Dict[str, Any]]:
        """List all mail folders with their paths (including all levels)."""
        return await self.email_client.list_mail_folders()

    async def load_emails_by_folder(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Load emails from a folder with optional days or top parameter (mutually exclusive)."""
        return await self.email_client.load_emails_by_folder(folder, days, top)

    async def browse_emails(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Browse emails from a folder with summary information."""
        return await self.email_client.browse_emails(folder, days, top)

    async def get_email_count(
        self, folder: str = "Inbox", filter_query: Optional[str] = None
    ) -> int:
        """Get email count for a folder."""
        return await self.email_client.get_email_count(folder, filter_query)

    async def get_email(
        self,
        email_id: str,
        emailNumber: int = 0,
        return_html: bool = False,
        download_attachments: bool = False,
        download_path: Optional[str] = None,
        attachment_names: Optional[List[str]] = None,
        multimodal_supported: bool = False,
    ) -> Dict[str, Any]:
        """Get email by ID."""
        return await self.email_client.get_email(
            email_id, emailNumber, return_html, download_attachments, download_path, attachment_names, multimodal_supported
        )

    async def search_emails(
        self,
        query: Optional[str] = None,
        search_type: Optional[str] = None,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        folder: str = "Inbox",
        top: int = 10,
        inference_classification: str = "focused",
    ) -> Dict[str, Any]:
        """Search or list emails by keywords, sender, recipient, subject, or body."""
        return await self.email_client.search_emails(
            query, search_type, start_date, end_date, folder, top, inference_classification
        )

    async def search_emails_by_sender(
        self,
        sender: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Search emails by sender."""
        return await self.email_client.search_emails_by_sender(
            sender, folder, top, days
        )

    async def search_emails_by_recipient(
        self,
        recipient: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Search emails by recipient."""
        return await self.email_client.search_emails_by_recipient(
            recipient, folder, top, days
        )

    async def search_emails_by_subject(
        self,
        subject: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Search emails by subject."""
        return await self.email_client.search_emails_by_subject(
            subject, folder, top, days
        )

    async def search_emails_by_body(
        self,
        body: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Search emails by body content."""
        return await self.email_client.search_emails_by_body(body, folder, top, days)

    async def send_message(self, message_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send a message."""
        return await self.email_client.send_message(message_data)

    async def forward_message(
        self, message_id: str, to_recipients: List[str], comment: Optional[str] = None
    ) -> Dict[str, Any]:
        """Forward a message."""
        return await self.email_client.forward_message(
            message_id, to_recipients, comment
        )

    async def reply_to_message(
        self, message_id: str, comment: str, reply_all: bool = False
    ) -> Dict[str, Any]:
        """Reply to a message."""
        return await self.email_client.reply_to_message(message_id, comment, reply_all)

    async def send_email(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        reply_to_message_id: Optional[str] = None,
        forward_to_message_id: Optional[str] = None,
        body_content_type: str = "Text",
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Send an email."""
        return await self.email_client.send_email(
            to_recipients,
            subject,
            body,
            cc_recipients,
            bcc_recipients,
            reply_to_message_id,
            forward_to_message_id,
            body_content_type,
            importance,
        )

    async def create_draft_message(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: Optional[List[str]] = None,
        body_content_type: str = "HTML",
    ) -> Dict[str, Any]:
        """Create a draft in the Drafts folder without sending."""
        return await self.email_client.create_draft_message(
            to_recipients,
            subject,
            body,
            cc_recipients,
            body_content_type,
        )

    async def batch_forward_emails(
        self,
        to_recipients: Optional[List[str]] = None,
        subject: str = "",
        body: str = "",
        email_ids: List[str] = [],
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        body_content_type: str = "Text",
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Batch forward emails."""
        return await self.email_client.batch_forward_emails(
            to_recipients,
            subject,
            body,
            email_ids,
            cc_recipients,
            bcc_recipients,
            body_content_type,
            importance,
        )

    async def create_folder(
        self, folder_name: str, parent_folder: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new mail folder."""
        return await self.email_client.create_folder(folder_name, parent_folder)

    async def delete_folder(self, folder_path: str) -> Dict[str, Any]:
        """Delete a mail folder."""
        return await self.email_client.delete_folder(folder_path)

    async def rename_folder(self, folder_path: str, new_name: str) -> Dict[str, Any]:
        """Rename a mail folder."""
        return await self.email_client.rename_folder(folder_path, new_name)

    async def get_folder_details(self, folder_path: str) -> Dict[str, Any]:
        """Get detailed information about a folder."""
        return await self.email_client.get_folder_details(folder_path)

    async def move_email_to_folder(
        self, email_id: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Move an email to a different folder."""
        return await self.email_client.move_email_to_folder(
            email_id, destination_folder
        )

    async def copy_email_to_folder(
        self, email_id: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Copy an email to a different folder."""
        return await self.email_client.copy_email_to_folder(
            email_id, destination_folder
        )

    async def move_all_emails_from_folder(
        self, source_folder: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Move all emails from one folder to another."""
        return await self.email_client.move_all_emails_from_folder(
            source_folder, destination_folder
        )

    async def delete_email(self, email_id: str) -> Dict[str, Any]:
        """Delete an email by moving it to Deleted Items."""
        return await self.email_client.delete_email(email_id)

    async def batch_delete_emails(self, email_ids: List[str]) -> Dict[str, Any]:
        """Delete multiple emails using batch operations."""
        return await self.email_client.batch_delete_emails(email_ids)

    async def move_folder(
        self, folder_path: str, destination_parent: str
    ) -> Dict[str, Any]:
        """Move a folder to a different parent folder."""
        return await self.email_client.move_folder(folder_path, destination_parent)

    async def archive_email(self, email_id: str) -> Dict[str, Any]:
        """Archive an email by moving it to the Archive folder."""
        return await self.email_client.archive_email(email_id)

    async def batch_archive_emails(self, email_ids: List[str]) -> Dict[str, Any]:
        """Archive multiple emails using batch operations."""
        return await self.email_client.batch_archive_emails(email_ids)

    async def flag_email(self, email_id: str, flag_status: str) -> Dict[str, Any]:
        """Flag or unflag an email."""
        return await self.email_client.flag_email(email_id, flag_status)

    async def batch_flag_emails(
        self, email_ids: List[str], flag_status: str
    ) -> Dict[str, Any]:
        """Flag multiple emails using batch operations."""
        return await self.email_client.batch_flag_emails(email_ids, flag_status)

    async def categorize_email(
        self, email_id: str, categories: List[str]
    ) -> Dict[str, Any]:
        """Add categories to an email."""
        return await self.email_client.categorize_email(email_id, categories)

    async def batch_categorize_emails(
        self, email_ids: List[str], categories: List[str]
    ) -> Dict[str, Any]:
        """Categorize multiple emails using batch operations."""
        return await self.email_client.batch_categorize_emails(email_ids, categories)

    # Template management methods - delegated to EmailClient
    async def create_template_from_email(
        self,
        email_id: str,
        template_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a template by copying an email to the Templates folder."""
        return await self.email_client.create_template_from_email(
            email_id, template_name
        )

    async def list_templates(
        self,
        top: int = 50,
    ) -> List[Dict[str, Any]]:
        """List all templates in the Templates folder."""
        return await self.email_client.list_templates(top)

    async def get_template(
        self,
        template_id: str,
        return_html: bool = False,
    ) -> Dict[str, Any]:
        """Get a template by ID."""
        return await self.email_client.get_template(template_id, return_html)

    async def update_template(
        self,
        template_id: str,
        subject: Optional[str] = None,
        body: Optional[str] = None,
        to_recipients: Optional[List[Dict[str, str]]] = None,
        cc_recipients: Optional[List[Dict[str, str]]] = None,
        bcc_recipients: Optional[List[Dict[str, str]]] = None,
    ) -> Dict[str, Any]:
        """Update a template."""
        return await self.email_client.update_template(
            template_id, subject, body, to_recipients, cc_recipients, bcc_recipients
        )

    async def delete_template(
        self,
        template_id: str,
    ) -> Dict[str, Any]:
        """Delete a template."""
        return await self.email_client.delete_template(template_id)

    async def send_template(
        self,
        template_id: str,
        to: Optional[List[str]] = None,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Send a template (creates a copy and sends it, preserving the original template)."""
        return await self.email_client.send_template(template_id, to, cc, bcc)

    # Calendar management methods - delegated to CalendarClient
    async def browse_events(
        self,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
        skip: int = 0,
    ) -> Dict[str, Any]:
        """Browse calendar events."""
        return await self.calendar_client.browse_events(start_date, end_date, top, skip)

    async def get_event(self, event_id: str) -> Dict[str, Any]:
        """Get event by ID."""
        return await self.calendar_client.get_event(event_id)

    async def search_events(
        self,
        query: str,
        search_type: str = "organizer",
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        top: int = 20,
    ) -> List[Dict[str, Any]]:
        """Search calendar events."""
        return await self.calendar_client.search_events(
            query, search_type, start_date, end_date, top
        )

    async def check_calendar_conflict(
        self,
        start_date: str,
        end_date: str,
        exclude_event_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Check for calendar conflicts in a time range.

        Args:
            start_date: Start datetime in UTC ISO format
            end_date: End datetime in UTC ISO format
            exclude_event_id: Optional event ID to exclude (for updates)

        Returns:
            Dictionary with has_conflict, conflicting_events, message
        """
        return await self.calendar_client.check_calendar_conflict(
            start_date, end_date, exclude_event_id
        )

    async def create_event(self, event_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a calendar event."""
        return await self.calendar_client.create_event(event_data)

    async def update_event(
        self, event_id: str, event_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Update a calendar event."""
        return await self.calendar_client.update_event(event_id, event_data)

    async def cancel_event(self, event_id: str, comment: Optional[str] = None) -> None:
        """Cancel a calendar event."""
        return await self.calendar_client.cancel_event(event_id, comment)

    async def delete_event(self, event_id: str) -> None:
        """Delete a calendar event from your calendar."""
        return await self.calendar_client.delete_event(event_id)

    async def forward_event(
        self,
        event_id: str,
        attendees: List[Dict[str, str]],
        comment: Optional[str] = None,
    ) -> None:
        """Forward a calendar event."""
        return await self.calendar_client.forward_event(event_id, attendees, comment)

    async def reply_to_event(
        self, event_id: str, comment: str, reply_all: bool = False
    ) -> None:
        """Reply to a calendar event."""
        return await self.calendar_client.reply_to_event(event_id, comment, reply_all)

    async def accept_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Accept a calendar event invitation."""
        return await self.calendar_client.accept_event(
            event_id, comment, send_response, series
        )

    async def decline_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Decline a calendar event invitation."""
        return await self.calendar_client.decline_event(
            event_id, comment, send_response, series
        )

    async def tentatively_accept_event(
        self,
        event_id: str,
        comment: Optional[str] = None,
        send_response: bool = True,
        series: bool = False,
    ) -> None:
        """Tentatively accept a calendar event invitation."""
        return await self.calendar_client.tentatively_accept_event(
            event_id, comment, send_response, series
        )

    async def propose_new_time(
        self,
        event_id: str,
        proposed_new_time: Dict[str, str],
        comment: Optional[str] = None,
        send_response: bool = True,
    ) -> None:
        """Decline an event and propose a new time to the organizer."""
        return await self.calendar_client.propose_new_time(
            event_id, proposed_new_time, comment, send_response
        )

    async def check_availability(
        self,
        schedules: List[str],
        start_time: Optional[Dict[str, str]],
        end_time: Optional[Dict[str, str]],
        availability_view_interval: Optional[int] = None,
        date: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Check availability of attendees for a given time range."""
        return await self.calendar_client.check_availability(
            schedules, start_time, end_time, availability_view_interval, date
        )

    # File management methods - delegated to FileClient
    async def get_drive_items(self, folder_path: str = "") -> List[Dict[str, Any]]:
        """Get drive items."""
        return await self.file_client.get_drive_items(folder_path)

    async def upload_file(self, file_path: str, file_content: bytes) -> Dict[str, Any]:
        """Upload a file."""
        return await self.file_client.upload_file(file_path, file_content)

    # Teams management methods - delegated to TeamsClient
    async def get_teams(self) -> List[Dict[str, Any]]:
        """Get teams."""
        return await self.teams_client.get_teams()

    async def get_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """Get team channels."""
        return await self.teams_client.get_team_channels(team_id)


# Global graph client instance
graph_client = GraphClient()

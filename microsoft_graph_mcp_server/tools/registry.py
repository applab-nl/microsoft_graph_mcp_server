"""Tool definitions and registry for Microsoft Graph MCP Server."""

import mcp.types as types
from typing import List

from ..config import settings


class ToolRegistry:
    """Central registry for all MCP tool definitions."""

    @staticmethod
    def get_all_tools() -> List[types.Tool]:
        """Get all available tools."""
        return [
            ToolRegistry.auth(),
            ToolRegistry.user_settings(),
            ToolRegistry.search_contacts(),
            ToolRegistry.manage_mail_folder(),
            ToolRegistry.manage_emails(),
            ToolRegistry.browse_email_cache(),
            ToolRegistry.search_emails(),
            ToolRegistry.get_email_content(),
            ToolRegistry.send_email(),
            ToolRegistry.create_draft(),
            ToolRegistry.browse_events(),
            ToolRegistry.get_event_detail(),
            ToolRegistry.search_events(),
            ToolRegistry.check_attendee_availability(),
            ToolRegistry.manage_event_as_organizer(),
            ToolRegistry.manage_event_as_attendee(),
            # ToolRegistry.list_files(),  # Disabled: OneDrive not needed
            # ToolRegistry.get_teams(),  # Disabled: Teams not needed
            # ToolRegistry.get_team_channels(),  # Disabled: Teams not needed
            # ToolRegistry.manage_templates(),  # Disabled: Templates not needed
        ]

    @staticmethod
    def user_settings() -> types.Tool:
        """User settings tool definition."""
        return types.Tool(
            name="user_settings",
            description="Manage user settings with two actions: 'init' to initialize settings when configuration is missing or corrupted (requires multimodal_supported parameter to indicate if LLM supports image processing), 'update' to update one or more settings (supports partial updates). Note: 'init' action attempts to get user timezone from Graph API if authenticated, but works without login. 'update' action is purely local and requires no authentication. MULTIMODAL DETECTION: If you are a multimodal LLM that can process images, please set multimodal_supported=true when calling init or update.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["init", "update"],
                        "description": "Action to perform: 'init' to initialize settings when configuration is missing (requires multimodal_supported), 'update' to update one or more settings (partial update supported)",
                    },
                    "multimodal_supported": {
                        "type": "boolean",
                        "description": "Set to true if the LLM supports multimodal capabilities (can process images). Used with both 'init' and 'update' actions. When true, email image attachments will include base64 content for the LLM to analyze.",
                    },
                    "page_size": {
                        "type": "integer",
                        "description": "Page size for user browsing (default: 5, recommended range: 3-10). Only used with 'update' action.",
                        "minimum": 1,
                        "maximum": 50,
                    },
                    "llm_page_size": {
                        "type": "integer",
                        "description": "Page size for LLM browsing (default: 20, recommended range: 10-50). Only used with 'update' action.",
                        "minimum": 1,
                        "maximum": 100,
                    },
                    "default_search_days": {
                        "type": "integer",
                        "description": "Default number of days to search for emails when not specified (default: 7). Only used with 'update' action.",
                        "minimum": 1,
                        "maximum": 365,
                    },
                    "max_search_days": {
                        "type": "integer",
                        "description": "Maximum allowed search range in days (default: 90). Only used with 'update' action.",
                        "minimum": 1,
                        "maximum": 365,
                    },
                    "timezone": {
                        "type": "string",
                        "description": "User timezone in IANA format. Examples: 'America/New_York', 'Asia/Shanghai', 'Europe/London', 'UTC'. Only used with 'update' action.",
                    },
                },
                "required": ["action"],
            },
        )

    @staticmethod
    def search_contacts() -> types.Tool:
        """Search contacts tool definition."""
        return types.Tool(
            name="search_contacts",
            description="FIND PEOPLE/CONTACTS ONLY. Search for people by name or email address in organization directory. Returns contact information (name, email, etc.). DO NOT use this to search email messages - use search_emails for that. Use this when you need to find information about a person, such as 'who is John Smith' or 'find contact with email john@company.com'. Default limit: 10. Note: If you encounter a rate limit error (429), the response will include a 'retry_after' field indicating how many seconds to wait before retrying. Returns: {success: boolean, contacts: array, count: integer, limit_reached: boolean, message: string, retry_after: integer (if rate limited)}.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (person's name or email address)",
                    }
                },
                "required": ["query"],
            },
        )

    @staticmethod
    def auth() -> types.Tool:
        """Auth tool definition."""
        return types.Tool(
            name="auth",
            description="Manage authentication with Microsoft Graph. Four simple actions:\n\n"
            "• 'start' - Start login flow, returns verification URL and code to enter in browser\n"
            "• 'complete' - Complete login after browser authentication (auto-loads device code)\n"
            "• 'check_status' - Check auth status: returns authenticated or need login\n"
            "• 'logout' - Clear all authentication tokens\n\n"
            "WHEN AUTHENTICATION NEEDED: Always try 'check_status' FIRST. If it fails (no token or expired), then call 'start'.\n\n"
            "WORKFLOW:\n"
            "1. Call auth action='start' → get URL and code\n"
            "2. Open URL in browser, enter the code, complete Microsoft login\n"
            "3. Call auth action='complete' → authentication finished\n\n"
            "AUTO-REFRESH: Access tokens auto-refresh when expired (no user action needed). "
            "Refresh tokens are valid for ~90 days.\n\n"
            "Returns:\n"
            "- check_status: {status: string, authenticated: boolean, message: string}\n"
            "- start: {status: string, verification_uri: string, user_code: string, message: string}\n"
            "- complete: {status: string, authenticated: boolean, message: string}\n"
            "- logout: {status: string, authenticated: false, message: string}",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["start", "complete", "check_status", "logout"],
                        "description": "Action: 'start' to begin login (get URL/code), 'complete' to finish login after browser auth, 'check_status' to check authentication status, 'logout' to sign out",
                    },
                },
                "required": ["action"],
            },
        )

    @staticmethod
    def manage_mail_folder() -> types.Tool:
        """Mail folder tool definition."""
        return types.Tool(
            name="manage_mail_folder",
            description="Manage mail folders. Supports list, create, delete, rename, get_details, and move operations. Returns: {success: boolean, message: string, path: string, displayName: string, totalItemCount: integer, unreadItemCount: integer, childFolderCount: integer}. Note: Invalid folder paths return appropriate error messages.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": [
                            "list",
                            "create",
                            "delete",
                            "rename",
                            "get_details",
                            "move",
                        ],
                        "description": "Action to perform: 'list' to list all mail folders, 'create' to create a new folder, 'delete' to delete a folder, 'rename' to rename a folder, 'get_details' to get folder information, 'move' to move a folder to a new parent",
                    },
                    "folder_path": {
                        "type": "string",
                        "description": "Path of the folder (e.g., 'Inbox', 'Archive/2024'). Required for delete, rename, get_details, and move actions",
                    },
                    "folder_name": {
                        "type": "string",
                        "description": "Name of the folder to create. Required for create action",
                    },
                    "parent_folder": {
                        "type": "string",
                        "description": "Optional parent folder path for create action (e.g., 'Inbox', 'Archive/2024'). If not provided, creates a top-level folder",
                    },
                    "new_name": {
                        "type": "string",
                        "description": "New name for the folder. Required for rename action",
                    },
                    "destination_parent": {
                        "type": "string",
                        "description": "Path of the destination parent folder (e.g., 'Archive', 'Sent Items'). Required for move action",
                    },
                },
                "required": ["action"],
            },
        )

    @staticmethod
    def manage_emails() -> types.Tool:
        """Manage emails tool definition."""
        return types.Tool(
            name="manage_emails",
            description="Manage emails with multiple actions. Supports moving, deleting, archiving, flagging, and categorizing emails. Actions include: move_single, move_all, delete_single, delete_multiple, delete_all, archive_single, archive_multiple, flag_single, flag_multiple, categorize_single, categorize_multiple. Returns: {success: boolean, message: string, moved_count: integer, failed_count: integer, errors: array}. Note: Invalid cache_number returns appropriate error message.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": [
                            "move_single",
                            "move_all",
                            "delete_single",
                            "delete_multiple",
                            "delete_all",
                            "archive_single",
                            "archive_multiple",
                            "flag_single",
                            "flag_multiple",
                            "categorize_single",
                            "categorize_multiple",
                        ],
                        "description": "Action to perform: 'move_single' to move a single email, 'move_all' to move all emails from a folder, 'delete_single' to delete a single email, 'delete_multiple' to delete multiple emails in ONE batch call (provide cache_numbers array with all items), 'delete_all' to delete all emails from a folder, 'archive_single' to archive a single email, 'archive_multiple' to archive multiple emails in ONE batch call, 'flag_single' to flag a single email, 'flag_multiple' to flag multiple emails in ONE batch call, 'categorize_single' to categorize a single email, 'categorize_multiple' to categorize multiple emails in ONE batch call",
                    },
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_email_cache (e.g., 1, 2, 3). Required for 'move_single', 'delete_single', 'archive_single', 'flag_single', and 'categorize_single' actions",
                    },
                    "cache_numbers": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": "List of cache numbers from browse_email_cache (e.g., [1, 2, 3]). Required for 'delete_multiple', 'archive_multiple', 'flag_multiple', and 'categorize_multiple' actions. IMPORTANT: Pass ALL cache numbers in a SINGLE call - do NOT call multiple times for individual items",
                    },
                    "source_folder": {
                        "type": "string",
                        "description": "Source folder path (e.g., 'Inbox', 'Archive/2024'). Required for 'move_all' and 'delete_all' actions",
                    },
                    "destination_folder": {
                        "type": "string",
                        "description": "Destination folder path (e.g., 'Archive/2024', 'Inbox/Projects'). Required for 'move_single' and 'move_all' actions",
                    },
                    "flag_status": {
                        "type": "string",
                        "enum": ["flagged", "complete"],
                        "description": "Flag status: 'flagged' to mark as flagged, 'complete' to mark as complete. Required for 'flag_single' and 'flag_multiple' actions",
                    },
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of category names to apply (e.g., ['Important', 'Work']). Required for 'categorize_single' and 'categorize_multiple' actions",
                    },
                },
                "required": ["action"],
            },
        )

    @staticmethod
    def browse_email_cache() -> types.Tool:
        """Browse email cache tool definition."""
        return types.Tool(
            name="browse_email_cache",
            description="Browse emails in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence. WORKFLOW: Use search_emails to load emails into the cache first. **ATTACHMENT INFO**: Each email now includes attachment details (id, name, size, contentType) so you can see what attachments are available before downloading. **IMAGE ATTACHMENTS**: When you see image attachments (contentType starts with 'image/'), use get_email_content to retrieve them. Image content will be included automatically for multimodal analysis (requires MULTIMODAL_SUPPORTED=true). You MUST then analyze the images - they often contain critical information not mentioned in text. Returns: {current_page: integer, total_pages: integer, count: integer, total_count: integer, emails: array (with attachments array), date_range: string, filter_date_range: string, timezone: string}.",
            inputSchema={
                "type": "object",
                "properties": {
                    "page_number": {
                        "type": "integer",
                        "description": "Page number to view (starts at 1)",
                        "minimum": 1,
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["user", "llm"],
                        "description": "Browsing mode. DEFAULT to 'llm' for most cases - use 'llm' when user wants you to analyze, summarize, or process emails (~20 per page for comprehensive view). Use 'user' ONLY when user EXPLICITLY requests to browse page-by-page themselves (~5 per page for human reading).",
                        "default": "llm",
                    },
                },
                "required": ["page_number"],
            },
        )

    @staticmethod
    def search_emails() -> types.Tool:
        """Search emails tool definition."""
        return types.Tool(
            name="search_emails",
            description="SEARCH EMAIL MESSAGES ONLY. Search or list email messages by keywords, sender, subject, or body. Returns matching email messages with summary information. DO NOT use this to find people/contacts - use search_contacts for that. If no search_type and query are provided, lists emails within the specified time range. All time parameters use your local timezone. PARAMETER PRECEDENCE (highest to lowest): 1) time_range (overrides all other time parameters), 2) start_date/end_date (overrides days), 3) days (used only if no other time parameters provided). **IMPORTANT FOR OLD EMAILS**: The 'days' parameter has a maximum limit (default 90 days). To search emails older than this limit, use start_date/end_date parameters instead - they have NO date range restriction and can search any historical emails. Example: start_date='2023-01-01', end_date='2023-12-31'. **FOCUSED INBOX**: By default, searches only 'focused' emails (Outlook's Focused Inbox). Use inference_classification='other' to search 'other' emails, or inference_classification='all' to search both. Returns: {success: boolean, emails: array, count: integer, date_range: string, filter_date_range: string, timezone: string}. Note: Rate limit errors (HTTP 429) return retry_after field with seconds to wait. Invalid folder paths return appropriate error messages.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query for email messages. For 'sender': email address only. For 'subject': subject text. For 'body': body text content. Required when search_type is provided. Optional - if not provided with search_type, lists emails within the time range",
                    },
                    "search_type": {
                        "type": "string",
                        "enum": ["sender", "subject", "body"],
                        "description": "Type of search to perform (optional). Options: 'sender' (search by sender email address only), 'subject' (search by subject text with exact substring matching), 'body' (search by body content with exact substring matching). If not provided, lists emails within the time range without filtering",
                    },
                    "folder": {
                        "type": "string",
                        "description": "Optional folder path to search (e.g., 'Inbox', 'Inbox/Projects', 'Archive/2024'). Default: Inbox",
                        "default": "Inbox",
                    },
                    "inference_classification": {
                        "type": "string",
                        "enum": ["focused", "other", "all"],
                        "description": "Filter by Outlook Focused Inbox classification. 'focused' (default): important emails shown in Focused tab. 'other': less important emails shown in Other tab. 'all': search both focused and other emails. Default: focused",
                        "default": "focused",
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in your local timezone (e.g., '2024-01-01' or '2024-01-01T14:30'). Overridden by time_range if both are provided. **NO DATE RANGE LIMIT**: Unlike 'days' parameter, start_date has no restriction - use this to search old emails beyond the default 90-day limit.",
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in your local timezone (e.g., '2024-12-31' or '2024-12-31T23:59'). Overridden by time_range if both are provided. **NO DATE RANGE LIMIT**: Unlike 'days' parameter, end_date has no restriction - use this to search old emails beyond the default 90-day limit. **SPECIAL VALUE**: Use 'now' to search from start_date until current time (e.g., start_date='2024-01-01', end_date='now' searches from Jan 1 to now).",
                    },
                    "time_range": {
                        "type": "string",
                        "description": "Time range type (case-insensitive). Optional, in your local timezone. Accepted values: 'today', 'tomorrow', 'this_week', 'next_week', 'this_month', 'next_month' (any case). HIGHEST PRIORITY: If provided, overrides start_date, end_date, and days. Returns a user-friendly display string in the response. Examples: 'today', 'Today', 'THIS_WEEK', 'Next_Month'",
                    },
                    "days": {
                        "type": "integer",
                        "description": f"Number of days to look back. Default: {settings.default_search_days}, maximum: {settings.max_search_days}. LOWEST PRIORITY: Used only when no time_range, start_date, or end_date are provided. **TIP**: For searching emails older than {settings.max_search_days} days, use start_date and end_date parameters instead.",
                        "default": settings.default_search_days,
                        "minimum": 1,
                        "maximum": settings.max_search_days,
                    },
                },
            },
        )

    @staticmethod
    def get_email_content() -> types.Tool:
        """Get email content tool definition."""
        return types.Tool(
            name="get_email_content",
            description="Get full email content by cache number. Use the cache number from browse_email_cache (e.g., 1, 2, 3) to retrieve complete email with body, attachments, and all details. Returns: {success: boolean, subject: string, from: string, to: array, cc: array, bcc: array, body: string, attachments: array, sent_date: string, received_date: string}. Note: Invalid cache_number returns appropriate error message. **ATTACHMENT DOWNLOAD**: Set download_attachments=true to download attachments to the workspace/attachments folder. Attachments are saved with their original names. Use attachment_names to download only specific attachments. **IMAGE ATTACHMENTS**: When an email has image attachments (contentType starts with 'image/'), image content will be included automatically for multimodal analysis (requires MULTIMODAL_SUPPORTED=true in config). **IMPORTANT**: If images appear in the response, you MUST analyze them immediately - they often contain critical information (screenshots, photos, diagrams, charts, instructions) not mentioned in the text body. Do NOT skip image analysis.",
            inputSchema={
                "type": "object",
                "properties": {
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_email_cache (e.g., 1, 2, 3)",
                    },
                    "return_html": {
                        "type": ["boolean", "string"],
                        "description": "If true, return full HTML body. If false (default), return plain text body. Note: This only affects email body format, not image attachments.",
                        "default": False,
                    },
                    "download_attachments": {
                        "type": "boolean",
                        "description": "If true, download email attachments to the workspace/attachments folder. Default: false. Attachments will include 'downloaded' and 'file_path' fields in the response.",
                        "default": False,
                    },
                    "download_path": {
                        "type": "string",
                        "description": "Optional custom path for downloading attachments. If not specified, defaults to workspace/attachments folder.",
                    },
                    "attachment_names": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Optional list of specific attachment names to download. If not specified, downloads all non-inline attachments.",
                    },
                },
                "required": ["cache_number"],
            },
        )

    @staticmethod
    def send_email() -> types.Tool:
        """Send email tool definition."""
        return types.Tool(
            name="send_email",
            description="Send emails directly without creating drafts. Supports three actions: 'send_new' to send a new email, 'reply' to reply to an existing email, and 'forward' to forward an existing email. All actions send emails immediately - no drafts are created. Supports multiple recipients, CC, and BCC. The htmlbody parameter accepts HTML format for rich email content. **RECOMMENDED FOR BCC**: Use bcc_csv_file parameter to provide BCC recipients from a CSV file - this is the preferred method for handling large BCC lists with automatic batching support (up to 500 recipients per batch by default). **NOTE**: For forward action, 'to' is optional when using 'bcc_csv_file' or 'bcc' - you can forward emails using only BCC recipients. Returns: {success: boolean, message: string, sent_count: integer, failed_count: integer, recipients: array}.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["send_new", "reply", "forward"],
                        "description": "Action to perform: 'send_new' to send a new email immediately (no draft), 'reply' to reply to existing email, 'forward' to forward existing email",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses (required for send_new, optional for reply and forward). For reply: if not provided, defaults to original sender and recipients; if provided (even if empty), replaces original recipients. For forward: optional when bcc_csv_file is provided.",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject (required for send_new, optional for reply/forward)",
                    },
                    "htmlbody": {
                        "type": "string",
                        "description": "Email body content in HTML format. IMPORTANT HTML FORMATTING RULES: (1) Do NOT use <br> between <p> tags - use </p><p> directly for paragraph separation. Using both creates excessive spacing in email clients. (2) Do NOT add newlines or whitespace between HTML block elements (<p>, <div>, <ul>, etc.) - keep HTML compact. (3) Use <p> for paragraphs, <strong>/<em> for emphasis, <ul>/<li> for lists. (4) Only use <br> for line breaks WITHIN a paragraph, not between paragraphs. Correct example: '<p>Hello,</p><p>This is <strong>important</strong>.</p><p>Best regards</p>'",
                    },
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_email_cache (required for reply/forward, e.g., 1, 2, 3)",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of CC recipient email addresses (optional). For reply: if not provided, defaults to original CC recipients; if provided (even if empty), replaces original CC recipients.",
                    },
                    "bcc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of BCC recipient email addresses (optional)",
                    },
                    "bcc_csv_file": {
                        "type": "string",
                        "description": "**PREFERRED METHOD FOR BCC**: Path to CSV file containing BCC recipients. This is the recommended approach for handling BCC recipients, especially for large lists. CSV must have a single column with header 'Email' or 'email'. The system automatically batches large BCC lists (up to 500 recipients per batch by default) and sends multiple emails as needed. Only available for 'forward' action. When both 'bcc' and 'bcc_csv_file' are provided, they are combined.",
                    },
                    "importance": {
                        "type": "string",
                        "enum": ["normal", "high", "low"],
                        "description": "Email importance level: 'normal' (default), 'high', or 'low' (optional)",
                    },
                },
                "required": ["action", "htmlbody"],
            },
        )

    @staticmethod
    def create_draft() -> types.Tool:
        """Create draft tool definition."""
        return types.Tool(
            name="create_draft",
            description="Create a draft email in the user's Drafts (Concepts) folder WITHOUT sending it. The user reviews and sends it from Outlook. Use this instead of send_email when the email should be saved for review rather than sent immediately. Returns: {success: boolean, id: string, webLink: string}.",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject",
                    },
                    "htmlbody": {
                        "type": "string",
                        "description": "Email body content in HTML format. Same formatting rules as send_email: <p> per paragraph, no <br> between paragraphs, compact HTML.",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of CC recipient email addresses (optional)",
                    },
                },
                "required": ["to", "subject", "htmlbody"],
            },
        )

    @staticmethod
    def browse_events() -> types.Tool:
        """Browse events tool definition."""
        return types.Tool(
            name="browse_events",
            description="Browse calendar events in the cache with pagination. Returns summary information with number column indicating position in cache. Use page_number to navigate. Automatically manages browsing state with disk cache for persistence. WORKFLOW: Use search_events to load events into the cache first. Returns: {current_page: integer, total_pages: integer, count: integer, total_count: integer, events: array, date_range: string, timezone: string}.",
            inputSchema={
                "type": "object",
                "properties": {
                    "page_number": {
                        "type": "integer",
                        "description": "Page number to view (starts at 1)",
                        "minimum": 1,
                    },
                    "mode": {
                        "type": "string",
                        "enum": ["user", "llm"],
                        "description": "Browsing mode. DEFAULT to 'llm' for most cases - use 'llm' when user wants you to analyze, summarize, or process events (~20 per page for comprehensive view). Use 'user' ONLY when user EXPLICITLY requests to browse page-by-page themselves (~5 per page for human reading).",
                        "default": "llm",
                    },
                },
                "required": ["page_number"],
            },
        )

    @staticmethod
    def get_event_detail() -> types.Tool:
        """Get event detail tool definition."""
        return types.Tool(
            name="get_event_detail",
            description="Get detailed information for a specific calendar event by its cache number. WORKFLOW: First call browse_events or search_events to get event list, then use this tool with cache number from results. Returns: Event object with id, subject, start, end, location, attendees, body, recurrence, and online meeting details. Note: Invalid cache_number returns appropriate error message.",
            inputSchema={
                "type": "object",
                "properties": {
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_events or search_events (e.g., 1, 2, 3)",
                    }
                },
                "required": ["cache_number"],
            },
        )

    @staticmethod
    def search_events() -> types.Tool:
        """Search events tool definition."""
        return types.Tool(
            name="search_events",
            description="Search or list calendar events by keywords. Returns matching events with summary information. If no query is provided, lists events within the specified time range. All time parameters use your local timezone. When using time_range, the response includes a user-friendly display string (e.g., 'Today', 'This Week', 'This Month'). Returns: {success: boolean, events: array, count: integer, date_range: string, timezone: string}. Note: Subject search uses exact substring matching, organizer search uses fuzzy matching.",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query. For 'subject': event title text. For 'organizer': organizer name. Optional - if not provided, lists events within the time range",
                    },
                    "search_type": {
                        "type": "string",
                        "description": "Field to search in (optional). Options: 'subject' (search by event title with exact substring matching), 'organizer' (search by organizer name/email with fuzzy matching). Default: 'subject'",
                        "enum": ["subject", "organizer"],
                    },
                    "start_date": {
                        "type": "string",
                        "description": "Start date in your local timezone (e.g., '2024-01-01' or '2024-01-01T14:30') (optional)",
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in your local timezone (e.g., '2024-12-31' or '2024-12-31T23:59') (optional)",
                    },
                    "time_range": {
                        "type": "string",
                        "description": "Time range type (case-insensitive, optional, in your local timezone). Accepted values: 'today', 'tomorrow', 'this_week', 'next_week', 'this_month', 'next_month' (any case). If provided, overrides start_date and end_date. Returns a user-friendly display string in the response. Examples: 'today', 'Today', 'THIS_WEEK', 'Next_Month'.",
                    },
                },
            },
        )

    @staticmethod
    def check_attendee_availability() -> types.Tool:
        """Check attendee availability tool definition."""
        return types.Tool(
            name="check_attendee_availability",
            description="Check availability of attendees for a given date. WORKFLOW: Typically use before calling manage_my_event with create action to find optimal meeting times. Automatically includes the organizer (you) in the availability check to ensure overlap-free time slots. Automatically calculates time range based on all attendees' working hours. Returns: {success: boolean, message: string, availability_view: string, schedule_items: array, top_slots: array}. Availability view string uses single-character codes for each time interval: 0=Free, 1=Tentative, 2=Busy, 3=Out of office (OOF), 4=Working elsewhere, ?=Unknown. Note: Supports up to 20 attendees total (mandatory + optional).",
            inputSchema={
                "type": "object",
                "properties": {
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of mandatory attendee email addresses to check availability for",
                    },
                    "optional_attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of optional attendee email addresses to check availability for (optional)",
                    },
                    "date": {
                        "type": "string",
                        "description": "Date in ISO format. Example: '2024-01-01'. The time range will be automatically calculated based on all attendees' working hours.",
                    },
                    "time_zone": {
                        "type": "string",
                        "description": "Timezone for the time range. Optional - defaults to user's mailbox settings. Example: 'India Standard Time', 'Pacific Standard Time', 'UTC'",
                    },
                    "availability_view_interval": {
                        "type": "integer",
                        "description": "Time interval in minutes for availability view. Optional, default: 30. Valid values: 5, 6, 10, 15, 30, 60",
                    },
                    "top_slots": {
                        "type": "integer",
                        "description": "Number of top time slots to display in the summary. Optional, default: 5",
                    },
                    "meeting_duration": {
                        "type": "integer",
                        "description": "Required meeting duration in minutes. Optional, default: 30. Used to find continuous free time slots that can accommodate the meeting. Valid values: 15, 30, 45, 60, 90, 120, etc.",
                    },
                },
                "required": ["attendees", "date"],
            },
        )

    @staticmethod
    def manage_event_as_attendee() -> types.Tool:
        """Manage event as attendee tool definition for responding to events organized by others."""
        return types.Tool(
            name="manage_event_as_attendee",
            description="Manage calendar events where you are an attendee (events organized by others). Actions: accept, decline, tentatively_accept, propose_new_time, delete_cancelled, email_attendees. WORKFLOW: Use cache_number from browse_events or search_events results. Returns: Response confirmation message with action status and updated event information. Note: If event is already responded to, returns appropriate error message. IMPORTANT: accept/decline/tentatively_accept/propose_new_time actions will automatically handle events where the organizer didn't request responses. For accept: updates the event to showAs='busy' and enables reminders. For tentatively_accept: updates the event to showAs='tentative' and enables reminders. For decline: deletes the event from your calendar. For propose_new_time: deletes the event from your calendar and instructs you to contact organizer directly (since proposals require responses). This matches Outlook's behavior. The delete_cancelled action is specifically for removing a CANCELLED event from your calendar (ONLY use when organizer has cancelled the event - do NOT use to decline invitations). The email_attendees action sends email to event attendees (to=required attendees, cc=optional attendees, organizer excluded from recipients).",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": [
                            "accept",
                            "decline",
                            "tentatively_accept",
                            "propose_new_time",
                            "delete_cancelled",
                            "email_attendees",
                        ],
                        "description": "Action to perform: 'accept' to accept event invitation, 'decline' to decline event invitation, 'tentatively_accept' to tentatively accept event invitation, 'propose_new_time' to decline and propose new time to organizer (if responses aren't requested, deletes the event and instructs to contact organizer directly), 'delete_cancelled' to remove a CANCELLED event from your calendar (ONLY use when organizer has cancelled the event - do NOT use to decline invitations), 'email_attendees' to send email to event attendees (to=required attendees, cc=optional attendees, organizer excluded from recipients)",
                    },
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_events or search_events (required for all actions, e.g., 1, 2, 3)",
                    },
                    "comment": {
                        "type": "string",
                        "description": "Optional comment for accept, decline, tentatively_accept, propose_new_time actions",
                    },
                    "send_response": {
                        "type": "boolean",
                        "description": "Whether to send response to organizer for accept, decline, tentatively_accept actions (optional, default: true)",
                    },
                    "series": {
                        "type": "boolean",
                        "description": "For accept/decline/tentatively_accept actions on recurring events: set to true to accept/decline entire series, or false (default) for single occurrence only",
                    },
                    "propose_new_time": {
                        "type": "object",
                        "description": "Propose a new time when using propose_new_time action (required for propose_new_time action). Note: If the organizer hasn't requested responses, the event will be deleted and you'll need to contact the organizer directly to suggest the new time.",
                        "properties": {
                            "dateTime": {
                                "type": "string",
                                "description": "Proposed new date and time in your local timezone. Example: '2024-12-31T14:30' or '2024-12-31 14:30'. The system will automatically convert to UTC using your timezone settings from your Microsoft 365 profile or .env configuration",
                            },
                            "timeZone": {
                                "type": "string",
                                "description": "Time zone for the proposed time. Optional - will use your Microsoft 365 profile timezone or .env configuration if not provided",
                            },
                        },
                        "required": ["dateTime"],
                    },
                    "email_subject": {
                        "type": "string",
                        "description": "Email subject for email_attendees action (optional, default: 'Re: Event')",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'to' recipient email addresses for email_attendees action (optional, defaults to required event attendees)",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'cc' recipient email addresses for email_attendees action (optional, defaults to optional event attendees)",
                    },
                    "body": {
                        "type": "string",
                        "description": "Email body content for email_attendees action (optional, defaults to event body content)",
                    },
                },
                "required": ["action", "cache_number"],
            },
        )

    @staticmethod
    def manage_event_as_organizer() -> types.Tool:
        """Manage event as organizer tool definition for managing user's own events."""
        return types.Tool(
            name="manage_event_as_organizer",
            description="Manage calendar events where you are the organizer (events you created). Actions: create, update, cancel, forward, email_attendees. WORKFLOW: For update, cancel, forward, and email_attendees actions, use cache number from browse_events or returned when creating an event. Returns: Event object with id, subject, start, end, location, attendees, body, recurrence, and online meeting details. Note: Conflict errors may occur when updating event times that overlap with existing events.",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["create", "update", "cancel", "forward", "email_attendees"],
                        "description": "Action to perform: 'create' to create a new calendar event, 'update' to update an existing event, 'cancel' to cancel an event and send cancellation notifications to attendees, 'forward' to forward event by adding new optional attendees, 'email_attendees' to send email to event attendees using event body as content (to=required attendees, cc=optional attendees)",
                    },
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_events or returned when creating an event (required for update, cancel, forward, email_attendees actions, e.g., 1, 2, 3)",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Event subject (required for create, optional for update)",
                    },
                    "start": {
                        "type": "string",
                        "description": "Start date and time in your local timezone. Example: '2024-01-01T14:30' or '2024-01-01 14:30'. The system will automatically convert to UTC using the timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration. Required for create, optional for update",
                    },
                    "end": {
                        "type": "string",
                        "description": "End date and time in your local timezone. Example: '2024-01-01T15:30' or '2024-01-01 15:30'. The system will automatically convert to UTC using the timezone parameter or your timezone settings from your Microsoft 365 profile or .env configuration. Required for create, optional for update",
                    },
                    "timezone": {
                        "type": "string",
                        "description": "Timezone for the event in IANA format. Examples: 'Asia/Singapore', 'America/New_York', 'Europe/London', 'UTC'. Optional for create and update actions - if not provided, will use your timezone settings from your Microsoft 365 profile or .env configuration",
                    },
                    "location": {
                        "type": "string",
                        "description": "Event location (optional for create, update)",
                    },
                    "body": {
                        "type": "string",
                        "description": "Event body content in HTML format (optional for create, update, reply). IMPORTANT: (1) Do NOT use <br> between <p> tags - use </p><p> directly for paragraph separation. (2) Do NOT add newlines or whitespace between HTML block elements - keep HTML compact. Use <p> for paragraphs, <strong>/<em> for emphasis.",
                    },
                    "body_content_type": {
                        "type": "string",
                        "enum": ["Text", "HTML"],
                        "description": "Body content type (optional for create, update, default: HTML)",
                    },
                    "attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of required attendee email addresses (optional for create, update, required for forward)",
                    },
                    "optional_attendees": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of optional attendee email addresses (optional for create, update)",
                    },
                    "isOnlineMeeting": {
                        "type": "boolean",
                        "description": "Whether to create the event as an online meeting (optional for create, update). If true, creates a Teams meeting when onlineMeetingProvider is 'teamsForBusiness'. For other providers (Zoom, Google Meet, etc.), set isOnlineMeeting to true and include the join link in the body field",
                    },
                    "onlineMeetingProvider": {
                        "type": "string",
                        "enum": [
                            "teamsForBusiness",
                            "skypeForBusiness",
                            "skypeForConsumer",
                            "unknown",
                        ],
                        "description": "Online meeting provider (optional for create, update). Use 'teamsForBusiness' for Teams meetings, 'skypeForBusiness' for Skype for Business, 'skypeForConsumer' for Skype Consumer, or 'unknown' for other providers. For 'unknown' or other providers (Zoom, Google Meet, etc.), include the join link in the body field. Requires isOnlineMeeting to be true",
                    },
                    "recurrence": {
                        "type": "object",
                        "description": "Recurrence pattern for the event (optional for create, update). Defines how the event repeats",
                        "properties": {
                            "pattern": {
                                "type": "object",
                                "description": "Recurrence pattern - how often the event repeats",
                                "properties": {
                                    "type": {
                                        "type": "string",
                                        "enum": [
                                            "daily",
                                            "weekly",
                                            "absoluteMonthly",
                                            "relativeMonthly",
                                            "absoluteYearly",
                                            "relativeYearly",
                                        ],
                                        "description": "The recurrence type: daily, weekly, absoluteMonthly (e.g., 'day 15 of every month'), relativeMonthly (e.g., 'second Tuesday of every month'), absoluteYearly (e.g., 'April 15 of every year'), relativeYearly (e.g., 'third Tuesday of April of every year')",
                                    },
                                    "interval": {
                                        "type": "integer",
                                        "description": "The interval between occurrences. For example, interval=2 for type='weekly' means every 2 weeks",
                                        "minimum": 1,
                                    },
                                    "daysOfWeek": {
                                        "type": "array",
                                        "items": {
                                            "type": "string",
                                            "enum": [
                                                "sunday",
                                                "monday",
                                                "tuesday",
                                                "wednesday",
                                                "thursday",
                                                "friday",
                                                "saturday",
                                            ],
                                        },
                                        "description": "Days of the week for weekly or relativeMonthly/relativeYearly patterns",
                                    },
                                    "dayOfMonth": {
                                        "type": "integer",
                                        "description": "Day of the month for absoluteMonthly or absoluteYearly patterns (1-31)",
                                        "minimum": 1,
                                        "maximum": 31,
                                    },
                                    "month": {
                                        "type": "integer",
                                        "description": "Month for absoluteYearly or relativeYearly patterns (1-12)",
                                        "minimum": 1,
                                        "maximum": 12,
                                    },
                                    "index": {
                                        "type": "string",
                                        "enum": [
                                            "first",
                                            "second",
                                            "third",
                                            "fourth",
                                            "last",
                                        ],
                                        "description": "Index for relativeMonthly or relativeYearly patterns (e.g., 'first', 'second', 'last')",
                                    },
                                },
                                "required": ["type", "interval"],
                            },
                            "range": {
                                "type": "object",
                                "description": "Recurrence range - how long the recurrence lasts",
                                "properties": {
                                    "type": {
                                        "type": "string",
                                        "enum": ["endDate", "noEnd", "numbered"],
                                        "description": "The range type: endDate (ends on a specific date), noEnd (never ends), numbered (ends after a specific number of occurrences)",
                                    },
                                    "startDate": {
                                        "type": "string",
                                        "description": "Start date of the recurrence in ISO format (e.g., '2024-01-01')",
                                    },
                                    "endDate": {
                                        "type": "string",
                                        "description": "End date of the recurrence in ISO format (e.g., '2024-12-31'). Required for type='endDate'",
                                    },
                                    "numberOfOccurrences": {
                                        "type": "integer",
                                        "description": "Number of occurrences. Required for type='numbered'",
                                        "minimum": 1,
                                    },
                                },
                                "required": ["type", "startDate"],
                            },
                        },
                        "required": ["pattern", "range"],
                    },
                    "comment": {
                        "type": "string",
                        "description": "Optional comment for cancel, forward actions",
                    },
                    "email_subject": {
                        "type": "string",
                        "description": "Email subject for email_attendees action (optional, default: 'Re: Event')",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'to' recipient email addresses for email_attendees action (optional, defaults to required event attendees)",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 'cc' recipient email addresses for email_attendees action (optional, defaults to optional event attendees)",
                    },
                },
                "required": ["action"],
            },
        )

    @staticmethod
    def list_files() -> types.Tool:
        """List files tool definition."""
        return types.Tool(
            name="list_files",
            description="List files and folders in OneDrive. Returns: Array of file/folder objects with id, name, type (file/folder), size, last_modified, and path information.",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "Folder path in OneDrive. Example: '/Documents', '/Projects', '/Archive'. Default: root folder",
                        "default": "",
                    }
                },
            },
        )

    @staticmethod
    def get_teams() -> types.Tool:
        """Get teams tool definition."""
        return types.Tool(
            name="get_teams",
            description="Get list of Teams that you are a member of. WORKFLOW: Use returned team_id with get_team_channels. Returns: Array of team objects with id, displayName, description, and memberCount.",
            inputSchema={"type": "object", "properties": {}},
        )

    @staticmethod
    def get_team_channels() -> types.Tool:
        """Get team channels tool definition."""
        return types.Tool(
            name="get_team_channels",
            description="Get channels for a specific Team. WORKFLOW: team_id comes from get_teams tool output. Returns: Array of channel objects with id, displayName, and channel_type.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {
                        "type": "string",
                        "description": "Team ID from get_teams output",
                    }
                },
                "required": ["team_id"],
            },
        )

    @staticmethod
    def manage_templates() -> types.Tool:
        """Manage email templates tool definition."""
        return types.Tool(
            name="manage_templates",
            description="Manage email templates stored as drafts in a Templates folder. Templates are draft emails that can be edited and sent. WORKFLOW: 1) User calls get with return_html=false to view simple text body, 2) User provides update instructions to LLM, 3) LLM calls get with return_html=true to retrieve full HTML, 4) LLM applies user's updates and calls update with complete updated HTML in htmlbody parameter, 5) User calls get with return_html=false to verify changes, 6) User gives command to send, 7) LLM calls send to send template (creates a copy and sends it, preserving original). Returns: {success: boolean, message: string, template_number: integer, subject: string, body: string}. Note: return_html parameter is for get action (simple text vs full HTML), htmlbody parameter is for update action (complete updated HTML).",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": [
                            "create_from_email",
                            "list",
                            "get",
                            "update",
                            "delete",
                            "send",
                        ],
                        "description": "Action to perform: 'create_from_email' to copy an existing email as a template, 'list' to browse templates with pagination, 'get' to view template details (simple text or full HTML), 'update' to edit template content, 'delete' to remove a template (soft delete - moves to Deleted Items folder), 'send' to send a template (creates a copy and sends it, preserving original)",
                    },
                    "cache_number": {
                        "type": "integer",
                        "description": "Cache number from browse_email_cache to copy as template. Required for 'create_from_email' action.",
                    },
                    "template_number": {
                        "type": "integer",
                        "description": "Template cache number. Required for 'get', 'update', 'delete', and 'send' actions.",
                    },
                    "subject": {
                        "type": "string",
                        "description": "Email subject (title). Optional for 'update' action.",
                    },
                    "to": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of recipient email addresses. Optional for 'update' action - if not provided, keeps existing recipients.",
                    },
                    "cc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of CC recipient email addresses. Optional for 'update' action.",
                    },
                    "bcc": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of BCC recipient email addresses. Optional for 'update' action.",
                    },
                    "htmlbody": {
                        "type": "string",
                        "description": "Email body content in HTML format. Optional for 'update' action - if not provided, keeps existing body. IMPORTANT HTML FORMATTING RULES: (1) Do NOT use <br> between <p> tags - use </p><p> directly for paragraph separation. (2) Do NOT add newlines or whitespace between HTML block elements - keep HTML compact. (3) Use <p> for paragraphs, <strong>/<em> for emphasis, <ul>/<li> for lists. (4) Only use <br> for line breaks WITHIN a paragraph. Note: When updating body, first call get with return_html=true to get the full HTML, then provide the complete updated HTML here.",
                    },
                    "return_html": {
                        "type": ["boolean", "string"],
                        "description": "For 'get' action: if true, returns full HTML body. If false (default), returns simple text body. Default: false",
                    },
                },
                "required": ["action"],
            },
        )

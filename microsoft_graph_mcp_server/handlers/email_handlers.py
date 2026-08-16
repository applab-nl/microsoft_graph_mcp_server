"""Email handlers for MCP tools."""

from datetime import datetime
from zoneinfo import ZoneInfo
from typing import Optional, List
import mcp.types as types
from .base import BaseHandler
from ..utils import read_bcc_from_csv, date_handler, normalize_email_html
from ..cache import email_cache, template_cache
from ..graph_client import graph_client
from ..config import settings
from ..clients.email_client import MAX_EMAIL_SEARCH_LIMIT
from ..validation import validate_cache_number, ValidationError, validate_email_address


class EmailHandler(BaseHandler):
    """Handler for email-related tools."""

    async def handle_browse_email_cache(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle browse_email_cache tool."""
        page_number = arguments["page_number"]
        mode = arguments.get("mode", "llm")

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_response(
                {
                    "message": "No emails in cache. Use load_emails_by_folder to load emails first.",
                    "emails": [],
                    "count": 0,
                }
            )

        page_size = settings.llm_page_size if mode == "llm" else settings.page_size
        start_idx = (page_number - 1) * page_size
        end_idx = start_idx + page_size
        page_emails = cached_emails[start_idx:end_idx]

        user_timezone = await graph_client.get_user_timezone()

        filtered_emails = []
        for idx, email in enumerate(page_emails):
            filtered_email = {
                k: v for k, v in email.items() if k not in ["id", "metadata"]
            }
            filtered_email["number"] = start_idx + idx + 1
            filtered_emails.append(filtered_email)

        return self._format_response(
            {
                "emails": filtered_emails,
                "count": len(filtered_emails),
                "total_count": total_count,
                "current_page": page_number,
                "total_pages": (total_count + page_size - 1) // page_size,
                "page_size": page_size,
                "mode": mode,
                "timezone": user_timezone,
            }
        )

    async def handle_search_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle search_emails tool. This operation clears and reloads the cache."""
        query = arguments.get("query")
        search_type = arguments.get("search_type")
        folder = arguments.get("folder", "Inbox")
        inference_classification = arguments.get("inference_classification", "focused")
        days = arguments.get("days")
        start_date = arguments.get("start_date")
        end_date = arguments.get("end_date")
        time_range = arguments.get("time_range")
        # Configurable via SEARCH_PAGE_SIZE; Iris injects it so this always
        # matches its own per-run message cap (see config.search_page_size).
        page_size = settings.search_page_size

        days_provided = "days" in arguments

        # Apply default days for performance optimization
        # KQL date filters are now embedded in $search query for server-side filtering
        if days is None:
            days = settings.default_search_days

        if days > settings.max_search_days:
            return self._format_error(
                f"Error: Days parameter must be {settings.max_search_days} or less."
            )

        # Clear cache and reload with search results - this is intentional for search operations
        email_cache.clear_cache()

        success, user_timezone, error = await self._handle_auth_error(
            lambda: graph_client.get_user_timezone(), "getting user timezone"
        )
        if not success:
            return self._format_error(error)

        today_date = date_handler.get_today_date(user_timezone)

        if days_provided:
            start_date, end_date = date_handler.get_filter_date_range(days)
            start_date_display = date_handler.format_date_with_weekday(
                start_date, user_timezone
            )
            end_date_display = date_handler.format_date_with_weekday(
                end_date, user_timezone
            )
        elif time_range:
            display_range, start_date_local, end_date_local = date_handler.parse_date_range(
                time_range, user_timezone
            )
            start_date = date_handler.parse_local_date_to_utc(start_date_local, user_timezone)
            end_date = date_handler.parse_local_date_to_utc(end_date_local, user_timezone)
            start_date_display = date_handler.format_date_with_weekday(
                start_date, user_timezone
            )
            end_date_display = date_handler.format_date_with_weekday(
                end_date, user_timezone
            )
        elif not start_date and not end_date:
            # Calculate date range using days parameter (default or user-specified)
            start_date, end_date = date_handler.get_filter_date_range(days)
            start_date_display = date_handler.format_date_with_weekday(
                start_date, user_timezone
            )
            end_date_display = date_handler.format_date_with_weekday(
                end_date, user_timezone
            )
        else:
            # Use explicit start_date and/or end_date
            if start_date:
                start_date = date_handler.parse_local_date_to_utc(
                    start_date, user_timezone
                )
                start_date_display = date_handler.format_date_with_weekday(
                    start_date, user_timezone
                )
            if end_date:
                # Handle special value "now" - use current time
                if end_date.lower() == "now":
                    end_date = datetime.now(ZoneInfo(user_timezone)).strftime("%Y-%m-%dT%H:%M:%S")
                    end_date_display = date_handler.format_date_with_weekday(
                        end_date, user_timezone
                    )
                else:
                    end_date = date_handler.parse_local_date_to_utc(end_date, user_timezone)
                    end_date_display = date_handler.format_date_with_weekday(
                        end_date, user_timezone
                    )

        success, result, error = await self._handle_auth_error(
            lambda: graph_client.search_emails(
                query, search_type, start_date, end_date, folder, page_size, inference_classification
            ),
            "searching emails",
        )
        if not success:
            return self._format_error(error)

        await email_cache.set_mode("search")
        await email_cache.update_search_state(
            query=query,
            folder=folder,
            top=page_size,
            days=days,
            search_type=search_type,
            total_count=result["count"],
            metadata=result["metadata"],
        )

        response_data = {
            "search_type": search_type,
            "query": query,
            "folder": folder,
            "inference_classification": inference_classification,
            "count": result["count"],
            "timezone": user_timezone,
            "today": today_date,
            "hint": f"Found {result['count']} emails. Use browse_email_cache to view the results.",
        }

        if time_range or start_date:
            response_data["start_date"] = start_date
            response_data["start_date_display"] = start_date_display
        if time_range or end_date:
            response_data["end_date"] = end_date
            response_data["end_date_display"] = end_date_display
        if time_range:
            response_data["time_range"] = display_range

        return self._format_response(response_data)

    async def handle_get_email_content(
        self, arguments: dict
    ) -> list[types.TextContent | types.ImageContent]:
        """Handle get_email_content tool.
        
        Returns text content and optionally image content for multimodal LLMs.
        """
        cache_number = arguments["cache_number"]
        return_html = arguments.get("return_html", False)
        download_attachments = arguments.get("download_attachments", False)
        download_path = arguments.get("download_path")
        attachment_names = arguments.get("attachment_names")

        cached_emails = email_cache.get_cached_emails()

        try:
            validate_cache_number(cache_number, len(cached_emails), "email_cache")
        except ValidationError as e:
            return self._format_response({"error": e.message})

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_response(
                {
                    "error": f"Cache number {cache_number} does not have a valid Graph ID in cache."
                }
            )

        # Get multimodal_supported from settings
        from ..config import settings
        multimodal_supported = settings.multimodal_supported
        import logging
        logging.info(f"[MULTIMODAL DEBUG] Handler: multimodal_supported={multimodal_supported} (from settings.multimodal_supported)")

        success, email_content, error = await self._handle_auth_error(
            lambda: graph_client.get_email(
                email_id, cache_number, return_html, download_attachments, download_path, attachment_names, multimodal_supported
            ),
            "getting email content",
        )
        if not success:
            return self._format_error(error)

        # Build response content list
        content_list: list[types.TextContent | types.ImageContent] = []
        
        # Add text content (email content as JSON)
        import json
        content_list.append(types.TextContent(
            type="text",
            text=json.dumps(email_content["content"], indent=2, ensure_ascii=False, default=str)
        ))
        
        # Add image content for multimodal LLMs
        if multimodal_supported:
            attachments = email_content["content"].get("attachments", [])
            for attachment in attachments:
                # Only include downloaded image attachments with base64 content
                if (attachment.get("downloaded") and 
                    attachment.get("base64_content") and 
                    attachment.get("contentType", "").startswith("image/")):
                    
                    content_list.append(types.ImageContent(
                        type="image",
                        data=attachment["base64_content"],
                        mimeType=attachment["contentType"]
                    ))
        
        return content_list

    async def handle_send_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle send_email tool with send_new, reply, and forward actions."""
        action = arguments.get("action")

        if action == "send_new":
            return await self._handle_send_new_email(arguments)
        elif action == "reply":
            return await self._handle_reply_email(arguments)
        elif action == "forward":
            return await self._handle_forward_email(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'send_new', 'reply', or 'forward'."
            )

    async def _handle_send_new_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle send_new email action."""
        to_recipients = arguments["to"]
        subject = arguments["subject"]
        body = normalize_email_html(arguments["htmlbody"])
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        importance = arguments.get("importance")

        from ..config import MAX_RECIPIENTS_LIMIT

        to_count = len(to_recipients) if to_recipients else 0
        cc_count = len(cc_recipients) if cc_recipients else 0
        bcc_count = len(bcc_recipients) if bcc_recipients else 0

        total_recipients = to_count + cc_count + bcc_count

        # If total recipients exceed limit, batch them
        if total_recipients > MAX_RECIPIENTS_LIMIT:
            return await self._handle_batch_send_new_email(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                importance=importance,
                total_recipients=total_recipients,
            )

        success, result, error = await self._handle_auth_error(
            lambda: graph_client.send_email(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body_content_type="HTML",
                importance=importance,
            ),
            "sending email",
        )
        if not success:
            return self._format_error(error)

        return self._format_response(f"Email sent successfully: {result}")

    async def handle_create_draft(self, arguments: dict) -> list[types.TextContent]:
        """Handle create_draft tool — save to the Drafts folder without sending."""
        to_recipients = arguments["to"]
        subject = arguments["subject"]
        body = normalize_email_html(arguments["htmlbody"])
        cc_recipients = arguments.get("cc")

        if not to_recipients:
            return self._format_error("create_draft: 'to' must contain at least one recipient.")
        try:
            for email in to_recipients + (cc_recipients or []):
                validate_email_address(email)
        except ValidationError as e:
            return self._format_error(f"create_draft: {e}")

        success, result, error = await self._handle_auth_error(
            lambda: graph_client.create_draft_message(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                cc_recipients=cc_recipients,
                body_content_type="HTML",
            ),
            "creating draft",
        )
        if not success:
            return self._format_error(error)

        # _format_response json.dumps the payload itself — pass the dict.
        return self._format_response(
            {
                "success": True,
                "id": result.get("id", ""),
                "webLink": result.get("webLink", ""),
            }
        )

    async def _handle_reply_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle reply email action.

        Behavior:
        - If 'to' and 'cc' are not provided in arguments, reply uses original recipients
          (original sender + original 'to' recipients in 'to' field, original 'cc' in 'cc' field)
        - If 'to' or 'cc' are provided in arguments (even if empty arrays),
          they replace the original recipients
        """
        cache_number = arguments["cache_number"]
        # Check if 'to' and 'cc' were explicitly provided in the arguments
        to_was_provided = 'to' in arguments
        cc_was_provided = 'cc' in arguments

        to_recipients = arguments.get("to")
        subject = arguments.get("subject")
        body = normalize_email_html(arguments.get("htmlbody"))
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        importance = arguments.get("importance")

        cached_emails = email_cache.get_cached_emails()
        if cache_number < 1 or cache_number > len(cached_emails):
            return self._format_error(
                f"Error: Cache number {cache_number} is out of range. Please use a number between 1 and {len(cached_emails)}."
            )

        email = cached_emails[cache_number - 1]
        email_id = email["id"]

        from ..config import MAX_RECIPIENTS_LIMIT

        # For reply action, implement the following logic:
        # 1. If 'to' and 'cc' are not provided at all, keep original values
        # 2. If 'to' or 'cc' are provided (even if empty), use the provided values

        # If 'to' was not provided, set to None to trigger default reply behavior
        if not to_was_provided:
            to_recipients = None

        # If 'cc' was not provided, set to None to trigger default reply behavior
        if not cc_was_provided:
            cc_recipients = None

        to_count = len(to_recipients) if to_recipients else 0
        cc_count = len(cc_recipients) if cc_recipients else 0
        bcc_count = len(bcc_recipients) if bcc_recipients else 0

        total_recipients = to_count + cc_count + bcc_count

        # If total recipients exceed limit, batch them
        if total_recipients > MAX_RECIPIENTS_LIMIT:
            return await self._handle_batch_reply_email(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                email_id=email_id,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                importance=importance,
                total_recipients=total_recipients,
            )

        success, result, error = await self._handle_auth_error(
            lambda: graph_client.send_email(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                reply_to_message_id=email_id,
                body_content_type="HTML",
                importance=importance,
            ),
            "sending reply email",
        )
        if not success:
            return self._format_error(error)

        return self._format_response(f"Reply email sent successfully: {result}")

    async def _handle_forward_email(self, arguments: dict) -> list[types.TextContent]:
        """Handle forward email action."""
        cache_number = arguments["cache_number"]
        to_recipients = arguments.get("to")
        subject = arguments.get("subject")
        body = normalize_email_html(arguments.get("htmlbody"))
        cc_recipients = arguments.get("cc")
        bcc_recipients = arguments.get("bcc")
        bcc_csv_file = arguments.get("bcc_csv_file")
        importance = arguments.get("importance")

        if not to_recipients and not bcc_csv_file and not bcc_recipients:
            return self._format_error(
                "Error: At least one of 'to', 'bcc', or 'bcc_csv_file' must be provided for forwarding email."
            )

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        if bcc_csv_file:
            try:
                csv_bcc = read_bcc_from_csv(bcc_csv_file)
                if bcc_recipients:
                    bcc_recipients = bcc_recipients + csv_bcc
                else:
                    bcc_recipients = csv_bcc
            except Exception as e:
                return self._format_error(f"Error reading BCC CSV file: {str(e)}")

        from ..config import MAX_RECIPIENTS_LIMIT

        to_count = len(to_recipients) if to_recipients else 0
        cc_count = len(cc_recipients) if cc_recipients else 0
        bcc_count = len(bcc_recipients) if bcc_recipients else 0

        total_recipients = to_count + cc_count + bcc_count

        # If total recipients exceed limit, batch them
        if total_recipients > MAX_RECIPIENTS_LIMIT:
            return await self._handle_batch_forward_email(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                email_id=email_id,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                importance=importance,
                total_recipients=total_recipients,
            )

        success, result, error = await self._handle_auth_error(
            lambda: graph_client.batch_forward_emails(
                to_recipients=to_recipients,
                subject=subject,
                body=body,
                email_ids=[email_id],
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body_content_type="HTML",
                importance=importance,
            ),
            "forwarding email",
        )
        if not success:
            return self._format_error(error)

        response_message = f"Email forwarded successfully: {result}"
        if bcc_recipients:
            response_message = (
                f"Email forwarded successfully to {bcc_count} BCC recipients: {result}"
            )

        return self._format_response(response_message)

    async def _handle_batch_forward_email(
        self,
        to_recipients: Optional[List[str]],
        subject: Optional[str],
        body: str,
        email_id: str,
        cc_recipients: Optional[List[str]],
        bcc_recipients: Optional[List[str]],
        importance: Optional[str],
        total_recipients: int,
    ) -> list[types.TextContent]:
        """Handle batch forwarding of emails when recipients exceed the limit.

        Splits recipients into batches and sends multiple emails.

        Args:
            to_recipients: List of TO recipients
            subject: Email subject
            body: Email body
            email_id: ID of the email to forward
            cc_recipients: List of CC recipients
            bcc_recipients: List of BCC recipients
            importance: Email importance
            total_recipients: Total number of recipients

        Returns:
            Response with batch send results
        """
        from ..config import MAX_RECIPIENTS_LIMIT

        sent_count = 0
        failed_count = 0
        errors = []

        # Combine all recipients into a single list
        all_recipients = []
        if to_recipients:
            all_recipients.extend([("to", email) for email in to_recipients])
        if cc_recipients:
            all_recipients.extend([("cc", email) for email in cc_recipients])
        if bcc_recipients:
            all_recipients.extend([("bcc", email) for email in bcc_recipients])

        # Split into batches
        batches = []
        current_batch = {"to": [], "cc": [], "bcc": []}
        current_count = 0

        for recipient_type, email in all_recipients:
            if current_count >= MAX_RECIPIENTS_LIMIT:
                batches.append(current_batch)
                current_batch = {"to": [], "cc": [], "bcc": []}
                current_count = 0

            current_batch[recipient_type].append(email)
            current_count += 1

        if current_batch["to"] or current_batch["cc"] or current_batch["bcc"]:
            batches.append(current_batch)

        # Send each batch
        for idx, batch in enumerate(batches, 1):
            batch_to_count = len(batch["to"])
            batch_cc_count = len(batch["cc"])
            batch_bcc_count = len(batch["bcc"])
            batch_total = batch_to_count + batch_cc_count + batch_bcc_count

            success, result, error = await self._handle_auth_error(
                lambda b=batch: graph_client.batch_forward_emails(
                    to_recipients=b["to"] if b["to"] else None,
                    subject=subject,
                    body=body,
                    email_ids=[email_id],
                    cc_recipients=b["cc"] if b["cc"] else None,
                    bcc_recipients=b["bcc"] if b["bcc"] else None,
                    body_content_type="HTML",
                    importance=importance,
                ),
                f"forwarding email batch {idx}/{len(batches)}",
            )

            if success:
                sent_count += batch_total
            else:
                failed_count += batch_total
                errors.append(f"Batch {idx} ({batch_total} recipients): {error}")

        # Prepare response
        response_data = {
            "success": True,
            "message": f"Email forwarding completed in {len(batches)} batches",
            "total_recipients": total_recipients,
            "total_batches": len(batches),
            "sent_count": sent_count,
            "failed_count": failed_count,
            "batch_size": MAX_RECIPIENTS_LIMIT,
        }

        if errors:
            response_data["errors"] = errors

        response_message = (
            f"Email forwarded in {len(batches)} batches: "
            f"{sent_count} sent, {failed_count} failed"
        )

        return self._format_response(response_data)

    async def _handle_batch_send_new_email(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: Optional[List[str]],
        bcc_recipients: Optional[List[str]],
        importance: Optional[str],
        total_recipients: int,
    ) -> list[types.TextContent]:
        """Handle batch sending of new emails when recipients exceed the limit.

        Splits recipients into batches and sends multiple emails.

        Args:
            to_recipients: List of TO recipients
            subject: Email subject
            body: Email body
            cc_recipients: List of CC recipients
            bcc_recipients: List of BCC recipients
            importance: Email importance
            total_recipients: Total number of recipients

        Returns:
            Response with batch send results
        """
        from ..config import MAX_RECIPIENTS_LIMIT

        sent_count = 0
        failed_count = 0
        errors = []

        # Combine all recipients into a single list
        all_recipients = []
        if to_recipients:
            all_recipients.extend([("to", email) for email in to_recipients])
        if cc_recipients:
            all_recipients.extend([("cc", email) for email in cc_recipients])
        if bcc_recipients:
            all_recipients.extend([("bcc", email) for email in bcc_recipients])

        # Split into batches
        batches = []
        current_batch = {"to": [], "cc": [], "bcc": []}
        current_count = 0

        for recipient_type, email in all_recipients:
            if current_count >= MAX_RECIPIENTS_LIMIT:
                batches.append(current_batch)
                current_batch = {"to": [], "cc": [], "bcc": []}
                current_count = 0

            current_batch[recipient_type].append(email)
            current_count += 1

        if current_batch["to"] or current_batch["cc"] or current_batch["bcc"]:
            batches.append(current_batch)

        # Send each batch
        for idx, batch in enumerate(batches, 1):
            batch_to_count = len(batch["to"])
            batch_cc_count = len(batch["cc"])
            batch_bcc_count = len(batch["bcc"])
            batch_total = batch_to_count + batch_cc_count + batch_bcc_count

            success, result, error = await self._handle_auth_error(
                lambda b=batch: graph_client.send_email(
                    to_recipients=b["to"] if b["to"] else None,
                    subject=subject,
                    body=body,
                    cc_recipients=b["cc"] if b["cc"] else None,
                    bcc_recipients=b["bcc"] if b["bcc"] else None,
                    body_content_type="HTML",
                    importance=importance,
                ),
                f"sending email batch {idx}/{len(batches)}",
            )

            if success:
                sent_count += batch_total
            else:
                failed_count += batch_total
                errors.append(f"Batch {idx} ({batch_total} recipients): {error}")

        # Prepare response
        response_data = {
            "success": True,
            "message": f"Email sending completed in {len(batches)} batches",
            "total_recipients": total_recipients,
            "total_batches": len(batches),
            "sent_count": sent_count,
            "failed_count": failed_count,
            "batch_size": MAX_RECIPIENTS_LIMIT,
        }

        if errors:
            response_data["errors"] = errors

        response_message = (
            f"Email sent in {len(batches)} batches: "
            f"{sent_count} sent, {failed_count} failed"
        )

        return self._format_response(response_data)

    async def _handle_batch_reply_email(
        self,
        to_recipients: Optional[List[str]],
        subject: Optional[str],
        body: str,
        email_id: str,
        cc_recipients: Optional[List[str]],
        bcc_recipients: Optional[List[str]],
        importance: Optional[str],
        total_recipients: int,
    ) -> list[types.TextContent]:
        """Handle batch replying to emails when recipients exceed the limit.

        Splits recipients into batches and sends multiple reply emails.

        Args:
            to_recipients: List of TO recipients
            subject: Email subject
            body: Email body
            email_id: ID of the email to reply to
            cc_recipients: List of CC recipients
            bcc_recipients: List of BCC recipients
            importance: Email importance
            total_recipients: Total number of recipients

        Returns:
            Response with batch send results
        """
        from ..config import MAX_RECIPIENTS_LIMIT

        sent_count = 0
        failed_count = 0
        errors = []

        # Combine all recipients into a single list
        all_recipients = []
        if to_recipients:
            all_recipients.extend([("to", email) for email in to_recipients])
        if cc_recipients:
            all_recipients.extend([("cc", email) for email in cc_recipients])
        if bcc_recipients:
            all_recipients.extend([("bcc", email) for email in bcc_recipients])

        # Split into batches
        batches = []
        current_batch = {"to": [], "cc": [], "bcc": []}
        current_count = 0

        for recipient_type, email in all_recipients:
            if current_count >= MAX_RECIPIENTS_LIMIT:
                batches.append(current_batch)
                current_batch = {"to": [], "cc": [], "bcc": []}
                current_count = 0

            current_batch[recipient_type].append(email)
            current_count += 1

        if current_batch["to"] or current_batch["cc"] or current_batch["bcc"]:
            batches.append(current_batch)

        # Send each batch
        for idx, batch in enumerate(batches, 1):
            batch_to_count = len(batch["to"])
            batch_cc_count = len(batch["cc"])
            batch_bcc_count = len(batch["bcc"])
            batch_total = batch_to_count + batch_cc_count + batch_bcc_count

            success, result, error = await self._handle_auth_error(
                lambda b=batch: graph_client.send_email(
                    to_recipients=b["to"] if b["to"] else None,
                    subject=subject,
                    body=body,
                    cc_recipients=b["cc"] if b["cc"] else None,
                    bcc_recipients=b["bcc"] if b["bcc"] else None,
                    reply_to_message_id=email_id,
                    body_content_type="HTML",
                    importance=importance,
                ),
                f"sending reply batch {idx}/{len(batches)}",
            )

            if success:
                sent_count += batch_total
            else:
                failed_count += batch_total
                errors.append(f"Batch {idx} ({batch_total} recipients): {error}")

        # Prepare response
        response_data = {
            "success": True,
            "message": f"Email reply completed in {len(batches)} batches",
            "total_recipients": total_recipients,
            "total_batches": len(batches),
            "sent_count": sent_count,
            "failed_count": failed_count,
            "batch_size": MAX_RECIPIENTS_LIMIT,
        }

        if errors:
            response_data["errors"] = errors

        response_message = (
            f"Email reply sent in {len(batches)} batches: "
            f"{sent_count} sent, {failed_count} failed"
        )

        return self._format_response(response_data)

    async def handle_manage_mail_folder(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle manage_mail_folder tool with list, create, delete, rename, get_details, and move actions."""
        action = arguments.get("action")

        if action == "list":
            return await self._handle_list_folders()
        elif action == "create":
            return await self._handle_create_folder(arguments)
        elif action == "delete":
            return await self._handle_delete_folder(arguments)
        elif action == "rename":
            return await self._handle_rename_folder(arguments)
        elif action == "get_details":
            return await self._handle_get_folder_details(arguments)
        elif action == "move":
            return await self._handle_move_folder(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'list', 'create', 'delete', 'rename', 'get_details', or 'move'."
            )

    async def _handle_list_folders(self) -> list[types.TextContent]:
        """Handle list folders action."""
        folders = await graph_client.list_mail_folders()
        return self._format_response(
            {
                "message": f"Found {len(folders)} mail folders",
                "folders": folders,
                "count": len(folders),
            }
        )

    async def _handle_create_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle create folder action."""
        folder_name = arguments["folder_name"]
        parent_folder = arguments.get("parent_folder")

        result = await graph_client.create_folder(folder_name, parent_folder)

        folder_path = f"{parent_folder}/{folder_name}" if parent_folder else folder_name

        folder_info = {
            "path": folder_path,
            "displayName": result.get("displayName", folder_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder '{folder_name}' created successfully",
                "folder": folder_info,
            }
        )

    async def _handle_delete_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle delete folder action."""
        folder_path = arguments["folder_path"]

        result = await graph_client.delete_folder(folder_path)

        return self._format_response(result)

    async def _handle_rename_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle rename folder action."""
        folder_path = arguments["folder_path"]
        new_name = arguments["new_name"]

        result = await graph_client.rename_folder(folder_path, new_name)

        if "/" in folder_path:
            parent_path = folder_path.rsplit("/", 1)[0]
            new_path = f"{parent_path}/{new_name}"
        else:
            new_path = new_name

        folder_info = {
            "path": new_path,
            "displayName": result.get("displayName", new_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder renamed to '{new_name}' successfully",
                "folder": folder_info,
            }
        )

    async def _handle_get_folder_details(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle get folder details action."""
        folder_path = arguments["folder_path"]

        result = await graph_client.get_folder_details(folder_path)

        folder_info = {
            "path": folder_path,
            "displayName": result.get("displayName", folder_path),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {"message": f"Folder details for '{folder_path}'", "folder": folder_info}
        )

    async def _handle_move_folder(self, arguments: dict) -> list[types.TextContent]:
        """Handle move folder action."""
        folder_path = arguments["folder_path"]
        destination_parent = arguments["destination_parent"]

        result = await graph_client.move_folder(folder_path, destination_parent)

        folder_name = folder_path.split("/")[-1]
        new_path = f"{destination_parent}/{folder_name}"

        folder_info = {
            "path": new_path,
            "displayName": result.get("displayName", folder_name),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "childFolderCount": result.get("childFolderCount", 0),
        }

        return self._format_response(
            {
                "message": f"Folder moved to '{destination_parent}' successfully",
                "folder": folder_info,
            }
        )

    async def handle_manage_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_emails tool with multiple actions."""
        action = arguments.get("action")

        if action == "move_single":
            return await self._handle_move_single_email(arguments)
        elif action == "move_all":
            return await self._handle_move_all_emails(arguments)
        elif action == "delete_single":
            return await self._handle_delete_single_email(arguments)
        elif action == "delete_multiple":
            return await self._handle_delete_multiple_emails(arguments)
        elif action == "delete_all":
            return await self._handle_delete_all_emails(arguments)
        elif action == "archive_single":
            return await self._handle_archive_single_email(arguments)
        elif action == "archive_multiple":
            return await self._handle_archive_multiple_emails(arguments)
        elif action == "flag_single":
            return await self._handle_flag_single_email(arguments)
        elif action == "flag_multiple":
            return await self._handle_flag_multiple_emails(arguments)
        elif action == "categorize_single":
            return await self._handle_categorize_single_email(arguments)
        elif action == "categorize_multiple":
            return await self._handle_categorize_multiple_emails(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'move_single', 'move_all', 'delete_single', 'delete_multiple', 'delete_all', 'archive_single', 'archive_multiple', 'flag_single', 'flag_multiple', 'categorize_single', or 'categorize_multiple'."
            )

    async def _handle_delete_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete single email action."""
        cache_number = arguments["cache_number"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.delete_email(email_id)
            await email_cache.remove_email(email_id)
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg or "ErrorItemNotFound" in error_msg:
                await email_cache.remove_email(email_id)
                return self._format_error(
                    f"Error: Cache number {cache_number} no longer exists in the mailbox. It may have been deleted or moved. The cache has been updated."
                )
            else:
                return self._format_error(f"Error deleting email: {error_msg}")

    async def _handle_delete_multiple_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete multiple emails action."""
        cache_numbers = arguments["cache_numbers"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        email_ids = []
        invalid_numbers = []

        for cache_number in cache_numbers:
            if cache_number < 1 or cache_number > total_count:
                invalid_numbers.append(cache_number)
                continue

            email = cached_emails[cache_number - 1]
            email_id = email.get("id")

            if not email_id:
                invalid_numbers.append(cache_number)
                continue

            email_ids.append(email_id)

        if invalid_numbers:
            return self._format_error(
                f"Error: Invalid cache numbers: {invalid_numbers}. Please use valid numbers from browse_email_cache (1-{total_count})."
            )

        if not email_ids:
            return self._format_error(
                "Error: No valid email IDs found. Please check the cache and try again."
            )

        result = await graph_client.batch_delete_emails(email_ids)

        deleted_count = result.get("deleted_count", 0)
        failed_count = result.get("failed_count", 0)
        errors = result.get("errors", [])

        if failed_count > 0:
            for error in errors:
                if "404" in error or "ErrorItemNotFound" in error:
                    for email_id in email_ids:
                        await email_cache.remove_email(email_id)
            return self._format_error(
                f"Error: Some emails could not be deleted. Deleted: {deleted_count}, Failed: {failed_count}. Errors: {errors}"
            )

        for email_id in email_ids:
            await email_cache.remove_email(email_id)

        return self._format_response(result)

    async def _handle_delete_all_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle delete all emails from folder action."""
        source_folder = arguments["source_folder"]

        result = await graph_client.delete_all_emails_from_folder(source_folder)

        current_mode = email_cache.get_mode()
        if current_mode == "list":
            cached_folder = email_cache.get_list_state()["folder"]
            if cached_folder == source_folder:
                email_cache.invalidate_list_state()
        elif current_mode == "search":
            cached_folder = email_cache.get_search_state()["folder"]
            if cached_folder == source_folder:
                email_cache.invalidate_search_state()

        return self._format_response(result)

    async def _handle_move_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle move single email action."""
        cache_number = arguments["cache_number"]
        destination_folder = arguments["destination_folder"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        result = await graph_client.move_email_to_folder(email_id, destination_folder)

        return self._format_response(result)

    async def _handle_move_all_emails(self, arguments: dict) -> list[types.TextContent]:
        """Handle move all emails action."""
        source_folder = arguments["source_folder"]
        destination_folder = arguments["destination_folder"]

        result = await graph_client.move_all_emails_from_folder(
            source_folder, destination_folder
        )

        return self._format_response(result)

    async def _handle_archive_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle archive single email action."""
        cache_number = arguments["cache_number"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.archive_email(email_id)
            await email_cache.remove_email(email_id)
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg or "ErrorItemNotFound" in error_msg:
                await email_cache.remove_email(email_id)
                return self._format_error(
                    f"Error: Cache number {cache_number} no longer exists in the mailbox. It may have been deleted or moved. The cache has been updated."
                )
            else:
                return self._format_error(f"Error archiving email: {error_msg}")

    async def _handle_archive_multiple_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle archive multiple emails action."""
        cache_numbers = arguments["cache_numbers"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        email_ids = []
        invalid_numbers = []

        for cache_number in cache_numbers:
            if cache_number < 1 or cache_number > total_count:
                invalid_numbers.append(cache_number)
                continue

            email = cached_emails[cache_number - 1]
            email_id = email.get("id")

            if not email_id:
                invalid_numbers.append(cache_number)
                continue

            email_ids.append(email_id)

        if invalid_numbers:
            return self._format_error(
                f"Error: Invalid cache numbers: {invalid_numbers}. Please use valid numbers from browse_email_cache (1-{total_count})."
            )

        if not email_ids:
            return self._format_error(
                "Error: No valid email IDs found. Please check the cache and try again."
            )

        result = await graph_client.batch_archive_emails(email_ids)

        for email_id in email_ids:
            await email_cache.remove_email(email_id)

        return self._format_response(result)

    async def _handle_flag_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle flag single email action."""
        cache_number = arguments["cache_number"]
        flag_status = arguments["flag_status"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        result = await graph_client.flag_email(email_id, flag_status)

        return self._format_response(result)

    async def _handle_flag_multiple_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle flag multiple emails action."""
        cache_numbers = arguments["cache_numbers"]
        flag_status = arguments["flag_status"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        email_ids = []
        invalid_numbers = []

        for cache_number in cache_numbers:
            if cache_number < 1 or cache_number > total_count:
                invalid_numbers.append(cache_number)
                continue

            email = cached_emails[cache_number - 1]
            email_id = email.get("id")

            if not email_id:
                invalid_numbers.append(cache_number)
                continue

            email_ids.append(email_id)

        if invalid_numbers:
            return self._format_error(
                f"Error: Invalid cache numbers: {invalid_numbers}. Please use valid numbers from browse_email_cache (1-{total_count})."
            )

        if not email_ids:
            return self._format_error(
                "Error: No valid email IDs found. Please check the cache and try again."
            )

        result = await graph_client.batch_flag_emails(email_ids, flag_status)

        return self._format_response(result)

    async def _handle_categorize_single_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle categorize single email action."""
        cache_number = arguments["cache_number"]
        categories = arguments["categories"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        result = await graph_client.categorize_email(email_id, categories)

        return self._format_response(result)

    async def _handle_categorize_multiple_emails(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle categorize multiple emails action."""
        cache_numbers = arguments["cache_numbers"]
        categories = arguments["categories"]

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        email_ids = []
        invalid_numbers = []

        for cache_number in cache_numbers:
            if cache_number < 1 or cache_number > total_count:
                invalid_numbers.append(cache_number)
                continue

            email = cached_emails[cache_number - 1]
            email_id = email.get("id")

            if not email_id:
                invalid_numbers.append(cache_number)
                continue

            email_ids.append(email_id)

        if invalid_numbers:
            return self._format_error(
                f"Error: Invalid cache numbers: {invalid_numbers}. Please use valid numbers from browse_email_cache (1-{total_count})."
            )

        if not email_ids:
            return self._format_error(
                "Error: No valid email IDs found. Please check the cache and try again."
            )

        result = await graph_client.batch_categorize_emails(email_ids, categories)

        return self._format_response(result)

    async def handle_manage_templates(self, arguments: dict) -> list[types.TextContent]:
        """Handle manage_templates tool with multiple actions."""
        action = arguments.get("action")

        if action == "create_from_email":
            return await self._handle_create_template_from_email(arguments)
        elif action == "list":
            return await self._handle_list_templates(arguments)
        elif action == "get":
            return await self._handle_get_template(arguments)
        elif action == "update":
            return await self._handle_update_template(arguments)
        elif action == "delete":
            return await self._handle_delete_template(arguments)
        elif action == "send":
            return await self._handle_send_template(arguments)
        else:
            return self._format_error(
                f"Invalid action: {action}. Must be 'create_from_email', 'list', 'get', 'update', 'delete', or 'send'."
            )

    async def _handle_create_template_from_email(
        self, arguments: dict
    ) -> list[types.TextContent]:
        """Handle create template from email action."""
        cache_number = arguments["cache_number"]
        template_name = arguments.get("template_name")

        cached_emails = email_cache.get_cached_emails()
        total_count = len(cached_emails)

        if total_count == 0:
            return self._format_error(
                "Error: No emails in cache. Use load_emails_by_folder or search_emails to load emails first."
            )

        if cache_number < 1 or cache_number > total_count:
            return self._format_error(
                f"Error: Invalid cache number: {cache_number}. Please use valid number from browse_email_cache (1-{total_count})."
            )

        email = cached_emails[cache_number - 1]
        email_id = email.get("id")

        if not email_id:
            return self._format_error(
                "Error: No valid email ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.create_template_from_email(
                email_id, template_name
            )
            await template_cache.add_template(
                {
                    "id": result["id"],
                    "subject": result["subject"],
                    "folder": result["folder"],
                }
            )
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg or "ErrorItemNotFound" in error_msg:
                return self._format_error(
                    f"Error: Cache number {cache_number} no longer exists in the mailbox. It may have been deleted or moved."
                )
            else:
                return self._format_error(f"Error creating template: {error_msg}")

    async def _handle_list_templates(self, arguments: dict) -> list[types.TextContent]:
        """Handle list templates action."""
        try:
            templates = await graph_client.list_templates()

            await template_cache.clear_cache()
            for template in templates:
                await template_cache.add_template(template)

            return self._format_response(
                {
                    "templates": templates,
                    "count": len(templates),
                }
            )
        except Exception as e:
            return self._format_error(f"Error listing templates: {str(e)}")

    async def _handle_get_template(self, arguments: dict) -> list[types.TextContent]:
        """Handle get template action."""
        template_number = arguments["template_number"]
        return_html = arguments.get("return_html", False)

        template = template_cache.get_template_by_number(template_number)

        if not template:
            return self._format_error(
                f"Error: Invalid template number: {template_number}. Please use valid number from list_templates."
            )

        template_id = template.get("id")

        if not template_id:
            return self._format_error(
                "Error: No valid template ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.get_template(template_id, return_html)
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg or "ErrorItemNotFound" in error_msg:
                await template_cache.remove_template(template_id)
                return self._format_error(
                    f"Error: Template number {template_number} no longer exists. It may have been deleted. The cache has been updated."
                )
            else:
                return self._format_error(f"Error getting template: {error_msg}")

    async def _handle_update_template(self, arguments: dict) -> list[types.TextContent]:
        """Handle update template action."""
        template_number = arguments["template_number"]
        subject = arguments.get("subject")
        htmlbody = normalize_email_html(arguments.get("htmlbody"))
        to = arguments.get("to")
        cc = arguments.get("cc")
        bcc = arguments.get("bcc")

        template = template_cache.get_template_by_number(template_number)

        if not template:
            return self._format_error(
                f"Error: Invalid template number: {template_number}. Please use valid number from list_templates."
            )

        template_id = template.get("id")

        if not template_id:
            return self._format_error(
                "Error: No valid template ID found. Please check the cache and try again."
            )

        try:
            to_recipients = [{"email": email} for email in to] if to else None
            cc_recipients = [{"email": email} for email in cc] if cc else None
            bcc_recipients = [{"email": email} for email in bcc] if bcc else None

            result = await graph_client.update_template(
                template_id,
                subject=subject,
                body=htmlbody,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
            )

            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Error updating template: {str(e)}")

    async def _handle_delete_template(self, arguments: dict) -> list[types.TextContent]:
        """Handle delete template action."""
        template_number = arguments["template_number"]

        template = template_cache.get_template_by_number(template_number)

        if not template:
            return self._format_error(
                f"Error: Invalid template number: {template_number}. Please use valid number from list_templates."
            )

        template_id = template.get("id")

        if not template_id:
            return self._format_error(
                "Error: No valid template ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.delete_template(template_id)
            await template_cache.remove_template(template_id)
            return self._format_response(result)
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg or "ErrorItemNotFound" in error_msg:
                await template_cache.remove_template(template_id)
                return self._format_error(
                    f"Error: Template number {template_number} no longer exists. It may have been deleted. The cache has been updated."
                )
            else:
                return self._format_error(f"Error deleting template: {error_msg}")

    async def _handle_send_template(self, arguments: dict) -> list[types.TextContent]:
        """Handle send template action."""
        template_number = arguments["template_number"]
        to = arguments.get("to")
        cc = arguments.get("cc")
        bcc = arguments.get("bcc")

        template = template_cache.get_template_by_number(template_number)

        if not template:
            return self._format_error(
                f"Error: Invalid template number: {template_number}. Please use valid number from list_templates."
            )

        template_id = template.get("id")

        if not template_id:
            return self._format_error(
                "Error: No valid template ID found. Please check the cache and try again."
            )

        try:
            result = await graph_client.send_template(template_id, to, cc, bcc)
            return self._format_response(result)
        except Exception as e:
            return self._format_error(f"Error sending template: {str(e)}")

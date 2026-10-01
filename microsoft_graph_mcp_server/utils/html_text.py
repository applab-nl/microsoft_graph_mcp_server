"""HTML -> plain text that keeps paragraph and line breaks."""

import html as _html
import re

_BLOCK_BREAK = re.compile(r"(?i)<\s*(br|/p|/div|/li|/h[1-6]|/tr)\s*/?>")
_DROP_BLOCKS = re.compile(r"(?is)<(head|style|script)[^>]*>.*?</\1>")
_TAGS = re.compile(r"<[^>]+>")


def html_to_text(content: str | None) -> str:
    """HTML -> plain text, keeping paragraph/line breaks (unlike the email helper)."""
    if not content:
        return ""
    s = _DROP_BLOCKS.sub("", content)
    s = _BLOCK_BREAK.sub("\n", s)
    s = _TAGS.sub("", s)
    s = _html.unescape(s).replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n")]
    return "\n".join(ln for ln in lines if ln)

"""
Claude API integration for PDF bill analysis
"""

import base64
import os

import anthropic
from anthropic.types import TextBlock

from .config import CLAUDE_MAX_TOKENS, CLAUDE_MODEL, EXTRACTION_PROMPT

client: anthropic.Anthropic = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY")
)


def analyze_bill_pdf(pdf_path: str) -> str | None:
    """Analyze a bill PDF using Claude AI to extract structured data.
    Returns the raw response text from Claude or None.
    """
    with open(pdf_path, "rb") as pdf_file:
        pdf_data: str = base64.standard_b64encode(pdf_file.read()).decode("utf-8")

    message: anthropic.types.Message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=CLAUDE_MAX_TOKENS,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_data,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )

    content_block = message.content[0]

    if not isinstance(content_block, TextBlock):
        return None

    if content_block.text == "error":
        return None

    return content_block.text

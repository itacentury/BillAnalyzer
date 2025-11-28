"""
Claude API integration for PDF bill analysis
"""

import base64
import os

import anthropic
from anthropic.types import TextBlock

from .config import CLAUDE_MAX_TOKENS, CLAUDE_MODEL, EXTRACTION_PROMPT

# Initialize Anthropic API client
client: anthropic.Anthropic = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY")
)


def analyze_bill_pdf(pdf_path: str) -> str:
    """Analyze a bill PDF using Claude AI to extract structured data.

    :param pdf_path: Path to the PDF file
    :type pdf_path: str
    :return: Raw response text from Claude (contains JSON in markdown format)
    :rtype: str
    :raises FileNotFoundError: If PDF file doesn't exist
    :raises anthropic.APIError: If Claude API call fails
    :raises TypeError: If Claude returns unexpected content type
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

    # Extract text from response, ensuring it's a TextBlock
    content_block = message.content[0]
    if isinstance(content_block, TextBlock):
        return content_block.text

    # Fallback for unexpected content types
    raise TypeError(f"Expected TextBlock, got {type(content_block).__name__}")

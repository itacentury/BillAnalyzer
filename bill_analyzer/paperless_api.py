"""
Paperless-ngx REST API integration
"""

import json
from typing import Any

import requests


def upload_to_paperless(  # pylint: disable=too-many-arguments,too-many-locals
    pdf_path: str,
    token: str,
    paperless_url: str,
    *,
    title: str | None = None,
    created: str | None = None,
    correspondent: int | None = None,
    document_type: int | None = None,
    storage_path: int | None = None,
    tags: list[int] | None = None,
    archive_serial_number: int | None = None,
    custom_fields: dict[int, Any] | list[int] | None = None,
) -> str:
    """Upload a PDF document to Paperless-ngx via REST API.

    :param pdf_path: Path to the PDF file to upload
    :type pdf_path: str
    :param token: Paperless-ngx API token for authentication
    :type token: str
    :param paperless_url: Base URL of the Paperless-ngx instance (e.g., "http://localhost:8000")
    :type paperless_url: str
    :param title: Optional title for the document
    :type title: str | None
    :param created: Optional creation date/time (e.g., "2016-04-19" or "2016-04-19 06:15:00+02:00")
    :type created: str | None
    :param correspondent: Optional correspondent ID
    :type correspondent: int | None
    :param document_type: Optional document type ID
    :type document_type: int | None
    :param storage_path: Optional storage path ID
    :type storage_path: int | None
    :param tags: Optional list of tag IDs
    :type tags: list[int] | None
    :param archive_serial_number: Optional archive serial number
    :type archive_serial_number: int | None
    :param custom_fields: Optional custom field assignments (dict mapping field id -> value,
        or list of field ids)
    :type custom_fields: dict[int, Any] | list[int] | None
    :return: UUID of the consumption task that will process the document
    :rtype: str
    :raises FileNotFoundError: If the PDF file doesn't exist
    :raises requests.HTTPError: If the API request fails
    :raises requests.RequestException: If network error occurs
    """
    # Construct the API endpoint
    endpoint: str = f"{paperless_url.rstrip('/')}/api/documents/post_document/"

    # Set up authentication header
    headers: dict[str, str] = {"Authorization": f"Token {token}"}

    # Open and upload the PDF file
    with open(pdf_path, "rb") as pdf_file:
        files: dict[str, Any] = {"document": pdf_file}

        # Build optional form data
        data: dict[str, Any] = {}

        if title is not None:
            data["title"] = title

        if created is not None:
            data["created"] = created

        if correspondent is not None:
            data["correspondent"] = correspondent

        if document_type is not None:
            data["document_type"] = document_type

        if storage_path is not None:
            data["storage_path"] = storage_path

        if tags is not None:
            # Tags need to be sent as comma-separated string
            data["tags"] = ",".join(str(tag_id) for tag_id in tags)

        if archive_serial_number is not None:
            data["archive_serial_number"] = archive_serial_number

        if custom_fields is not None:
            # Custom fields must be sent as JSON string
            data["custom_fields"] = json.dumps(custom_fields)

        # Make the POST request
        response: requests.Response = requests.post(
            endpoint, headers=headers, files=files, data=data, timeout=30
        )

        # Raise exception for HTTP errors
        response.raise_for_status()

        # Return the task UUID (API returns JSON string with the UUID)
        task_uuid: str | dict[str, Any] = response.json()

        # Handle both string UUID and dict responses
        if isinstance(task_uuid, str):
            return task_uuid

        # Fallback: if it's a dict, try to extract task_id or return as string
        return str(task_uuid.get("task_id", task_uuid))

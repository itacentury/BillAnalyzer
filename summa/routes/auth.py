"""Authentication endpoints: login, logout and session status.

These routes stay reachable without a session (see the public allowlist in
:mod:`summa.auth`) — otherwise there would be no way to obtain one.
"""

from typing import Any

from flask import Blueprint, jsonify, request

from summa import config
from summa.auth import end_session, is_authenticated, start_session, verify_password
from summa.helpers import ApiResponse, error_response

auth_bp: Blueprint = Blueprint("auth", __name__, url_prefix="/api/auth")


@auth_bp.route("/login", methods=["POST"])
def login() -> ApiResponse:
    """Verify the password and start a session on success."""
    payload: Any = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return error_response("Request body must be a JSON object", 400)

    password: Any = payload.get("password")
    if not isinstance(password, str):
        return error_response("Field 'password' must be a string", 400)

    if not verify_password(password):
        return error_response("Invalid password", 401)

    start_session(remember=payload.get("remember") is True)
    return jsonify({"success": True, "authed": True})


@auth_bp.route("/logout", methods=["POST"])
def logout() -> ApiResponse:
    """Clear the session cookie. Always succeeds."""
    end_session()
    return jsonify({"success": True, "authed": False})


@auth_bp.route("/me", methods=["GET"])
def me() -> ApiResponse:
    """Report authentication status for the client-side login gate.

    Answers 200 in every case (never 401) so the gate can ask on each page load
    without producing console errors. ``enabled`` tells the client whether the
    password feature exists at all, which is what hides the sign-out control on
    an unauthenticated deployment.
    """
    enabled: bool = config.auth_enabled()
    return jsonify({"authed": not enabled or is_authenticated(), "enabled": enabled})

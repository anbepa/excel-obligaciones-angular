"""Session management using FastMCP Context for authentication."""

import inspect
import json
from collections.abc import Callable
from typing import TypeVar

from fastmcp import Context

from .errors import ToolError
from .models import UserState


class SessionManager:
    """Manages user sessions, mapping user_id to UserState."""

    def __init__(self):
        self._sessions: dict[str, UserState] = {}

    def get_session(self, user_id: str) -> UserState:
        """Get or create a session for the given user_id."""
        if user_id not in self._sessions:
            self._sessions[user_id] = UserState(user_id=user_id)
        return self._sessions[user_id]

    def remove_session(self, user_id: str) -> bool:
        """Remove a user session and close the workbook. Returns True if session existed."""
        if user_id in self._sessions:
            state = self._sessions[user_id]
            if state.workbook is not None:
                state.workbook.close()
            del self._sessions[user_id]
            return True
        return False

    def has_session(self, user_id: str) -> bool:
        """Check if a session exists for the given user_id."""
        return user_id in self._sessions


# Global session manager instance
sessions = SessionManager()


def _get_user_id_from_context(ctx: Context) -> str | None:
    """Extract user_id from FastMCP Context via HTTP headers."""
    try:
        request = ctx.get_http_request()
        return request.headers.get("X-User-ID")
    except Exception:
        return None


def _copy_wrapper_metadata(fn: Callable, wrapper: Callable) -> None:
    """
    Copy function metadata and build wrapper signature for session decorators.

    Removes the first parameter (wb/state) and adds ctx: Context as the first param.
    """
    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    wrapper.__module__ = fn.__module__
    wrapper.__qualname__ = fn.__qualname__

    sig = inspect.signature(fn)
    params = list(sig.parameters.values())

    ctx_param = inspect.Parameter(
        "ctx", inspect.Parameter.POSITIONAL_OR_KEYWORD, annotation=Context
    )

    new_params = [ctx_param] + params[1:]
    wrapper.__signature__ = sig.replace(parameters=new_params)

    wrapper.__annotations__ = {"ctx": Context}
    for p in params[1:]:
        if p.annotation is not inspect.Parameter.empty:
            wrapper.__annotations__[p.name] = p.annotation
    if sig.return_annotation is not inspect.Parameter.empty:
        wrapper.__annotations__["return"] = sig.return_annotation


def _is_successful_result(result) -> bool:
    """Check if a tool result indicates success via JSON {"status": "success"}."""
    if not isinstance(result, str):
        return False
    try:
        return json.loads(result).get("status") == "success"
    except (json.JSONDecodeError, AttributeError):
        return False


T = TypeVar("T")


def _workbook_decorator(
    marks_dirty: bool = False,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """
    Factory for workbook injection decorators.

    Args:
        marks_dirty: If True, marks session as dirty on successful operations.
    """

    def decorator(fn: Callable[..., T]) -> Callable[..., T]:
        def wrapper(ctx: Context, *args, **kwargs) -> T:
            user_id = _get_user_id_from_context(ctx)
            if user_id is None:
                return ToolError(
                    "Not authenticated - missing X-User-ID header", code="AUTH_ERROR"
                ).to_json()

            state = sessions.get_session(user_id)
            if state.workbook is None:
                return ToolError(
                    "No workbook loaded. Call init_session first.", code="NO_WORKBOOK"
                ).to_json()

            result = fn(state.workbook, *args, **kwargs)

            if marks_dirty:  # and _is_successful_result(result):
                state.workbook.is_dirty = True

            return result

        _copy_wrapper_metadata(fn, wrapper)
        return wrapper

    return decorator


def with_workbook(fn: Callable[..., T]) -> Callable[..., T]:
    """
    Decorator that injects the ManagedWorkbook from the authenticated session.

    The decorated function should have `mwb: ManagedWorkbook` as its first parameter.
    A `ctx: Context` parameter is added for FastMCP to inject.
    Both parameters are HIDDEN from the MCP schema.

    Returns error message string if not authenticated or no workbook loaded.
    """
    return _workbook_decorator(marks_dirty=False)(fn)


def with_workbook_mutation(fn: Callable[..., T]) -> Callable[..., T]:
    """
    Decorator for mutation tools that modify the workbook.

    Same as with_workbook, but also marks the session as dirty after
    a successful operation (detects JSON {"status": "success"}).
    """
    return _workbook_decorator(marks_dirty=True)(fn)


def with_session(fn: Callable[..., T]) -> Callable[..., T]:
    """
    Decorator that injects the session state from the authenticated context.

    The decorated function should have `state: UserState` as its first parameter.
    A `ctx: Context` parameter is added for FastMCP to inject.
    Both are HIDDEN from the MCP schema (ctx is handled by FastMCP, state by us).
    """

    def wrapper(ctx: Context, *args, **kwargs) -> T:
        user_id = _get_user_id_from_context(ctx)
        if user_id is None:
            return ToolError(
                "Not authenticated - missing X-User-ID header", code="AUTH_ERROR"
            ).to_json()

        state = sessions.get_session(user_id)
        return fn(state, *args, **kwargs)

    _copy_wrapper_metadata(fn, wrapper)
    return wrapper

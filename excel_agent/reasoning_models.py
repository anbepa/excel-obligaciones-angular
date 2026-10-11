"""Custom model wrappers for handling provider-specific features like reasoning tokens."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Literal, overload

from pydantic_ai.messages import ModelMessage, ModelResponse, TextPart
from pydantic_ai.models import Model, ModelRequestParameters
from pydantic_ai.models.openai import OpenAIChatModel, OpenAIChatModelSettings
from pydantic_ai.settings import ModelSettings
from pydantic_ai.exceptions import UnexpectedModelBehavior

try:
    from openai import AsyncStream, NOT_GIVEN, APIStatusError
    from openai.types import chat
    from openai.types.chat import ChatCompletionChunk
except ImportError:
    raise ImportError("openai package is required for OpenRouter model support")

logger = logging.getLogger(__name__)


@dataclass(init=False)
class OpenRouterReasoningModel(OpenAIChatModel):
    """
    OpenRouter-compatible model that properly handles reasoning tokens.

    This model intercepts API calls to:
    1. Capture `reasoning_details` from API responses
    2. Inject them back into subsequent requests

    This is required for models like Gemini 3 Pro that use thought signatures
    with tool calls.

    Usage:
        model = OpenRouterReasoningModel('google/gemini-3-pro-preview')
        agent = Agent(model, ...)

    See: https://openrouter.ai/docs/guides/best-practices/reasoning-tokens
    """

    _reasoning_details_by_turn: list[list[dict[str, Any]]] = field(
        default_factory=list, repr=False
    )
    """Cache of reasoning_details indexed by assistant turn number."""

    def __init__(
        self,
        model_name: str,
        *,
        profile=None,
        settings: ModelSettings | None = None,
    ):
        """Initialize an OpenRouter model with reasoning token support.

        Args:
            model_name: The name of the model (e.g., 'google/gemini-3-pro-preview')
            profile: Optional model profile
            settings: Optional default model settings
        """
        # Initialize parent with openrouter provider
        super().__init__(
            model_name, provider="openrouter", profile=profile, settings=settings
        )
        self._reasoning_details_by_turn = []

    def clear_conversation(self):
        """Clear the reasoning cache when starting a new conversation."""
        self._reasoning_details_by_turn = []

    @overload
    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: Literal[True],
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> AsyncStream[ChatCompletionChunk]: ...

    @overload
    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: Literal[False],
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> chat.ChatCompletion: ...

    async def _completions_create(
        self,
        messages: list[ModelMessage],
        stream: bool,
        model_settings: OpenAIChatModelSettings,
        model_request_parameters: ModelRequestParameters,
    ) -> chat.ChatCompletion | AsyncStream[ChatCompletionChunk]:
        """Override to inject reasoning_details into messages."""

        # First, let the parent build the OpenAI messages
        openai_messages = await self._map_messages(messages)

        # Inject reasoning_details into assistant messages
        openai_messages = self._inject_reasoning_details(openai_messages)

        # Now make the actual API call using the parent's logic
        # We need to replicate the parent's _completions_create but with our modified messages
        tools = self._get_tools(model_request_parameters)
        web_search_options = self._get_web_search_options(model_request_parameters)

        from pydantic_ai.profiles.openai import OpenAIModelProfile

        if not tools:
            tool_choice = None
        elif (
            not model_request_parameters.allow_text_output
            and OpenAIModelProfile.from_profile(
                self.profile
            ).openai_supports_tool_choice_required
        ):
            tool_choice = "required"
        else:
            tool_choice = "auto"

        response_format = None
        if model_request_parameters.output_mode == "native":
            output_object = model_request_parameters.output_object
            assert output_object is not None
            response_format = self._map_json_schema(output_object)
        elif (
            model_request_parameters.output_mode == "prompted"
            and self.profile.supports_json_object_output
        ):
            response_format = {"type": "json_object"}

        unsupported_model_settings = OpenAIModelProfile.from_profile(
            self.profile
        ).openai_unsupported_model_settings
        for setting in unsupported_model_settings:
            model_settings.pop(setting, None)  # type: ignore

        from pydantic_ai.models import get_user_agent
        from pydantic_ai import ModelHTTPError

        try:
            extra_headers = model_settings.get("extra_headers", {})
            extra_headers.setdefault("User-Agent", get_user_agent())

            response = await self.client.chat.completions.create(
                model=self._model_name,
                messages=openai_messages,
                parallel_tool_calls=model_settings.get(
                    "parallel_tool_calls", NOT_GIVEN
                ),
                tools=tools or NOT_GIVEN,
                tool_choice=tool_choice or NOT_GIVEN,
                stream=stream,
                stream_options={"include_usage": True} if stream else NOT_GIVEN,
                stop=model_settings.get("stop_sequences", NOT_GIVEN),
                max_completion_tokens=model_settings.get("max_tokens", NOT_GIVEN),
                timeout=model_settings.get("timeout", NOT_GIVEN),
                response_format=response_format or NOT_GIVEN,
                seed=model_settings.get("seed", NOT_GIVEN),
                reasoning_effort=model_settings.get(
                    "openai_reasoning_effort", NOT_GIVEN
                ),
                user=model_settings.get("openai_user", NOT_GIVEN),
                web_search_options=web_search_options or NOT_GIVEN,
                service_tier=model_settings.get("openai_service_tier", NOT_GIVEN),
                prediction=model_settings.get("openai_prediction", NOT_GIVEN),
                temperature=model_settings.get("temperature", NOT_GIVEN),
                top_p=model_settings.get("top_p", NOT_GIVEN),
                presence_penalty=model_settings.get("presence_penalty", NOT_GIVEN),
                frequency_penalty=model_settings.get("frequency_penalty", NOT_GIVEN),
                logit_bias=model_settings.get("logit_bias", NOT_GIVEN),
                logprobs=model_settings.get("openai_logprobs", NOT_GIVEN),
                top_logprobs=model_settings.get("openai_top_logprobs", NOT_GIVEN),
                extra_headers=extra_headers,
                extra_body=model_settings.get("extra_body"),
            )

            # Capture reasoning_details from non-streaming response
            if not stream and isinstance(response, chat.ChatCompletion):
                self._capture_reasoning_details(response)

            return response

        except APIStatusError as e:
            if (status_code := e.status_code) >= 400:
                raise ModelHTTPError(
                    status_code=status_code, model_name=self.model_name, body=e.body
                ) from e
            raise

    def _inject_reasoning_details(self, messages: list[dict]) -> list[dict]:
        """Inject stored reasoning_details into assistant messages."""
        if not self._reasoning_details_by_turn:
            return messages

        modified_messages = []
        assistant_turn = 0

        for msg in messages:
            if msg.get("role") == "assistant":
                if assistant_turn < len(self._reasoning_details_by_turn):
                    reasoning_details = self._reasoning_details_by_turn[assistant_turn]
                    if reasoning_details:
                        msg = dict(msg)
                        msg["reasoning_details"] = reasoning_details
                assistant_turn += 1
            modified_messages.append(msg)

        return modified_messages

    def _capture_reasoning_details(self, response: chat.ChatCompletion) -> None:
        """Capture reasoning_details from the response for future turns."""
        if not response.choices:
            return

        choice = response.choices[0]
        message = choice.message

        # OpenRouter adds reasoning_details as an extra field
        reasoning_details = getattr(message, "reasoning_details", None)
        if reasoning_details:
            self._reasoning_details_by_turn.append(reasoning_details)
        else:
            # Track the turn even without reasoning_details
            self._reasoning_details_by_turn.append([])

    async def request(
        self,
        messages: list[ModelMessage],
        model_settings: ModelSettings | None,
        model_request_parameters: ModelRequestParameters,
    ) -> ModelResponse:
        """
        Override request to catch UnexpectedModelBehavior and convert it to feedback.

        This allows the agent to see the error and potentially adjust, rather than
        crashing entirely. Similar to how retry_prompt_to_user_message handles
        schema validation errors.
        """
        try:
            return await super().request(
                messages, model_settings, model_request_parameters
            )
        except UnexpectedModelBehavior as e:
            # Log the error for debugging
            logger.warning(f"Model returned unexpected response: {e}")

            # Return a ModelResponse with error text so the agent can see and adjust
            error_message = (
                f"The model API returned an invalid response that could not be processed.\n\n"
                f"Error details: {e}\n\n"
                f"Please try a different approach or simplify your request."
            )
            return ModelResponse(
                parts=[TextPart(content=error_message)],
                model_name=self.model_name,
                timestamp=datetime.now(timezone.utc),
            )


def create_openrouter_model(model_name: str, **kwargs) -> OpenRouterReasoningModel:
    """
    Create an OpenRouter model with reasoning token support.

    This should be used for OpenRouter reasoning models like Gemini
    that require thought signature preservation.

    Args:
        model_name: Model name without the 'openrouter:' prefix
                   (e.g., 'google/gemini-3-pro-preview')
        **kwargs: Additional arguments passed to the model

    Returns:
        OpenRouterReasoningModel instance

    Example:
        model = create_openrouter_model('google/gemini-3-pro-preview')
        agent = Agent(model, system_prompt="...")
    """
    return OpenRouterReasoningModel(model_name, **kwargs)


# List of models known to require reasoning token preservation
REASONING_MODELS = [
    "google/gemini-3-pro-preview",
]


def needs_reasoning_support(model_string: str) -> bool:
    """Check if a model needs reasoning token support."""
    if not model_string.startswith("openrouter:"):
        return False

    model_name = model_string.split(":", 1)[1]

    # Check against known reasoning models
    for reasoning_model in REASONING_MODELS:
        if model_name.startswith(reasoning_model.split("/")[0] + "/gemini"):
            # All Gemini models through OpenRouter need this
            return True
        if model_name == reasoning_model:
            return True

    return False


def get_model_for_config(model_string: str, **kwargs) -> Model | str:
    """
    Get the appropriate model for a configuration string.

    For OpenRouter reasoning models, returns an OpenRouterReasoningModel.
    For other models, returns the string to let pydantic-ai handle it.

    Args:
        model_string: Model string (e.g., 'openrouter:google/gemini-3-pro-preview')
        **kwargs: Additional arguments for model creation

    Returns:
        Either a Model instance or the original string
    """
    if needs_reasoning_support(model_string):
        model_name = model_string.split(":", 1)[1]
        return create_openrouter_model(model_name, **kwargs)

    return model_string

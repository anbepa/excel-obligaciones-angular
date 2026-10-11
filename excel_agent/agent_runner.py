"""Agent runner for Excel Agent."""

import asyncio
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Optional

from pydantic_ai import Agent, messages as ai_messages
from pydantic_ai.mcp import MCPServerSSE

from excel_mcp.excel_server import init_session
from excel_mcp.helpers import save_workbook_async
from .reasoning_models import get_model_for_config
from .config import ExperimentConfig, TraceCollector


def retry_prompt_to_user_message(
    messages: list[ai_messages.ModelMessage],
) -> list[ai_messages.ModelMessage]:
    """
    Convert RetryPrompt parts (triggered by schema validation failures, missing tool calls, etc.)
    into plain user messages so that the model can see the validation error details and self-correct.
    """
    transformed: list[ai_messages.ModelMessage] = []

    for message in messages:
        if isinstance(message, ai_messages.ModelRequest):
            new_parts: list[ai_messages.ModelRequestPart] = []
            for part in message.parts:
                if isinstance(part, ai_messages.RetryPromptPart):
                    error_text = part.model_response()
                    tool_label = (
                        f"Tool '{part.tool_name}'" if part.tool_name else "Tool call"
                    )
                    user_message = (
                        f"{tool_label} failed validation.\n\n"
                        f"{error_text}\n\n"
                        "Please fix the tool arguments and issue another tool call."
                    )
                    new_parts.append(ai_messages.UserPromptPart(content=user_message))
                else:
                    new_parts.append(part)

            transformed.append(replace(message, parts=new_parts))
        else:
            transformed.append(message)

    return transformed


@dataclass
class TaskInput:
    """Input for an Excel manipulation task."""

    instruction: str
    input_file: str
    output_file: str


class ExcelAgentRunner:
    """
    Runs the Excel agent and manages agent configuration.
    """

    def __init__(
        self,
        config: Optional[ExperimentConfig] = None,
        trace_collector: Optional[TraceCollector] = None,
        mcp_server_url: str = None,
    ):
        self.config = config or ExperimentConfig()
        self.trace_collector = trace_collector or TraceCollector()
        self.mcp_server_url = mcp_server_url
        assert self.mcp_server_url is not None

    async def run_excel_agent(self, inputs: TaskInput) -> str:
        """Run the Excel agent on the given task."""
        # Create a unique session ID for this evaluation
        session_id = f"eval_{Path(inputs.output_file).stem}"
        prompt = self.config.format_prompt(inputs.instruction, "input.xlsx")

        # Initialize session
        init_result = init_session(session_id, inputs.output_file)
        if "Error" in init_result:
            return f"Session init failed: {init_result}"

        # Create client connection for the agent to use
        client = MCPServerSSE(
            url=self.mcp_server_url,
            headers={"X-User-ID": session_id},
            timeout=self.config.timeout_tool,
        )

        try:
            async with asyncio.timeout(self.config.timeout):
                async with client:
                    # Use reasoning-aware model for OpenRouter models that need it
                    # This handles preserving reasoning_details/thought_signatures
                    model = get_model_for_config(self.config.model)

                    agent = Agent(
                        model,
                        system_prompt=self.config.system_prompt,
                        toolsets=[client],
                        history_processors=[retry_prompt_to_user_message],
                    )

                    result = await agent.run(
                        prompt,
                        model_settings={"timeout": self.config.timeout_request},
                    )

                    await save_workbook_async(session_id)
                    return result.output

        except asyncio.TimeoutError:
            return f"error: Agent run timed out after {self.config.timeout}s"

        except Exception as e:
            return f"error: {e}"

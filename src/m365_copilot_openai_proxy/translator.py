from __future__ import annotations

import json
from typing import Iterable

from .models import (
    AnthropicMessagesRequest,
    ContentPart,
    OpenAIChatRequest,
    ToolCall,
    TranslatedRequest,
)


def flatten_content(content: str | list[ContentPart] | None) -> str:
    if content is None:
        return ""
    if isinstance(content, str):
        return content
    return "".join(part.text or "" for part in content if part.type == "text")


def _join_lines(lines: Iterable[str]) -> str:
    return "\n".join(line for line in lines if line).strip()


def _format_tools_prompt(tools) -> str | None:
    """Format tool definitions into a system-level prompt so the model knows about available tools."""
    if not tools:
        return None
    tool_descriptions = []
    for tool in tools:
        func = tool.function
        desc = f"- {func.name}: {func.description or 'No description'}"
        if func.parameters:
            props = func.parameters.get("properties", {})
            required = func.parameters.get("required", [])
            param_parts = []
            for pname, pdef in props.items():
                ptype = pdef.get("type", "any")
                pdesc = pdef.get("description", "")
                req_flag = " (required)" if pname in required else ""
                param_parts.append(f"    - {pname}: {ptype}{req_flag} — {pdesc}")
            if param_parts:
                desc += "\n  Parameters:\n" + "\n".join(param_parts)
        tool_descriptions.append(desc)
    return (
        "IMPORTANT: You have access to tools that can read, write, and edit files on the user's local machine. "
        "You MUST use these tools to perform file operations — do NOT just describe what you would do or say you cannot access files. "
        "The tools are REAL and FUNCTIONAL. You CAN and SHOULD use them directly.\n\n"
        "To call a tool, output a JSON code block with the EXACT format:\n"
        "```tool_call\n"
        '{"name": "<tool_name>", "arguments": {<key-value pairs>}}\n'
        "```\n\n"
        "CRITICAL RULES:\n"
        "- When asked to read a file, you MUST call the Read tool — do NOT say you cannot access it.\n"
        "- When asked to create/write a file, you MUST call the Write tool — do NOT just output the content in a code block.\n"
        "- When asked to edit a file, you MUST call the Edit tool — do NOT just show the changes.\n"
        "- ALWAYS call tools for file operations. NEVER claim you 'cannot access local files' or 'are in a sandbox'.\n\n"
        "Examples:\n"
        "To read a file:\n"
        "```tool_call\n"
        '{"name": "Read", "arguments": {"file_path": "/path/to/file"}}\n'
        "```\n\n"
        "To write a file:\n"
        "```tool_call\n"
        '{"name": "Write", "arguments": {"file_path": "/path/to/file", "content": "file content here"}}\n'
        "```\n\n"
        "To edit a file:\n"
        "```tool_call\n"
        '{"name": "Edit", "arguments": {"file_path": "/path/to/file", "old_string": "text to replace", "new_string": "replacement text"}}\n'
        "```\n\n"
        "Available tools:\n" + "\n".join(tool_descriptions)
    )


def _format_tool_results(tool_calls: list[ToolCall] | None, content: str, name: str | None, tool_call_id: str | None) -> str:
    """Format a tool role message into human-readable text."""
    parts = []
    if name:
        parts.append(f"Tool result from {name}")
    elif tool_call_id:
        parts.append(f"Tool result (id: {tool_call_id})")
    if content:
        parts.append(content)
    return "\n".join(parts)


def translate_openai_request(request: OpenAIChatRequest) -> TranslatedRequest:
    system_lines: list[str] = []
    transcript_lines: list[str] = []
    prompt = ""

    # Inject tool definitions into system context
    tools_prompt = _format_tools_prompt(request.tools)
    if tools_prompt:
        system_lines.append(tools_prompt)

    for index, message in enumerate(request.messages):
        is_last = index == len(request.messages) - 1

        # Handle tool result messages
        if message.role == "tool":
            text = _format_tool_results(
                message.tool_calls,
                flatten_content(message.content),
                message.name,
                message.tool_call_id,
            )
            transcript_lines.append(f"Tool: {text}")
            continue

        text = flatten_content(message.content).strip()

        # Handle assistant messages with tool_calls
        if message.role == "assistant" and message.tool_calls:
            tool_call_texts = []
            for tc in message.tool_calls:
                try:
                    args = json.loads(tc.function.arguments) if isinstance(tc.function.arguments, str) else tc.function.arguments
                    tool_call_texts.append(f"Assistant called tool: {tc.function.name}({json.dumps(args, ensure_ascii=False)})")
                except (json.JSONDecodeError, TypeError):
                    tool_call_texts.append(f"Assistant called tool: {tc.function.name}({tc.function.arguments})")
            if text:
                tool_call_texts.insert(0, f"Assistant: {text}")
            transcript_lines.append("\n".join(tool_call_texts))
            continue

        if not text:
            continue
        if message.role in {"system", "developer"}:
            system_lines.append(text)
            continue
        if is_last:
            if message.role != "user":
                raise ValueError("The final OpenAI message must be a user message.")
            prompt = text
            continue
        transcript_lines.append(f"{message.role.capitalize()}: {text}")

    if not prompt:
        raise ValueError("A final user message is required.")

    additional_context: list[str] = []
    system_text = _join_lines(system_lines)
    if system_text:
        additional_context.append(f"System instructions:\n{system_text}")
    transcript_text = _join_lines(transcript_lines)
    if transcript_text:
        additional_context.append(f"Prior conversation transcript:\n{transcript_text}")
    return TranslatedRequest(prompt=prompt, additional_context=additional_context)


def translate_responses_request(request: "OpenAIResponsesRequest") -> TranslatedRequest:
    from .models import OpenAIResponsesRequest
    instructions = request.instructions or ""
    if isinstance(request.input, str):
        return TranslatedRequest(
            prompt=request.input,
            additional_context=[f"System instructions:\n{instructions}"] if instructions else [],
        )
    # input is a list of message dicts
    system_lines: list[str] = []
    if instructions:
        system_lines.append(instructions)
    transcript_lines: list[str] = []
    prompt = ""
    items = request.input
    for index, item in enumerate(items):
        role = item.get("role", "") if isinstance(item, dict) else ""
        content = item.get("content", "") if isinstance(item, dict) else str(item)
        if isinstance(content, list):
            content = "".join(p.get("text", "") for p in content if isinstance(p, dict) and p.get("type") in ("text", "input_text"))
        text = content.strip()
        if not text:
            continue
        is_last = index == len(items) - 1
        if role in {"system", "developer"}:
            system_lines.append(text)
            continue
        if is_last:
            if role != "user":
                raise ValueError("The final Responses input message must be a user message.")
            prompt = text
            continue
        transcript_lines.append(f"{role.capitalize()}: {text}")
    if not prompt:
        raise ValueError("No user message found in input.")
    additional_context: list[str] = []
    system_text = _join_lines(system_lines)
    if system_text:
        additional_context.append(f"System instructions:\n{system_text}")
    transcript_text = _join_lines(transcript_lines)
    if transcript_text:
        additional_context.append(f"Prior conversation transcript:\n{transcript_text}")
    return TranslatedRequest(prompt=prompt, additional_context=additional_context)


def translate_anthropic_request(
    request: AnthropicMessagesRequest,
) -> TranslatedRequest:
    system_text = flatten_content(request.system).strip()
    transcript_lines: list[str] = []
    prompt = ""

    for index, message in enumerate(request.messages):
        text = flatten_content(message.content).strip()
        if not text:
            continue
        is_last = index == len(request.messages) - 1
        if is_last:
            if message.role != "user":
                raise ValueError("The final Anthropic message must be a user message.")
            prompt = text
            continue
        transcript_lines.append(f"{message.role.capitalize()}: {text}")

    if not prompt:
        raise ValueError("A final user message is required.")

    additional_context: list[str] = []
    if system_text:
        additional_context.append(f"System instructions:\n{system_text}")
    transcript_text = _join_lines(transcript_lines)
    if transcript_text:
        additional_context.append(f"Prior conversation transcript:\n{transcript_text}")
    return TranslatedRequest(prompt=prompt, additional_context=additional_context)

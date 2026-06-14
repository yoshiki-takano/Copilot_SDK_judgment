import argparse
import json
import os
import re
import shutil
import sys
import asyncio
import inspect
from pathlib import Path
from typing import Any

import chardet

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    from copilot import CopilotClient
    try:
        from copilot import ExternalServerConfig, SubprocessConfig
    except Exception:
        ExternalServerConfig = None
        SubprocessConfig = None
    try:
        from copilot.session import PermissionHandler
    except Exception:
        PermissionHandler = None
    try:
        from copilot._sdk_protocol_version import get_sdk_protocol_version
    except Exception:
        from copilot.sdk_protocol_version import get_sdk_protocol_version
except Exception:
    CopilotClient = None
    ExternalServerConfig = None
    SubprocessConfig = None
    PermissionHandler = None
    get_sdk_protocol_version = None


DEFAULT_MAX_RETRIES = 5
DEFAULT_MODEL = "gpt-5"
MIN_SUPPORTED_SDK_PROTOCOL = 3
TOKEN_KEYS = {"GITHUB_COPILOT_TOKEN", "COPILOT_GITHUB_TOKEN", "GITHUB_TOKEN"}
_TOKEN_KEY_PATTERN = re.compile(r"^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*[:=]\s*(.+)$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Copilot SDK content generation helper")
    parser.add_argument("claims_text", help="Claims text")
    parser.add_argument("auth_token_path", help="Path to GitHub token text file")
    parser.add_argument("prompt_path", help="Path to prompt template text file")
    parser.add_argument("title_text", help="Title text")
    parser.add_argument("legacy_model", nargs="?", help="Legacy positional model argument")
    parser.add_argument("--model", dest="model_name", help="Copilot model name")
    parser.add_argument(
        "--stage",
        choices=["screening", "extract"],
        default="extract",
        help="Stage hint for fallback handling",
    )
    parser.add_argument(
        "--web-search",
        action="store_true",
        help="Compatibility flag. Copilot SDK path currently ignores this option.",
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=DEFAULT_MAX_RETRIES,
        help="Maximum retry count when generation fails",
    )
    parser.add_argument(
        "--reasoning-effort",
        choices=["low", "medium", "high", "xhigh"],
        help="Reasoning effort for models that support it",
    )
    parser.add_argument("--cli-path", help="Optional path to copilot CLI binary")
    parser.add_argument("--cli-url", help="Optional URL for existing copilot CLI server")
    parser.add_argument("--working-directory", help="Optional working directory for session tools")
    parser.add_argument(
        "--inputs-json",
        help="JSON payload for replacing [入力1], [入力2], ... placeholders",
    )
    parser.add_argument(
        "--inputs-json-file",
        help="Path to JSON payload file for replacing [入力1], [入力2], ... placeholders",
    )
    return parser.parse_args()


def read_auth_token(path: Path) -> str:
    if not str(path).strip():
        return ""
    if path.is_dir():
        return ""
    try:
        raw = path.read_text(encoding="utf-8", errors="replace")
        token = extract_token_value(raw)
        return token
    except FileNotFoundError:
        return ""
    except Exception:
        return ""


def extract_token_value(raw_text: str) -> str:
    if not raw_text:
        return ""
    text = raw_text.replace("\ufeff", "").strip()
    if not text:
        return ""

    fallback = ""
    for raw_line in text.replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        m = _TOKEN_KEY_PATTERN.match(line)
        if m:
            key = m.group(1).strip().upper()
            value = m.group(2).strip()
            if " #" in value:
                value = value.split(" #", 1)[0].rstrip()
            if value and value[0] in {'"', "'"} and value[-1:] == value[0]:
                value = value[1:-1].strip()
            if value.lower().startswith("bearer "):
                value = value[7:].strip()
            if key in TOKEN_KEYS and value:
                return value
            if value and not fallback:
                fallback = value
            continue

        candidate = line
        if candidate.lower().startswith("bearer "):
            candidate = candidate[7:].strip()
        if candidate and not fallback:
            fallback = candidate

    return fallback.strip()


def detect_and_read_file(file_path: Path) -> str:
    try:
        raw_data = file_path.read_bytes()
        detected_encoding = chardet.detect(raw_data).get("encoding") or "utf-8"
        return raw_data.decode(detected_encoding)
    except Exception as exc:
        print(f"Error reading prompt file: {exc}", file=sys.stderr)
        sys.exit(1)


def strip_code_fences(s: str) -> str:
    text = s.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if len(lines) >= 2:
            body = lines[1:]
            if body and body[-1].strip().startswith("```"):
                body = body[:-1]
            return "\n".join(body).strip()
    return text


def try_parse_json_text(text: str):
    candidates = [text, strip_code_fences(text)]
    for c in candidates:
        try:
            return json.loads(c)
        except Exception:
            pass
    return None


def normalize_screening_payload(parsed_obj: Any) -> dict[str, Any]:
    if not isinstance(parsed_obj, dict):
        return {
            "is_relevant": False,
            "reason": "model output was not JSON object",
        }
    is_relevant = bool(parsed_obj.get("is_relevant", False))
    reason = parsed_obj.get("reason", "理由のキーが見つかりません")
    if not isinstance(reason, str):
        reason = json.dumps(reason, ensure_ascii=False)
    return {"is_relevant": is_relevant, "reason": reason}


def normalize_extract_payload(parsed_obj: Any, raw_text: str) -> dict[str, Any]:
    if isinstance(parsed_obj, dict):
        return parsed_obj
    return {
        "_json_error": "model output was not valid JSON object",
        "_raw_text": raw_text,
    }


def build_fallback_payload(stage: str, error_text: str, raw_text: str = "") -> dict[str, Any]:
    msg = f"Copilot generation failed: {error_text}".strip()
    if raw_text:
        return {
            "_json_error": msg,
            "_raw_text": raw_text,
        }
    if stage == "screening":
        return {
            "is_relevant": False,
            "reason": msg,
        }
    return {
        "_json_error": msg,
        "_raw_text": "",
    }


def find_copilot_cli_path(explicit_path: str | None = None) -> str | None:
    if explicit_path:
        return explicit_path

    for name in ("copilot", "copilot.exe"):
        found = shutil.which(name)
        if found:
            return found

    if os.name == "nt":
        local_appdata = Path.home() / "AppData" / "Local"
        winget_root = local_appdata / "Microsoft" / "WinGet" / "Packages"
        if winget_root.exists():
            for candidate in winget_root.glob("GitHub.Copilot*/*copilot.exe"):
                if candidate.is_file():
                    return str(candidate)

    return None


def create_copilot_client(client_config: Any, client_options: dict[str, Any]) -> Any:
    """Create a Copilot client while tolerating SDK constructor differences."""
    if client_config is not None:
        return CopilotClient(client_config)

    if not client_options:
        return CopilotClient()

    try:
        sig = inspect.signature(CopilotClient)
        accepted: dict[str, Any] = {}
        has_var_kw = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in sig.parameters.values())
        if has_var_kw:
            accepted = dict(client_options)
        else:
            for key, value in client_options.items():
                if key in sig.parameters:
                    accepted[key] = value
        if accepted:
            return CopilotClient(**accepted)
        return CopilotClient()
    except TypeError:
        # Final fallback for older/newer SDKs with incompatible kwargs.
        return CopilotClient()


def parse_inputs_json(raw: str | None, file_path: str | None = None) -> list[dict[str, str]]:
    if file_path:
        try:
            raw = Path(file_path).read_text(encoding="utf-8")
        except Exception:
            raw = None
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return []

    items = data.get("inputs") if isinstance(data, dict) else data
    if not isinstance(items, list):
        return []

    out: list[dict[str, str]] = []
    for idx, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            continue
        placeholder = str(item.get("placeholder") or f"[入力{idx}]")
        column = str(item.get("column") or "")
        value = "" if item.get("value") is None else str(item.get("value"))
        out.append({"placeholder": placeholder, "column": column, "value": value})
    return out


def build_prompt(prompt_template: str, args: argparse.Namespace) -> str:
    inputs = parse_inputs_json(args.inputs_json, args.inputs_json_file)
    prompt = prompt_template

    for idx, item in enumerate(inputs, start=1):
        prompt = prompt.replace(item["placeholder"], item["value"])
        prompt = prompt.replace(f"[入力{idx}]", item["value"])

    input1 = inputs[0]["value"] if len(inputs) >= 1 else args.claims_text
    input2 = inputs[1]["value"] if len(inputs) >= 2 else args.title_text
    inputs_json = json.dumps(inputs, ensure_ascii=False)

    return (
        prompt
        .replace("{{Claims}}", input1)
        .replace("{{Title}}", input2)
        .replace("{{InputsJson}}", inputs_json)
    )


def _event_type_to_str(event_type: Any) -> str:
    value = getattr(event_type, "value", None)
    if isinstance(value, str):
        return value
    return str(event_type) if event_type is not None else ""


def _extract_response_text(response: Any) -> str:
    if response is None:
        return ""

    data = getattr(response, "data", None)
    if data is not None:
        content = getattr(data, "content", None)
        if isinstance(content, str):
            return content

    if isinstance(response, dict):
        data_obj = response.get("data")
        if isinstance(data_obj, dict):
            content = data_obj.get("content")
            if isinstance(content, str):
                return content
        content = response.get("content")
        if isinstance(content, str):
            return content

    return ""


async def _send_prompt_and_get_text(session: Any, prompt: str, timeout_sec: int = 300) -> str:
    if hasattr(session, "send_and_wait"):
        try:
            response = await session.send_and_wait(prompt, timeout=timeout_sec)
        except TypeError:
            response = await session.send_and_wait({"prompt": prompt})
        return _extract_response_text(response)

    done = asyncio.Event()
    message_holder = {"content": ""}

    def on_event(event: Any):
        event_type = _event_type_to_str(getattr(event, "type", None))
        data = getattr(event, "data", None)
        if event_type == "assistant.message":
            content = getattr(data, "content", None)
            if isinstance(content, str):
                message_holder["content"] = content
        elif event_type == "session.idle":
            done.set()

    unsubscribe = session.on(on_event)
    try:
        await session.send({"prompt": prompt})
        await asyncio.wait_for(done.wait(), timeout=timeout_sec)
    finally:
        if callable(unsubscribe):
            unsubscribe()

    return message_holder["content"]


async def generate_content_with_retry(args: argparse.Namespace) -> int:
    if CopilotClient is None:
        fallback = build_fallback_payload(args.stage, "github-copilot-sdk is not installed")
        print(json.dumps(fallback, ensure_ascii=False))
        return 1

    sdk_protocol = get_sdk_protocol_version() if get_sdk_protocol_version else 0
    if sdk_protocol < MIN_SUPPORTED_SDK_PROTOCOL:
        fallback = build_fallback_payload(
            args.stage,
            "github-copilot-sdk is too old for the installed Copilot CLI "
            f"(SDK protocol={sdk_protocol}, required>={MIN_SUPPORTED_SDK_PROTOCOL}). "
            "Run: C:\\D\\copilot_sdk_env\\Scripts\\python.exe -m pip install --upgrade --force-reinstall -r requirements.copilot.txt",
        )
        print(json.dumps(fallback, ensure_ascii=False))
        return 1

    model_name = args.model_name or args.legacy_model or DEFAULT_MODEL
    prompt_template = detect_and_read_file(Path(args.prompt_path))
    final_prompt = build_prompt(prompt_template, args)
    auth_token = read_auth_token(Path(args.auth_token_path))

    client_options: dict[str, Any] = {}
    if auth_token:
        client_options["github_token"] = auth_token
        client_options["use_logged_in_user"] = False
    cli_path = find_copilot_cli_path(args.cli_path)

    client_config = None
    if args.cli_url and ExternalServerConfig is not None:
        client_config = ExternalServerConfig(url=args.cli_url)
    elif cli_path and SubprocessConfig is not None:
        client_config = SubprocessConfig(
            cli_path=cli_path,
            github_token=client_options.get("github_token"),
            use_logged_in_user=client_options.get("use_logged_in_user"),
        )

    client = create_copilot_client(client_config, client_options)
    try:
        await client.start()
    except RuntimeError as exc:
        fallback = build_fallback_payload(
            args.stage,
            f"{exc}. If this is a protocol mismatch, update github-copilot-sdk with: "
            "C:\\D\\copilot_sdk_env\\Scripts\\python.exe -m pip install --upgrade --force-reinstall -r requirements.copilot.txt",
        )
        print(json.dumps(fallback, ensure_ascii=False))
        return 1

    if args.web_search:
        print("[WARN] --web-search is ignored in Copilot SDK mode.", file=sys.stderr)

        session = None
    try:
        session_config: dict[str, Any] = {"model": model_name}
        if args.reasoning_effort:
            session_config["reasoning_effort"] = args.reasoning_effort
        if args.working_directory:
            session_config["working_directory"] = args.working_directory

        if PermissionHandler is not None:
            session = await client.create_session(
                on_permission_request=PermissionHandler.approve_all,
                **session_config,
            )
        else:
            session = await client.create_session(session_config)

        last_error = "unknown error"
        last_raw_text = ""
        for attempt in range(args.max_retries):
            try:
                text = await _send_prompt_and_get_text(session, final_prompt)
                if not text:
                    raise RuntimeError("No valid text in response")
                parsed = try_parse_json_text(text)
                if parsed is None:
                    last_raw_text = text
                    raise ValueError("model output was not valid JSON")

                if args.stage == "screening":
                    out = normalize_screening_payload(parsed)
                else:
                    out = normalize_extract_payload(parsed, text)
                print(json.dumps(out, ensure_ascii=False))
                return 0
            except Exception as exc:
                last_error = str(exc)
                print(f"Attempt {attempt + 1}: Error generating content: {exc}", file=sys.stderr)
                if attempt < args.max_retries - 1:
                    await asyncio.sleep(2 ** attempt)

        fallback = build_fallback_payload(args.stage, last_error, last_raw_text)
        print(json.dumps(fallback, ensure_ascii=False))
        return 1
    finally:
        if session is not None:
            try:
                if hasattr(session, "disconnect"):
                    await session.disconnect()
                else:
                    await session.destroy()
            except Exception:
                pass
        await client.stop()


def main() -> int:
    args = parse_args()
    return asyncio.run(generate_content_with_retry(args))


if __name__ == "__main__":
    sys.exit(main())

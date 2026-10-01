"""Shared helpers for token handling, prompt building, JSON parsing, and Copilot SDK calls."""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from dataclasses import dataclass
from importlib import metadata
from pathlib import Path
from typing import Any, Callable

TOKEN_KEYS = ("GITHUB_COPILOT_TOKEN", "COPILOT_GITHUB_TOKEN", "GITHUB_TOKEN")
# The Streamlit app passes the token to child processes through this variable instead of a file.
TOKEN_ENV_VAR = "COPILOT_RUNNER_TOKEN"

MIN_SUPPORTED_SDK_PROTOCOL = 3
DEFAULT_REQUEST_TIMEOUT_SEC = 300
DEFAULT_MAX_RETRIES = 5
CLIENT_START_TIMEOUT_SEC = 120
CLIENT_STOP_TIMEOUT_SEC = 30
MAX_BACKOFF_SEC = 30

_TOKEN_KEY_PATTERN = re.compile(r"^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*[:=]\s*(.+)$")


def log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# Token
# ---------------------------------------------------------------------------

def extract_token_value(raw_text: str) -> str:
    """Extract a usable token from plain text or KEY=VALUE style input."""
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


def resolve_token(token_file: str | None) -> str:
    """Token from an explicit file, otherwise from TOKEN_ENV_VAR. Empty means logged-in user."""
    if token_file and token_file.strip():
        path = Path(token_file)
        if not path.is_file():
            raise FileNotFoundError(f"GitHubトークンのファイルが見つかりません: {token_file}")
        return extract_token_value(read_text_auto(path))
    return extract_token_value(os.environ.get(TOKEN_ENV_VAR, ""))


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------

def read_text_auto(path: str | Path) -> str:
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    try:
        import chardet

        detected = chardet.detect(raw).get("encoding")
        if detected:
            return raw.decode(detected, errors="replace")
    except Exception:
        pass
    return raw.decode("utf-8", errors="replace")


def strip_prompt_comments(text: str) -> str:
    """Remove template comment lines starting with // so they are never sent to the model."""
    return "\n".join(line for line in text.splitlines() if not line.lstrip().startswith("//"))


def read_prompt_text(path: str | Path) -> str:
    return strip_prompt_comments(read_text_auto(path))


_PLACEHOLDER_RE = re.compile(r"\[入力\s*(\d+)\]|\{\{(Claims|Title|InputsJson)\}\}")


def build_prompt(template: str, values: list[str]) -> str:
    """Replace [入力N] (and legacy {{...}}) in one pass so input text is never re-substituted."""
    inputs_json = json.dumps(
        [{"placeholder": f"[入力{i}]", "value": v} for i, v in enumerate(values, start=1)],
        ensure_ascii=False,
    )

    def repl(m: re.Match) -> str:
        if m.group(1):
            idx = int(m.group(1))
            return values[idx - 1] if 1 <= idx <= len(values) else m.group(0)
        name = m.group(2)
        if name == "Claims":
            return values[0] if values else ""
        if name == "Title":
            return values[1] if len(values) > 1 else ""
        return inputs_json

    return _PLACEHOLDER_RE.sub(repl, template)


# ---------------------------------------------------------------------------
# JSON / screening helpers
# ---------------------------------------------------------------------------

_FENCE_RE = re.compile(r"^\s*```(?:json|JSON)?\s*([\s\S]*?)\s*```\s*$")


def strip_code_fences(text: str) -> str:
    m = _FENCE_RE.match(text.strip())
    return m.group(1) if m else text


def parse_json_text(text: str) -> tuple[Any, str | None]:
    """Return (parsed, None) or (None, error message)."""
    candidates = [text.strip(), strip_code_fences(text).strip()]
    start, end = text.find("{"), text.rfind("}")
    if 0 <= start < end:
        candidates.append(text[start:end + 1])
    last_error = "empty response"
    for candidate in candidates:
        if not candidate:
            continue
        try:
            return json.loads(candidate), None
        except json.JSONDecodeError as exc:
            last_error = str(exc)
    return None, last_error


_TRUE_WORDS = {"true", "yes", "y", "1", "該当", "はい"}
_FALSE_WORDS = {"false", "no", "n", "0", "非該当", "いいえ"}


def parse_bool(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        word = value.strip().lower()
        if word in _TRUE_WORDS:
            return True
        if word in _FALSE_WORDS:
            return False
    return None


def validate_screening(data: Any) -> str | None:
    if not isinstance(data, dict):
        return "出力がJSONオブジェクトではありません"
    if parse_bool(data.get("is_relevant")) is None:
        return f"is_relevant が真偽値ではありません: {data.get('is_relevant')!r}"
    return None


def interpret_screening(data: dict[str, Any]) -> tuple[bool, str]:
    """Call only after validate_screening() succeeded."""
    reason = data.get("reason", "理由のキーが見つかりません")
    if not isinstance(reason, str):
        reason = json.dumps(reason, ensure_ascii=False)
    return bool(parse_bool(data.get("is_relevant"))), reason


def unwrap_extraction(data: Any) -> dict[str, Any] | None:
    if isinstance(data, dict):
        return data
    if isinstance(data, list) and len(data) == 1 and isinstance(data[0], dict):
        return data[0]
    return None


def validate_extraction(data: Any) -> str | None:
    if unwrap_extraction(data) is None:
        return "出力がJSONオブジェクト1つではありません"
    return None


# ---------------------------------------------------------------------------
# Copilot SDK
# ---------------------------------------------------------------------------

@dataclass
class GenerationResult:
    ok: bool
    data: Any = None
    raw_text: str = ""
    error: str = ""


def normalize_model(model: str | None) -> str | None:
    value = (model or "").strip()
    return value or None


def sdk_info() -> dict[str, Any]:
    info: dict[str, Any] = {}
    try:
        info["version"] = metadata.version("github-copilot-sdk")
    except Exception:
        info["version"] = "unknown"
    try:
        from copilot._sdk_protocol_version import get_sdk_protocol_version

        info["protocol"] = int(get_sdk_protocol_version())
    except Exception as exc:
        info["error"] = f"github-copilot-sdk を読み込めません: {exc}"
    return info


def _reject_all_permissions(request: Any, invocation: dict[str, str]) -> Any:
    from copilot.generated.rpc import PermissionDecisionReject

    return PermissionDecisionReject(feedback="Tool use is disabled in this workflow.")


def _extract_response_text(event: Any) -> str:
    data = getattr(event, "data", None)
    content = getattr(data, "content", None)
    return content if isinstance(content, str) else ""


class CopilotRunner:
    """One Copilot runtime per process; each request gets a fresh, tool-less session."""

    def __init__(
        self,
        token: str = "",
        cli_path: str | None = None,
        cli_url: str | None = None,
        request_timeout: float = DEFAULT_REQUEST_TIMEOUT_SEC,
        max_retries: int = DEFAULT_MAX_RETRIES,
        reasoning_effort: str | None = None,
    ) -> None:
        self.token = token
        self.cli_path = cli_path
        self.cli_url = cli_url
        self.request_timeout = request_timeout
        self.max_retries = max(1, max_retries)
        self.reasoning_effort = reasoning_effort
        self.client: Any = None

    async def __aenter__(self) -> "CopilotRunner":
        info = sdk_info()
        if "error" in info:
            raise RuntimeError(info["error"])
        if info["protocol"] < MIN_SUPPORTED_SDK_PROTOCOL:
            raise RuntimeError(
                f"github-copilot-sdk が古すぎます (protocol={info['protocol']}, 必要>={MIN_SUPPORTED_SDK_PROTOCOL})。"
                f" 更新: \"{sys.executable}\" -m pip install --upgrade -r requirements.txt"
            )

        from copilot import CopilotClient, RuntimeConnection

        kwargs: dict[str, Any] = {}
        if self.cli_url:
            kwargs["connection"] = RuntimeConnection.for_uri(self.cli_url)
        elif self.cli_path:
            kwargs["connection"] = RuntimeConnection.for_stdio(path=self.cli_path)
        if self.token:
            kwargs["github_token"] = self.token
            kwargs["use_logged_in_user"] = False

        self.client = CopilotClient(**kwargs)
        try:
            await asyncio.wait_for(self.client.start(), CLIENT_START_TIMEOUT_SEC)
        except BaseException:
            await self._shutdown()
            raise
        return self

    async def __aexit__(self, *exc_info: Any) -> None:
        await self._shutdown()

    async def _shutdown(self) -> None:
        if self.client is None:
            return
        client, self.client = self.client, None
        try:
            await asyncio.wait_for(client.stop(), CLIENT_STOP_TIMEOUT_SEC)
        except Exception:
            try:
                await client.force_stop()
            except Exception:
                pass

    async def ask(self, prompt: str, model: str | None) -> str:
        session = await self.client.create_session(
            on_permission_request=_reject_all_permissions,
            model=normalize_model(model),
            reasoning_effort=self.reasoning_effort,
            available_tools=[],
            skip_custom_instructions=True,
            enable_config_discovery=False,
        )
        try:
            event = await session.send_and_wait(prompt, timeout=self.request_timeout)
        finally:
            try:
                await session.disconnect()
            except Exception:
                pass
            try:
                await self.client.delete_session(session.session_id)
            except Exception:
                pass
        return _extract_response_text(event)

    async def generate_json(
        self,
        prompt: str,
        model: str | None,
        validate: Callable[[Any], str | None] | None = None,
    ) -> GenerationResult:
        last_error = "unknown error"
        last_raw = ""
        for attempt in range(1, self.max_retries + 1):
            try:
                text = await self.ask(prompt, model)
                if not text.strip():
                    raise RuntimeError("応答が空です")
                data, parse_error = parse_json_text(text)
                if parse_error is not None:
                    last_raw = text
                    raise ValueError(f"JSONとして解析できません: {parse_error}")
                validation_error = validate(data) if validate else None
                if validation_error:
                    last_raw = text
                    raise ValueError(validation_error)
                return GenerationResult(ok=True, data=data, raw_text=text)
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                log(f"[RETRY] attempt {attempt}/{self.max_retries}: {last_error}")
                if attempt < self.max_retries:
                    await asyncio.sleep(min(2 ** (attempt - 1), MAX_BACKOFF_SEC))
        return GenerationResult(ok=False, raw_text=last_raw, error=last_error)


async def fetch_status(token: str, cli_path: str | None = None, cli_url: str | None = None) -> dict[str, Any]:
    result: dict[str, Any] = {"python": sys.executable, "sdk": sdk_info()}
    try:
        async with CopilotRunner(token=token, cli_path=cli_path, cli_url=cli_url) as runner:
            status = await runner.client.get_status()
            result["cli"] = {"version": status.version, "protocol": status.protocol_version}
            auth = await runner.client.get_auth_status()
            result["auth"] = {
                "authenticated": bool(auth.isAuthenticated),
                "login": auth.login,
                "type": auth.authType,
                "message": auth.statusMessage,
            }
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
    return result


async def fetch_models(token: str, cli_path: str | None = None, cli_url: str | None = None) -> list[dict[str, str]]:
    async with CopilotRunner(token=token, cli_path=cli_path, cli_url=cli_url) as runner:
        models = await runner.client.list_models()
    return [{"id": m.id, "name": m.name} for m in models]

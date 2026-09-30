"""
Copilot SDK command-line helper.

  単発生成（後方互換）:
    Copilot_Calling.py <claims_text> <auth_token_path> <prompt_path> <title_text> [model]
        [--inputs-json JSON | --inputs-json-file PATH] [--stage screening|extract]
  状態確認:    Copilot_Calling.py --status [--token-file PATH]
  モデル一覧:  Copilot_Calling.py --list-models [--token-file PATH]

トークンは auth_token_path / --token-file、なければ環境変数 COPILOT_RUNNER_TOKEN、
どちらも無ければログイン済みユーザーを使用する。
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
from pathlib import Path
from typing import Any

from copilot_common import (
    DEFAULT_MAX_RETRIES,
    DEFAULT_REQUEST_TIMEOUT_SEC,
    CopilotRunner,
    build_prompt,
    fetch_models,
    fetch_status,
    interpret_screening,
    read_prompt_text,
    resolve_token,
    unwrap_extraction,
    validate_extraction,
    validate_screening,
)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DEFAULT_MODEL = "auto"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Copilot SDK content generation helper")
    parser.add_argument("claims_text", nargs="?", default="", help="Claims text ([入力1] の既定値)")
    parser.add_argument("auth_token_path", nargs="?", default="", help="Path to GitHub token text file")
    parser.add_argument("prompt_path", nargs="?", default="", help="Path to prompt template text file")
    parser.add_argument("title_text", nargs="?", default="", help="Title text ([入力2] の既定値)")
    parser.add_argument("legacy_model", nargs="?", help="Legacy positional model argument")
    parser.add_argument("--status", action="store_true", help="SDK / CLI / 認証状態をJSONで出力")
    parser.add_argument("--list-models", action="store_true", help="利用可能なモデル一覧をJSONで出力")
    parser.add_argument("--token-file", help="GitHub token file (auth_token_path の代替)")
    parser.add_argument("--model", dest="model_name", help="Copilot model name")
    parser.add_argument("--stage", choices=["screening", "extract"], default="extract")
    parser.add_argument("--web-search", action="store_true", help="Compatibility flag (ignored).")
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES)
    parser.add_argument("--request-timeout", type=float, default=DEFAULT_REQUEST_TIMEOUT_SEC)
    parser.add_argument("--reasoning-effort", choices=["low", "medium", "high", "xhigh"])
    parser.add_argument("--cli-path", help="Optional path to copilot CLI binary")
    parser.add_argument("--cli-url", help="Optional URL for existing copilot CLI server")
    parser.add_argument("--inputs-json", help="JSON payload for [入力1], [入力2], ... placeholders")
    parser.add_argument("--inputs-json-file", help="Path to JSON payload file for placeholders")
    return parser.parse_args()


def parse_input_values(args: argparse.Namespace) -> list[str]:
    raw = args.inputs_json
    if args.inputs_json_file:
        raw = Path(args.inputs_json_file).read_text(encoding="utf-8")
    if raw:
        data = json.loads(raw)
        items = data.get("inputs") if isinstance(data, dict) else data
        if isinstance(items, list):
            values: dict[int, str] = {}
            for pos, item in enumerate(items, start=1):
                if not isinstance(item, dict):
                    continue
                placeholder = str(item.get("placeholder") or "")
                digits = "".join(ch for ch in placeholder if ch.isdigit())
                idx = int(digits) if digits else pos
                values[idx] = "" if item.get("value") is None else str(item["value"])
            if values:
                return [values.get(i, "") for i in range(1, max(values) + 1)]
    return [args.claims_text, args.title_text]


def emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False))


async def generate(args: argparse.Namespace, token: str) -> int:
    if args.web_search:
        print("[WARN] --web-search is ignored in Copilot SDK mode.", file=sys.stderr)
    if not args.prompt_path:
        emit({"_json_error": "prompt_path is required", "_raw_text": ""})
        return 1

    prompt = build_prompt(read_prompt_text(args.prompt_path), parse_input_values(args))
    model = args.model_name or args.legacy_model or DEFAULT_MODEL
    validate = validate_screening if args.stage == "screening" else validate_extraction

    try:
        async with CopilotRunner(
            token=token,
            cli_path=args.cli_path,
            cli_url=args.cli_url,
            request_timeout=args.request_timeout,
            max_retries=args.max_retries,
            reasoning_effort=args.reasoning_effort,
        ) as runner:
            result = await runner.generate_json(prompt, model, validate)
    except Exception as exc:
        emit({"_json_error": f"Copilot generation failed: {type(exc).__name__}: {exc}", "_raw_text": ""})
        return 1

    if not result.ok:
        emit({"_json_error": f"Copilot generation failed: {result.error}", "_raw_text": result.raw_text})
        return 1
    if args.stage == "screening":
        is_relevant, reason = interpret_screening(result.data)
        emit({"is_relevant": is_relevant, "reason": reason})
    else:
        emit(unwrap_extraction(result.data))
    return 0


def main() -> int:
    args = parse_args()
    try:
        token = resolve_token(args.token_file or args.auth_token_path)
    except FileNotFoundError as exc:
        emit({"error": str(exc)} if (args.status or args.list_models) else {"_json_error": str(exc), "_raw_text": ""})
        return 1

    if args.status:
        emit(asyncio.run(fetch_status(token, args.cli_path, args.cli_url)))
        return 0
    if args.list_models:
        try:
            emit({"models": asyncio.run(fetch_models(token, args.cli_path, args.cli_url))})
            return 0
        except Exception as exc:
            emit({"error": f"{type(exc).__name__}: {exc}"})
            return 2
    return asyncio.run(generate(args, token))


if __name__ == "__main__":
    sys.exit(main())

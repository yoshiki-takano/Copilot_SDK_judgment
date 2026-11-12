from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:
    tk = None
    filedialog = None

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from copilot_common import (
    MIN_SUPPORTED_SDK_PROTOCOL,
    TOKEN_ENV_VAR,
    TOKEN_KEYS,
    extract_token_value,
    read_prompt_text,
)
from pipeline_common import EXIT_OK, EXIT_ROW_ERRORS, MERGE_DONE_MARKER, OUTPUT_MARKER

APP_DIR = Path(__file__).resolve().parent

MODEL_CANDIDATES = [
    "auto",
    "gpt-5.5",
    "gpt-6-sol",
    "gpt-6-luna",
    "claude-sonnet-5",
    "claude-opus-5.5",
]

SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_SPLIT_PARTS = 50
MAX_WORKERS = 10
SESSION_DIR_MAX_AGE_SEC = 24 * 60 * 60


@dataclass
class RunConfig:
    workspace: Path
    mode: str
    python_exe: str
    base_excel: Path | None
    base_name: str
    sheet_name: str
    part_count: int
    input_columns: list[str]
    prompt1: str
    prompt2: str
    token: str = field(repr=False)
    auth_label: str
    output_root: str
    parts_dir: str
    metadata_path: str
    outdir: str
    model_screening: str
    model_extraction: str


def norm(path_text: str) -> str:
    return path_text.replace("\\", "/")


def is_streamlit_cloud() -> bool:
    cloud_flag = str(os.getenv("STREAMLIT_SHARING_MODE", "")).strip().lower()
    runtime = str(os.getenv("STREAMLIT_RUNTIME", "")).strip().lower()
    return (
        cloud_flag in {"1", "true", "yes"}
        or runtime == "cloud"
        or norm(str(Path.cwd())).startswith("/mount/src")
    )


# ---------------------------------------------------------------------------
# Per-session storage (kept outside the workspace so uploads are not synced or shared)
# ---------------------------------------------------------------------------

def session_runtime_dir() -> Path:
    d = Path(tempfile.gettempdir()) / "copilot_sdk_runner" / st.session_state["_session_id"]
    d.mkdir(parents=True, exist_ok=True)
    return d


def _latest_mtime(path: Path) -> float:
    latest = path.stat().st_mtime
    for child in path.rglob("*"):
        try:
            latest = max(latest, child.stat().st_mtime)
        except OSError:
            continue
    return latest


def cleanup_stale_session_dirs() -> None:
    """Delete other sessions' folders whose newest file is older than SESSION_DIR_MAX_AGE_SEC."""
    root = Path(tempfile.gettempdir()) / "copilot_sdk_runner"
    if not root.is_dir():
        return
    own = st.session_state.get("_session_id", "")
    cutoff = time.time() - SESSION_DIR_MAX_AGE_SEC
    for d in root.iterdir():
        try:
            if d.name != own and d.is_dir() and _latest_mtime(d) < cutoff:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            continue


def default_output_root(workspace: Path, is_cloud: bool) -> str:
    if is_cloud:
        return norm(str(session_runtime_dir() / "out"))
    return norm(str(workspace / "out"))


def sync_upload(uploaded_file, path_key: str) -> None:
    """Mirror an uploader's value into session state, saving to disk only when the file changes."""
    id_key = f"{path_key}_id"
    if uploaded_file is None:
        st.session_state[path_key] = ""
        st.session_state.pop(id_key, None)
        return
    file_id = getattr(uploaded_file, "file_id", None) or f"{uploaded_file.name}:{uploaded_file.size}"
    current = st.session_state.get(path_key, "")
    if st.session_state.get(id_key) == file_id and current and Path(current).is_file():
        return
    out_dir = session_runtime_dir() / "inputs" / path_key.strip("_")
    out_dir.mkdir(parents=True, exist_ok=True)
    save_path = out_dir / Path(uploaded_file.name).name
    save_path.write_bytes(uploaded_file.getvalue())
    st.session_state[path_key] = norm(str(save_path))
    st.session_state[id_key] = file_id


def pick_directory_dialog(initial_dir: str) -> str:
    if tk is None or filedialog is None:
        raise RuntimeError("tkinter が利用できないため、ダイアログを開けません")
    root = tk.Tk()
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            title="出力先フォルダを選択",
            initialdir=initial_dir or str(Path.cwd()),
        )
    finally:
        root.destroy()
    return norm(str(Path(selected).expanduser())) if selected else ""


# ---------------------------------------------------------------------------
# Authentication (token is held in memory and passed to child processes via env)
# ---------------------------------------------------------------------------

def _safe_get_streamlit_secret(key: str) -> str:
    try:
        return str(st.secrets.get(key, ""))
    except Exception:
        return ""


def secrets_token() -> str:
    for key in TOKEN_KEYS:
        token = extract_token_value(_safe_get_streamlit_secret(key))
        if token:
            return token
    return ""


def sync_token_upload(uploaded_file) -> None:
    if uploaded_file is None:
        st.session_state["_uploaded_token"] = ""
        return
    raw = uploaded_file.getvalue().decode("utf-8", errors="replace")
    st.session_state["_uploaded_token"] = extract_token_value(raw)


def current_auth() -> tuple[str, str]:
    uploaded = st.session_state.get("_uploaded_token", "")
    if uploaded:
        return uploaded, "token (uploaded file)"
    secret = secrets_token()
    if secret:
        return secret, "token (Streamlit Secrets)"
    return "", "logged-in user"


def child_env(token: str) -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    if token:
        env[TOKEN_ENV_VAR] = token
    else:
        env.pop(TOKEN_ENV_VAR, None)
    return env


# ---------------------------------------------------------------------------
# Copilot helper subprocesses
# ---------------------------------------------------------------------------

def _run_helper_json(workspace: Path, helper_args: list[str], token: str, timeout: int) -> dict[str, Any]:
    result = subprocess.run(
        [sys.executable, str(workspace / "Copilot_Calling.py"), *helper_args],
        cwd=str(workspace),
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=child_env(token),
    )
    lines = [line for line in result.stdout.splitlines() if line.strip()]
    if not lines:
        raise RuntimeError(result.stderr.strip() or f"exit code {result.returncode}")
    payload = json.loads(lines[-1])
    if not isinstance(payload, dict):
        raise RuntimeError(f"unexpected output: {lines[-1][:200]}")
    return payload


def fetch_copilot_status(workspace: Path, token: str) -> dict[str, Any]:
    try:
        return _run_helper_json(workspace, ["--status"], token, timeout=180)
    except Exception as exc:
        return {"error": str(exc)}


def fetch_copilot_models(workspace: Path, token: str) -> list[str]:
    payload = _run_helper_json(workspace, ["--list-models"], token, timeout=180)
    if payload.get("error"):
        raise RuntimeError(str(payload["error"]))
    ids: list[str] = []
    for item in payload.get("models") or []:
        model_id = str(item.get("id") or "").strip() if isinstance(item, dict) else ""
        if model_id and model_id not in ids:
            ids.append(model_id)
    return ids


# ---------------------------------------------------------------------------
# Prompt analysis
# ---------------------------------------------------------------------------

def script_for_mode(mode: str) -> str:
    if mode == "both":
        return "main_parallel.py"
    if mode == "screening":
        return "step1only_main_parallel.py"
    if mode == "extraction":
        return "step2only_main_parallel.py"
    raise ValueError(f"Unsupported mode: {mode}")


def required_prompts(mode: str) -> set[str]:
    if mode == "both":
        return {"prompt1", "prompt2"}
    if mode == "screening":
        return {"prompt1"}
    if mode == "extraction":
        return {"prompt2"}
    return set()


def required_prompt_paths(mode: str, prompt1_path: str, prompt2_path: str) -> list[str]:
    needed = required_prompts(mode)
    paths = []
    if "prompt1" in needed and prompt1_path:
        paths.append(prompt1_path)
    if "prompt2" in needed and prompt2_path:
        paths.append(prompt2_path)
    return paths


_INPUT_PLACEHOLDER_RE = re.compile(r"\[入力\s*(\d+)\]")
_INPUT_BRACE_RE = re.compile(r"\{\{[^}]+\}\}")
_INPUT_LABEL_PREFIX_RE = re.compile(r"^【入力\s*\d+】\s*")


def _extract_input_mapping_from_prompt(prompt_text: str) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for raw_line in prompt_text.replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        for m in _INPUT_PLACEHOLDER_RE.finditer(line):
            idx = int(m.group(1))
            if idx in mapping:
                # Keep the first definition; later prose mentioning [入力n] must not overwrite it.
                continue
            # "【入力1】Article Title: [入力1]" -> "Article Title"; the label is for the model only.
            col = _INPUT_LABEL_PREFIX_RE.sub("", line[: m.start()].strip())
            col = col.strip(":：- ")
            if col:
                mapping[idx] = col
    return mapping


def _extract_brace_columns_from_prompt(prompt_text: str) -> list[str]:
    columns: list[str] = []
    for raw_line in prompt_text.replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if line and _INPUT_BRACE_RE.search(line):
            col = line.split("{{", 1)[0].strip().strip(":：- ")
            if col:
                columns.append(col)
    return columns


def detect_input_columns_from_prompts(mode: str, prompt1_path: str, prompt2_path: str) -> list[str]:
    merged: dict[int, str] = {}
    brace_columns: list[str] = []
    for p in required_prompt_paths(mode, prompt1_path, prompt2_path):
        try:
            text = read_prompt_text(p)
        except OSError:
            continue
        for idx, col in _extract_input_mapping_from_prompt(text).items():
            merged.setdefault(idx, col)
        brace_columns.extend(_extract_brace_columns_from_prompt(text))

    ordered = [merged[k] for k in sorted(merged)] + brace_columns
    deduped: list[str] = []
    for col in ordered:
        if col not in deduped:
            deduped.append(col)
    return deduped


def _normalize_column_name(name: str) -> str:
    # Normalize punctuation/spacing variants so prompt labels can match Excel headers.
    text = name.strip().lower()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"[\s\u3000]+", "", text)
    text = re.sub(r"[-‐‑‒–—―ー_()]", "", text)
    return text


def resolve_detected_columns_to_headers(detected_columns: list[str], headers: list[str]) -> list[str]:
    norm_to_headers: dict[str, list[str]] = {}
    for h in headers:
        norm_to_headers.setdefault(_normalize_column_name(h), []).append(h)

    resolved: list[str] = []
    for col in detected_columns:
        if col in headers:
            resolved.append(col)
            continue
        candidates = norm_to_headers.get(_normalize_column_name(col), [])
        # Keep original when unresolved/ambiguous; validation will show a clear error.
        resolved.append(candidates[0] if len(candidates) == 1 else col)
    return resolved


INPUT_SLOT_KEY = "_input_col_slot_{}"


def detect_placeholder_defaults(mode: str, prompt1_path: str, prompt2_path: str) -> tuple[int, dict[int, str]]:
    """Return (max [入力N] index, auto-detected column per index) across required prompts."""
    max_idx = 0
    defaults: dict[int, str] = {}
    for p in required_prompt_paths(mode, prompt1_path, prompt2_path):
        try:
            text = read_prompt_text(p)
        except OSError:
            continue
        for m in _INPUT_PLACEHOLDER_RE.finditer(text):
            max_idx = max(max_idx, int(m.group(1)))
        for idx, col in _extract_input_mapping_from_prompt(text).items():
            defaults.setdefault(idx, col)
    return max_idx, defaults


def selected_slot_columns(slot_count: int) -> list[str]:
    return [str(st.session_state.get(INPUT_SLOT_KEY.format(n)) or "") for n in range(1, slot_count + 1)]


def _file_digest(path_text: str) -> str:
    if not path_text or not Path(path_text).is_file():
        return ""
    return hashlib.sha1(Path(path_text).read_bytes()).hexdigest()


_SHEET_LABEL_RE = re.compile(r"^[\s#*\-【\[]*作業対象シート名?\s*[】\]]?\s*[:：]?\s*(.*)$")
_SHEET_QUOTED_RE = re.compile(r"[「『\"“](.+?)[」』\"”]")


def detect_sheet_from_prompts(mode: str, prompt1_path: str, prompt2_path: str) -> str:
    """Return the sheet name written as 「作業対象シート：」(same or next line) in required prompts."""
    for p in required_prompt_paths(mode, prompt1_path, prompt2_path):
        try:
            lines = read_prompt_text(p).splitlines()
        except OSError:
            continue
        for i, line in enumerate(lines):
            m = _SHEET_LABEL_RE.match(line)
            if not m:
                continue
            for cand in [m.group(1)] + lines[i + 1:i + 3]:
                cand = cand.strip()
                if cand:
                    quoted = _SHEET_QUOTED_RE.search(cand)
                    return (quoted.group(1) if quoted else cand).strip()
    return ""


def prompt_sheet_signature(mode: str, prompt1_path: str, prompt2_path: str, excel_path: str) -> tuple:
    return (mode, _file_digest(prompt1_path), _file_digest(prompt2_path), excel_path)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def validate_excel_path(src_path: Path) -> None:
    if not src_path.exists():
        raise ValueError(f"Excel file not found: {src_path}")
    if not src_path.is_file():
        raise ValueError(f"Base Excel path is not a file: {src_path}")
    if src_path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        allowed = ", ".join(sorted(SUPPORTED_EXCEL_SUFFIXES))
        raise ValueError(f"Unsupported Excel format: {src_path.suffix}. Supported: {allowed}")


def _open_workbook_read_only(src_path: Path):
    validate_excel_path(src_path)
    try:
        return load_workbook(filename=str(src_path), data_only=True, read_only=True)
    except InvalidFileException as exc:
        raise ValueError("Excel形式が不正です。Excelで開ける .xlsx/.xlsm を指定してください。") from exc


def read_excel_headers(src_path: Path, sheet_name: str | None) -> tuple[str, list[str]]:
    wb = _open_workbook_read_only(src_path)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers = [str(v).strip() for v in first_row if v is not None and str(v).strip()]
        return ws.title, headers
    finally:
        wb.close()


def read_headers_safe(base_excel: Path | None, sheet_name: str) -> list[str] | None:
    if base_excel is None or not base_excel.is_file():
        return None
    try:
        return read_excel_headers(base_excel, sheet_name)[1]
    except Exception:
        return None


def read_excel_sheet_names(src_path: Path) -> list[str]:
    wb = _open_workbook_read_only(src_path)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _row_has_value(row: tuple[Any, ...]) -> bool:
    return any(v is not None and str(v).strip() != "" for v in row)


def split_excel(src_path: Path, parts_dir: Path, sheet_name: str | None, part_count: int) -> tuple[str, int, str]:
    """Split data rows evenly into up to part_count files. Returns (base_name, parts_made, sheet_used)."""
    if src_path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("Input must be .xlsx or .xlsm for split output compatibility")
    if part_count <= 0:
        raise ValueError("part_count must be > 0")

    wb = _open_workbook_read_only(src_path)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
        used_sheet_name = ws.title
        base_name = src_path.stem

        last_data_row = 1
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
            if _row_has_value(row):
                last_data_row = row_idx
        data_rows = last_data_row - 1
        if data_rows < 1:
            return base_name, 0, used_sheet_name

        parts = min(part_count, data_rows)
        sizes = [data_rows // parts + (1 if i < data_rows % parts else 0) for i in range(parts)]
        parts_dir.mkdir(parents=True, exist_ok=True)

        rows = ws.iter_rows(min_row=1, max_row=last_data_row, values_only=True)
        header = list(next(rows))
        for part_idx, size in enumerate(sizes, start=1):
            wb_out = Workbook(write_only=True)
            ws_out = wb_out.create_sheet(title=used_sheet_name)
            ws_out.append(header)
            for _ in range(size):
                ws_out.append(list(next(rows)))
            wb_out.save(str(parts_dir / f"{base_name}_Part{part_idx}.xlsx"))
            wb_out.close()
        return base_name, parts, used_sheet_name
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Commands / metadata
# ---------------------------------------------------------------------------

def _part_args(cfg: RunConfig, part_number: int) -> list[str]:
    args = [
        str(cfg.workspace / script_for_mode(cfg.mode)),
        "--input", str(Path(cfg.parts_dir) / f"{cfg.base_name}_Part{part_number}.xlsx"),
        "--sheet", cfg.sheet_name,
    ]
    for col in cfg.input_columns:
        args += ["--input-column", col]
    needed = required_prompts(cfg.mode)
    if "prompt1" in needed:
        args += ["--prompt1", cfg.prompt1]
    if "prompt2" in needed:
        args += ["--prompt2", cfg.prompt2]
    args += [
        "--model-stage1", cfg.model_screening,
        "--model-stage2", cfg.model_extraction,
        "--outdir", cfg.outdir,
        "--no-ui",
    ]
    return args


def _merge_args(cfg: RunConfig) -> list[str]:
    return [
        str(cfg.workspace / script_for_mode(cfg.mode)),
        "--merge-only",
        "--merge-base", f"{cfg.base_name}_Part{cfg.part_count}",
        "--sheet", cfg.sheet_name,
        "--outdir", cfg.outdir,
        "--parts", str(cfg.part_count),
        "--no-ui",
        "--run-metadata", cfg.metadata_path,
    ]


def build_part_command(cfg: RunConfig, part_number: int) -> list[str]:
    return [cfg.python_exe, *_part_args(cfg, part_number)]


def build_merge_command(cfg: RunConfig) -> list[str]:
    return [cfg.python_exe, *_merge_args(cfg)]


def build_tasks_payload(cfg: RunConfig) -> dict[str, Any]:
    """VS Code tasks.json. The token is never written; tasks use the logged-in user or COPILOT_RUNNER_TOKEN."""
    if cfg.part_count <= 0:
        raise ValueError("No split parts found")

    def task(label: str, args: list[str]) -> dict[str, Any]:
        return {
            "label": label,
            "type": "process",
            "command": cfg.python_exe,
            "args": args,
            "options": {"cwd": str(cfg.workspace)},
            "problemMatcher": [],
        }

    labels = [f"part{n}" for n in range(1, cfg.part_count + 1)]
    tasks = [task(label, _part_args(cfg, n)) for n, label in enumerate(labels, start=1)]
    tasks.append(task("merge-only", _merge_args(cfg)))
    tasks.append({"label": "run-all-parallel", "dependsOn": labels})
    tasks.append({
        "label": "run-all-parallel-then-merge",
        "dependsOn": ["run-all-parallel", "merge-only"],
        "dependsOrder": "sequence",
    })
    return {"version": "2.0.0", "tasks": tasks}


def _prompt_meta(path_text: str) -> tuple[str, str]:
    if not path_text or not Path(path_text).is_file():
        return "", ""
    return _file_digest(path_text), read_prompt_text(path_text)


def write_run_files(cfg: RunConfig) -> None:
    Path(cfg.outdir).mkdir(parents=True, exist_ok=True)
    metadata_path = Path(cfg.metadata_path)
    metadata_path.parent.mkdir(parents=True, exist_ok=True)
    needed = required_prompts(cfg.mode)
    p1_sha, p1_text = _prompt_meta(cfg.prompt1) if "prompt1" in needed else ("", "")
    p2_sha, p2_text = _prompt_meta(cfg.prompt2) if "prompt2" in needed else ("", "")
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "mode": cfg.mode,
        "parts": cfg.part_count,
        "base_excel_path": norm(str(cfg.base_excel)) if cfg.base_excel else "",
        "base_name": cfg.base_name,
        "sheet_name": cfg.sheet_name,
        "input_columns": cfg.input_columns,
        "prompt1_path": norm(cfg.prompt1) if "prompt1" in needed else "",
        "prompt1_sha1": p1_sha,
        "prompt1_text": p1_text,
        "prompt2_path": norm(cfg.prompt2) if "prompt2" in needed else "",
        "prompt2_sha1": p2_sha,
        "prompt2_text": p2_text,
        "auth": cfg.auth_label,
        "out_dir": norm(cfg.outdir),
        "model_screening": cfg.model_screening,
        "model_extraction": cfg.model_extraction,
        "python_command": cfg.python_exe,
    }
    metadata_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tasks_path = metadata_path.parent / "tasks.json"
    tasks_path.write_text(json.dumps(build_tasks_payload(cfg), ensure_ascii=False, indent=2), encoding="utf-8")


_ANSI_ESCAPE = re.compile(r"\x1b\[[0-9;]*[mGKHF]")


def _decode_output(raw: bytes) -> str:
    for enc in ("utf-8", "cp932"):
        try:
            return _ANSI_ESCAPE.sub("", raw.decode(enc))
        except UnicodeDecodeError:
            continue
    return _ANSI_ESCAPE.sub("", raw.decode("utf-8", errors="replace"))


def run_command(args: list[str], cwd: Path, env: dict[str, str]) -> dict[str, Any]:
    started = datetime.now()
    try:
        proc = subprocess.run(args, cwd=str(cwd), stdin=subprocess.DEVNULL, capture_output=True, env=env)
        returncode, stdout, stderr = proc.returncode, _decode_output(proc.stdout), _decode_output(proc.stderr)
    except OSError as exc:
        returncode, stdout, stderr = -1, "", f"起動に失敗しました: {exc}"
    return {
        "args": args,
        "returncode": returncode,
        "stdout": stdout,
        "stderr": stderr,
        "started_at": started.isoformat(timespec="seconds"),
        "finished_at": datetime.now().isoformat(timespec="seconds"),
    }


def marker_path(stdout: str, marker: str) -> Path | None:
    for line in reversed(stdout.splitlines()):
        if line.startswith(marker):
            path = Path(line[len(marker):].strip())
            return path if path.is_file() else None
    return None


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def validate_config(cfg: RunConfig) -> list[str]:
    errs: list[str] = []
    if not cfg.token:
        errs.append("Copilotトークンが確認できません。Token file upload または Streamlit Secrets でトークンを設定してください")
    if cfg.base_excel is None:
        errs.append("分析対象の Excel をアップロードしてください")
    elif not cfg.base_excel.is_file():
        errs.append("分析対象のExcelが存在しません。再アップロードしてください")
    if not cfg.output_root.strip():
        errs.append("出力フォルダを指定してください")
    else:
        try:
            Path(cfg.output_root).mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            errs.append(f"出力フォルダを作成できません: {exc}")
    if "prompt1" in required_prompts(cfg.mode) and not (cfg.prompt1 and Path(cfg.prompt1).is_file()):
        errs.append("Stage 1 prompt をアップロードしてください")
    if "prompt2" in required_prompts(cfg.mode) and not (cfg.prompt2 and Path(cfg.prompt2).is_file()):
        errs.append("Stage 2 prompt をアップロードしてください")
    if not cfg.input_columns:
        errs.append("Prompt から入力列を特定できません")
    elif any(not col for col in cfg.input_columns):
        unselected = [f"[入力{i}]" for i, col in enumerate(cfg.input_columns, start=1) if not col]
        errs.append(f"入力列を選択してください: {', '.join(unselected)}")
    elif cfg.base_excel is not None and cfg.base_excel.is_file():
        try:
            _, headers = read_excel_headers(cfg.base_excel, cfg.sheet_name)
            missing = [col for col in cfg.input_columns if col not in headers]
            if missing:
                errs.append(f"入力ファイルに対象列がありません: {', '.join(missing)}")
        except Exception as exc:
            errs.append(f"Excelヘッダ確認に失敗しました: {exc}")
    return errs


def get_config_from_ui() -> RunConfig:
    workspace = Path(st.session_state.workspace)
    mode = st.session_state.get("mode", "screening")
    base_excel_text = st.session_state.get("_uploaded_base_excel_path", "").strip()
    base_excel = Path(base_excel_text) if base_excel_text else None
    prompt1 = st.session_state.get("_uploaded_prompt1_path", "").strip()
    prompt2 = st.session_state.get("_uploaded_prompt2_path", "").strip()
    base_name = st.session_state.get("_base_name", "").strip() or (base_excel.stem if base_excel else "")
    sheet_name = st.session_state.get("_sheet_name", "Sheet1").strip() or "Sheet1"
    output_root = st.session_state.get("_output_root", "").strip()
    output_root = norm(str(Path(output_root).expanduser().resolve())) if output_root else ""
    root = Path(output_root) if output_root else workspace / "out"

    headers = read_headers_safe(base_excel, sheet_name)
    slot_count, _ = detect_placeholder_defaults(mode, prompt1, prompt2)
    if slot_count and headers is not None:
        columns = selected_slot_columns(slot_count)
    else:
        columns = detect_input_columns_from_prompts(mode, prompt1, prompt2)
        if columns and headers is not None:
            columns = resolve_detected_columns_to_headers(columns, headers)

    token, auth_label = current_auth()
    return RunConfig(
        workspace=workspace,
        mode=mode,
        python_exe=st.session_state.python_exe,
        base_excel=base_excel,
        base_name=base_name,
        sheet_name=sheet_name,
        part_count=int(st.session_state.get("_part_count", 0)),
        input_columns=columns,
        prompt1=prompt1,
        prompt2=prompt2,
        token=token,
        auth_label=auth_label,
        output_root=output_root,
        parts_dir=norm(str(root / "parts")),
        metadata_path=norm(str(root / ".vscode" / "copilot_run_metadata.json")),
        outdir=norm(str(root / "outputs")),
        model_screening=st.session_state.get("model_screening", MODEL_CANDIDATES[0]),
        model_extraction=st.session_state.get("model_extraction", MODEL_CANDIDATES[0]),
    )


def prepare_split(cfg: RunConfig, split_parts: int) -> RunConfig:
    base_name, made, used_sheet = split_excel(cfg.base_excel, Path(cfg.parts_dir), cfg.sheet_name, split_parts)
    cfg.base_name = base_name
    cfg.sheet_name = used_sheet
    cfg.part_count = made
    st.session_state["_base_name"] = base_name
    st.session_state["_sheet_name"] = used_sheet
    st.session_state["_part_count"] = made
    return cfg


def init_state() -> None:
    if "_session_id" not in st.session_state:
        st.session_state["_session_id"] = uuid.uuid4().hex
        cleanup_stale_session_dirs()
    is_cloud = is_streamlit_cloud()
    output_root = default_output_root(APP_DIR, is_cloud)
    defaults = {
        "workspace": str(APP_DIR),
        "python_exe": norm(sys.executable),
        "mode": "screening",
        "_is_cloud": is_cloud,
        # backing keys — not bound to widgets, safe to update programmatically
        "_base_name": "",
        "_sheet_name": "Sheet1",
        "_part_count": 0,
        "_uploaded_base_excel_path": "",
        "_uploaded_prompt1_path": "",
        "_uploaded_prompt2_path": "",
        "_uploaded_token": "",
        "_output_root": output_root,
        "_output_root_widget": output_root,
        "_model_options": list(MODEL_CANDIDATES),
        "_model_status": "静的候補を使用中",
        "_model_refresh_notice": "",
        "_notice": "",
        "model_screening": MODEL_CANDIDATES[0],
        "model_extraction": MODEL_CANDIDATES[0],
        "_split_parts": 10,
        "_max_workers": MAX_WORKERS,
        "_downloads": [],
        "last_logs": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ---------------------------------------------------------------------------
# Callbacks (run before the script so widget-bound keys can be updated)
# ---------------------------------------------------------------------------

def _on_output_root_change() -> None:
    text = st.session_state.get("_output_root_widget", "").strip()
    if text:
        st.session_state["_output_root"] = norm(str(Path(text).expanduser()))


def _on_pick_output_root() -> None:
    try:
        selected = pick_directory_dialog(st.session_state.get("_output_root", ""))
    except Exception as exc:
        st.session_state["_notice"] = f"フォルダ選択ダイアログを開けませんでした: {exc}"
        return
    if selected:
        st.session_state["_output_root"] = selected
        st.session_state["_output_root_widget"] = selected


def _on_refresh_models() -> None:
    token, _ = current_auth()
    try:
        fetched = fetch_copilot_models(Path(st.session_state.workspace), token)
    except Exception as exc:
        st.session_state["_model_status"] = f"動的取得失敗（現在の候補を継続）: {exc}"
        return
    if not fetched:
        st.session_state["_model_status"] = "動的取得結果が空のため現在の候補を継続"
        return
    options = ["auto"] + [m for m in fetched if m != "auto"]
    st.session_state["_model_options"] = options
    for key in ("model_screening", "model_extraction"):
        if st.session_state.get(key) not in options:
            st.session_state[key] = options[0]
    st.session_state["_model_status"] = f"動的取得: {len(fetched)}件"
    st.session_state["_model_refresh_notice"] = f"モデル再取得に成功しました（{len(fetched)}件）"


def _show_download(path: Path) -> None:
    st.download_button(
        label=f"{path.name} をダウンロード",
        data=path.read_bytes(),
        file_name=path.name,
        mime=XLSX_MIME,
        icon=":material/download:",
        on_click="ignore",
        key=f"_dl_{path.name}",
    )


def _describe_rc(rc: int) -> str:
    if rc == EXIT_OK:
        return "成功"
    if rc == EXIT_ROW_ERRORS:
        return "完了（一部の行でエラー）"
    return f"失敗 (rc={rc})"


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def render_workspace_tab(ws_dir: Path, is_cloud: bool) -> None:
    st.subheader("ワークスペース設定")
    st.markdown("**実行環境**")
    st.caption("プログラムファイルの保存されているフォルダが自動的に選択されます")
    e1, e2 = st.columns(2)
    with e1:
        st.caption(f"Workspace: {norm(str(ws_dir))}")
    with e2:
        st.caption(f"Python: {st.session_state.python_exe}")

    st.markdown("**出力フォルダ**")
    st.caption("生成物は指定したフォルダ内に保存されます")
    notice = st.session_state.get("_notice", "")
    if notice:
        st.warning(notice)
        st.session_state["_notice"] = ""

    out_text_col, out_btn_col = st.columns([6, 1], vertical_alignment="bottom")
    with out_text_col:
        st.text_input(
            "Output root folder",
            key="_output_root_widget",
            on_change=_on_output_root_change,
            disabled=is_cloud,
            help="生成物は <出力先>/.vscode/copilot_run_metadata.json, <出力先>/parts, <出力先>/outputs に保存されます",
        )
    with out_btn_col:
        st.button(
            "参照",
            key="_pick_output_root",
            on_click=_on_pick_output_root,
            disabled=is_cloud or tk is None or filedialog is None,
        )
    if is_cloud:
        st.caption("Cloudではセッション専用の一時フォルダに出力されます")

    output_root_preview = Path(st.session_state.get("_output_root", ""))
    st.caption(f"Metadata: {norm(str(output_root_preview / '.vscode' / 'copilot_run_metadata.json'))}")
    st.caption(f"Parts: {norm(str(output_root_preview / 'parts'))}")
    st.caption(f"Outputs: {norm(str(output_root_preview / 'outputs'))}")

    sync_token_upload(st.file_uploader("Token file upload", type=["txt"], key="_upload_token"))
    _, auth_label = current_auth()
    if auth_label == "token (Streamlit Secrets)":
        st.success("Copilot token: Streamlit Secrets から読み込み済み")
    elif auth_label == "token (uploaded file)":
        st.success("Copilot token: アップロードされたファイルから読み込み済み（ディスクには保存しません）")
    else:
        st.warning("Copilotトークンが未設定です。設定しない限り実行できません")


    st.markdown("**Copilot SDK・CLI・認証状態確認**")
    if st.button("統合状態を確認"):
        token, _ = current_auth()
        with st.spinner("状態確認中..."):
            status = fetch_copilot_status(ws_dir, token)

        sdk = status.get("sdk") or {}
        if sdk.get("error"):
            st.error(f"SDK: {sdk['error']}")
        elif sdk:
            protocol = sdk.get("protocol")
            if isinstance(protocol, int) and protocol >= MIN_SUPPORTED_SDK_PROTOCOL:
                st.success(f"SDK OK: version={sdk.get('version')}, protocol={protocol}")
            else:
                st.warning(
                    f"SDK: 更新が必要です (version={sdk.get('version')}, protocol={protocol}, "
                    f"要件: >= {MIN_SUPPORTED_SDK_PROTOCOL})"
                )

        cli = status.get("cli") or {}
        if cli:
            st.success(f"CLI OK: version={cli.get('version')}, protocol={cli.get('protocol')}")

        auth = status.get("auth") or {}
        if auth:
            if auth.get("authenticated"):
                st.success(f"認証 OK: {auth.get('login')} ({auth.get('type')})")
            else:
                st.error(f"未認証: {auth.get('message') or 'トークンまたはログイン状態を確認してください'}")

        if status.get("error"):
            st.error(f"Copilot runtime: {status['error']}")


def render_input_section() -> None:
    st.markdown("**入力データ**")
    sync_upload(
        st.file_uploader("分析対象Excel upload", type=["xlsx", "xlsm"], key="_upload_base_excel"),
        "_uploaded_base_excel_path",
    )

    sheet_names: list[str] = []
    sheet_err = ""
    base_excel_path = st.session_state.get("_uploaded_base_excel_path", "").strip()
    if base_excel_path:
        try:
            sheet_names = read_excel_sheet_names(Path(base_excel_path))
        except Exception as exc:
            sheet_err = str(exc)

    current_sheet = st.session_state.get("_sheet_name", "Sheet1")
    mode_now = st.session_state.get("mode", "screening")
    p1_now = st.session_state.get("_uploaded_prompt1_path", "").strip()
    p2_now = st.session_state.get("_uploaded_prompt2_path", "").strip()
    prompt_sheet = detect_sheet_from_prompts(mode_now, p1_now, p2_now)
    sheet_sig = prompt_sheet_signature(mode_now, p1_now, p2_now, base_excel_path)
    if st.session_state.get("_prompt_sheet_signature") != sheet_sig:
        st.session_state["_prompt_sheet_signature"] = sheet_sig
        # Apply only when prompt/Excel changes so a manual sheet choice is kept afterwards.
        if prompt_sheet and prompt_sheet in sheet_names:
            st.session_state["_sheet_name"] = prompt_sheet
            st.session_state["_sheet_name_widget"] = prompt_sheet
            current_sheet = prompt_sheet

    if prompt_sheet and sheet_names:
        if prompt_sheet in sheet_names:
            st.caption(f"Prompt の作業対象シート: 「{prompt_sheet}」")
        else:
            st.warning(f"Prompt の作業対象シート「{prompt_sheet}」が Excel に見つからないため、自動選択しませんでした")

    if sheet_err:
        st.warning(sheet_err)
    elif len(sheet_names) == 1:
        st.session_state["_sheet_name"] = sheet_names[0]
        st.caption(f"Sheet: {sheet_names[0]}（1シートのため自動固定）")
    elif not sheet_names:
        st.caption("Sheet: Base Excel アップロード後に自動設定")
    else:
        if current_sheet not in sheet_names:
            st.session_state["_sheet_name"] = sheet_names[0]
        if st.session_state.get("_sheet_name_widget") not in sheet_names:
            st.session_state["_sheet_name_widget"] = st.session_state["_sheet_name"]

        def _sync_sheet_name() -> None:
            st.session_state["_sheet_name"] = st.session_state["_sheet_name_widget"]

        st.selectbox("Sheet", options=sheet_names, key="_sheet_name_widget", on_change=_sync_sheet_name)

    st.markdown("**分析設定**")
    st.caption("stage 1: screening (ノイズ落とし), stage 2: extraction (stage 1の該当に対して要素抽出)")
    mode_col, _ = st.columns([1, 4])
    with mode_col:
        st.selectbox(
            "Mode",
            ["screening", "extraction", "both"],
            key="mode",
            format_func=lambda m: {
                "screening": "stage 1",
                "extraction": "stage 2",
                "both": "stage 1 + stage 2",
            }.get(m, m),
        )
    mode = st.session_state.mode

    st.markdown("**プロンプト**")
    p1, p2 = st.columns(2)
    with p1:
        sync_upload(
            st.file_uploader("Stage 1 prompt upload", type=["txt"], key="_upload_prompt1"),
            "_uploaded_prompt1_path",
        )
        if not st.session_state.get("_uploaded_prompt1_path") and mode in {"screening", "both"}:
            st.warning("Stage 1 prompt が必要です")
    with p2:
        sync_upload(
            st.file_uploader("Stage 2 prompt upload", type=["txt"], key="_upload_prompt2"),
            "_uploaded_prompt2_path",
        )
        if not st.session_state.get("_uploaded_prompt2_path") and mode in {"extraction", "both"}:
            st.warning("Stage 2 prompt が必要です")

    # The sheet selector renders above the uploaders, so rerun once to apply a newly uploaded prompt's sheet.
    if prompt_sheet_signature(
        mode,
        st.session_state.get("_uploaded_prompt1_path", "").strip(),
        st.session_state.get("_uploaded_prompt2_path", "").strip(),
        base_excel_path,
    ) != st.session_state.get("_prompt_sheet_signature"):
        st.rerun()

    st.markdown("**モデル選択**")
    notice = st.session_state.get("_model_refresh_notice", "")
    if notice:
        st.success(notice)
        st.session_state["_model_refresh_notice"] = ""
    model_options = st.session_state.get("_model_options", MODEL_CANDIDATES)
    for key in ("model_screening", "model_extraction"):
        if st.session_state.get(key) not in model_options:
            st.session_state[key] = model_options[0]
    m1, m2, m_btn = st.columns([2, 2, 1], vertical_alignment="bottom")
    if mode in {"screening", "both"}:
        with m1:
            st.selectbox("Model for Stage 1", model_options, key="model_screening")
    if mode in {"extraction", "both"}:
        with m2 if mode == "both" else m1:
            st.selectbox("Model for Stage 2", model_options, key="model_extraction")
    with m_btn:
        st.button("モデル再取得", on_click=_on_refresh_models)
    st.caption(st.session_state.get("_model_status", ""))

    st.markdown("**入力列**")
    prompt1_path = st.session_state.get("_uploaded_prompt1_path", "").strip()
    prompt2_path = st.session_state.get("_uploaded_prompt2_path", "").strip()
    slot_count, slot_defaults = detect_placeholder_defaults(mode, prompt1_path, prompt2_path)
    sheet_for_cols = st.session_state.get("_sheet_name", "Sheet1")
    excel_for_cols = st.session_state.get("_uploaded_base_excel_path", "").strip()
    headers = read_headers_safe(Path(excel_for_cols), sheet_for_cols) if excel_for_cols else None

    if slot_count == 0:
        detected_cols = detect_input_columns_from_prompts(mode, prompt1_path, prompt2_path)
        st.warning("Prompt に [入力n] が見つからないため、入力列を選択できません")
        if detected_cols:
            st.caption("Auto input columns: " + ", ".join(detected_cols))
    elif headers is None:
        st.caption("分析対象Excel をアップロードすると、[入力n] ごとに列を選択できます")
        auto_text = ", ".join(f"[入力{n}]={slot_defaults.get(n, '（未検出）')}" for n in range(1, slot_count + 1))
        st.caption("Auto input columns: " + auto_text)
    else:
        excel_stat = Path(excel_for_cols).stat()
        signature = (
            mode,
            _file_digest(prompt1_path),
            _file_digest(prompt2_path),
            excel_for_cols,
            excel_stat.st_size,
            excel_stat.st_mtime_ns,
            sheet_for_cols,
        )
        if st.session_state.get("_input_col_signature") != signature:
            st.session_state["_input_col_signature"] = signature
            resolved = resolve_detected_columns_to_headers(
                [slot_defaults.get(n, "") for n in range(1, slot_count + 1)], headers
            )
            for n, col in enumerate(resolved, start=1):
                key = INPUT_SLOT_KEY.format(n)
                if col in headers:
                    st.session_state[key] = col
                else:
                    st.session_state.pop(key, None)

        st.caption("[入力n] ごとに使用するExcel列を選択してください（初期値は Prompt から自動検出）")
        slot_cols = st.columns(min(slot_count, 3))
        for n in range(1, slot_count + 1):
            key = INPUT_SLOT_KEY.format(n)
            if key in st.session_state and st.session_state[key] not in headers:
                st.session_state.pop(key)
            extra = {} if key in st.session_state else {"index": None}
            auto = slot_defaults.get(n)
            with slot_cols[(n - 1) % len(slot_cols)]:
                st.selectbox(
                    f"[入力{n}]",
                    headers,
                    key=key,
                    placeholder="列を選択",
                    help=f"Prompt から検出: {auto}" if auto else "Prompt から列名を検出できませんでした",
                    **extra,
                )

    st.markdown("**チェック**")
    if st.button("Preflight check"):
        cfg = get_config_from_ui()
        errors = validate_config(cfg)
        if errors:
            st.error("\n".join(f"- {item}" for item in errors))
        else:
            used_sheet, excel_headers = read_excel_headers(cfg.base_excel, cfg.sheet_name)
            st.session_state["_sheet_name"] = used_sheet
            st.session_state["_base_name"] = cfg.base_excel.stem
            cfg.base_name = cfg.base_excel.stem
            st.success("設定チェックOK")
            st.caption(f"ヘッダ確認: シート={used_sheet}, 列数={len(excel_headers)}, 認証={cfg.auth_label}")
            st.code(" ".join(build_part_command(cfg, 1)), language="bash")


def _run_all_parts(cfg: RunConfig, part_numbers: list[int], workers: int, stop_on_error: bool) -> dict[str, Any]:
    results: dict[str, Any] = {}
    progress = st.progress(0.0)
    status = st.empty()
    history: list[str] = []
    env = child_env(cfg.token)

    def push_status(msg: str) -> None:
        history.append(msg)
        # Keep the latest lines visible while processing many parts.
        status.markdown("  \n".join(history[-30:]))

    def record(n: int, result: dict[str, Any]) -> None:
        results[f"part{n}"] = result
        progress.progress(len(results) / len(part_numbers))
        push_status(f"part{n}: {_describe_rc(result['returncode'])}")

    if workers <= 1:
        for n in part_numbers:
            push_status(f"part{n}: 実行中...")
            result = run_command(build_part_command(cfg, n), cfg.workspace, env)
            record(n, result)
            if stop_on_error and result["returncode"] not in (EXIT_OK, EXIT_ROW_ERRORS):
                push_status("エラーのため以降のPartを中止しました")
                break
    else:
        with ThreadPoolExecutor(max_workers=workers) as exe:
            futures = {}
            for n in part_numbers:
                push_status(f"part{n}: 実行中...")
                futures[exe.submit(run_command, build_part_command(cfg, n), cfg.workspace, env)] = n
            for fut in as_completed(futures):
                record(futures[fut], fut.result())
    return results


def render_execute_section(is_cloud: bool) -> None:
    st.subheader("実行")
    cfg = get_config_from_ui()

    c1, c2 = st.columns(2)
    with c1:
        st.number_input(
            "分割数 (Parts)",
            min_value=1,
            max_value=MAX_SPLIT_PARTS,
            key="_split_parts",
            help="実行前に分析対象Excelをこの数に自動分割します（データ行数が少ない場合は行数まで）",
        )
    with c2:
        st.slider(
            "Parallel workers",
            min_value=1,
            max_value=MAX_WORKERS,
            key="_max_workers",
            disabled=is_cloud,
            help="同時に実行するPart数",
        )
    if is_cloud:
        st.caption("Streamlit Cloud では逐次実行を使用します")
    stop_on_error = st.checkbox("逐次実行時にエラーで停止", value=True)
    split_parts = int(st.session_state["_split_parts"])
    workers = 1 if is_cloud else int(st.session_state["_max_workers"])

    st.markdown("Part to run (分割したファイルの一部を実行)")
    part_col, run_col, _ = st.columns([1.25, 1.15, 3.6])
    with part_col:
        selected_part = st.selectbox(
            "Part to run",
            options=list(range(1, split_parts + 1)),
            index=0,
            label_visibility="collapsed",
        )
    with run_col:
        run_selected = st.button("Run selected part", width="stretch")

    st.markdown("Run all parts (分割したファイルの全てを実行し結合する)")
    run_all = st.button("Run all parts", width="stretch")

    action = "selected" if run_selected else "all" if run_all else ""
    if action:
        st.session_state["_downloads"] = []
        errors = validate_config(cfg)
        if errors:
            st.error("\n".join(f"- {item}" for item in errors))
            return
        try:
            cfg = prepare_split(cfg, split_parts)
        except Exception as exc:
            st.error(f"Split失敗: {exc}")
            return
        part_numbers = list(range(1, cfg.part_count + 1))
        if not part_numbers:
            st.error("実行対象のPartがありません（データ行がありません）")
            return
        st.info(f"Split完了: {len(part_numbers)} parts")
        write_run_files(cfg)

        if action == "selected":
            if selected_part not in part_numbers:
                st.error(f"選択したPart{selected_part}は存在しません。利用可能: {part_numbers}")
                return
            with st.spinner(f"part{selected_part}: 実行中..."):
                result = run_command(build_part_command(cfg, selected_part), cfg.workspace, child_env(cfg.token))
            st.session_state.last_logs = {f"part{selected_part}": result}
            rc = result["returncode"]
            if rc == EXIT_OK:
                st.success(f"part{selected_part} 成功")
            elif rc == EXIT_ROW_ERRORS:
                st.warning(f"part{selected_part} 完了（一部の行でエラー。1_Judgment / 2_Status 列の ERROR を確認してください）")
            else:
                st.error(f"part{selected_part} 失敗 (rc={rc})")
            output = marker_path(result["stdout"], OUTPUT_MARKER)
            if output:
                st.session_state["_downloads"] = [str(output)]
        else:
            results = _run_all_parts(cfg, part_numbers, workers, stop_on_error)
            ok_codes = {EXIT_OK, EXIT_ROW_ERRORS}
            failed = [k for k, v in results.items() if v["returncode"] not in ok_codes]
            not_run = [f"part{n}" for n in part_numbers if f"part{n}" not in results]
            partial = [k for k, v in results.items() if v["returncode"] == EXIT_ROW_ERRORS]
            if failed or not_run:
                if failed:
                    st.error(f"失敗: {', '.join(sorted(failed))}")
                if not_run:
                    st.error(f"未実行: {', '.join(not_run)}")
                st.warning("失敗・未実行のPartがあるためMergeは実行しません")
            else:
                if partial:
                    st.warning(
                        f"一部の行でエラー: {', '.join(sorted(partial))}"
                        "（結合結果の 1_Judgment / 2_Status 列の ERROR を確認してください）"
                    )
                merge_result = run_command(build_merge_command(cfg), cfg.workspace, child_env(cfg.token))
                results["merge"] = merge_result
                if merge_result["returncode"] == EXIT_OK:
                    st.success("全part完了 + merge成功")
                    merged = marker_path(merge_result["stdout"], MERGE_DONE_MARKER)
                    if merged:
                        st.session_state["_downloads"] = [str(merged)]
                    else:
                        st.warning("結合ファイルのパスを取得できませんでした。Latest logs を確認してください")
                else:
                    st.error("全part完了 / merge失敗")
            st.session_state.last_logs = results

    for path_text in st.session_state.get("_downloads", []):
        path = Path(path_text)
        if path.is_file():
            _show_download(path)

    if st.session_state.last_logs:
        with st.expander("Latest logs", expanded=False):
            for key in sorted(st.session_state.last_logs):
                log = st.session_state.last_logs[key]
                st.markdown(f"**{key}**")
                st.write(f"returncode: {log['returncode']}")
                st.code(log["stdout"] or "(no stdout)")
                if log["stderr"]:
                    st.code(log["stderr"])


def main() -> None:
    st.set_page_config(page_title="Copilot SDK Runner", layout="wide")
    init_state()
    is_cloud = bool(st.session_state.get("_is_cloud", False))
    ws_dir = Path(st.session_state.workspace)

    st.title("Copilot SDK Streamlit Runner")
    if is_cloud:
        st.caption("Streamlit Cloud モード: ヘッドレス実行 + Secrets ベースの認証を推奨します")
    else:
        st.caption("ローカル環境でCopilot SDK/CLIを実行します")

    tab_wsp, tab_cfg = st.tabs(["Workspace", "Execute"])
    with tab_wsp:
        render_workspace_tab(ws_dir, is_cloud)
    with tab_cfg:
        st.subheader("実行設定")
        render_input_section()
        render_execute_section(is_cloud)


if __name__ == "__main__":
    main()

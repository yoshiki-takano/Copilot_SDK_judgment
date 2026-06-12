# Copilot SDK 実行フォルダ共有ガイド

## 1. 配布時の基本方針

- 推奨配布方法: Git リポジトリ共有
- 代替配布方法: ZIP 配布（`.venv` や出力ファイルは除外）
- 機密情報（トークン等）は配布物に含めない

## 2. 受け取り側の初回セットアップ

Python 3.11 以上を用意し、このフォルダで以下を実行します。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.copilot.txt
```

既存の社内 Python を使う場合は、その環境に `requirements.copilot.txt` をインストールしてください。

### 2.1 前提条件（重要）

- `github-copilot-sdk` は必須
	- 本プロジェクトは `Copilot_Calling.py` で Copilot SDK クライアントを使用します。
- Copilot CLI は原則必須（推奨）
	- モデル取得・実行時に CLI 連携を前提とした経路を使用します。
	- 通常は `copilot` / `copilot.exe` が PATH 上にある状態を想定します。
- GitHub Copilot が利用可能なアカウント権限が必要

確認コマンド例:

```powershell
.\.venv\Scripts\python.exe -c "from copilot import CopilotClient; print('copilot sdk ok')"
copilot --version
```

`copilot --version` が失敗する場合は、GitHub Copilot CLI のインストール/パス設定を確認してください。

## 3. トークンの扱い

- `github_copilot_token.txt` は共有しない
- 受け取り側が自分の権限のトークンファイルをローカルで用意する
- Streamlit UI では `Token file upload` で指定する
- `run_from_env.ps1` 利用時は `.env` の `TOKEN_FILE` を指定する

## 4. 実行方法（推奨: Streamlit UI）

```powershell
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

UI で実行する場合の要点:

- `Config` タブで Excel / Prompt / Token / 出力先を指定
- `モデル再取得` でモデル候補を更新
- `Execute` タブで `Run selected part` または `Run all parts` を実行

## 5. 実行方法（CLI ラッパー: run_from_env.ps1）

`.env.example` を `.env` にコピーして値を編集します。

```powershell
Copy-Item .env.example .env
```

主な設定項目:

- `PYTHON_EXE`: 使用 Python
- `TOKEN_FILE`: トークンファイルパス
- `OUTPUT_DIR`: 出力先
- `BASE_NAME`: `parts` のベース名
- `PARTS`: パート数
- `MODE`: `screening` / `extraction` / `both`
- `MODEL_STAGE1`, `MODEL_STAGE2`: モデルID

実行例:

```powershell
# 1パート実行
.\run_from_env.ps1 -Part 1

# 全パート実行
.\run_from_env.ps1 -RunAll

# マージのみ
.\run_from_env.ps1 -MergeOnly
```

## 6. 共有前チェックリスト

- `github_copilot_token.txt` や `.env` を含めていない
- `.venv` / `out` / `parts` / `outputs` / `.streamlit_runtime` を含めていない
- 受け取り側向けに `requirements.copilot.txt` が同梱されている
- 起動確認コマンドを README に明記している

## 7. 主なファイル

- `streamlit_app.py`: メインUI
- `main_parallel.py`: Stage1+Stage2 実行
- `step1only_main_parallel.py`: Stage1 実行
- `step2only_main_parallel.py`: Stage2 実行
- `Copilot_Calling.py`: Copilot SDK 呼び出し
- `rev_JSON_Fillout in Excel.py`: JSON 結果の Excel 展開
- `run_from_env.ps1`: `.env` ベースの CLI 実行
- `requirements.copilot.txt`: 依存関係

## 8. 配布用ZIPの含める/除外する一覧

### 8.1 含める（ZIPに入れる）

- `streamlit_app.py`
- `main_parallel.py`
- `step1only_main_parallel.py`
- `step2only_main_parallel.py`
- `Copilot_Calling.py`
- `rev_JSON_Fillout in Excel.py`
- `make_tasks_json.py`
- `run_from_env.ps1`
- `requirements.copilot.txt`
- `.env.example`
- `README_共有用.md`
- `schemas/`
- `manual_assets/`（必要時のみ）
- `テスト用データ/`（検証用に共有する場合のみ）

### 8.2 除外する（ZIPに入れない）

- `.venv/`
- `venv/`
- `env/`
- `__pycache__/`
- `.streamlit_runtime/`
- `out/`
- `outputs/`
- `parts/`
- `.git/`
- `.vscode/`
- `.env`
- `github_copilot_token.txt`
- `copilot_run_metadata.json`
- `Gemini_Calling.py`

### 8.3 PowerShellでZIPを作る例

以下は、配布対象だけを一時フォルダへ集めてから ZIP 化する例です。

```powershell
$root = Get-Location
$stage = Join-Path $root "_share_stage"
$zip = Join-Path $root "Copilot_SDK_share.zip"

if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
if (Test-Path $zip) { Remove-Item $zip -Force }
New-Item -ItemType Directory -Path $stage | Out-Null

$include = @(
	"streamlit_app.py",
	"main_parallel.py",
	"step1only_main_parallel.py",
	"step2only_main_parallel.py",
	"Copilot_Calling.py",
	"Gemini_Calling.py",
	"rev_JSON_Fillout in Excel.py",
	"make_tasks_json.py",
	"run_from_env.ps1",
	"requirements.copilot.txt",
	".env.example",
	"README_共有用.md",
	"schemas",
	"manual_assets"
)

foreach ($item in $include) {
	if (Test-Path $item) {
		Copy-Item $item -Destination $stage -Recurse -Force
	}
}

Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zip
Remove-Item $stage -Recurse -Force

Write-Host "Created: $zip"
```

### 8.4 配布前の最終確認

- ZIP内に `.env` と `github_copilot_token.txt` が入っていない
- ZIP内に `out/`, `parts/`, `outputs/` が入っていない
- 受け取り側が `README_共有用.md` の手順だけで起動できる

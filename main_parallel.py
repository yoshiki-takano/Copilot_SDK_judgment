# Stage 1（スクリーニング）+ Stage 2（詳細抽出）を分割済みExcel 1本に対して実行する。
# Stage 2 は Stage 1 で該当（True）となった行だけに実行される。
# オプションは pipeline_common.parse_args を参照（--merge-only / --merge-after 等）。
import sys

from pipeline_common import main

if __name__ == "__main__":
    sys.exit(main(run_stage1=True, run_stage2=True))

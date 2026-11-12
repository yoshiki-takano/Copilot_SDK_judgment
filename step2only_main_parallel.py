# Stage 2（詳細抽出）のみを分割済みExcel 1本の全行に対して実行する。
# オプションは pipeline_common.parse_args を参照（--merge-only / --merge-after 等）。
import sys

from pipeline_common import main

if __name__ == "__main__":
    sys.exit(main(run_stage1=False, run_stage2=True))

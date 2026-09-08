#!/usr/bin/env python3
"""Code Agent 统一入口：导出最新数据、执行翻译、按需生成上传格式。"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


BASE = Path(os.environ.get("I18N_WORKDIR") or Path(__file__).resolve().parents[1]).resolve()
CODE_DIR = BASE / "04_代码"
RAW_DIR = BASE / "01_原始数据"


def run_step(args: list[str]) -> None:
    print("+ " + " ".join(args), flush=True)
    subprocess.run(args, cwd=str(BASE), check=True)


def snapshot_xlsx() -> set[Path]:
    if not RAW_DIR.exists():
        return set()
    return {path.resolve() for path in RAW_DIR.glob("*.xlsx") if not path.name.startswith("~$")}


def export_files(type_id: int | None, empty_langs: list[str] | None) -> list[Path]:
    before = snapshot_xlsx()
    cmd = [sys.executable, str(CODE_DIR / "export_from_api.py")]
    if type_id:
        cmd += ["--type", str(type_id)]
    if empty_langs:
        cmd += ["--empty-langs", *empty_langs]
    run_step(cmd)
    after = snapshot_xlsx()
    new_files = sorted(after - before, key=lambda path: path.stat().st_mtime)
    if not new_files:
        raise RuntimeError("导出完成但未发现新增 Excel，请检查导出日志和 01_原始数据 目录")
    return new_files


def translate_files(files: list[Path], args: argparse.Namespace) -> None:
    cmd = [
        sys.executable,
        str(CODE_DIR / "translate_multi_site.py"),
        *[str(path) for path in files],
        "--langs",
        *args.langs,
        "--engine",
        args.engine,
        "--dify-workers",
        str(args.dify_workers),
        "--google-workers",
        str(args.google_workers),
        "--google-interval",
        str(args.google_interval),
    ]
    if args.all:
        cmd.append("--all")
    run_step(cmd)


def build_uploads(langs: list[str], run_date: str | None) -> None:
    if not run_date:
        raise RuntimeError("生成上传格式需要 --run-date，且 01_原始数据 中必须有对应日期的全量导出文件")
    for lang in langs:
        run_step([
            sys.executable,
            str(CODE_DIR / "build_upload_for_lang_20260612.py"),
            "--lang",
            lang,
            "--run-date",
            run_date,
        ])


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "action",
        choices=["export", "translate", "export-translate", "build-upload", "all"],
        help="执行阶段：导出、翻译、导出后翻译、生成上传格式，或全流程",
    )
    parser.add_argument("files", nargs="*", help="translate 阶段的输入 Excel；不传则使用 01_原始数据 下现有文件")
    parser.add_argument("--langs", nargs="+", required=True, help="目标语种列名，如 de fr lv")
    parser.add_argument("--type", type=int, choices=[1, 2, 3], help="导出文案类型：1=前端key，2=提示语，3=错误码")
    parser.add_argument("--empty-langs", nargs="+", help="只导出这些语种为空的 key；通常与 --langs 相同")
    parser.add_argument("--all", action="store_true", help="全量翻译，不只补空")
    parser.add_argument("--engine", choices=["hybrid", "google", "dify"], default="hybrid")
    parser.add_argument("--dify-workers", type=int, default=4)
    parser.add_argument("--google-workers", type=int, default=2)
    parser.add_argument("--google-interval", type=float, default=0.8)
    parser.add_argument("--run-date", help="生成上传格式使用的源文件日期，如 20260907")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    BASE.mkdir(parents=True, exist_ok=True)

    if args.action in {"export", "export-translate", "all"}:
        exported = export_files(args.type, args.empty_langs)
        print("本次导出文件：", flush=True)
        for path in exported:
            print(f"  {path}", flush=True)
    else:
        exported = []

    if args.action in {"translate", "export-translate", "all"}:
        if args.action == "translate":
            files = [Path(path).resolve() for path in args.files] if args.files else sorted(snapshot_xlsx())
        else:
            files = exported
        if not files:
            raise RuntimeError("没有可翻译的 Excel 文件")
        translate_files(files, args)

    if args.action in {"build-upload", "all"}:
        build_uploads(args.langs, args.run_date)

    print("执行完成", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())

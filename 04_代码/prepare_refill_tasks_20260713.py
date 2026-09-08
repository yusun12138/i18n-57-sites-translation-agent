#!/usr/bin/env python3
"""从最新全量 SHEIN 导出中生成本次补翻任务。

规则：
- 只处理 en 有值且目标语种为空的行。
- 第 1 列 id 是系统字段；后面的 id 才是印尼语语种列。
- en/en_au/en_gb/zh-hk/zh 不作为本次目标语种。
- nb 仅生成缺失清单，后续不自动翻译。
"""

import csv
import json
import os
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl


BASE = Path(os.environ.get("I18N_WORKDIR") or Path(__file__).resolve().parents[1]).resolve()
RAW = BASE / "01_原始数据"
PROCESS = BASE / "02_过程数据"
TASK_DIR = PROCESS / "refill_tasks_by_lang_20260713"
EXCLUDED = BASE / "04_代码" / "excluded_keyids.txt"
RUN_DATE = "20260713"
RAW_TS = "20260713_092731"

TARGET_LANGS = [
    "nb",
    "ru",
    "ar",
    "lt",
    "id",
    "el-gr",
    "nl",
    "sl",
    "it",
    "hu",
    "tr",
    "pt-pt",
    "lv",
    "uk",
    "pl",
    "ro",
    "es",
    "fr",
    "sr",
    "de",
    "et",
    "zh-cn",
    "zh-tw",
]


def norm(value):
    return "" if value is None else str(value).strip()


def is_empty(value):
    return norm(value).lower() in {"", "none", "null"}


def load_excluded():
    if not EXCLUDED.exists():
        return set()
    return {
        line.strip()
        for line in EXCLUDED.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    }


def latest(pattern):
    files = sorted(RAW.glob(pattern), key=lambda path: path.stat().st_mtime)
    if not files:
        raise FileNotFoundError(pattern)
    return files[-1]


def input_files():
    files = [
        latest(f"type1_前端key_platform{platform}_full_{RAW_TS}.xlsx")
        for platform in [1, 2, 3, 4, 5]
    ]
    files.append(latest(f"type2_提示语_full_{RAW_TS}.xlsx"))
    files.append(latest(f"type3_错误码_full_{RAW_TS}.xlsx"))
    return files


def language_indexes(headers):
    indexes = {}
    for idx, header in enumerate(headers):
        if idx == 0 and header == "id":
            continue
        if header in TARGET_LANGS:
            indexes[header] = idx
    missing_headers = [lang for lang in TARGET_LANGS if lang not in indexes]
    if missing_headers:
        raise ValueError(f"导出表缺少目标语种列: {missing_headers}")
    return indexes


def main():
    excluded = load_excluded()
    tasks_by_lang = {lang: {} for lang in TARGET_LANGS}
    file_stats = []
    excluded_hits = Counter()
    rows_with_en = 0
    total_rows = 0

    for path in input_files():
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        ws.reset_dimensions()
        rows = ws.iter_rows(values_only=True)
        headers = [norm(value) for value in next(rows)]
        lower_headers = [header.lower() for header in headers]
        key_idx = next(i for i, header in enumerate(headers) if header in {"keyId", "key ID", "key_id"})
        en_idx = lower_headers.index("en")
        lang_idx = language_indexes(headers)

        file_counter = Counter()
        file_total = 0
        file_with_en = 0

        for row in rows:
            total_rows += 1
            file_total += 1
            key = norm(row[key_idx] if key_idx < len(row) else "")
            en = norm(row[en_idx] if en_idx < len(row) else "")
            if not key or not en:
                continue
            rows_with_en += 1
            file_with_en += 1
            if key in excluded:
                for lang, idx in lang_idx.items():
                    value = row[idx] if idx < len(row) else None
                    if is_empty(value):
                        excluded_hits[lang] += 1
                continue

            for lang, idx in lang_idx.items():
                value = row[idx] if idx < len(row) else None
                if not is_empty(value):
                    continue
                tasks_by_lang[lang][key] = {
                    "keyId": key,
                    "en": en,
                    "source_file": path.name,
                }
                file_counter[lang] += 1

        file_stats.append(
            {
                "file": path.name,
                "rows": file_total,
                "rows_with_en": file_with_en,
                "missing_by_lang": dict(sorted(file_counter.items())),
            }
        )

    PROCESS.mkdir(parents=True, exist_ok=True)
    TASK_DIR.mkdir(parents=True, exist_ok=True)

    summary = {}
    task_files = {}
    all_rows = []
    for lang in TARGET_LANGS:
        rows = sorted(tasks_by_lang[lang].values(), key=lambda item: item["keyId"])
        task_path = TASK_DIR / f"{lang}_refill_tasks_{RUN_DATE}.json"
        task_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        task_files[lang] = str(task_path)
        summary[lang] = {
            "missing_rows": len(rows),
            "split_file_count_if_25000": (len(rows) + 24999) // 25000,
            "excluded_missing_rows": excluded_hits[lang],
        }
        for item in rows:
            all_rows.append(
                {
                    "lang": lang,
                    "keyId": item["keyId"],
                    "en": item["en"],
                    "source_file": item["source_file"],
                }
            )

    csv_path = PROCESS / f"refill_tasks_{RUN_DATE}.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["lang", "keyId", "en", "source_file"])
        writer.writeheader()
        writer.writerows(all_rows)

    report = {
        "run_date": RUN_DATE,
        "raw_timestamp": RAW_TS,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_files": [str(path) for path in input_files()],
        "target_langs": TARGET_LANGS,
        "excluded_key_count": len(excluded),
        "total_rows": total_rows,
        "rows_with_en": rows_with_en,
        "summary": summary,
        "file_stats": file_stats,
        "task_files": task_files,
        "task_csv": str(csv_path),
    }
    report_path = PROCESS / f"refill_task_report_{RUN_DATE}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(
        json.dumps(
            {
                "report": str(report_path),
                "task_csv": str(csv_path),
                "summary": summary,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()

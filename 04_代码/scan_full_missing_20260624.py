#!/usr/bin/env python3
"""扫描最新全量导出：统计有 en 但语种为空的 key。"""

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
EXCLUDED = BASE / "04_代码" / "excluded_keyids.txt"
RUN_DATE = "20260624"

META_COLS = {
    "id",
    "keyId",
    "en",
    "所属产品",
    "备注",
    "产品创建人",
    "key创建人",
    "key最后修改人",
}


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


def latest(path_pattern):
    files = sorted(RAW.glob(path_pattern), key=lambda p: p.stat().st_mtime)
    if not files:
        raise FileNotFoundError(path_pattern)
    return files[-1]


def latest_input_files():
    files = [
        latest(f"type1_前端key_platform{platform}_full_{RUN_DATE}_*.xlsx")
        for platform in [1, 2, 3, 4, 5]
    ]
    files.append(latest(f"type2_提示语_full_{RUN_DATE}_*.xlsx"))
    type3_files = sorted(
        list(RAW.glob(f"type3_错误码_full_{RUN_DATE}_*.xlsx"))
        + list(RAW.glob(f"type3_错误码_{RUN_DATE}_*.xlsx")),
        key=lambda p: p.stat().st_mtime,
    )
    if not type3_files:
        raise FileNotFoundError(f"type3_错误码*{RUN_DATE}*.xlsx")
    files.append(type3_files[-1])
    return files


def main():
    excluded = load_excluded()
    files = latest_input_files()

    file_stats = []
    lang_missing = Counter()
    lang_missing_excluding_excluded = Counter()
    lang_row_missing = Counter()
    key_missing_counts = Counter()
    missing_rows = []
    lang_samples = defaultdict(list)
    total_rows = 0
    rows_with_en = 0
    rows_with_any_missing = 0
    rows_with_any_missing_excluding_excluded = 0
    missing_cells = 0
    missing_cells_excluding_excluded = 0
    excluded_missing_rows = 0
    all_langs = []

    for path in files:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        ws.reset_dimensions()
        iterator = ws.iter_rows(values_only=True)
        headers = [norm(v) for v in next(iterator)]
        lower_headers = [h.lower() for h in headers]
        key_idx = next(i for i, h in enumerate(headers) if h in {"keyId", "key ID", "key_id"})
        en_idx = lower_headers.index("en")

        lang_indexes = []
        for idx, header in enumerate(headers):
            if idx == 0 and header == "id":
                continue
            if not header or header in META_COLS:
                continue
            lang_indexes.append((idx, header))
        if not all_langs:
            all_langs = [lang for _, lang in lang_indexes]

        file_rows = 0
        file_rows_with_en = 0
        file_missing_cells = 0
        file_rows_with_any_missing = 0
        file_lang_missing = Counter()

        for row in iterator:
            file_rows += 1
            total_rows += 1
            key = norm(row[key_idx] if key_idx < len(row) else "")
            en = norm(row[en_idx] if en_idx < len(row) else "")
            if not key or not en:
                continue
            rows_with_en += 1
            file_rows_with_en += 1
            missing_langs = [
                lang
                for idx, lang in lang_indexes
                if idx >= len(row) or is_empty(row[idx])
            ]
            if not missing_langs:
                continue

            rows_with_any_missing += 1
            file_rows_with_any_missing += 1
            missing_count = len(missing_langs)
            missing_cells += missing_count
            file_missing_cells += missing_count
            key_missing_counts[key] += missing_count
            file_lang_missing.update(missing_langs)
            lang_missing.update(missing_langs)
            for lang in missing_langs:
                lang_row_missing[lang] += 1
                if len(lang_samples[lang]) < 20:
                    lang_samples[lang].append(
                        {
                            "file": path.name,
                            "keyId": key,
                            "en": en[:220],
                        }
                    )

            is_excluded = key in excluded
            if is_excluded:
                excluded_missing_rows += 1
            else:
                rows_with_any_missing_excluding_excluded += 1
                missing_cells_excluding_excluded += missing_count
                lang_missing_excluding_excluded.update(missing_langs)

            missing_rows.append(
                {
                    "file": path.name,
                    "keyId": key,
                    "en": en,
                    "missing_count": missing_count,
                    "missing_langs": "|".join(missing_langs),
                    "excluded": "Y" if is_excluded else "",
                }
            )

        file_stats.append(
            {
                "file": path.name,
                "rows": file_rows,
                "rows_with_en": file_rows_with_en,
                "rows_with_any_missing": file_rows_with_any_missing,
                "missing_cells": file_missing_cells,
                "missing_by_lang": dict(sorted(file_lang_missing.items())),
            }
        )

    missing_csv = PROCESS / f"full_missing_rows_{RUN_DATE}.csv"
    with open(missing_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["file", "keyId", "en", "missing_count", "missing_langs", "excluded"],
        )
        writer.writeheader()
        writer.writerows(missing_rows)

    report = {
        "run_date": RUN_DATE,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "input_files": [str(path) for path in files],
        "language_count": len(all_langs),
        "languages": all_langs,
        "summary": {
            "total_rows": total_rows,
            "rows_with_en": rows_with_en,
            "rows_with_any_missing": rows_with_any_missing,
            "missing_cells": missing_cells,
            "excluded_missing_rows": excluded_missing_rows,
            "rows_with_any_missing_excluding_excluded": rows_with_any_missing_excluding_excluded,
            "missing_cells_excluding_excluded": missing_cells_excluding_excluded,
        },
        "missing_by_lang": dict(lang_missing.most_common()),
        "missing_by_lang_excluding_excluded": dict(lang_missing_excluding_excluded.most_common()),
        "top_keys_by_missing_count": key_missing_counts.most_common(50),
        "file_stats": file_stats,
        "samples_by_lang": dict(lang_samples),
        "missing_rows_csv": str(missing_csv),
    }

    report_path = PROCESS / f"full_missing_report_{RUN_DATE}.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({
        "report": str(report_path),
        "missing_rows_csv": str(missing_csv),
        "summary": report["summary"],
        "top_missing_langs": dict(lang_missing.most_common(20)),
        "top_missing_langs_excluding_excluded": dict(lang_missing_excluding_excluded.most_common(20)),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

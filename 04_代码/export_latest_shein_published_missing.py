#!/usr/bin/env python3
"""Export latest SHEIN published copy keys and scan non-zh empty translations."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
import requests


BASE_DIR = Path(os.environ.get("I18N_WORKDIR") or Path(__file__).resolve().parents[1]).resolve()
RAW_ROOT = BASE_DIR / "01_原始数据" / "latest_shein_published"
PROCESS_ROOT = BASE_DIR / "02_过程数据" / "latest_shein_published"

MCP_PROXY = os.environ.get("MCP_PROXY") or "http://127.0.0.1:30001/rest"
GW_NAME = os.environ.get("LCC_GW_NAME") or "language-configuration-center"
MCP_DC = os.environ.get("MCP_PROXY_DC") or "central_p2_dc"
ULP_COOKIE = os.environ.get("ULP_COOKIE") or "ulp_token={{ulp_token}}"

KEYS_PATH = f"/{GW_NAME}/headless/api/v1/keys"
EXPORT_PATH = f"/{GW_NAME}/headless/api/v1/keys/export"
EXPORT_TASK_PATH = f"/{GW_NAME}/headless/api/v1/keys/export-task"

POLL_MAX = 120
POLL_INTERVAL = 5

# "Already published" includes published to prod and gray-published content.
# Guide enum: 1=已发布, 2=已发布灰度, 3=未发布, 4=已更新未发布.
DEFAULT_RELEASE_STATUSES = [1, 2]

TASKS = [
    {"tab": "前端key", "platform_label": "APP", "type": 1, "brand": 1, "platform": [1]},
    {"tab": "前端key", "platform_label": "H5", "type": 1, "brand": 1, "platform": [2]},
    {"tab": "前端key", "platform_label": "PWA", "type": 1, "brand": 1, "platform": [3]},
    {"tab": "前端key", "platform_label": "PC", "type": 1, "brand": 1, "platform": [4]},
    {"tab": "前端key", "platform_label": "PAY", "type": 1, "brand": 1, "platform": [5]},
    {"tab": "提示语", "platform_label": "B2C", "type": 2, "brand": 2, "platform": [6]},
    {"tab": "提示语", "platform_label": "社区", "type": 2, "brand": 2, "platform": [7]},
    {"tab": "错误码", "platform_label": "SHEIN", "type": 3, "brand": 3, "platform": [12]},
]

META_HEADERS = {
    "id",
    "keyid",
    "key id",
    "key_id",
    "en",
    "zh",
    "所属产品",
    "备注",
    "产品创建人",
    "key创建人",
    "key最后修改人",
    "创建人",
    "修改人",
    "发布状态",
    "同步状态",
    "平台",
    "品牌",
    "类型",
}

LANG_HEADER_RE = re.compile(r"^[a-z]{2}(?:[-_][a-z]{2})?$")


def log(message: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}", flush=True)


def norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_empty(value: Any) -> bool:
    return norm(value).lower() in {"", "none", "null", "nan"}


def headers(proxy_path: str) -> dict[str, str]:
    return {
        "Content-Type": "application/json",
        "mcp-proxy-path": proxy_path,
        "mcp-proxy-dc": MCP_DC,
        "Cookie": ULP_COOKIE,
    }


def api_count(task: dict[str, Any], release_statuses: list[int]) -> int:
    payload = {
        "page": 1,
        "pageSize": 1,
        "searchType": 2,
        "type": task["type"],
        "brand": task["brand"],
        "platform": task["platform"],
        "keyword": "",
        "keyId": [],
        "projects": [],
        "emptyLangList": [],
        "syncStatus": [],
        "releaseStatus": release_statuses,
    }
    response = requests.post(
        MCP_PROXY,
        headers=headers(KEYS_PATH),
        data=json.dumps(payload, ensure_ascii=False),
        timeout=30,
    )
    response.raise_for_status()
    body = response.json()
    if body.get("code") != "0":
        raise RuntimeError(f"count failed: {body}")
    info = body.get("info") or {}
    if "total" in info:
        return int(info["total"])
    meta = info.get("meta")
    if isinstance(meta, dict) and "count" in meta:
        return int(meta["count"])
    return len(info.get("data") or info.get("items") or [])


def submit_export(task: dict[str, Any], release_statuses: list[int]) -> str:
    payload = {
        "fileType": 1,
        "searchType": 2,
        "type": task["type"],
        "brand": task["brand"],
        "platform": task["platform"],
        "keyword": "",
        "keyId": [],
        "projects": [],
        "emptyLangList": [],
        "releaseStatus": release_statuses,
    }
    response = requests.post(
        MCP_PROXY,
        headers=headers(EXPORT_PATH),
        data=json.dumps(payload, ensure_ascii=False),
        timeout=60,
    )
    response.raise_for_status()
    body = response.json()
    if body.get("code") != "0":
        raise RuntimeError(f"submit export failed: {body}")
    return body["info"]["taskId"]


def poll_export_task(task_id: str) -> str:
    time.sleep(3)
    for idx in range(POLL_MAX):
        response = requests.get(
            MCP_PROXY,
            headers=headers(f"{EXPORT_TASK_PATH}?taskId={task_id}"),
            timeout=30,
        )
        response.raise_for_status()
        body = response.json()
        if body.get("code") != "0":
            raise RuntimeError(f"poll export failed: {body}")
        info = body.get("info") or {}
        status = info.get("status")
        if status == 3:
            return info.get("url") or ""
        if status == 4:
            raise RuntimeError(f"export task failed: taskId={task_id}")
        log(f"  [{idx + 1}/{POLL_MAX}] taskId={task_id} status={status}")
        time.sleep(POLL_INTERVAL)
    raise TimeoutError(f"export task timeout: taskId={task_id}")


def download_file(url: str, out_path: Path) -> None:
    if not url:
        raise RuntimeError("empty download url")
    last_error: Exception | None = None
    for attempt in range(1, 7):
        try:
            response = requests.get(url, timeout=180, stream=True)
            response.raise_for_status()
            with out_path.open("wb") as handle:
                for chunk in response.iter_content(1024 * 128):
                    if chunk:
                        handle.write(chunk)
            log(f"  saved {out_path.name} ({out_path.stat().st_size // 1024} KB)")
            return
        except Exception as exc:  # noqa: BLE001 - keep retry diagnostics in CLI script.
            last_error = exc
            if out_path.exists():
                out_path.unlink()
            log(f"  download retry {attempt}/6 failed: {exc}")
            time.sleep(5 * attempt)
    raise RuntimeError(f"download failed: {last_error}")


def export_latest(run_id: str, release_statuses: list[int]) -> list[dict[str, Any]]:
    raw_dir = RAW_ROOT / run_id
    raw_dir.mkdir(parents=True, exist_ok=True)
    exports: list[dict[str, Any]] = []

    for task in TASKS:
        label = f"{task['tab']}_{task['platform_label']}"
        log("=" * 72)
        log(f"export {label}: type={task['type']} brand={task['brand']} platform={task['platform']} releaseStatus={release_statuses}")
        total = api_count(task, release_statuses)
        status_1 = api_count(task, [1])
        status_2 = api_count(task, [2])
        log(f"  count included={total}, status1={status_1}, status2={status_2}")
        export_info = {
            **task,
            "label": label,
            "release_statuses": release_statuses,
            "api_count": total,
            "status_1_count": status_1,
            "status_2_count": status_2,
            "file": "",
            "rows": 0,
        }
        if total:
            task_id = submit_export(task, release_statuses)
            log(f"  taskId={task_id}")
            url = poll_export_task(task_id)
            out_name = f"{label}_published_{run_id}.xlsx"
            out_path = raw_dir / out_name
            download_file(url, out_path)
            export_info["task_id"] = task_id
            export_info["file"] = str(out_path)
        exports.append(export_info)
    return exports


def find_header_index(headers: list[str], candidates: set[str]) -> int | None:
    lowered = [header.lower() for header in headers]
    for idx, header in enumerate(headers):
        if header in candidates or lowered[idx] in candidates:
            return idx
    return None


def language_columns(headers: list[str]) -> list[tuple[int, str]]:
    cols: list[tuple[int, str]] = []
    for idx, header in enumerate(headers):
        h = norm(header)
        if not h:
            continue
        low = h.lower()
        if h in META_HEADERS or low in META_HEADERS:
            continue
        if low in {"en", "zh"}:
            continue
        if LANG_HEADER_RE.fullmatch(low):
            cols.append((idx, h))
    return cols


def scan_exports(run_id: str, exports: list[dict[str, Any]]) -> dict[str, Any]:
    process_dir = PROCESS_ROOT / run_id
    process_dir.mkdir(parents=True, exist_ok=True)

    detail_by_key: dict[tuple[str, str], dict[str, Any]] = {}
    tab_summary: dict[str, dict[str, Any]] = defaultdict(lambda: {
        "tab": "",
        "exported_keys": 0,
        "missing_keys": 0,
        "missing_cells": 0,
        "language_count": 0,
    })
    platform_summary: list[dict[str, Any]] = []
    lang_summary = Counter()
    lang_by_tab: dict[str, Counter] = defaultdict(Counter)
    languages_seen: set[str] = set()

    for export_info in exports:
        file_path = export_info.get("file")
        if not file_path:
            platform_summary.append({
                **{k: export_info[k] for k in ("tab", "platform_label", "api_count", "status_1_count", "status_2_count")},
                "rows": 0,
                "missing_keys": 0,
                "missing_cells": 0,
            })
            continue
        path = Path(file_path)
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True)
        try:
            raw_headers = next(rows)
        except StopIteration:
            continue
        headers = [norm(v) for v in raw_headers]
        key_idx = find_header_index(headers, {"keyid", "key id", "key_id"})
        en_idx = find_header_index(headers, {"en"})
        product_idx = find_header_index(headers, {"所属产品", "project", "projects"})
        remark_idx = find_header_index(headers, {"备注", "remark"})
        if key_idx is None:
            raise RuntimeError(f"keyId column not found in {path}")

        lang_cols = language_columns(headers)
        languages_seen.update(lang for _, lang in lang_cols)
        tab = export_info["tab"]
        tab_summary[tab]["tab"] = tab
        tab_summary[tab]["language_count"] = max(tab_summary[tab]["language_count"], len(lang_cols))

        row_count = 0
        platform_missing_keys = 0
        platform_missing_cells = 0

        for row_number, row in enumerate(rows, start=2):
            row_count += 1
            key_id = norm(row[key_idx] if key_idx < len(row) else "")
            if not key_id:
                continue
            missing_langs = [
                lang
                for idx, lang in lang_cols
                if idx >= len(row) or is_empty(row[idx])
            ]
            if not missing_langs:
                continue

            en_value = norm(row[en_idx] if en_idx is not None and en_idx < len(row) else "")
            product = norm(row[product_idx] if product_idx is not None and product_idx < len(row) else "")
            remark = norm(row[remark_idx] if remark_idx is not None and remark_idx < len(row) else "")

            detail_key = (tab, key_id)
            record = detail_by_key.setdefault(detail_key, {
                "tab": tab,
                "platform": set(),
                "keyId": key_id,
                "en": en_value,
                "missing_langs": set(),
                "source_files": set(),
                "source_rows": [],
                "product": product,
                "remark": remark,
            })
            record["platform"].add(export_info["platform_label"])
            record["missing_langs"].update(missing_langs)
            record["source_files"].add(path.name)
            record["source_rows"].append(row_number)
            if not record["en"] and en_value:
                record["en"] = en_value
            if not record["product"] and product:
                record["product"] = product
            if not record["remark"] and remark:
                record["remark"] = remark

            platform_missing_keys += 1
            platform_missing_cells += len(missing_langs)
            lang_summary.update(missing_langs)
            lang_by_tab[tab].update(missing_langs)

        export_info["rows"] = row_count
        tab_summary[tab]["exported_keys"] += row_count
        platform_summary.append({
            **{k: export_info[k] for k in ("tab", "platform_label", "api_count", "status_1_count", "status_2_count")},
            "rows": row_count,
            "missing_keys": platform_missing_keys,
            "missing_cells": platform_missing_cells,
        })

    details = []
    for record in detail_by_key.values():
        missing_langs = sorted(record["missing_langs"])
        details.append({
            "tab": record["tab"],
            "platform": "|".join(sorted(record["platform"])),
            "keyId": record["keyId"],
            "missing_lang_count": len(missing_langs),
            "missing_langs": "|".join(missing_langs),
            "en": truncate_cell(record["en"], 1200),
            "product": truncate_cell(record["product"], 500),
            "remark": truncate_cell(record["remark"], 500),
            "source_file": "|".join(sorted(record["source_files"])),
            "source_rows": "|".join(str(v) for v in record["source_rows"][:20]),
        })

    details.sort(key=lambda row: (row["tab"], row["platform"], -row["missing_lang_count"], row["keyId"]))

    for row in details:
        tab_summary[row["tab"]]["missing_keys"] += 1
        tab_summary[row["tab"]]["missing_cells"] += row["missing_lang_count"]

    tab_rows = [
        {
            "tab": row["tab"],
            "exported_keys": row["exported_keys"],
            "missing_keys": row["missing_keys"],
            "missing_key_rate": row["missing_keys"] / row["exported_keys"] if row["exported_keys"] else 0,
            "missing_cells": row["missing_cells"],
            "language_count_excluding_en_zh": row["language_count"],
        }
        for row in sorted(tab_summary.values(), key=lambda item: item["tab"])
    ]
    lang_rows = [
        {
            "language": lang,
            "missing_key_count": count,
            **{f"{tab}_missing": lang_by_tab[tab].get(lang, 0) for tab in sorted(tab_summary.keys())},
        }
        for lang, count in lang_summary.most_common()
    ]

    outputs = {
        "run_id": run_id,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "release_statuses": exports[0]["release_statuses"] if exports else [],
        "release_status_note": "1=已发布, 2=已发布灰度; 3/4 未纳入",
        "raw_dir": str(RAW_ROOT / run_id),
        "language_count_excluding_en_zh": len(languages_seen),
        "languages_excluding_en_zh": sorted(languages_seen),
        "summary": {
            "exported_keys": sum(row["rows"] for row in platform_summary),
            "missing_keys": len(details),
            "missing_cells": sum(row["missing_lang_count"] for row in details),
        },
        "tab_summary": tab_rows,
        "platform_summary": platform_summary,
        "language_summary": lang_rows,
        "missing_details": details,
        "exports": exports,
    }

    json_path = process_dir / f"shein_published_missing_{run_id}.json"
    csv_detail_path = process_dir / f"shein_published_missing_detail_{run_id}.csv"
    csv_tab_path = process_dir / f"shein_published_tab_summary_{run_id}.csv"
    csv_lang_path = process_dir / f"shein_published_lang_summary_{run_id}.csv"

    json_path.write_text(json.dumps(outputs, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(csv_detail_path, details, [
        "tab", "platform", "keyId", "missing_lang_count", "missing_langs",
        "en", "product", "remark", "source_file", "source_rows",
    ])
    write_csv(csv_tab_path, tab_rows, [
        "tab", "exported_keys", "missing_keys", "missing_key_rate",
        "missing_cells", "language_count_excluding_en_zh",
    ])
    write_csv(csv_lang_path, lang_rows, ["language", "missing_key_count", *[f"{tab}_missing" for tab in sorted(tab_summary.keys())]])

    outputs["files"] = {
        "json": str(json_path),
        "detail_csv": str(csv_detail_path),
        "tab_csv": str(csv_tab_path),
        "lang_csv": str(csv_lang_path),
    }
    json_path.write_text(json.dumps(outputs, ensure_ascii=False, indent=2), encoding="utf-8")
    return outputs


def truncate_cell(value: str, limit: int) -> str:
    if len(value) <= limit:
        return value
    return value[: limit - 3] + "..."


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--release-statuses",
        nargs="+",
        type=int,
        default=DEFAULT_RELEASE_STATUSES,
        help="Release status filter. Defaults to 1 2.",
    )
    parser.add_argument("--run-id", default=datetime.now().strftime("%Y%m%d_%H%M%S"))
    parser.add_argument("--scan-only-manifest", help="Skip API export and scan an existing manifest JSON.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.scan_only_manifest:
        manifest = json.loads(Path(args.scan_only_manifest).read_text(encoding="utf-8"))
        outputs = scan_exports(manifest["run_id"], manifest["exports"])
    else:
        exports = export_latest(args.run_id, args.release_statuses)
        manifest = {
            "run_id": args.run_id,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "exports": exports,
        }
        manifest_path = PROCESS_ROOT / args.run_id / f"export_manifest_{args.run_id}.json"
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        outputs = scan_exports(args.run_id, exports)
    log("=" * 72)
    log("done")
    print(json.dumps({
        "run_id": outputs["run_id"],
        "summary": outputs["summary"],
        "files": outputs["files"],
        "raw_dir": outputs["raw_dir"],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""基于上传前对账结果，只重跑需要补翻的 key，并按语种输出一个 Excel。"""

import argparse
import hashlib
import json
import os
import re
import threading
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import requests
from deep_translator import GoogleTranslator
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


BASE = Path(os.environ.get("I18N_WORKDIR") or Path(__file__).resolve().parents[1]).resolve()
PROCESS = BASE / "02_过程数据"
TASK_DIR = PROCESS / "retry_tasks_by_lang_20260624"
OUTPUT_DIR = BASE / "05_上传格式" / "20260624_补翻"
JSONL = PROCESS / "retry_translations_20260624.jsonl"
QC_REPORT = PROCESS / "retry_translation_qc_20260624.json"
RUN_DATE = "20260624"

DIFY_PROXY = os.environ.get("DIFY_PROXY") or "http://127.0.0.1:30001/rest"
DIFY_API_KEY = (os.environ.get("DIFY_API_KEY") or "").strip()
DIFY_WORKERS = 4
DIFY_BATCH = 8
GOOGLE_MIN_INTERVAL = 0.15

LANG_TO_BCP47 = {
    "ar": "ar-SA",
    "de": "de-DE",
    "el-gr": "el-GR",
    "es": "es-ES",
    "et": "et-EE",
    "fr": "fr-FR",
    "hu": "hu-HU",
    "it": "it-IT",
    "lt": "lt-LT",
    "lv": "lv-LV",
    "nb": "nb-NO",
    "nl": "nl-NL",
    "pl": "pl-PL",
    "pt-pt": "pt-PT",
    "ro": "ro-RO",
    "ru": "ru-RU",
    "sl": "sl-SI",
    "sr": "sr-RS",
    "tr": "tr-TR",
    "uk": "uk-UA",
    "zh-cn": "zh-CN",
    "zh-tw": "zh-TW",
}
BCP47_TO_LANG = {value: key for key, value in LANG_TO_BCP47.items()}
GOOGLE_LANG = {
    "ar": "ar",
    "de": "de",
    "el-gr": "el",
    "es": "es",
    "et": "et",
    "fr": "fr",
    "hu": "hu",
    "it": "it",
    "lt": "lt",
    "lv": "lv",
    "nb": "no",
    "nl": "nl",
    "pl": "pl",
    "pt-pt": "pt",
    "ro": "ro",
    "ru": "ru",
    "sl": "sl",
    "sr": "sr",
    "tr": "tr",
    "uk": "uk",
    "zh-cn": "zh-CN",
    "zh-tw": "zh-TW",
}

NOTICE = (
    "模版使用注意事项：\n"
    "1. 模版注意事项内容无需删掉\n"
    "2. 必填字段：key ID（key ID/code ID）；en（系统中原英文文案）；更新语种（需要更新的语种文案，语种名称需和系统中保持一致）\n"
    "3. 字段解释：\n"
    "key ID——需要更新的key ID或code ID，需和系统中的ID保持一致，上传时会进行校验\n"
    "en——需要更新的key在系统中的原英文文案\n"
    "更新语种——需要更新的语种名称，将字段替换为系统中的语种名称，上传时会进行校验\n"
    "4. 表格中仅填写，ID、原英文文案及需要更新的语种文案，其余语种文案不要填写，更新时会覆盖！\n"
    "例如，需要更新日语文案，则表格内容共三列：key ID、en、ja\n"
    "如果需要更新英语、韩语、阿语文案，则表格内容共五列：key ID、en、en、ko、ar（第一列en为原英文文案，第二列en为需要更新的en文案）\n"
    "5. 备注字段如果为空，不会覆盖线上数据。如果填写值，会覆盖线上数据。"
)

HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
RISK_FILL = PatternFill("solid", fgColor="FFF2CC")
FORMULA_PREFIXES = ("=", "+", "-", "@")
PH_RE = re.compile(r"\{[^}]+\}|%\d+\$[sdf]|%[sdf]|%\d+")
TAG_RE = re.compile(r"<[^>]+>")
ESC_RE = re.compile(r"\\['\u2018\u2019]")
CURLY_RE = re.compile(r"\{([^}]*)\}")
CYRILLIC_RE = re.compile(r"[\u0400-\u04FF]")

SR_CYR_TO_LAT = {
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Ђ": "Đ", "Е": "E", "Ж": "Ž",
    "З": "Z", "И": "I", "Ј": "J", "К": "K", "Л": "L", "Љ": "Lj", "М": "M", "Н": "N",
    "Њ": "Nj", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T", "Ћ": "Ć", "У": "U",
    "Ф": "F", "Х": "H", "Ц": "C", "Ч": "Č", "Џ": "Dž", "Ш": "Š",
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ђ": "đ", "е": "e", "ж": "ž",
    "з": "z", "и": "i", "ј": "j", "к": "k", "л": "l", "љ": "lj", "м": "m", "н": "n",
    "њ": "nj", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "ћ": "ć", "у": "u",
    "ф": "f", "х": "h", "ц": "c", "ч": "č", "џ": "dž", "ш": "š",
    "Ё": "Jo", "Й": "J", "Щ": "Šč", "Ъ": "", "Ы": "Y", "Ь": "", "Э": "E", "Ю": "Ju", "Я": "Ja",
    "ё": "jo", "й": "j", "щ": "šč", "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "ju", "я": "ja",
}
CYR_UPPER = set("АБВГДЂЕЖЗИЈКЛЉМНЊОПРСТЋУФХЦЧЏШЁЙЩЪЫЬЭЮЯ")
CYR_LOWER = set("абвгдђежзијклљмнњопрстћуфхцчџшёйщъыьэюя")
DIGRAPH_UPPER = {"Љ": "LJ", "Њ": "NJ", "Џ": "DŽ"}

write_lock = threading.Lock()
google_lock = threading.Lock()
google_next_at = 0.0


def log(message):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}", flush=True)


def md5(text):
    return hashlib.md5(text.encode("utf-8")).hexdigest()


def escape_curly(text):
    return CURLY_RE.sub(lambda match: f"[[CURLY_{match.group(1)}]]", text)


def unescape_curly(text):
    return re.sub(r"\[\[CURLY_([^\]]*)\]\]", lambda match: "{" + match.group(1) + "}", text)


def protect(text):
    mapping = []

    def repl(match):
        token = f"ZXSHEIN{len(mapping)}ZX"
        mapping.append((token, match.group(0)))
        return token

    text = TAG_RE.sub(repl, text)
    text = PH_RE.sub(repl, text)
    return text, mapping


def restore(text, mapping):
    for token, raw in mapping:
        text = text.replace(token, raw)
    return text


def wait_google_slot():
    global google_next_at
    with google_lock:
        now = time.time()
        if now < google_next_at:
            time.sleep(google_next_at - now)
        google_next_at = time.time() + GOOGLE_MIN_INTERVAL


def google_translate(en, lang):
    safe, mapping = protect(en)
    wait_google_slot()
    text = GoogleTranslator(source="en", target=GOOGLE_LANG[lang]).translate(safe)
    text = restore(str(text or ""), mapping)
    if lang == "sr":
        text = latinize_sr(text)
    return text


def latinize_sr(text):
    if not isinstance(text, str):
        return text
    out = []
    chars = list(text)
    for idx, char in enumerate(chars):
        mapped = SR_CYR_TO_LAT.get(char)
        if mapped is None:
            out.append(char)
            continue
        if char in DIGRAPH_UPPER:
            prev_char = chars[idx - 1] if idx > 0 else ""
            next_char = chars[idx + 1] if idx + 1 < len(chars) else ""
            if (prev_char in CYR_UPPER or prev_char not in CYR_LOWER) and next_char in CYR_UPPER:
                mapped = DIGRAPH_UPPER[char]
        out.append(mapped)
    return "".join(out)


def qc_issues(en, translation, lang):
    en = str(en or "")
    translation = str(translation or "")
    issues = []
    if not translation.strip():
        issues.append("EMPTY")
    if sorted(PH_RE.findall(en)) != sorted(PH_RE.findall(translation)):
        issues.append("PH_MISMATCH")
    if ESC_RE.findall(en) != ESC_RE.findall(translation):
        issues.append("ESCAPED_SINGLE_QUOTE_MISMATCH")
    if TAG_RE.findall(en) != TAG_RE.findall(translation):
        issues.append("TAG_SEQUENCE_MISMATCH")
    if len(en) >= 1200 and len(translation) < len(en) * 0.45:
        issues.append("POSSIBLE_TRUNCATED_LONG_TEXT")
    if lang == "sr" and CYRILLIC_RE.search(translation):
        issues.append("SR_CYRILLIC_REMAINING")
    return issues


def load_tasks(langs=None):
    tasks_by_lang = {}
    selected = set(langs or [])
    for path in sorted(TASK_DIR.glob("*_retry_tasks_20260624.json")):
        lang = path.name.replace("_retry_tasks_20260624.json", "")
        if selected and lang not in selected:
            continue
        if lang not in LANG_TO_BCP47:
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        tasks_by_lang[lang] = data
    return tasks_by_lang


def load_done_jsonl():
    done = {}
    if not JSONL.exists():
        return done
    with open(JSONL, encoding="utf-8") as f:
        for line in f:
            try:
                record = json.loads(line)
            except Exception:
                continue
            if record.get("status") != "ok":
                continue
            key = record.get("en_hash")
            if not key:
                continue
            if key not in done:
                done[key] = {
                    "en_hash": key,
                    "en": record.get("en", ""),
                    "translations": {},
                    "status": "ok",
                }
            done[key]["translations"].update(record.get("translations") or {})
    return done


def history_files_for_lang(lang):
    patterns = [f"{lang}_full_*.jsonl", f"{lang}_cache.jsonl"]
    special = {
        "ar": ["ar_ar_*.jsonl"],
        "de": ["de_de_*.jsonl"],
        "es": ["es_es_*.jsonl"],
        "fr": ["fr_fr_*.jsonl"],
        "it": ["it_it_*.jsonl", "it_translated_*_it.jsonl"],
        "pt-pt": ["pt_unified_full_*.jsonl"],
        "uk": ["uk_uk_*.jsonl"],
    }
    patterns.extend(special.get(lang, []))
    files = []
    seen = set()
    for pattern in patterns:
        for path in sorted(PROCESS.glob(pattern)):
            if path not in seen:
                seen.add(path)
                files.append(path)
    return files


def load_history_cache(lang):
    cache = {}
    for path in history_files_for_lang(lang):
        with open(path, encoding="utf-8") as f:
            for line in f:
                try:
                    record = json.loads(line)
                except Exception:
                    continue
                if record.get("status") != "ok":
                    continue
                en = record.get("en_full") or record.get("en") or ""
                translation = record.get("translation") or record.get("translated") or ""
                if lang == "sr":
                    translation = latinize_sr(translation)
                if en and translation and not qc_issues(en, translation, lang):
                    cache[en] = translation
    return cache


def hydrate_done_from_history(tasks_by_lang, done):
    appended = 0
    by_en = defaultdict(dict)
    for lang, tasks in tasks_by_lang.items():
        cache = load_history_cache(lang)
        if not cache:
            continue
        hit = 0
        for task in tasks:
            en = task["en"]
            key = md5(en)
            if lang in (done.get(key, {}).get("translations") or {}):
                continue
            translation = cache.get(en)
            if not translation:
                continue
            if key not in done:
                done[key] = {"en_hash": key, "en": en, "translations": {}, "status": "ok"}
            done[key]["translations"][lang] = translation
            by_en[en][lang] = translation
            hit += 1
        if hit:
            log(f"历史缓存命中 {lang}: {hit}")

    for en, translations in by_en.items():
        append_jsonl({
            "en_hash": md5(en),
            "en": en,
            "translations": translations,
            "requested_langs": sorted(translations),
            "missing_langs": [],
            "status": "ok",
            "engine": "history_cache",
            "ts": datetime.now(timezone.utc).isoformat(),
        })
        appended += 1
    if appended:
        log(f"历史缓存写入 JSONL 唯一 EN 数：{appended}")
    return done


def append_jsonl(record):
    with write_lock:
        with open(JSONL, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")


def call_dify(en, langs):
    if not DIFY_API_KEY:
        raise RuntimeError("缺少环境变量 DIFY_API_KEY，不能调用 Dify 翻译接口")
    safe, mapping = protect(en[:4000])
    bcp47_list = [LANG_TO_BCP47[lang] for lang in langs]
    results = {}
    headers = {
        "Authorization": f"Bearer {DIFY_API_KEY}",
        "Content-Type": "application/json",
        "mcp-proxy-path": "/aihelper-dify-api2/v1/workflows/run",
        "mcp-proxy-env": "prod",
        "mcp-proxy-dc": "aliyun_dc",
    }
    for idx in range(0, len(bcp47_list), DIFY_BATCH):
        batch = bcp47_list[idx:idx + DIFY_BATCH]
        payload = {
            "inputs": {
                "OriginalLanguage": "English",
                "OriginalText": escape_curly(safe),
                "TargetLanguage": ",".join(batch),
            },
            "response_mode": "blocking",
            "user": "retry-fill-20260624",
        }
        last_error = None
        for attempt in range(4):
            try:
                response = requests.post(
                    DIFY_PROXY,
                    headers=headers,
                    data=json.dumps(payload, ensure_ascii=False),
                    timeout=90,
                )
                if response.status_code == 429:
                    time.sleep(8 * (attempt + 1))
                    continue
                response.raise_for_status()
                output = response.json().get("data", {}).get("outputs", {}).get("output", [])
                for item in output:
                    code = item.get("target_language_code")
                    if not code:
                        continue
                    lang = BCP47_TO_LANG.get(code)
                    if not lang:
                        continue
                    text = restore(unescape_curly(item.get("translated_text", "") or ""), mapping)
                    if lang == "sr":
                        text = latinize_sr(text)
                    results[lang] = text
                break
            except Exception as exc:
                last_error = exc
                time.sleep(3 * (attempt + 1))
        else:
            raise RuntimeError(f"Dify failed: {last_error}")
    return results


def translate_hybrid(en, langs):
    translations = {}
    fallback_langs = []
    google_errors = {}
    for lang in langs:
        try:
            text = google_translate(en, lang)
            issues = qc_issues(en, text, lang)
            if issues:
                fallback_langs.append(lang)
                google_errors[lang] = ";".join(issues)
            else:
                translations[lang] = text
        except Exception as exc:
            fallback_langs.append(lang)
            google_errors[lang] = str(exc)
    if fallback_langs:
        fallback = call_dify(en, fallback_langs)
        translations.update(fallback)
    return translations, google_errors


def build_pending(tasks_by_lang, done):
    pending = {}
    for lang, tasks in tasks_by_lang.items():
        for task in tasks:
            en = task["en"]
            key = md5(en)
            existing = done.get(key, {}).get("translations", {})
            if lang not in existing:
                pending.setdefault(en, set()).add(lang)
    return {en: sorted(langs) for en, langs in pending.items()}


def translate_pending(pending, done, limit=0):
    items = sorted(pending.items(), key=lambda item: (len(item[0]), len(item[1])))
    if limit:
        items = items[:limit]
    counters = {"ok": 0, "error": 0}
    log(f"待调用 Dify 的唯一 EN 数：{len(items)}")
    if not items:
        return counters

    def work(en, langs):
        translations, google_errors = translate_hybrid(en, langs)
        missing = [lang for lang in langs if lang not in translations]
        return {
            "en_hash": md5(en),
            "en": en,
            "translations": translations,
            "requested_langs": langs,
            "missing_langs": missing,
            "google_errors": google_errors,
            "status": "ok",
            "ts": datetime.now(timezone.utc).isoformat(),
        }

    with ThreadPoolExecutor(max_workers=DIFY_WORKERS) as executor:
        futures = {executor.submit(work, en, langs): (en, langs) for en, langs in items}
        for idx, future in enumerate(as_completed(futures), 1):
            en, langs = futures[future]
            try:
                record = future.result()
                append_jsonl(record)
                done[record["en_hash"]] = record
                counters["ok"] += 1
            except Exception as exc:
                record = {
                    "en_hash": md5(en),
                    "en": en,
                    "translations": {},
                    "requested_langs": langs,
                    "status": "error",
                    "error": str(exc),
                    "ts": datetime.now(timezone.utc).isoformat(),
                }
                append_jsonl(record)
                counters["error"] += 1
            if idx <= 5 or idx % 20 == 0:
                log(f"翻译进度 {idx}/{len(items)} ok={counters['ok']} error={counters['error']}")
    return counters


def set_text(cell, value):
    text = "" if value is None else str(value)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    cell.value = text
    cell.data_type = "s"
    cell.number_format = "@"
    if cell.value.startswith(FORMULA_PREFIXES):
        cell.quotePrefix = True


def build_workbook(lang, tasks, done):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = lang
    ws.freeze_panes = "A3"
    ws.merge_cells("A1:D1")
    set_text(ws["A1"], NOTICE)
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 165

    for col, header in enumerate(["key ID", "en", lang, "备注"], 1):
        cell = ws.cell(2, col)
        set_text(cell, header)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    stats = {
        "lang": lang,
        "rows": 0,
        "ok_rows": 0,
        "risk_rows": 0,
        "missing_translation_rows": 0,
        "error_types": defaultdict(int),
        "sr_cyrillic_rows": 0,
        "formula_cell_count": 0,
        "formula_prefix_without_quote_count": 0,
        "samples": [],
    }
    seen_keys = set()
    duplicate_keys = []

    for row_idx, task in enumerate(tasks, 3):
        key = task["keyId"]
        en = task["en"]
        record = done.get(md5(en), {})
        translation = (record.get("translations") or {}).get(lang, "")
        if lang == "sr":
            translation = latinize_sr(translation)
        issues = qc_issues(en, translation, lang)
        remark = ""
        if issues:
            remark = f"质检风险：{';'.join(issues)}；请人工确认后上传"
            for issue in issues:
                stats["error_types"][issue] += 1
            if not translation:
                stats["missing_translation_rows"] += 1
            stats["risk_rows"] += 1
            if len(stats["samples"]) < 30:
                stats["samples"].append({
                    "keyId": key,
                    "en": en[:160],
                    "translation": translation[:160],
                    "issues": issues,
                })
        else:
            stats["ok_rows"] += 1

        if lang == "sr" and CYRILLIC_RE.search(translation):
            stats["sr_cyrillic_rows"] += 1

        values = [key, en, translation, remark]
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row_idx, col_idx)
            set_text(cell, value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if remark and col_idx in (3, 4):
                cell.fill = RISK_FILL
            if cell.data_type == "f":
                stats["formula_cell_count"] += 1
            if isinstance(cell.value, str) and cell.value.startswith(FORMULA_PREFIXES) and not cell.quotePrefix:
                stats["formula_prefix_without_quote_count"] += 1

        stats["rows"] += 1
        if key in seen_keys:
            duplicate_keys.append(key)
        seen_keys.add(key)

    stats["duplicate_key_count"] = len(duplicate_keys)
    stats["error_types"] = dict(stats["error_types"])
    for col, width in {"A": 28, "B": 62, "C": 62, "D": 46}.items():
        ws.column_dimensions[col].width = width

    path = OUTPUT_DIR / f"{lang}_补翻_{RUN_DATE}.xlsx"
    wb.save(path)
    stats["output"] = str(path)
    return stats


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--langs", nargs="*", default=None)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--skip-translate", action="store_true")
    args = parser.parse_args()

    tasks_by_lang = load_tasks(args.langs)
    log("任务行数：" + json.dumps({lang: len(tasks) for lang, tasks in tasks_by_lang.items()}, ensure_ascii=False))
    done = load_done_jsonl()
    done = hydrate_done_from_history(tasks_by_lang, done)
    pending = build_pending(tasks_by_lang, done)
    if not args.skip_translate:
        counters = translate_pending(pending, done, args.limit)
        log("翻译完成：" + json.dumps(counters, ensure_ascii=False))
        done = load_done_jsonl()

    report = {
        "run_date": RUN_DATE,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "jsonl": str(JSONL),
        "output_dir": str(OUTPUT_DIR),
        "langs": {},
    }
    for lang, tasks in tasks_by_lang.items():
        stats = build_workbook(lang, tasks, done)
        report["langs"][lang] = stats
        log(f"已生成 {lang}: rows={stats['rows']} ok={stats['ok_rows']} risk={stats['risk_rows']} -> {stats['output']}")

    QC_REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "qc_report": str(QC_REPORT),
        "output_dir": str(OUTPUT_DIR),
        "langs": {
            lang: {
                "rows": stats["rows"],
                "ok_rows": stats["ok_rows"],
                "risk_rows": stats["risk_rows"],
                "missing_translation_rows": stats["missing_translation_rows"],
                "output": stats["output"],
            }
            for lang, stats in report["langs"].items()
        },
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

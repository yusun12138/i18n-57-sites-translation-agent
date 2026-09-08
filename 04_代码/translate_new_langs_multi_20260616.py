#!/usr/bin/env python3
import argparse
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import requests

from translate_single_lang_full_20260612 import (
    DIFY_HEADERS,
    DIFY_PROXY,
    FILES,
    PROCESS,
    RUN_DATE,
    escape_curly,
    google_translate,
    load_excluded,
    md5,
    protect,
    qc_ok,
    qc_issues,
    restore,
    should_copy_source,
    unescape_curly,
)

LANG_CONFIG = {
    'nb': ('no', 'nb-NO'),
    'sl': ('sl', 'sl-SI'),
    'lv': ('lv', 'lv-LV'),
    'sr': ('sr', 'sr-RS'),
    'et': ('et', 'et-EE'),
}

FLUSH_EVERY = 500
WORKERS = 3


def log(msg):
    print(f'[{datetime.now().strftime("%H:%M:%S")}] {msg}', flush=True)


def norm(value):
    return '' if value is None else str(value).strip()


def find_lang_col(headers, lang):
    matches = [idx for idx, header in enumerate(headers) if header == lang]
    if lang == 'id' and matches:
        return matches[-1]
    return matches[0] if matches else None


def build_needed_and_cache(langs, task_cache):
    if task_cache.exists():
        with open(task_cache, encoding='utf-8') as f:
            data = json.load(f)
        return (
            {lang: set(data['needed'].get(lang, [])) for lang in langs},
            {lang: dict(data['cache'].get(lang, {})) for lang in langs},
            data.get('source_stats', []),
        )

    excluded = load_excluded()
    needed = {lang: set() for lang in langs}
    cache = {lang: {} for lang in langs}
    source_stats = []

    for path in FILES:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        ws.reset_dimensions()
        rows = ws.iter_rows(values_only=True)
        headers = [norm(v).lower() for v in next(rows)]
        key_i = next(i for i, h in enumerate(headers) if h in ('keyid', 'key id', 'key_id'))
        en_i = headers.index('en')
        lang_cols = {lang: find_lang_col(headers, lang) for lang in langs}
        stats = {
            'file': path.name,
            'rows': 0,
            'excluded_seen': 0,
            'missing_added': {lang: 0 for lang in langs},
            'cache_added': {lang: 0 for lang in langs},
        }

        for row in rows:
            stats['rows'] += 1
            key = norm(row[key_i])
            en = norm(row[en_i])
            if not key or not en:
                continue
            if key in excluded:
                stats['excluded_seen'] += 1
                continue

            for lang, col_idx in lang_cols.items():
                tr = '' if col_idx is None else norm(row[col_idx])
                if tr and en not in cache[lang] and qc_ok(en, tr):
                    cache[lang][en] = tr
                    stats['cache_added'][lang] += 1
                if not tr:
                    needed[lang].add(en)
                    stats['missing_added'][lang] += 1

        source_stats.append(stats)

    with open(task_cache, 'w', encoding='utf-8') as f:
        json.dump(
            {
                'langs': langs,
                'needed': {lang: sorted(values) for lang, values in needed.items()},
                'cache': cache,
                'source_stats': source_stats,
            },
            f,
            ensure_ascii=False,
        )
    return needed, cache, source_stats


def load_jsonl_done(lang):
    jsonl = PROCESS / f'{lang}_full_{RUN_DATE}.jsonl'
    done = {}
    blocked = {}
    stats = {'ok': 0, 'blocked': 0, 'google': 0, 'dify_fallback': 0, 'copy_source': 0, 'lines': 0}
    if not jsonl.exists():
        return done, blocked, stats

    with open(jsonl, encoding='utf-8') as f:
        for line in f:
            stats['lines'] += 1
            try:
                rec = json.loads(line)
            except Exception:
                continue
            en = rec.get('en_full') or ''
            tr = rec.get('translation') or ''
            status = rec.get('status')
            engine = rec.get('engine')
            if engine in stats:
                stats[engine] += 1
            if status == 'ok' and en and qc_ok(en, tr):
                done[en] = tr
                stats['ok'] += 1
            elif en:
                blocked[en] = {'translation': tr, 'issues': rec.get('issues') or ['JSONL_BLOCKED']}
                if status == 'blocked':
                    stats['blocked'] += 1
    return done, blocked, stats


class JsonlWriters:
    def __init__(self, langs):
        self.paths = {lang: PROCESS / f'{lang}_full_{RUN_DATE}.jsonl' for lang in langs}
        self.buffers = {lang: [] for lang in langs}
        self.lock = threading.Lock()

    def append(self, lang, record):
        with self.lock:
            self.buffers[lang].append(record)
            if len(self.buffers[lang]) >= FLUSH_EVERY:
                self._flush_locked(lang)

    def _flush_locked(self, lang):
        records = self.buffers[lang]
        if not records:
            return 0
        with open(self.paths[lang], 'a', encoding='utf-8') as f:
            for record in records:
                f.write(json.dumps(record, ensure_ascii=False) + '\n')
        count = len(records)
        self.buffers[lang] = []
        return count

    def flush_all(self):
        flushed = {}
        with self.lock:
            for lang in self.buffers:
                count = self._flush_locked(lang)
                if count:
                    flushed[lang] = count
        return flushed


def dify_translate_many(en, lang_list):
    safe, mapping = protect(en)
    target_codes = [LANG_CONFIG[lang][1] for lang in lang_list]
    payload = {
        'inputs': {
            'OriginalLanguage': 'English',
            'OriginalText': escape_curly(safe),
            'TargetLanguage': ','.join(target_codes),
        },
        'response_mode': 'blocking',
        'user': 'new-langs-google-first-multi-20260616',
    }
    last_error = None
    for attempt in range(2):
        try:
            r = requests.post(
                DIFY_PROXY,
                headers=DIFY_HEADERS,
                data=json.dumps(payload, ensure_ascii=False),
                timeout=180,
            )
            r.raise_for_status()
            out = r.json().get('data', {}).get('outputs', {}).get('output', [])
            by_code = {}
            for item in out:
                code = item.get('target_language_code')
                if code:
                    tr = unescape_curly(item.get('translated_text', '') or '')
                    by_code[code] = restore(tr, mapping)
            return {lang: by_code.get(LANG_CONFIG[lang][1], '') for lang in lang_list}
        except Exception as exc:
            last_error = exc
            time.sleep(3 * (attempt + 1))
    raise RuntimeError(f'Dify failed: {last_error}')


def make_record(en, translation, status, engine, **extra):
    record = {
        'en_hash': md5(en),
        'en_full': en,
        'translation': translation,
        'status': status,
        'engine': engine,
        'ts': datetime.now(timezone.utc).isoformat(),
    }
    record.update(extra)
    return record


def translate_en(en, langs):
    if should_copy_source(en):
        return {
            lang: make_record(en, en, 'ok', 'copy_source')
            for lang in langs
        }

    records = {}
    fallback_langs = []
    google_errors = {}

    for lang in langs:
        google_code = LANG_CONFIG[lang][0]
        try:
            tr = google_translate(en, google_code)
            issues = qc_issues(en, tr)
            if not issues:
                records[lang] = make_record(en, tr, 'ok', 'google')
            else:
                fallback_langs.append(lang)
                google_errors[lang] = ';'.join(issues)
        except Exception as exc:
            fallback_langs.append(lang)
            google_errors[lang] = str(exc)

    if fallback_langs:
        try:
            fallback = dify_translate_many(en, fallback_langs)
            for lang in fallback_langs:
                tr = fallback.get(lang, '')
                issues = qc_issues(en, tr)
                if not issues:
                    records[lang] = make_record(
                        en,
                        tr,
                        'ok',
                        'dify_fallback',
                        google_error=google_errors.get(lang, ''),
                    )
                else:
                    records[lang] = make_record(
                        en,
                        tr,
                        'blocked',
                        'dify_fallback',
                        google_error=google_errors.get(lang, ''),
                        issues=issues,
                    )
        except Exception as exc:
            for lang in fallback_langs:
                records[lang] = make_record(
                    en,
                    '',
                    'blocked',
                    'dify_fallback',
                    google_error=google_errors.get(lang, ''),
                    issues=[f'DIFY_FAILED: {exc}'],
                )
    return records


def translate_en_dify_primary(en, langs):
    if should_copy_source(en):
        return {
            lang: make_record(en, en, 'ok', 'copy_source')
            for lang in langs
        }

    records = {}
    try:
        translations = dify_translate_many(en, langs)
        for lang in langs:
            tr = translations.get(lang, '')
            issues = qc_issues(en, tr)
            if not issues:
                records[lang] = make_record(en, tr, 'ok', 'dify_primary')
            else:
                records[lang] = make_record(en, tr, 'blocked', 'dify_primary', issues=issues)
    except Exception as exc:
        for lang in langs:
            records[lang] = make_record(en, '', 'blocked', 'dify_primary', issues=[f'DIFY_FAILED: {exc}'])
    return records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--langs', nargs='+', default=list(LANG_CONFIG), choices=sorted(LANG_CONFIG))
    parser.add_argument('--workers', type=int, default=WORKERS)
    parser.add_argument('--mode', choices=['google-first', 'dify-primary'], default='google-first')
    parser.add_argument('--limit', type=int, default=0)
    args = parser.parse_args()
    langs = list(dict.fromkeys(args.langs))
    task_cache = PROCESS / f'new_langs_{"_".join(langs)}_full_{RUN_DATE}_tasks.json'

    needed, cache, source_stats = build_needed_and_cache(langs, task_cache)
    done_by_lang = {}
    jsonl_stats = {}
    pending_by_en = {}
    for lang in langs:
        done, _blocked, stats = load_jsonl_done(lang)
        done_by_lang[lang] = done
        jsonl_stats[lang] = stats
        for en in sorted(needed[lang]):
            if en in cache[lang] or en in done:
                continue
            pending_by_en.setdefault(en, []).append(lang)

    log('source_stats=' + json.dumps(source_stats, ensure_ascii=False))
    for lang in langs:
        log(
            f'{lang}: needed={len(needed[lang])} full_cache={len(cache[lang])} '
            f'jsonl_done={len(done_by_lang[lang])} pending={sum(1 for v in pending_by_en.values() if lang in v)} '
            f'jsonl_stats={jsonl_stats[lang]}'
        )

    total = len(pending_by_en)
    if args.limit and total > args.limit:
        pending_by_en = dict(list(pending_by_en.items())[:args.limit])
        total = len(pending_by_en)
        log(f'limit enabled: pending trimmed to {total}')
    if total == 0:
        log('no pending en')
        return

    writers = JsonlWriters(langs)
    counters = {
        lang: {'ok': 0, 'google': 0, 'dify_fallback': 0, 'dify_primary': 0, 'copy_source': 0, 'blocked': 0}
        for lang in langs
    }
    translate_func = translate_en_dify_primary if args.mode == 'dify-primary' else translate_en

    try:
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {
                executor.submit(translate_func, en, en_langs): (en, en_langs)
                for en, en_langs in pending_by_en.items()
            }
            for idx, future in enumerate(as_completed(futures), 1):
                en, en_langs = futures[future]
                records = future.result()
                for lang in en_langs:
                    rec = records[lang]
                    writers.append(lang, rec)
                    if rec['status'] == 'ok':
                        counters[lang]['ok'] += 1
                    else:
                        counters[lang]['blocked'] += 1
                    engine = rec.get('engine')
                    if engine in counters[lang]:
                        counters[lang][engine] += 1
                if idx <= 3 or idx % 500 == 0:
                    flushed = writers.flush_all()
                    log(f'progress {idx}/{total} flushed={flushed} counters={json.dumps(counters, ensure_ascii=False)}')
    finally:
        flushed = writers.flush_all()
        if flushed:
            log(f'final_flush={flushed}')

    log(f'finished counters={json.dumps(counters, ensure_ascii=False)}')


if __name__ == '__main__':
    main()

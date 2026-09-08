#!/usr/bin/env python3
"""
57站多语言翻译主脚本
- EN文本去重，相同EN只调一次API
- 12并发worker，压缩排队等待
- PH_MISMATCH硬性阻断（不写入，记录重译队列）
- 强制文本格式写入，防=开头被Excel解析为公式
- JSONL断点续跑，keyed by hash(en)
用法:
  python3 translate_multi_site.py --langs de fr ar       # 只填空格(老语种)
  python3 translate_multi_site.py --langs lt lv et --all # 全量(新语种)
  python3 translate_multi_site.py site1.xlsx --langs de  # 指定文件
"""

import argparse, hashlib, json, os, random, re, sys, threading, time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import requests
try:
    from deep_translator import GoogleTranslator
except Exception:
    GoogleTranslator = None
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

# ── 配置 ──────────────────────────────────────────────────────────────────────
DIFY_PROXY   = os.environ.get('DIFY_PROXY') or 'http://127.0.0.1:30001/rest'
DIFY_API_KEY = (os.environ.get('DIFY_API_KEY') or '').strip()
DIFY_WORKERS = 6
DIFY_BATCH   = 8   # Dify单次最多8种语言
DEFAULT_ENGINE = 'hybrid'  # hybrid=Google优先，Dify兜底
GOOGLE_WORKERS = 3
GOOGLE_MIN_INTERVAL = 0.35
GOOGLE_RETRIES = 4

BASE_DIR    = Path(os.environ.get('I18N_WORKDIR') or Path(__file__).resolve().parents[1]).resolve()
PROCESS_DIR = BASE_DIR / '02_过程数据'
OUTPUT_DIR  = BASE_DIR / '03_产出'
EXCLUDED_KEYIDS_PATH = BASE_DIR / '04_代码' / 'excluded_keyids.txt'
PROCESS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ── 语种列名 → BCP-47 映射 ───────────────────────────────────────────────────
LANG_MAP = {
    'zh': 'zh-CN', 'zh-cn': 'zh-CN', 'zh-hk': 'zh-HK', 'zh-tw': 'zh-TW',
    'en_au': 'en-AU', 'en_gb': 'en-GB', 'en_in': 'en-IN',
    'de': 'de-DE', 'fr': 'fr-FR', 'es': 'es-ES', 'es-mx': 'es-MX',
    'it': 'it-IT', 'pt': 'pt-PT', 'pt-br': 'pt-BR',
    'ar': 'ar-SA', 'he': 'he-IL', 'tr': 'tr-TR',
    'ru': 'ru-RU', 'uk': 'uk-UA', 'pl': 'pl-PL',
    'cs': 'cs-CZ', 'sk': 'sk-SK', 'ro': 'ro-RO',
    'nl': 'nl-NL', 'sv': 'sv-SE', 'da': 'da-DK',
    'fi': 'fi-FI', 'nb': 'nb-NO', 'hu': 'hu-HU',
    'el': 'el-GR', 'bg': 'bg-BG', 'hr': 'hr-HR',
    'lt': 'lt-LT', 'lv': 'lv-LV', 'et': 'et-EE',
    'sl': 'sl-SI', 'sr': 'sr-RS', 'ka': 'ka-GE',
    'id': 'id-ID', 'ms': 'ms-MY', 'th': 'th-TH',
    'vi': 'vi-VN', 'ja': 'ja-JP', 'ko': 'ko-KR',
}

# deep-translator / Google 语言代码映射
GOOGLE_LANG = {
    'zh': 'zh-CN', 'zh-cn': 'zh-CN', 'zh-hk': 'zh-TW', 'zh-tw': 'zh-TW',
    'en_au': 'en', 'en_gb': 'en', 'en_in': 'en',
    'de': 'de', 'fr': 'fr', 'es': 'es', 'es-mx': 'es',
    'it': 'it', 'pt': 'pt', 'pt-br': 'pt',
    'ar': 'ar', 'he': 'iw', 'tr': 'tr',
    'ru': 'ru', 'uk': 'uk', 'pl': 'pl',
    'cs': 'cs', 'sk': 'sk', 'ro': 'ro',
    'nl': 'nl', 'sv': 'sv', 'da': 'da',
    'fi': 'fi', 'nb': 'no', 'hu': 'hu',
    'el': 'el', 'bg': 'bg', 'hr': 'hr',
    'lt': 'lt', 'lv': 'lv', 'et': 'et',
    'sl': 'sl', 'sr': 'sr', 'ka': 'ka',
    'id': 'id', 'ms': 'ms', 'th': 'th',
    'vi': 'vi', 'ja': 'ja', 'ko': 'ko',
}

# ── 颜色 ─────────────────────────────────────────────────────────────────────
FILL_RED    = PatternFill('solid', fgColor='FFCCCC')
FILL_YELLOW = PatternFill('solid', fgColor='FFFFCC')
FILL_GREEN  = PatternFill('solid', fgColor='CCFFCC')

# ── 正则 ─────────────────────────────────────────────────────────────────────
_PH_RE         = re.compile(r'\{[^}]+\}|%\d+\$[sdf]|%[sdf]|%\d+')
_HTML_TAG_RE   = re.compile(r'<(/?)([a-zA-Z][a-zA-Z0-9]*)[^>]*/?>', re.IGNORECASE)
_TAG_TOKEN_RE   = re.compile(r'<[^>]+>')
_DOUBLE_SP_RE  = re.compile(r'  +')
_CURLY_RE      = re.compile(r'\{([^}]*)\}')
_FORMULA_START = re.compile(r'^[=+\-@]')  # 防公式头
_ESCAPED_SINGLE_QUOTE_RE = re.compile(r"\\['\u2018\u2019]")
LONG_TEXT_MIN_LEN = 1200
LONG_TEXT_MIN_RATIO = 0.45

VOID_TAGS = {'br', 'hr', 'img', 'input', 'meta', 'link'}


def log(msg): print(f'[{datetime.now().strftime("%H:%M:%S")}] {msg}', flush=True)

def en_hash(text: str) -> str:
    return hashlib.md5(text.encode('utf-8')).hexdigest()

def load_excluded_keyids() -> set:
    if not EXCLUDED_KEYIDS_PATH.exists():
        return set()
    with open(EXCLUDED_KEYIDS_PATH, encoding='utf-8') as f:
        return {
            line.strip()
            for line in f
            if line.strip() and not line.lstrip().startswith('#')
        }

EXCLUDED_KEYIDS = load_excluded_keyids()

def escape_curly(text):  return _CURLY_RE.sub(lambda m: f'[[CURLY_{m.group(1)}]]', text)
def unescape_curly(text): return re.sub(r'\[\[CURLY_([^\]]*)\]\]', lambda m: '{'+m.group(1)+'}', text)

def protect_placeholders(text: str) -> tuple:
    """把 HTML 标签和完整占位符换成稳定 token，避免翻译引擎改写格式内容。"""
    mapping = []

    def repl_tag(match):
        token = f'ZXHTML{len(mapping)}ZX'
        mapping.append((token, match.group(0)))
        return token

    def repl_placeholder(match):
        token = f'ZXPH{len(mapping)}ZX'
        mapping.append((token, match.group(0)))
        return token

    text = _TAG_TOKEN_RE.sub(repl_tag, text)
    text = _PH_RE.sub(repl_placeholder, text)
    return text, mapping

def restore_placeholders(text: str, mapping: list) -> str:
    for token, placeholder in mapping:
        text = text.replace(token, placeholder)
    return text

def safe_cell_value(val: str):
    """强制文本：防止=开头被Excel解析为公式"""
    if isinstance(val, str) and _FORMULA_START.match(val):
        return "'" + val   # 加单引号前缀，openpyxl不会渲染为公式
    return val

def write_text_cell(ws, row, col, val):
    """写入单元格并强制文本格式"""
    cell = ws.cell(row=row, column=col, value=safe_cell_value(val))
    cell.number_format = '@'
    return cell


# ── QC ───────────────────────────────────────────────────────────────────────

def _placeholders(text: str) -> list:
    return sorted(_PH_RE.findall(text))

def _escaped_single_quotes(text: str) -> list:
    """反斜杠转义单引号序列必须与 EN 完全一致，避免上传后解析冲突。"""
    return _ESCAPED_SINGLE_QUOTE_RE.findall(text)

def _tag_tokens(text: str) -> list:
    """富文本/HTML/伪标签必须保持完整序列，不能被翻译引擎改写属性或标签内容。"""
    return _TAG_TOKEN_RE.findall(text)

def _html_balanced(text: str) -> bool:
    stack = []
    for closing, tag in _HTML_TAG_RE.findall(text):
        tag = tag.lower()
        if tag in VOID_TAGS:
            continue
        if not closing:
            stack.append(tag)
        else:
            if stack and stack[-1] == tag:
                stack.pop()
            else:
                return False
    return len(stack) == 0

def qc_check(en: str, translated: str) -> tuple:
    """返回 (issues: list, is_block_error: bool)
    is_block_error=True 时不允许写入产出文件（占位符硬性阻断）
    """
    issues = []
    block = False
    if not translated.strip():
        return ['EMPTY'], True
    en_ph = _placeholders(en)
    tr_ph = _placeholders(translated)
    if en_ph != tr_ph:
        issues.append(f'PH_MISMATCH en={en_ph} got={tr_ph}')
        block = True   # 红线2：占位符不一致，硬性阻断
    en_quotes = _escaped_single_quotes(en)
    tr_quotes = _escaped_single_quotes(translated)
    if en_quotes != tr_quotes:
        issues.append(f'ESCAPED_SINGLE_QUOTE_MISMATCH en={en_quotes} got={tr_quotes}')
        block = True
    en_tags = _tag_tokens(en)
    tr_tags = _tag_tokens(translated)
    if en_tags != tr_tags:
        issues.append('TAG_SEQUENCE_MISMATCH')
        block = True
    if len(en) >= LONG_TEXT_MIN_LEN and len(translated) < len(en) * LONG_TEXT_MIN_RATIO:
        issues.append(
            f'POSSIBLE_TRUNCATED_LONG_TEXT en_len={len(en)} got_len={len(translated)}'
        )
        block = True
    if '<' in en and not _html_balanced(translated):
        issues.append('HTML_UNBALANCED')
    if _DOUBLE_SP_RE.search(translated):
        issues.append('DOUBLE_SPACE')
    return issues, block


# ── Dify API ──────────────────────────────────────────────────────────────────

def call_dify_batch(en_text: str, bcp47_list: list) -> dict:
    """调用Dify翻译一个EN文本到多语言，返回 {bcp47: translated_text}"""
    if not DIFY_API_KEY:
        raise RuntimeError('缺少环境变量 DIFY_API_KEY，不能调用 Dify 翻译接口')
    safe, ph_mapping = protect_placeholders(en_text)
    results = {}
    for i in range(0, len(bcp47_list), DIFY_BATCH):
        batch = bcp47_list[i:i + DIFY_BATCH]
        payload = {
            'inputs': {
                'OriginalLanguage': 'English',
                'OriginalText': safe,
                'TargetLanguage': ','.join(batch),
            },
            'response_mode': 'blocking',
            'user': 'app-57sites-fill',
        }
        headers = {
            'Authorization': f'Bearer {DIFY_API_KEY}',
            'Content-Type': 'application/json',
            'mcp-proxy-path': '/aihelper-dify-api2/v1/workflows/run',
            'mcp-proxy-env': 'prod',
            'mcp-proxy-dc': 'aliyun_dc',
        }
        for attempt in range(5):
            try:
                r = requests.post(DIFY_PROXY, headers=headers,
                                  data=json.dumps(payload, ensure_ascii=False),
                                  timeout=120)
                if r.status_code == 429:
                    wait = 10 * (attempt + 1)
                    log(f'  429限流，等待{wait}s（第{attempt+1}次）')
                    time.sleep(wait)
                    continue
                r.raise_for_status()
                out = r.json().get('data', {}).get('outputs', {}).get('output', [])
                for item in out:
                    code = item.get('target_language_code', '')
                    text = restore_placeholders(str(item.get('translated_text', '') or ''), ph_mapping)
                    results[code] = text
                break
            except requests.exceptions.Timeout:
                wait = 15 * (attempt + 1)
                log(f'  超时，等待{wait}s（第{attempt+1}次）')
                time.sleep(wait)
    return results


_google_lock = threading.Lock()
_google_next_at = 0.0


def wait_google_slot(min_interval: float):
    """全局限速，避免 Google 短时间并发触发 429。"""
    global _google_next_at
    with _google_lock:
        now = time.time()
        if now < _google_next_at:
            time.sleep(_google_next_at - now)
        _google_next_at = time.time() + min_interval + random.uniform(0, 0.12)


def call_google_one(en_text: str, lang_col: str, min_interval: float) -> str:
    """调用 Google 翻译单条文本；失败抛出异常，由上层决定是否 Dify 兜底。"""
    if GoogleTranslator is None:
        raise RuntimeError('未安装 deep-translator，无法使用 Google 引擎')
    google_code = GOOGLE_LANG.get(lang_col)
    if not google_code:
        raise ValueError(f'未知Google语言代码: {lang_col}')

    safe, ph_mapping = protect_placeholders(en_text)
    last_error = None
    for attempt in range(GOOGLE_RETRIES):
        try:
            wait_google_slot(min_interval)
            tr = GoogleTranslator(source='en', target=google_code).translate(safe)
            tr = restore_placeholders(str(tr or ''), ph_mapping)
            if tr.strip():
                return tr
            raise ValueError('Google返回空翻译')
        except Exception as e:
            last_error = e
            wait = min(45, (2 ** attempt) + random.uniform(0.5, 1.8))
            time.sleep(wait)
    raise RuntimeError(f'Google重试失败: {last_error}')


def qc_filter(en_text: str, lang_col: str, translated: str, engine: str):
    issues, block = qc_check(en_text, translated or '')
    if block:
        return None, {
            'lang': lang_col,
            'engine': engine,
            'translated': translated or '',
            'issues': issues,
        }
    return translated, None


def add_qc_issue(qc_issues: list, row_idx: int, keyid: str, lang_col: str,
                 issues: list, en_text: str, translated: str = ''):
    qc_issues.append({
        'row': row_idx,
        'keyId': keyid,
        'lang': lang_col,
        'issues': ';'.join(issues),
        'en': en_text,
        'translated': translated,
    })


def write_qc_sheet(wb, qc_issues: list):
    if '_QC_issues' in wb.sheetnames:
        del wb['_QC_issues']
    if not qc_issues:
        return
    ws_qc = wb.create_sheet('_QC_issues')
    headers = ['row', 'keyId', 'lang', 'issues', 'en', 'translated']
    ws_qc.append(headers)
    for item in qc_issues:
        ws_qc.append([item.get(h, '') for h in headers])
    for cell in ws_qc[1]:
        cell.fill = FILL_YELLOW


def load_legacy_lang_cache(lang_col: str) -> dict:
    """读取旧脚本按语种保存的缓存，格式为 {'en': 'translated'}。"""
    cache = {}
    for path in PROCESS_DIR.glob(f'{lang_col}*.jsonl'):
        if path.stat().st_size == 0:
            continue
        with open(path, encoding='utf-8') as f:
            for line in f:
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                en = str(rec.get('en') or '').strip()
                tr = str(rec.get('translated') or '').strip()
                if rec.get('status') == 'ok' and en and tr:
                    cache[en] = tr
                translations = rec.get('translations') or {}
                tr2 = str(translations.get(lang_col) or '').strip()
                if rec.get('status') in ('ok', 'partial') and en and tr2:
                    cache[en] = tr2
    return cache


def load_excel_lang_cache(lang_col: str) -> dict:
    """读取历史 Excel 产物中的同语种译文，作为本地翻译记忆。"""
    cache = {}
    paths = set()
    for pattern in (f'{lang_col}*.xlsx', f'*_{lang_col}_*.xlsx', f'*_{lang_col}.xlsx'):
        paths.update(OUTPUT_DIR.glob(pattern))
    for path in sorted(paths):
        if path.name.startswith('~$') or path.name.startswith('.~'):
            continue
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            try:
                header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            except StopIteration:
                continue
            headers = [str(h or '').strip() for h in header_row]
            if 'en' not in headers or lang_col not in headers:
                continue
            en_idx = headers.index('en')
            lang_idx = headers.index(lang_col)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) <= max(en_idx, lang_idx):
                    continue
                en = str(row[en_idx] or '').strip()
                tr = str(row[lang_idx] or '').strip()
                if en and tr:
                    cache[en] = tr
    return cache


# ── 加载语种映射表 ─────────────────────────────────────────────────────────────

def get_bcp47(col_name: str) -> str:
    """Excel列名 → BCP-47代码"""
    return LANG_MAP.get(col_name.lower().strip(), '')


# ── 主流程 ────────────────────────────────────────────────────────────────────

def translate_file(filepath: Path, target_langs: list, fill_all: bool = False,
                   engine: str = DEFAULT_ENGINE, dify_workers: int = DIFY_WORKERS,
                   google_workers: int = GOOGLE_WORKERS,
                   google_interval: float = GOOGLE_MIN_INTERVAL):
    log(f'\n{"="*60}')
    log(f'处理文件: {filepath.name}  目标语种: {target_langs}  全量={fill_all}  引擎={engine}')

    wb, ws, col_map, tasks, headers, keyid_col, en_col = \
        load_excel(filepath, target_langs, fill_all)

    if not tasks:
        log('  无需翻译（所有格均已有内容），跳过')
        return

    log(f'  待翻译行数: {len(tasks)}')

    stem = filepath.stem
    jsonl_path = PROCESS_DIR / f'{stem}_{"_".join(sorted(target_langs))}.jsonl'
    done: dict = defaultdict(dict)
    blocked_cache: dict = defaultdict(dict)
    for lang_col in target_langs:
        local_cache = {}
        local_cache.update(load_excel_lang_cache(lang_col))
        local_cache.update(load_legacy_lang_cache(lang_col))
        accepted_legacy = blocked_legacy = 0
        for en_text, tr in local_cache.items():
            h = en_hash(en_text)
            clean, blocked = qc_filter(en_text, lang_col, tr, 'legacy')
            if blocked:
                blocked_cache[h][lang_col] = blocked
                blocked_legacy += 1
                continue
            done[h][lang_col] = clean
            accepted_legacy += 1
        if local_cache:
            log(f'  历史缓存({lang_col}): 可用={accepted_legacy} 阻断={blocked_legacy}')
    if jsonl_path.exists():
        with open(jsonl_path, encoding='utf-8') as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    h = rec['en_hash']
                    en_text = rec.get('en') or ''
                    if rec.get('status') in ('ok', 'partial'):
                        for lang_col, tr in (rec.get('translations') or {}).items():
                            if str(tr or '').strip():
                                clean, blocked = qc_filter(en_text, lang_col, tr, 'jsonl')
                                if blocked:
                                    blocked_cache[h][lang_col] = blocked
                                else:
                                    done[h][lang_col] = clean
                    for item in rec.get('blocked') or []:
                        lang_col = item.get('lang', '')
                        if lang_col:
                            blocked_cache[h][lang_col] = item
                except Exception:
                    pass
        log(f'  已有缓存: {len(done)} 条EN翻译')

    en_to_tasks: dict = defaultdict(list)
    for t in tasks:
        en_to_tasks[t['en']].append(t)

    pending_en = []
    for en_text, en_tasks in en_to_tasks.items():
        h = en_hash(en_text)
        all_missing = set()
        for t in en_tasks:
            all_missing.update(t['missing_langs'])
        missing_after_cache = [lc for lc in sorted(all_missing) if not done[h].get(lc)]
        if missing_after_cache:
            pending_en.append((en_text, missing_after_cache))

    log(f'  唯一EN文本: {len(en_to_tasks)}  去重后待翻译: {len(pending_en)}')

    accepted: dict = defaultdict(dict)
    blocked_current: dict = defaultdict(dict)
    google_failures: dict = defaultdict(dict)
    total_pairs = sum(len(missing_langs) for _, missing_langs in pending_en)

    if engine in ('google', 'hybrid') and pending_en:
        google_jobs = [
            (en_text, lang_col)
            for en_text, missing_langs in pending_en
            for lang_col in missing_langs
            if lang_col in GOOGLE_LANG
        ]
        log(f'  Google待翻译: {len(google_jobs)} 格  workers={google_workers}  interval={google_interval}s')
        with ThreadPoolExecutor(max_workers=google_workers) as executor:
            futures = {
                executor.submit(call_google_one, en_text, lang_col, google_interval): (en_text, lang_col)
                for en_text, lang_col in google_jobs
            }
            for i, future in enumerate(as_completed(futures), 1):
                en_text, lang_col = futures[future]
                h = en_hash(en_text)
                try:
                    tr = future.result()
                    clean, blocked = qc_filter(en_text, lang_col, tr, 'google')
                    if blocked:
                        google_failures[h][lang_col] = blocked
                    else:
                        accepted[h][lang_col] = clean
                except Exception as e:
                    google_failures[h][lang_col] = {
                        'lang': lang_col, 'engine': 'google',
                        'translated': '', 'issues': [f'GOOGLE_FAILED: {e}'],
                    }
                if i % 100 == 0 or i <= 3:
                    log(f'  Google进度 [{i}/{len(google_jobs)}] 可用={sum(len(v) for v in accepted.values())}')

    fallback_en = []
    for en_text, missing_langs in pending_en:
        h = en_hash(en_text)
        remain = [lc for lc in missing_langs if not accepted[h].get(lc)]
        if remain:
            fallback_en.append((en_text, remain))

    if engine in ('dify', 'hybrid') and fallback_en:
        log(f'  Dify兜底待翻译: {len(fallback_en)} 条EN/{sum(len(v) for _, v in fallback_en)} 格  workers={dify_workers}')

        def do_one(item):
            en_text, missing_langs = item
            bcp47_list = [LANG_MAP[lc] for lc in missing_langs if lc in LANG_MAP]
            translations = call_dify_batch(en_text, bcp47_list)
            return en_text, missing_langs, translations

        with ThreadPoolExecutor(max_workers=dify_workers) as executor:
            futures = {executor.submit(do_one, item): item for item in fallback_en}
            for i, future in enumerate(as_completed(futures), 1):
                en_text, missing_langs = futures[future]
                h = en_hash(en_text)
                try:
                    _, _, translations = future.result()
                    for lang_col in missing_langs:
                        bcp47 = LANG_MAP.get(lang_col, '')
                        tr = translations.get(bcp47, '')
                        clean, blocked = qc_filter(en_text, lang_col, tr, 'dify')
                        if blocked:
                            blocked_current[h][lang_col] = blocked
                        else:
                            accepted[h][lang_col] = clean
                    if i % 50 == 0 or i <= 3:
                        log(f'  Dify进度 [{i}/{len(fallback_en)}] 可用={sum(len(v) for v in accepted.values())}')
                except Exception as e:
                    for lang_col in missing_langs:
                        blocked_current[h][lang_col] = {
                            'lang': lang_col, 'engine': 'dify',
                            'translated': '', 'issues': [f'DIFY_FAILED: {e}'],
                        }

    if engine == 'google':
        for h, lang_map in google_failures.items():
            for lang_col, item in lang_map.items():
                if not accepted[h].get(lang_col):
                    blocked_current[h][lang_col] = item
    elif engine == 'hybrid':
        for en_text, missing_langs in pending_en:
            h = en_hash(en_text)
            for lang_col in missing_langs:
                if not accepted[h].get(lang_col) and lang_col not in blocked_current[h]:
                    blocked_current[h][lang_col] = google_failures[h].get(lang_col, {
                        'lang': lang_col, 'engine': 'hybrid',
                        'translated': '', 'issues': ['EMPTY_OR_NOT_TRANSLATED'],
                    })

    f_out = open(jsonl_path, 'a', encoding='utf-8')
    success = fail = blocked_count = 0
    for en_text, missing_langs in pending_en:
        h = en_hash(en_text)
        if accepted[h]:
            done[h].update(accepted[h])
        for lang_col, item in blocked_current[h].items():
            blocked_cache[h][lang_col] = item

        blocked_list = [blocked_current[h][lc] for lc in missing_langs if lc in blocked_current[h]]
        available = {lc: done[h][lc] for lc in missing_langs if done[h].get(lc)}
        if len(available) == len(missing_langs):
            status = 'ok'
            success += 1
        elif available:
            status = 'partial'
            success += 1
            blocked_count += len(blocked_list)
        else:
            status = 'blocked'
            fail += 1
            blocked_count += len(blocked_list) or len(missing_langs)
        rec = {
            'en_hash': h, 'en': en_text[:200],
            'translations': available,
            'blocked': blocked_list,
            'status': status,
            'ts': datetime.now(timezone.utc).isoformat(),
        }
        f_out.write(json.dumps(rec, ensure_ascii=False) + '\n')
    f_out.close()
    log(f'  翻译阶段完成: EN成功/部分成功={success} EN失败={fail} 阻断格={blocked_count}/{total_pairs}')

    # 回填到产出目录
    import shutil
    out_path = OUTPUT_DIR / filepath.name
    if not out_path.exists():
        shutil.copy2(str(filepath), str(out_path))
    wb_out = openpyxl.load_workbook(str(out_path), data_only=True)
    ws_out = wb_out.active
    out_headers = [str(ws_out.cell(row=1, column=c).value or '').strip()
                   for c in range(1, ws_out.max_column + 1)]
    out_col_map = {h: i+1 for i, h in enumerate(out_headers)}

    filled = skip_existing = skip_no_trans = 0
    qc_issues = []
    for task in tasks:
        row_idx = task['row_idx']
        h = en_hash(task['en'])
        translations = done.get(h, {})
        for lang_col in task['missing_langs']:
            if lang_col not in out_col_map:
                continue
            col_idx = out_col_map[lang_col]
            existing = str(ws_out.cell(row=row_idx, column=col_idx).value or '').strip()
            if existing and not fill_all:
                skip_existing += 1
                continue
            tr = translations.get(lang_col, '')
            if not tr:
                cell = ws_out.cell(row=row_idx, column=col_idx)
                cell.fill = FILL_RED
                blocked = blocked_cache.get(h, {}).get(lang_col, {})
                issues = blocked.get('issues') or ['EMPTY_OR_NOT_TRANSLATED']
                translated = blocked.get('translated') or ''
                cell.comment = Comment(';'.join(issues), 'translate_multi_site')
                add_qc_issue(qc_issues, row_idx, task['keyid'], lang_col,
                             issues, task['en'], translated)
                skip_no_trans += 1
                continue
            write_text_cell(ws_out, row_idx, col_idx, tr)
            ws_out.cell(row=row_idx, column=col_idx).fill = FILL_GREEN
            filled += 1

    # 红线1：验证EN列未被修改
    wb_in = openpyxl.load_workbook(str(filepath), data_only=True)
    ws_in = wb_in.active
    en_col_out = out_col_map.get('en', en_col)
    mismatch_en = 0
    for row_idx in range(2, ws_in.max_row + 1):
        v_in = str(ws_in.cell(row=row_idx, column=en_col).value or '')
        v_out = str(ws_out.cell(row=row_idx, column=en_col_out).value or '')
        if v_in != v_out:
            mismatch_en += 1
    if mismatch_en:
        log(f'  [红线1警告] EN列有{mismatch_en}行与原文不一致！')

    write_qc_sheet(wb_out, qc_issues)
    wb_out.save(str(out_path))
    total_check = filled + skip_existing + skip_no_trans
    log(f'  回填: 写入={filled} 跳过已有={skip_existing} 无翻译={skip_no_trans} 校验={total_check}')
    log(f'  QC问题: {len(qc_issues)} 条（红色单元格 + _QC_issues sheet）')
    log(f'  输出: {out_path}')


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='57站多语言翻译')
    parser.add_argument('files', nargs='*',
                        help='输入Excel路径，不填则处理01_原始数据/下所有xlsx')
    parser.add_argument('--langs', nargs='+', required=True,
                        help='目标语种列名，如 de fr ar（与Excel列名一致）')
    parser.add_argument('--all', dest='fill_all', action='store_true',
                        help='全量模式：不管现有内容是否为空都翻译（用于新语种）')
    parser.add_argument('--engine', choices=['hybrid', 'google', 'dify'],
                        default=DEFAULT_ENGINE,
                        help='翻译引擎：hybrid=Google优先+Dify兜底，google=仅Google，dify=仅Dify')
    parser.add_argument('--dify-workers', type=int, default=DIFY_WORKERS,
                        help=f'Dify并发数，默认{DIFY_WORKERS}')
    parser.add_argument('--google-workers', type=int, default=GOOGLE_WORKERS,
                        help=f'Google并发数，默认{GOOGLE_WORKERS}')
    parser.add_argument('--google-interval', type=float, default=GOOGLE_MIN_INTERVAL,
                        help=f'Google全局请求间隔秒数，默认{GOOGLE_MIN_INTERVAL}')
    args = parser.parse_args()

    unknown = [l for l in args.langs if l not in LANG_MAP]
    if unknown:
        print(f'[错误] 以下语种未在映射表中: {unknown}')
        print(f'已知语种: {sorted(LANG_MAP.keys())}')
        sys.exit(1)

    if args.files:
        files = [Path(f) for f in args.files]
    else:
        files = sorted((BASE_DIR / '01_原始数据').glob('*.xlsx'))
        if not files:
            print(f'[错误] {BASE_DIR}/01_原始数据 下无xlsx文件')
            sys.exit(1)

    log(f'待处理文件: {[f.name for f in files]}')
    log(f'目标语种: {args.langs}')
    log(f'模式: {"全量" if args.fill_all else "仅填空"}')
    log(f'引擎: {args.engine}  Dify workers={args.dify_workers}  Google workers={args.google_workers}')

    for f in files:
        if not f.exists():
            log(f'[跳过] 文件不存在: {f}')
            continue
        translate_file(f, args.langs, fill_all=args.fill_all,
                       engine=args.engine,
                       dify_workers=args.dify_workers,
                       google_workers=args.google_workers,
                       google_interval=args.google_interval)

    log('\n全部完成！')
    log(f'产出目录: {OUTPUT_DIR}')

# ── 读取Excel，提取待翻译任务 ──────────────────────────────────────────────────

def load_excel(filepath: Path, target_langs: list, fill_all: bool = False):
    """
    读取输入Excel，返回:
      - wb: Workbook对象
      - ws: 活动Sheet
      - col_map: {列名: 列索引(1-based)}
      - tasks: [{row_idx, keyid, en, missing_langs:[col_name,...]}]
    """
    wb = openpyxl.load_workbook(str(filepath), data_only=True)
    ws = wb.active
    headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        headers.append(str(v).strip() if v is not None else '')

    # 找 keyId 和 en 列
    keyid_col = None
    en_col = None
    for i, h in enumerate(headers):
        if h.lower() in ('keyid', 'key id', 'key_id'):
            keyid_col = i + 1
        if h.lower() == 'en':
            en_col = i + 1
    if keyid_col is None or en_col is None:
        raise ValueError(f'{filepath.name}: 找不到keyId或en列，当前列: {headers}')

    col_map = {h: i+1 for i, h in enumerate(headers)}

    tasks = []
    for row_idx in range(2, ws.max_row + 1):
        keyid = str(ws.cell(row=row_idx, column=keyid_col).value or '').strip()
        en_raw = ws.cell(row=row_idx, column=en_col).value
        en_val = str(en_raw or '').strip()
        if not keyid or not en_val or en_val.lower() == 'none':
            continue
        if keyid in EXCLUDED_KEYIDS:
            continue

        missing = []
        for lang_col in target_langs:
            if lang_col not in col_map:
                continue
            existing = str(ws.cell(row=row_idx, column=col_map[lang_col]).value or '').strip()
            if fill_all or not existing:
                missing.append(lang_col)

        if missing:
            tasks.append({
                'row_idx': row_idx,
                'keyid': keyid,
                'en': en_val,
                'missing_langs': missing,
            })
    return wb, ws, col_map, tasks, headers, keyid_col, en_col


if __name__ == '__main__':
    main()

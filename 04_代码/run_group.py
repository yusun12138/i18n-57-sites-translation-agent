#!/usr/bin/env python3
"""
分组翻译脚本：拉空key → 导出Excel → Google批量翻译 → QC → 存产出
用法:
  python3 run_group.py --langs ar uk        # 第1组
  python3 run_group.py --langs de es fr it  # 第2组
"""

import argparse, json, os, re, sys, time, threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from deep_translator import GoogleTranslator
from openpyxl.styles import PatternFill

# ── 配置 ──────────────────────────────────────────────────────────────────────
MCP_PROXY  = os.environ.get('MCP_PROXY') or 'http://127.0.0.1:30001/rest'
KEYS_PATH  = '/language-configuration-center/headless/api/v1/keys'
EXPORT_PATH= '/language-configuration-center/headless/api/v1/keys/export'
TASK_PATH  = '/language-configuration-center/headless/api/v1/keys/export-task'
MCP_DC     = os.environ.get('MCP_PROXY_DC') or 'central_p2_dc'
ULP_COOKIE = os.environ.get('ULP_COOKIE') or 'ulp_token={{ulp_token}}'

TYPE_CONFIG = {
    1: (1, [1,2,3,4,5]),
    2: (2, [6,7]),
    3: (3, [12]),
}

BASE_DIR    = Path(os.environ.get('I18N_WORKDIR') or Path(__file__).resolve().parents[1]).resolve()
PROCESS_DIR = BASE_DIR / '02_过程数据'
OUTPUT_DIR  = BASE_DIR / '03_产出'
PROCESS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

BATCH_SIZE   = 50    # Google每批条数（Dify模式不用）
WORKERS      = 5     # 并发数
EXPORT_LIMIT = 5000

# 翻译引擎：'google' 或 'dify'
ENGINE = 'dify'

# Dify 配置
DIFY_KEY = (os.environ.get('DIFY_API_KEY') or '').strip()
DIFY_HEADERS = {
    'Authorization': f'Bearer {DIFY_KEY}',
    'Content-Type': 'application/json',
    'mcp-proxy-path': '/aihelper-dify-api2/v1/workflows/run',
    'mcp-proxy-env': 'prod',
    'mcp-proxy-dc': 'aliyun_dc',
}

# 列名 → Dify BCP-47
DIFY_LANG = {
    'ar':'ar-SA', 'de':'de-DE', 'el-gr':'el-GR', 'es':'es-ES', 'fr':'fr-FR',
    'hu':'hu-HU', 'id':'id-ID', 'it':'it-IT', 'nl':'nl-NL', 'pl':'pl-PL',
    'pt-br':'pt-BR', 'pt-pt':'pt-PT', 'ro':'ro-RO', 'ru':'ru-RU', 'tr':'tr-TR',
    'uk':'uk-UA', 'zh':'zh-CN', 'zh-cn':'zh-CN', 'zh-hk':'zh-HK',
    'zh-tw':'zh-TW', 'ko':'ko-KR', 'ja':'ja-JP', 'th':'th-TH', 'vi':'vi-VN',
    'sv':'sv-SE', 'da':'da-DK', 'fi':'fi-FI', 'cs-cz':'cs-CZ',
    'sk':'sk-SK', 'bg':'bg-BG', 'hr':'hr-HR', 'he':'he-IL', 'ms':'ms-MY',
    'nb':'nb-NO', 'lt':'lt-LT', 'lv':'lv-LV', 'et':'et-EE', 'sl':'sl-SI',
    'sr':'sr-RS', 'ka':'ka-GE',
}

FILL_GREEN  = PatternFill('solid', fgColor='CCFFCC')
FILL_YELLOW = PatternFill('solid', fgColor='FFFFCC')
FILL_RED    = PatternFill('solid', fgColor='FFCCCC')

_PH_RE = re.compile(r'\{[^}]+\}|%\d+\$[sdf]|%[sdf]|%\d+')

# deep-translator 语言代码映射（列名→Google代码）
GOOGLE_LANG = {
    'ar':'ar', 'de':'de', 'el-gr':'el', 'es':'es', 'fr':'fr',
    'hu':'hu', 'id':'id', 'it':'it', 'nl':'nl', 'pl':'pl',
    'pt-br':'pt', 'pt-pt':'pt', 'ro':'ro', 'ru':'ru', 'tr':'tr',
    'uk':'uk', 'zh':'zh-CN', 'zh-cn':'zh-CN', 'zh-hk':'zh-TW',
    'zh-tw':'zh-TW', 'ko':'ko', 'ja':'ja', 'th':'th', 'vi':'vi',
    'sv':'sv', 'da':'da', 'fi':'fi', 'nl':'nl', 'cs-cz':'cs',
    'sk':'sk', 'bg':'bg', 'hr':'hr', 'he':'iw', 'ms':'ms',
    'uk':'uk',
}


def log(msg, lang=''):
    prefix = f'[{lang}]' if lang else ''
    print(f'[{datetime.now().strftime("%H:%M:%S")}]{prefix} {msg}', flush=True)


def mcp_headers(path):
    return {
        'Content-Type': 'application/json',
        'mcp-proxy-path': path,
        'mcp-proxy-dc': MCP_DC,
        'Cookie': ULP_COOKIE,
    }


# ── Step 1+2：按筛选条件直接导出（3个type分别导出再合并）────────────────────

def export_by_filter(lang: str) -> Path:
    """直接用 emptyLangList 筛选条件导出，不需要先收集keyId"""
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    all_rows = []
    header = None

    for type_val, (brand, platforms) in TYPE_CONFIG.items():
        payload = {
            'fileType': 1, 'searchType': 2,
            'type': type_val, 'brand': brand, 'platform': platforms,
            'keyId': [], 'projects': [],
            'emptyLangList': [lang],
            'releaseStatus': [],
        }
        r = requests.post(MCP_PROXY, headers=mcp_headers(EXPORT_PATH),
                          data=json.dumps(payload), timeout=60)
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != '0':
            raise ValueError(f'提交导出失败 type={type_val}: {resp}')
        task_id = resp['info']['taskId']
        log(f'  type={type_val} taskId={task_id}', lang)

        # 轮询
        time.sleep(3)
        url = None
        for i in range(60):
            r2 = requests.get(MCP_PROXY,
                              headers=mcp_headers(f'{TASK_PATH}?taskId={task_id}'),
                              timeout=30)
            info = r2.json()['info']
            if info['status'] == 3:
                url = info['url']
                break
            elif info['status'] == 4:
                raise ValueError(f'导出任务失败 type={type_val} taskId={task_id}')
            time.sleep(5)
        if not url:
            raise TimeoutError(f'导出超时 type={type_val} taskId={task_id}')

        # 下载并读取行
        r3 = requests.get(url, timeout=120, stream=True)
        tmp_path = PROCESS_DIR / f'{lang}_type{type_val}_{ts}.xlsx'
        with open(tmp_path, 'wb') as f:
            for chunk in r3.iter_content(8192):
                f.write(chunk)

        wb_tmp = openpyxl.load_workbook(str(tmp_path))
        ws_tmp = wb_tmp.active
        if header is None:
            header = [ws_tmp.cell(1, c).value for c in range(1, ws_tmp.max_column+1)]
        for row in ws_tmp.iter_rows(min_row=2, values_only=True):
            all_rows.append(row)
        cnt = ws_tmp.max_row - 1
        log(f'  type={type_val} 下载 {cnt} 行', lang)
        tmp_path.unlink()

    if not all_rows:
        log('无数据', lang)
        return None

    # 合并写入单个Excel
    out_path = OUTPUT_DIR / f'{lang}_{ts}_raw.xlsx'
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.append(header)
    for row in all_rows:
        ws_out.append(list(row))
    wb_out.save(str(out_path))
    log(f'合并导出: {len(all_rows)} 行 → {out_path.name}', lang)
    return out_path


# ── Step 3：Google批量翻译 + QC ───────────────────────────────────────────────

def placeholders(text: str) -> list:
    return sorted(_PH_RE.findall(str(text)))


def translate_and_fill(excel_path: Path, lang: str) -> Path:
    google_code = GOOGLE_LANG.get(lang)
    if not google_code:
        raise ValueError(f'未知Google语言代码: {lang}')

    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    headers = [str(ws.cell(1, c).value or '').strip()
               for c in range(1, ws.max_column+1)]

    # 找 en 列和目标语种列
    en_col = next((i+1 for i, h in enumerate(headers) if h.lower() == 'en'), None)
    # 语种列名匹配（支持 el-gr / el_gr / el 等变体）
    lang_col = None
    for i, h in enumerate(headers):
        if h.lower().replace('_','-') == lang.lower().replace('_','-'):
            lang_col = i+1
            break
    if not en_col or not lang_col:
        raise ValueError(f'找不到en列或{lang}列，当前列: {headers}')

    log(f'en列={en_col} {lang}列={lang_col} 总行={ws.max_row-1}', lang)

    # 收集待翻译行
    pending = []  # (row_idx, en_text)
    for row in range(2, ws.max_row+1):
        en_val = str(ws.cell(row, en_col).value or '').strip()
        lang_val = str(ws.cell(row, lang_col).value or '').strip()
        if en_val and not lang_val:
            pending.append((row, en_val))

    log(f'待翻译: {len(pending)} 行', lang)
    if not pending:
        log('无需翻译，跳过', lang)
        return excel_path

    # JSONL断点：按语种固定命名，不带时间戳，重跑永远命中缓存
    jsonl_path = PROCESS_DIR / f'{lang}_cache.jsonl'
    cache = {}
    if jsonl_path.exists():
        with open(jsonl_path, encoding='utf-8') as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    if rec.get('status') == 'ok':
                        cache[rec['en']] = rec['translated']
                except Exception:
                    pass
        log(f'缓存已有: {len(cache)} 条', lang)

    # 过滤已缓存
    todo = [(row, en) for row, en in pending if en not in cache]
    log(f'去除缓存后待翻译: {len(todo)} 行', lang)

    # 批量翻译
    f_out = open(jsonl_path, 'a', encoding='utf-8')
    success = fail = 0
    _CURLY = re.compile(r'\{([^}]*)\}')
    _DBL   = re.compile(r'\{\{([^}]*)\}\}')

    def esc(t):
        t = _DBL.sub(lambda m: f'[[DBL_{m.group(1)}]]', t)
        t = _CURLY.sub(lambda m: f'[[CURLY_{m.group(1)}]]', t)
        return t
    def unesc(t):
        t = re.sub(r'\[\[CURLY_([^\]]*)\]\]', lambda m: '{'+m.group(1)+'}', t)
        t = re.sub(r'\[\[DBL_([^\]]*)\]\]',   lambda m: '{{'+m.group(1)+'}}', t)
        return t

    def trans_one(item):
        if not DIFY_KEY:
            raise RuntimeError('缺少环境变量 DIFY_API_KEY，不能调用 Dify 翻译接口')
        row, en = item
        bcp47 = DIFY_LANG.get(lang, '')
        if not bcp47:
            raise ValueError(f'未知Dify语言代码: {lang}')
        payload = {
            'inputs': {
                'OriginalLanguage': 'English',
                'OriginalText': esc(en[:2000]),
                'TargetLanguage': bcp47,
            },
            'response_mode': 'blocking',
            'user': 'run-group',
        }
        for attempt in range(5):
            try:
                r = requests.post(MCP_PROXY, headers=DIFY_HEADERS,
                                  data=json.dumps(payload, ensure_ascii=False),
                                  timeout=120)
                if r.status_code == 429:
                    time.sleep(10*(attempt+1)); continue
                r.raise_for_status()
                out = r.json().get('data',{}).get('outputs',{}).get('output',[])
                for item2 in out:
                    if item2.get('target_language_code') == bcp47:
                        return row, en, unesc(item2.get('translated_text',''))
                if out:
                    return row, en, unesc(out[0].get('translated_text',''))
                return row, en, ''
            except requests.exceptions.Timeout:
                time.sleep(15*(attempt+1))
        raise Exception('Dify重试5次仍失败')

    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures = {ex.submit(trans_one, item): item for item in todo}
        for i, future in enumerate(as_completed(futures), 1):
            row, en = futures[future]
            try:
                row, en, tr = future.result()
                cache[en] = tr
                rec = {'en': en, 'translated': tr, 'lang': lang,
                       'status': 'ok', 'ts': datetime.now().isoformat()}
                f_out.write(json.dumps(rec, ensure_ascii=False) + '\n')
                f_out.flush()
                success += 1
                if i % 100 == 0 or i <= 3:
                    log(f'  [{i}/{len(todo)}] {en[:30]} → {tr[:30]}', lang)
            except Exception as e:
                fail += 1
                log(f'  [{i}/{len(todo)}] ✗ {en[:30]}: {e}', lang)

    f_out.close()
    log(f'翻译完成: 成功={success} 失败={fail}', lang)

    # 写回Excel + QC
    qc_errors = qc_warns = filled = 0
    for row, en in pending:
        tr = cache.get(en, '')
        if not tr:
            ws.cell(row, lang_col).fill = FILL_RED
            qc_errors += 1
            continue

        # 红线：占位符数量必须一致
        en_ph = placeholders(en)
        tr_ph = placeholders(tr)
        if en_ph != tr_ph:
            # 标红，不写入
            ws.cell(row, lang_col).value = f'[PH_MISMATCH]{tr}'
            ws.cell(row, lang_col).fill = FILL_RED
            qc_errors += 1
            continue

        # 写入，强制文本格式防公式
        cell = ws.cell(row, lang_col)
        cell.value = tr
        cell.number_format = '@'
        cell.fill = FILL_GREEN
        filled += 1

        # 双空格警告
        if '  ' in tr:
            cell.fill = FILL_YELLOW
            qc_warns += 1

    log(f'写入: {filled}格  PH错误(红): {qc_errors}  双空格(黄): {qc_warns}', lang)

    # 验证EN列未被修改（抽查前10行）
    wb_orig = openpyxl.load_workbook(str(excel_path), data_only=True)
    ws_orig = wb_orig.active
    mismatch = sum(1 for r in range(2, min(12, ws.max_row+1))
                   if str(ws.cell(r, en_col).value or '') !=
                      str(ws_orig.cell(r, en_col).value or ''))
    if mismatch:
        log(f'[红线警告] EN列有{mismatch}行不一致！', lang)

    out_path = OUTPUT_DIR / f'{lang}_translated_{excel_path.stem}.xlsx'
    wb.save(str(out_path))
    log(f'产出: {out_path}', lang)
    return out_path


# ── 主流程：多语种并行 ─────────────────────────────────────────────────────────

def process_lang(lang: str):
    log(f'开始处理', lang)
    try:
        excel_path = export_by_filter(lang)
        if not excel_path:
            log('无空key，跳过', lang)
            return lang, None, 0, 0
        out_path = translate_and_fill(excel_path, lang)
        log(f'完成！产出: {out_path.name}', lang)
        wb = openpyxl.load_workbook(str(excel_path))
        cnt = wb.active.max_row - 1
        return lang, out_path, cnt, 0
    except Exception as e:
        log(f'失败: {e}', lang)
        import traceback; traceback.print_exc()
        return lang, None, 0, 1


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--langs', nargs='+', required=True,
                        help='目标语种列名，如 ar uk')
    args = parser.parse_args()

    for lang in args.langs:
        if lang not in GOOGLE_LANG:
            print(f'[错误] {lang} 未在GOOGLE_LANG映射表中')
            sys.exit(1)

    log(f'目标语种: {args.langs}，{len(args.langs)}个并行处理')
    t_start = time.time()

    results = []
    with ThreadPoolExecutor(max_workers=1) as ex:  # 语种间串行，防Google限流
        futures = {ex.submit(process_lang, lang): lang for lang in args.langs}
        for f in as_completed(futures):
            results.append(f.result())

    print(f'\n{"="*60}')
    print(f'全部完成！总耗时: {time.time()-t_start:.0f}s')
    for lang, path, cnt, err in results:
        if path:
            print(f'  {lang}: {cnt}条 → {path.name}')
        else:
            print(f'  {lang}: 失败或无数据')
    print(f'产出目录: {OUTPUT_DIR}')


if __name__ == '__main__':
    main()

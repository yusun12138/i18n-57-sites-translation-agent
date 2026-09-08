#!/usr/bin/env python3
"""
从多语言配置中心 Headless API 导出数据（生产环境）
- 覆盖三类文案：前端key(type=1) / 提示语(type=2) / 错误码(type=3)
- 品牌：SHEIN
- 导出流程：提交异步任务 → 轮询状态 → 下载OSS文件

用法:
  python3 export_from_api.py                        # 导出全部三类
  python3 export_from_api.py --type 1               # 只导前端key
  python3 export_from_api.py --type 1 --empty-langs da fi
"""

import argparse, json, os, sys, time
from datetime import datetime
from pathlib import Path

import requests

# ── MCP Proxy 配置（生产环境）────────────────────────────────────────────────
MCP_PROXY   = os.environ.get('MCP_PROXY') or 'http://127.0.0.1:30001/rest'
GW_NAME     = os.environ.get('LCC_GW_NAME') or 'language-configuration-center'
MCP_DC      = os.environ.get('MCP_PROXY_DC') or 'central_p2_dc'
# mcp-proxy 自动从本地 kite 登录态获取真实 ULP token
ULP_COOKIE  = os.environ.get('ULP_COOKIE') or 'ulp_token={{ulp_token}}'

KEYS_PATH        = f'/{GW_NAME}/headless/api/v1/keys'
EXPORT_PATH      = f'/{GW_NAME}/headless/api/v1/keys/export'
EXPORT_TASK_PATH = f'/{GW_NAME}/headless/api/v1/keys/export-task'

# ── 文案类型配置（SHEIN品牌）─────────────────────────────────────────────────
# type → (label, brand, platform列表)
# 枚举来源：接入指南 brandPlatformMapping
TYPE_CONFIG = {
    1: ('前端key',  1, [1, 2, 3, 4, 5]),   # SHEIN APP/H5/PWA/PC/PAY
    2: ('提示语',   2, [6, 7]),              # SHEIN B2C/社区
    3: ('错误码',   3, [12]),               # SHEIN 错误码
}

PAGE_SIZE    = 100   # 接口最大100
POLL_MAX     = 60    # 最多轮询60次（5分钟）
POLL_INTERVAL = 5    # 每5秒轮询一次

BASE_DIR    = Path(os.environ.get('I18N_WORKDIR') or Path(__file__).resolve().parents[1]).resolve()
OUTPUT_DIR  = BASE_DIR / '01_原始数据'
OUTPUT_DIR.mkdir(exist_ok=True)


def log(msg): print(f'[{datetime.now().strftime("%H:%M:%S")}] {msg}', flush=True)


def _headers(proxy_path: str) -> dict:
    return {
        'Content-Type': 'application/json',
        'mcp-proxy-path': proxy_path,
        'mcp-proxy-dc': MCP_DC,
        'Cookie': ULP_COOKIE,
    }


def count_keys(type_val: int, empty_langs: list = None) -> int:
    """查询某类文案的总数"""
    label, brand, platforms = TYPE_CONFIG[type_val]
    payload = {
        'page': 1, 'pageSize': 1,
        'searchType': 2,
        'type': type_val,
        'brand': brand,
        'platform': platforms,
        'keyId': [], 'projects': [],
        'emptyLangList': empty_langs or [],
        'syncStatus': [], 'releaseStatus': [],
    }
    r = requests.post(MCP_PROXY, headers=_headers(KEYS_PATH),
                      data=json.dumps(payload), timeout=30)
    r.raise_for_status()
    resp = r.json()
    if resp.get('code') != '0':
        raise ValueError(f"查询失败: {resp}")
    return resp['info']['total']


def submit_export(type_val: int, empty_langs: list = None) -> str:
    """提交导出任务，返回 taskId"""
    label, brand, platforms = TYPE_CONFIG[type_val]
    payload = {
        'fileType': 1,            # Excel
        'searchType': 2,
        'type': type_val,
        'brand': brand,
        'platform': platforms,
        'keyId': [], 'projects': [],
        'emptyLangList': empty_langs or [],
        'releaseStatus': [],
    }
    r = requests.post(MCP_PROXY, headers=_headers(EXPORT_PATH),
                      data=json.dumps(payload), timeout=60)
    r.raise_for_status()
    resp = r.json()
    if resp.get('code') != '0':
        raise ValueError(f"提交导出失败: {resp}")
    task_id = resp['info']['taskId']
    log(f'  任务已提交 taskId={task_id}')
    return task_id


def poll_export_task(task_id: str) -> str:
    """轮询导出任务，返回下载 URL"""
    log(f'  开始轮询（每{POLL_INTERVAL}s，最多{POLL_MAX}次）...')
    time.sleep(3)  # 首次等3s
    for i in range(POLL_MAX):
        r = requests.get(MCP_PROXY,
                         headers=_headers(f'{EXPORT_TASK_PATH}?taskId={task_id}'),
                         timeout=30)
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != '0':
            raise ValueError(f"查询任务状态失败: {resp}")
        status = resp['info']['status']
        if status == 3:
            url = resp['info'].get('url', '')
            log(f'  导出完成！下载地址: {url[:80]}...')
            return url
        elif status == 4:
            raise ValueError('导出任务失败（status=4），建议缩小范围后重试')
        else:
            log(f'  [{i+1}/{POLL_MAX}] 处理中（status={status}）...')
            time.sleep(POLL_INTERVAL)
    raise TimeoutError(f'轮询超时（{POLL_MAX * POLL_INTERVAL}s），taskId={task_id}')


def download_file(url: str, out_path: Path):
    """从 OSS URL 下载文件"""
    log(f'  下载中: {out_path.name}')
    r = requests.get(url, timeout=120, stream=True)
    r.raise_for_status()
    with open(out_path, 'wb') as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    size_kb = out_path.stat().st_size // 1024
    log(f'  已保存: {out_path}（{size_kb} KB）')


def export_type(type_val: int, empty_langs: list = None):
    """导出单类文案：提交任务 → 轮询 → 下载"""
    label, brand, platforms = TYPE_CONFIG[type_val]
    log(f'\n{"="*60}')
    log(f'导出类型: {label}（type={type_val}, brand={brand}, platforms={platforms}）')
    if empty_langs:
        log(f'语种为空筛选: {empty_langs}')

    # 先查一下总数，给用户预期
    try:
        total = count_keys(type_val, empty_langs)
        log(f'  符合条件的 key 总数: {total}')
        if total == 0:
            log('  无数据，跳过')
            return
    except Exception as e:
        log(f'  [警告] 查询总数失败（{e}），继续导出')

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    lang_suffix = ('_empty_' + '_'.join(empty_langs)) if empty_langs else ''
    out_path = OUTPUT_DIR / f'type{type_val}_{label}{lang_suffix}_{ts}.xlsx'

    task_id = submit_export(type_val, empty_langs)
    url = poll_export_task(task_id)
    download_file(url, out_path)


def main():
    parser = argparse.ArgumentParser(description='从文案系统Headless API导出多语言数据')
    parser.add_argument('--type', type=int, choices=[1, 2, 3],
                        help='文案类型：1=前端key，2=提示语，3=错误码。不填则导出全部')
    parser.add_argument('--empty-langs', nargs='+', metavar='LANG',
                        help='只导指定语种为空的key，如 --empty-langs da fi lv')
    args = parser.parse_args()

    types = [args.type] if args.type else [1, 2, 3]

    log(f'导出类型: {[TYPE_CONFIG[t][0] for t in types]}')
    log(f'语种为空筛选: {args.empty_langs or "不筛选（全量）"}')
    log(f'输出目录: {OUTPUT_DIR}')
    log(f'注意：使用生产环境，mcp-proxy 自动注入 ulp_token')

    for t in types:
        try:
            export_type(t, args.empty_langs)
        except Exception as e:
            log(f'[错误] type={t} 导出失败: {e}')
            sys.exit(1)

    log('\n全部导出完成！')
    log(f'文件位于: {OUTPUT_DIR}')
    log('下一步: python3 translate_multi_site.py --langs <语种列名>')


if __name__ == '__main__':
    main()

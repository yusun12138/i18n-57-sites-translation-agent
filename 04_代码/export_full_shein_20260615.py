#!/usr/bin/env python3
import json
import os
import time
from datetime import datetime
from pathlib import Path

import requests

MCP_PROXY = os.environ.get('MCP_PROXY') or 'http://127.0.0.1:30001/rest'
GW_NAME = os.environ.get('LCC_GW_NAME') or 'language-configuration-center'
MCP_DC = os.environ.get('MCP_PROXY_DC') or 'central_p2_dc'
ULP_COOKIE = os.environ.get('ULP_COOKIE') or 'ulp_token={{ulp_token}}'

KEYS_PATH = f'/{GW_NAME}/headless/api/v1/keys'
EXPORT_PATH = f'/{GW_NAME}/headless/api/v1/keys/export'
EXPORT_TASK_PATH = f'/{GW_NAME}/headless/api/v1/keys/export-task'

BASE_DIR = Path(os.environ.get('I18N_WORKDIR') or Path(__file__).resolve().parents[1]).resolve()
OUTPUT_DIR = BASE_DIR / '01_原始数据'
OUTPUT_DIR.mkdir(exist_ok=True)

POLL_MAX = 90
POLL_INTERVAL = 5


def log(msg):
    print(f'[{datetime.now().strftime("%H:%M:%S")}] {msg}', flush=True)


def headers(proxy_path):
    return {
        'Content-Type': 'application/json',
        'mcp-proxy-path': proxy_path,
        'mcp-proxy-dc': MCP_DC,
        'Cookie': ULP_COOKIE,
    }


def count_keys(type_val, brand, platforms):
    payload = {
        'page': 1,
        'pageSize': 1,
        'searchType': 2,
        'type': type_val,
        'brand': brand,
        'platform': platforms,
        'keyId': [],
        'projects': [],
        'emptyLangList': [],
        'syncStatus': [],
        'releaseStatus': [],
    }
    r = requests.post(MCP_PROXY, headers=headers(KEYS_PATH), data=json.dumps(payload), timeout=30)
    r.raise_for_status()
    resp = r.json()
    if resp.get('code') != '0':
        raise ValueError(f'查询失败: {resp}')
    info = resp.get('info') or {}
    if 'total' in info:
        return info['total']
    if isinstance(info.get('meta'), dict) and 'count' in info['meta']:
        return info['meta']['count']
    raise KeyError(f'响应中没有 total/count: {resp}')


def submit_export(type_val, brand, platforms):
    payload = {
        'fileType': 1,
        'searchType': 2,
        'type': type_val,
        'brand': brand,
        'platform': platforms,
        'keyId': [],
        'projects': [],
        'emptyLangList': [],
        'releaseStatus': [],
    }
    r = requests.post(MCP_PROXY, headers=headers(EXPORT_PATH), data=json.dumps(payload), timeout=60)
    r.raise_for_status()
    resp = r.json()
    if resp.get('code') != '0':
        raise ValueError(f'提交导出失败: {resp}')
    return resp['info']['taskId']


def poll_task(task_id):
    time.sleep(3)
    for i in range(POLL_MAX):
        r = requests.get(
            MCP_PROXY,
            headers=headers(f'{EXPORT_TASK_PATH}?taskId={task_id}'),
            timeout=30,
        )
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != '0':
            raise ValueError(f'查询任务失败: {resp}')
        info = resp['info']
        status = info['status']
        if status == 3:
            return info.get('url', '')
        if status == 4:
            raise ValueError(f'导出任务失败 status=4 taskId={task_id}')
        log(f'  [{i + 1}/{POLL_MAX}] 处理中 status={status}')
        time.sleep(POLL_INTERVAL)
    raise TimeoutError(f'导出轮询超时 taskId={task_id}')


def download(url, out_path):
    last_error = None
    for attempt in range(1, 7):
        try:
            r = requests.get(url, timeout=180, stream=True)
            r.raise_for_status()
            with open(out_path, 'wb') as f:
                for chunk in r.iter_content(1024 * 128):
                    if chunk:
                        f.write(chunk)
            break
        except Exception as exc:
            last_error = exc
            if out_path.exists():
                out_path.unlink()
            log(f'  下载失败 attempt={attempt}/6: {exc}')
            time.sleep(5 * attempt)
    else:
        raise last_error
    log(f'  已保存 {out_path.name} ({out_path.stat().st_size // 1024} KB)')


def export_one(label, type_val, brand, platforms, out_name):
    log('=' * 60)
    log(f'导出 {label}: type={type_val}, brand={brand}, platforms={platforms}')
    total = count_keys(type_val, brand, platforms)
    log(f'  符合条件 key 总数: {total}')
    if total == 0:
        return None
    task_id = submit_export(type_val, brand, platforms)
    log(f'  taskId={task_id}')
    url = poll_task(task_id)
    out_path = OUTPUT_DIR / out_name
    download(url, out_path)
    return out_path


def main():
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    outputs = []
    for platform in [1, 2, 3, 4, 5]:
        outputs.append(
            export_one(
                f'type1_前端key_platform{platform}',
                1,
                1,
                [platform],
                f'type1_前端key_platform{platform}_full_{ts}.xlsx',
            )
        )
    outputs.append(export_one('type2_提示语', 2, 2, [6, 7], f'type2_提示语_full_{ts}.xlsx'))
    outputs.append(export_one('type3_错误码', 3, 3, [12], f'type3_错误码_full_{ts}.xlsx'))
    log('今日全量导出完成:')
    for path in outputs:
        if path:
            log(f'  {path}')


if __name__ == '__main__':
    main()

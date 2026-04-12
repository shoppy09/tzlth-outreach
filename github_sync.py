#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub 資料同步模組
將 targets.xlsx 的變更自動 commit 回 GitHub repo
確保 Render 重啟後資料不遺失

環境變數需求：
  GITHUB_TOKEN  - GitHub Personal Access Token（repo scope）
  GITHUB_REPO   - 格式: owner/repo（例如 shoppy09/tzlth-outreach）
"""

import os
import base64
import json
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO  = os.environ.get("GITHUB_REPO", "shoppy09/tzlth-outreach")
API_BASE     = "https://api.github.com"

def _headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
        "User-Agent": "tzlth-outreach-sync"
    }

def _api(method: str, path: str, body: dict = None):
    url = f"{API_BASE}/repos/{GITHUB_REPO}{path}"
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, headers=_headers(), method=method)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"GitHub API {method} {path} → {e.code}: {e.read().decode()}")

def push_file(local_path: Path, repo_path: str, message: str = None) -> bool:
    """把本地檔案推送到 GitHub repo。回傳是否成功。"""
    if not GITHUB_TOKEN:
        return False  # 沒有 token，靜默跳過
    try:
        # 讀取本地檔案
        content_b64 = base64.b64encode(local_path.read_bytes()).decode()
        # 取得目前 GitHub 版本的 SHA（更新時需要）
        sha = None
        try:
            info = _api("GET", f"/contents/{repo_path}")
            sha = info.get("sha")
        except Exception:
            pass  # 檔案不存在也 OK（新建）
        # 上傳
        body = {
            "message": message or f"sync: 自動更新 {repo_path}（{datetime.now().strftime('%Y-%m-%d %H:%M')}）",
            "content": content_b64,
            "branch": "master"
        }
        if sha:
            body["sha"] = sha
        _api("PUT", f"/contents/{repo_path}", body)
        return True
    except Exception as e:
        print(f"[github_sync] 同步失敗（不影響本地操作）: {e}")
        return False

def sync_targets(data_file: Path) -> bool:
    """同步 targets.xlsx 到 GitHub。"""
    return push_file(data_file, "data/targets.xlsx",
                     f"sync: 外展資料更新（{datetime.now().strftime('%Y-%m-%d %H:%M')}）")

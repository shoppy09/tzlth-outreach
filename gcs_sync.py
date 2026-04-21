#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GCS 資料同步模組
取代原 github_sync.py，將 targets.xlsx 持久化到 Google Cloud Storage
Cloud Run 重啟後從 GCS 恢復資料

環境變數需求：
  GCS_SA_KEY_B64  - Service Account JSON key（base64 編碼）
  GCS_BUCKET      - GCS bucket 名稱（預設：tzlth-outreach-data）
"""

import os
import json
import base64
from pathlib import Path

def _get_client():
    """從環境變數建立 GCS client"""
    key_b64 = os.environ.get('GCS_SA_KEY_B64', '')
    if not key_b64:
        return None, None
    try:
        from google.oauth2 import service_account
        from google.cloud import storage
        sa_info = json.loads(base64.b64decode(key_b64))
        creds = service_account.Credentials.from_service_account_info(
            sa_info,
            scopes=['https://www.googleapis.com/auth/cloud-platform']
        )
        client = storage.Client(credentials=creds, project=sa_info['project_id'])
        return client, sa_info['project_id']
    except Exception as e:
        print(f'[gcs_sync] 初始化失敗: {e}')
        return None, None

BUCKET_NAME = os.environ.get('GCS_BUCKET', 'tzlth-outreach-data')

def download_targets(local_path: Path) -> bool:
    """從 GCS 下載 targets.xlsx 到本地路徑"""
    client, _ = _get_client()
    if not client:
        return False
    try:
        local_path.parent.mkdir(parents=True, exist_ok=True)
        bucket = client.bucket(BUCKET_NAME)
        blob = bucket.blob('data/targets.xlsx')
        if not blob.exists():
            print('[gcs_sync] GCS 無資料檔（首次部署）')
            return False
        blob.download_to_filename(str(local_path))
        print(f'[gcs_sync] 從 GCS 下載成功 → {local_path}')
        return True
    except Exception as e:
        print(f'[gcs_sync] 下載失敗: {e}')
        return False

def upload_targets(local_path: Path):
    """上傳 targets.xlsx 到 GCS（覆寫）"""
    client, _ = _get_client()
    if not client:
        return
    try:
        bucket = client.bucket(BUCKET_NAME)
        blob = bucket.blob('data/targets.xlsx')
        blob.upload_from_filename(str(local_path))
    except Exception as e:
        print(f'[gcs_sync] 上傳失敗: {e}')

def upload_pdf(local_path: Path, gcs_filename: str):
    """上傳 PDF 到 GCS outputs/ 資料夾（背景持久化）"""
    client, _ = _get_client()
    if not client:
        return
    try:
        bucket = client.bucket(BUCKET_NAME)
        blob = bucket.blob(f'outputs/{gcs_filename}')
        blob.upload_from_filename(str(local_path))
    except Exception as e:
        print(f'[gcs_sync] PDF 上傳失敗: {e}')

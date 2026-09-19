# -*- coding: utf-8 -*-
"""
Vercel Serverless Function 진입점
파일명: api/index.py
"""

import os
import sys

# 프로젝트 루트 디렉터리를 sys.path에 추가
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from web_app import app

# Vercel Serverless WSGI 앱 핸들러
# (Vercel @vercel/python 빌더가 app 인스턴스를 찾아 실행합니다)

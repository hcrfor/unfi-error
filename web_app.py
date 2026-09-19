# -*- coding: utf-8 -*-
"""
산림 도시 조사 데이터 오류 검수 로컬 웹 애플리케이션
파일명: web_app.py
작성자: Antigravity (시니어 풀스택 개발자)
설명:
    - Flask 기반의 경량 로컬 웹 서버
    - 직관적인 REST API 및 폴더 선택 대화상자(Tkinter) 연동
    - 결과 엑셀 보고서 다운로드 및 탐색기 열기 지원
"""

import os
import sys
import threading
import webbrowser

# 윈도우 콘솔 UTF-8 출력 호환성 설정
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

from flask import Flask, render_template, request, jsonify, send_file
import excel_inspector

app = Flask(__name__)

# 기본 작업 디렉터리 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REF_EXCEL_FILE = os.path.join(BASE_DIR, "2026 도시 조사결과표(한성안).xlsx")


@app.route('/')
def index():
    """메인 대시보드 페이지 렌더링"""
    return render_template('index.html')


@app.route('/api/select-folder', methods=['POST'])
def select_folder():
    """
    윈도우 기본 폴더 선택 대화상자(Tkinter FileDialog)를 호출합니다.
    """
    selected_path = None

    def ask_folder_dialog():
        nonlocal selected_path
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()           # 메인 윈도우 숨김
            root.attributes('-topmost', True)  # 화면 최상단으로 띄움
            chosen = filedialog.askdirectory(
                title="검수할 엑셀 파일이 포함된 폴더를 선택하세요",
                initialdir=BASE_DIR
            )
            root.destroy()
            if chosen:
                # 윈도우 역슬래시 경로 정리
                selected_path = os.path.normpath(chosen)
        except Exception as e:
            print(f"Tkinter 다이얼로그 호출 오류: {e}")

    # 별도 스레드에서 UI 다이얼로그 실행
    t = threading.Thread(target=ask_folder_dialog)
    t.start()
    t.join(timeout=60)  # 최대 60초 대기

    if selected_path:
        return jsonify({"success": True, "folder_path": selected_path})
    return jsonify({"success": False, "folder_path": None, "message": "폴더 선택이 취소되었거나 지원되지 않습니다."})


@app.route('/api/start-inspection', methods=['POST'])
def start_inspection():
    """
    지정된 폴더에 대해 정밀 오류 검수를 수행하고 결과를 반환합니다.
    """
    data = request.get_json() or {}
    target_folder = data.get('target_folder', '').strip()

    if not target_folder:
        return jsonify({"success": False, "error": "검수할 폴더 경로를 입력해주세요."}), 400

    # 상대 경로인 경우 절대 경로로 변환
    if not os.path.isabs(target_folder):
        target_folder = os.path.normpath(os.path.join(BASE_DIR, target_folder))

    if not os.path.exists(target_folder):
        return jsonify({"success": False, "error": f"폴더가 존재하지 않습니다: {target_folder}"}), 404

    try:
        # 검수 엔진 모듈 최신 상태로 강제 리로드
        import importlib
        importlib.reload(excel_inspector)

        # 사용자 기본 다운로드 폴더 경로 (C:\Users\<사용자>\Downloads)
        downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(downloads_dir):
            downloads_dir = BASE_DIR

        # 검수 엔진 실행 (결과 파일은 다운로드 폴더에 자동 생성)
        result = excel_inspector.run_inspection(
            target_dir=target_folder,
            ref_file=REF_EXCEL_FILE,
            output_dir=downloads_dir
        )
        return jsonify({"success": True, "data": result})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/download-report', methods=['GET'])
def download_report():
    """
    생성된 오류 검수 결과 엑셀 파일을 다운로드합니다.
    (다운로드 폴더 우선 탐색)
    """
    folder = request.args.get('folder', '').strip()
    filename = request.args.get('filename', '').strip()

    if not filename:
        return "파일명이 필요합니다.", 400

    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    cand_path = os.path.normpath(os.path.join(downloads_dir, filename))

    # 다운로드 폴더에 없으면 전달된 folder에서 탐색
    if not os.path.exists(cand_path) and folder:
        if not os.path.isabs(folder):
            folder = os.path.normpath(os.path.join(BASE_DIR, folder))
        cand_path = os.path.normpath(os.path.join(folder, filename))

    if not os.path.exists(cand_path):
        return f"요청한 결과 파일을 찾을 수 없습니다: {filename}", 404

    return send_file(
        cand_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    """
    결과 파일이 저장된 다운로드 폴더를 윈도우 파일 탐색기에서 엽니다.
    """
    data = request.get_json() or {}
    folder_path = data.get('folder_path', '').strip()

    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')

    if not folder_path or not os.path.exists(folder_path):
        folder_path = downloads_dir if os.path.exists(downloads_dir) else BASE_DIR

    try:
        os.startfile(folder_path)
        return jsonify({"success": True, "opened_path": folder_path})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


def open_browser():
    """서버 구동 후 기본 브라우저를 자동으로 엽니다."""
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == '__main__':
    # 0.8초 후 브라우저 자동 오픈
    threading.Timer(0.8, open_browser).start()
    print("=" * 60)
    print(" 산림 도시 조사 엑셀 데이터 오류 검수 시스템 웹 서버 시작")
    print(" 접속 주소: http://127.0.0.1:5000")
    print(" 종료하려면 터미널에서 Ctrl+C를 누르세요.")
    print("=" * 60)
    app.run(host='127.0.0.1', port=5000, debug=False)

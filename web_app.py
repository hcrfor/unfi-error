# -*- coding: utf-8 -*-
"""
산림 도시 조사 데이터 오류 검수 웹 애플리케이션
파일명: web_app.py
작성자: Antigravity (시니어 풀스택 개발자)
설명:
    - Flask 기반의 반응형 웹 검수 시스템
    - Vercel Serverless 클라우드 환경 및 로컬 데스크톱 환경 동시 지원
    - 웹 드래그 앤 드롭 파일/폴더 업로드 검수 API 지원
    - 결과 엑셀 보고서 생성 및 안전한 다운로드 제공
"""

import os
import sys
import uuid
import tempfile
import threading
import webbrowser
from werkzeug.utils import secure_filename

# 윈도우 콘솔 UTF-8 출력 호환성 설정
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

from flask import Flask, render_template, request, jsonify, send_file
import excel_inspector

# 기본 작업 디렉터리 설정 (프로젝트 루트)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REF_EXCEL_FILE = os.path.join(BASE_DIR, "2026 도시 조사결과표(한성안).xlsx")

# Vercel 환경 감지 (VERCEL 환경변수가 있으면 클라우드 환경)
IS_VERCEL = bool(os.environ.get('VERCEL') or os.environ.get('NOW_REGION'))

# Flask 앱 인스턴스 생성 (정적 파일 및 템플릿 경로 명시)
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)

# 최대 파일 업로드 크기 (50MB)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


def get_default_output_dir():
    """결과 보고서를 저장할 안전한 디렉터리를 반환합니다."""
    if IS_VERCEL:
        return tempfile.gettempdir()
    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    if os.path.exists(downloads_dir):
        return downloads_dir
    return tempfile.gettempdir()


@app.route('/')
def index():
    """메인 대시보드 페이지 렌더링"""
    return render_template('index.html', is_vercel=IS_VERCEL)


@app.route('/api/upload-and-inspect', methods=['POST'])
def upload_and_inspect():
    """
    [웹 클라우드 지원] 사용자가 브라우저에서 드래그 앤 드롭 또는 파일 선택으로
    업로드한 엑셀 파일들을 받아 임시 공간에서 검수하고 결과를 반환합니다.
    """
    uploaded_files = request.files.getlist('files')
    if not uploaded_files or len(uploaded_files) == 0:
        return jsonify({"success": False, "error": "업로드된 엑셀 파일이 없습니다."}), 400

    # 임시 작업 폴더 생성
    session_id = uuid.uuid4().hex[:8]
    temp_dir = os.path.join(tempfile.gettempdir(), f"unfi_inspect_{session_id}")
    os.makedirs(temp_dir, exist_ok=True)

    saved_count = 0
    for f in uploaded_files:
        if not f.filename:
            continue
        fname = os.path.basename(f.filename)
        # 엑셀 파일(.xlsx) 및 임시 파일 제외
        if fname.lower().endswith('.xlsx') and not fname.startswith('~$'):
            save_path = os.path.join(temp_dir, fname)
            f.save(save_path)
            saved_count += 1

    if saved_count == 0:
        return jsonify({"success": False, "error": "유효한 .xlsx 엑셀 파일이 없습니다."}), 400

    try:
        output_dir = get_default_output_dir()
        result = excel_inspector.run_inspection(
            target_dir=temp_dir,
            ref_file=REF_EXCEL_FILE,
            output_dir=output_dir
        )
        result['output_dir'] = output_dir
        return jsonify({"success": True, "data": result})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": f"검수 처리 중 오류 발생: {str(e)}"}), 500


@app.route('/api/sample-inspection', methods=['POST'])
def sample_inspection():
    """
    내장된 excel_1-4 샘플 데이터를 이용해 웹에서 1클릭으로 즉시 검수를 체험합니다.
    """
    sample_dir = os.path.join(BASE_DIR, 'excel_1-4')
    if not os.path.exists(sample_dir):
        return jsonify({"success": False, "error": "내장 샘플 폴더(excel_1-4)를 찾을 수 없습니다."}), 404

    try:
        output_dir = get_default_output_dir()
        result = excel_inspector.run_inspection(
            target_dir=sample_dir,
            ref_file=REF_EXCEL_FILE,
            output_dir=output_dir
        )
        result['output_dir'] = output_dir
        return jsonify({"success": True, "data": result})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": f"샘플 검수 중 오류 발생: {str(e)}"}), 500


@app.route('/api/select-folder', methods=['POST'])
def select_folder():
    """
    [로컬 PC 전용] 윈도우 기본 폴더 선택 대화상자(Tkinter FileDialog)를 호출합니다.
    (Vercel 클라우드에서는 지원되지 않음을 안내)
    """
    if IS_VERCEL:
        return jsonify({
            "success": False,
            "folder_path": None,
            "message": "클라우드(Vercel) 환경에서는 상단의 '파일/폴더 업로드' 기능을 이용해 주세요."
        })

    selected_path = None

    def ask_folder_dialog():
        nonlocal selected_path
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            chosen = filedialog.askdirectory(
                title="검수할 엑셀 파일이 포함된 폴더를 선택하세요",
                initialdir=BASE_DIR
            )
            root.destroy()
            if chosen:
                selected_path = os.path.normpath(chosen)
        except Exception as e:
            print(f"Tkinter 다이얼로그 오류: {e}")

    t = threading.Thread(target=ask_folder_dialog)
    t.start()
    t.join(timeout=60)

    if selected_path:
        return jsonify({"success": True, "folder_path": selected_path})
    return jsonify({"success": False, "folder_path": None, "message": "폴더 선택이 취소되었습니다."})


@app.route('/api/start-inspection', methods=['POST'])
def start_inspection():
    """
    [로컬 PC 폴더 경로 검수] 지정된 폴더 경로에 대해 정밀 오류 검수를 수행합니다.
    """
    data = request.get_json() or {}
    target_folder = data.get('target_folder', '').strip()

    if not target_folder:
        return jsonify({"success": False, "error": "검수할 폴더 경로를 입력해주세요."}), 400

    if not os.path.isabs(target_folder):
        target_folder = os.path.normpath(os.path.join(BASE_DIR, target_folder))

    if not os.path.exists(target_folder):
        return jsonify({"success": False, "error": f"폴더가 존재하지 않습니다: {target_folder}"}), 404

    try:
        output_dir = get_default_output_dir()
        result = excel_inspector.run_inspection(
            target_dir=target_folder,
            ref_file=REF_EXCEL_FILE,
            output_dir=output_dir
        )
        result['output_dir'] = output_dir
        return jsonify({"success": True, "data": result})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/download-report', methods=['GET'])
def download_report():
    """
    생성된 오류 검수 결과 엑셀 보고서 파일을 다운로드합니다.
    """
    folder = request.args.get('folder', '').strip()
    filename = request.args.get('filename', '').strip()

    if not filename:
        return "파일명이 필요합니다.", 400

    # 탐색 후보 디렉터리 목록
    candidate_dirs = [
        folder if folder and os.path.exists(folder) else None,
        get_default_output_dir(),
        tempfile.gettempdir(),
        os.path.join(os.path.expanduser('~'), 'Downloads'),
        BASE_DIR
    ]

    target_path = None
    for d in candidate_dirs:
        if d:
            cand = os.path.normpath(os.path.join(d, filename))
            if os.path.exists(cand):
                target_path = cand
                break

    if not target_path or not os.path.exists(target_path):
        return f"요청한 결과 파일을 찾을 수 없습니다: {filename}", 404

    return send_file(
        target_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    """
    [로컬 PC 전용] 결과 파일이 저장된 폴더를 파일 탐색기에서 엽니다.
    """
    if IS_VERCEL:
        return jsonify({
            "success": False,
            "error": "웹(Vercel) 환경에서는 상단의 '결과 보고서 다운로드' 버튼을 이용해 주세요."
        })

    data = request.get_json() or {}
    folder_path = data.get('folder_path', '').strip()
    downloads_dir = get_default_output_dir()

    if not folder_path or not os.path.exists(folder_path):
        folder_path = downloads_dir

    try:
        os.startfile(folder_path)
        return jsonify({"success": True, "opened_path": folder_path})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


def open_browser():
    """로컬 구동 시 브라우저 자동 오픈"""
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == '__main__':
    if not IS_VERCEL:
        threading.Timer(0.8, open_browser).start()
    print("=" * 60)
    print(" 산림 도시 조사 엑셀 데이터 오류 검수 시스템 웹 서버 시작")
    print(" 접속 주소: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5000, debug=False)

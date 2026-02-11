"""
Flask 웹 애플리케이션: 방화벽 정책 검증

로컬 웹 서버에서 실행하여 파일을 업로드하고 정책을 검증합니다.
"""

import sys
from pathlib import Path

# 상위 디렉터리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from flask import Flask, render_template, request, send_file, jsonify, session
from werkzeug.utils import secure_filename
import os
from datetime import datetime
import pandas as pd
from core import parse_policy_file, parse_target_file, validate_policy_changes
from core.utils import get_summary_dict
from core.vendor import PaloaltoParser, SECUIParser
from rich.console import Console
import tempfile
import shutil
import uuid

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = os.urandom(24)  # 세션 암호화용
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 최대 100MB 파일 업로드
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

console = Console()

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """파일 확장자 확인"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_upload_dir():
    """세션별 업로드 디렉터리 경로 반환 (SecureCookieSession에는 sid가 없으므로 _upload_id 사용)"""
    if '_upload_id' not in session:
        session['_upload_id'] = str(uuid.uuid4())
    return Path(app.config['UPLOAD_FOLDER']) / session['_upload_id']


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    """SECUI 파일의 시트 목록 가져오기"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '파일이 없습니다.'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '파일을 선택해주세요.'}), 400
        
        # 임시 저장
        upload_dir = get_upload_dir()
        upload_dir.mkdir(parents=True, exist_ok=True)
        temp_path = upload_dir / f"temp_{secure_filename(file.filename)}"
        file.save(str(temp_path))
        
        # 시트 목록 가져오기 (openpyxl 사용 - Excel 설치 없이 동작)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(temp_path), read_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception:
            sheets = SECUIParser.get_sheets(str(temp_path))
        
        # 임시 파일 삭제
        temp_path.unlink()
        
        return jsonify({'sheets': sheets})
    except Exception as e:
        return jsonify({'error': f'시트 목록 가져오기 오류: {str(e)}'}), 500


@app.route('/upload', methods=['POST'])
def upload_files():
    """파일 업로드 및 검증 처리"""
    try:
        # 세션 초기화 (대용량 데이터 제거, 쿠키 크기 제한 방지)
        session.pop('results', None)
        session.pop('report_file', None)
        session.pop('report_filename', None)
        session.pop('summary', None)
        
        # 벤더 정보 가져오기
        vendor = request.form.get('vendor', 'Paloalto')
        
        # 파일 확인
        if vendor == 'SECUI':
            if 'running_file' not in request.files:
                return jsonify({'error': 'Running 정책 파일은 필수입니다.'}), 400
            running_file = request.files['running_file']
            running_sheet = request.form.get('running_sheet')
            candidate_sheet = request.form.get('candidate_sheet')
            candidate_file = request.files.get('candidate_file')
            use_same_file = not candidate_file or not (getattr(candidate_file, 'filename', None) or '').strip()

            if running_file.filename == '':
                return jsonify({'error': 'Running 정책 파일을 선택해주세요.'}), 400
            if not running_sheet or not candidate_sheet:
                return jsonify({'error': '작업 전/후 시트를 선택해주세요.'}), 400
            if not allowed_file(running_file.filename):
                return jsonify({'error': 'Excel 파일만 업로드 가능합니다 (.xlsx, .xls)'}), 400

            upload_dir = get_upload_dir()
            upload_dir.mkdir(parents=True, exist_ok=True)
            running_path = upload_dir / secure_filename(running_file.filename)
            running_file.save(str(running_path))

            if use_same_file:
                candidate_path = running_path
            else:
                if not allowed_file(candidate_file.filename):
                    return jsonify({'error': 'Candidate 파일은 Excel(.xlsx, .xls)만 가능합니다.'}), 400
                candidate_path = upload_dir / secure_filename(candidate_file.filename)
                candidate_file.save(str(candidate_path))
        else:
            # Paloalto는 두 파일 필요
            if 'running_file' not in request.files or 'candidate_file' not in request.files:
                return jsonify({'error': 'Running 정책 파일과 Candidate 정책 파일은 필수입니다.'}), 400
            
            running_file = request.files['running_file']
            candidate_file = request.files['candidate_file']
            
            if running_file.filename == '' or candidate_file.filename == '':
                return jsonify({'error': '파일을 선택해주세요.'}), 400
            
            if not (allowed_file(running_file.filename) and allowed_file(candidate_file.filename)):
                return jsonify({'error': 'Excel 파일만 업로드 가능합니다 (.xlsx, .xls)'}), 400
            
            # 파일 저장
            upload_dir = get_upload_dir()
            upload_dir.mkdir(parents=True, exist_ok=True)
            
            running_path = upload_dir / secure_filename(running_file.filename)
            candidate_path = upload_dir / secure_filename(candidate_file.filename)
            
            running_file.save(str(running_path))
            candidate_file.save(str(candidate_path))
        
        # 대상 파일 저장
        target_files = request.files.getlist('target_files')
        target_paths = []
        for target_file in target_files:
            if target_file.filename and allowed_file(target_file.filename):
                target_path = upload_dir / secure_filename(target_file.filename)
                target_file.save(str(target_path))
                target_paths.append(target_path)
        
        # 정책 파일 파싱
        console.print(f"[cyan]Running 정책 파싱 중... (벤더: {vendor})[/cyan]")
        if vendor == 'SECUI':
            running_policy_data = SECUIParser.parse_policy_file(str(running_path), running_sheet)
        else:
            running_policy_data = PaloaltoParser.parse_policy_file(str(running_path))
        console.print(f"[green]✓ {len(running_policy_data)}개 정책 발견[/green]")
        
        console.print(f"[cyan]Candidate 정책 파싱 중... (벤더: {vendor})[/cyan]")
        if vendor == 'SECUI':
            candidate_policy_data = SECUIParser.parse_policy_file(str(candidate_path), candidate_sheet)
        else:
            candidate_policy_data = PaloaltoParser.parse_policy_file(str(candidate_path))
        console.print(f"[green]✓ {len(candidate_policy_data)}개 정책 발견[/green]")
        
        # 대상 정책 목록 로드
        target_policies = []
        if target_paths:
            for target_path in target_paths:
                console.print(f"[cyan]대상 정책 파일 파싱 중: {target_path.name}[/cyan]")
                try:
                    policies = parse_target_file(str(target_path))
                    target_policies.extend(policies)
                    console.print(f"[green]✓ {len(policies)}개 정책 발견[/green]")
                except Exception as e:
                    console.print(f"[yellow]경고: {target_path.name} 파싱 실패 - {e}[/yellow]")
            
            target_policies = list(dict.fromkeys(target_policies))
        
        # 정책 검증
        if len(target_policies) == 0:
            return jsonify({
                'error': '대상 정책이 없습니다. 대상 정책 파일을 업로드하거나 검증을 건너뛰세요.',
                'warning': True
            }), 200
        
        if running_policy_data.empty or candidate_policy_data.empty:
            return jsonify({'error': '정책 데이터가 비어있습니다.'}), 400
        
        console.print(f"[cyan]{len(target_policies)}개 대상 정책 검증 중...[/cyan]")
        validation_results = validate_policy_changes(
            running_policy_data,
            candidate_policy_data,
            target_policies
        )
        
        # 리포트 저장 (날짜+시간으로 파일명 중복 방지)
        report_filename = datetime.now().strftime("%Y-%m-%d_%H%M%S") + "_validation_report.xlsx"
        report_path = upload_dir / report_filename
        validation_results.to_excel(str(report_path), index=False, engine='openpyxl')
        
        # 결과 요약 생성
        summary = get_summary_dict(validation_results)
        
        # 세션에는 파일명·요약만 저장 (쿠키 4KB 제한 방지, 대용량 results 제외)
        session['report_filename'] = report_filename
        session['summary'] = summary
        
        return jsonify({
            'success': True,
            'summary': summary,
            'report_filename': report_filename
        })
        
    except Exception as e:
        console.print(f"[red]오류 발생: {e}[/red]")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'처리 중 오류가 발생했습니다: {str(e)}'}), 500


@app.route('/download')
def download_report():
    """리포트 다운로드"""
    report_filename = session.get('report_filename')
    if not report_filename:
        return jsonify({'error': '리포트 파일을 찾을 수 없습니다.'}), 404
    report_path = get_upload_dir() / report_filename
    if not report_path.exists():
        return jsonify({'error': '리포트 파일을 찾을 수 없습니다.'}), 404
    
    return send_file(
        str(report_path),
        as_attachment=True,
        download_name=report_filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/results')
def show_results():
    """결과 페이지 (요약만 세션에 있음, 상세는 리포트 다운로드로 확인)"""
    summary = session.get('summary')
    report_filename = session.get('report_filename')
    if not summary or not report_filename:
        return render_template('index.html', error='검증 결과가 없습니다.')
    return render_template('index.html', summary_only=summary, report_filename=report_filename)


if __name__ == '__main__':
    # 임시 디렉터리 정리 함수
    def cleanup_temp_files():
        """임시 파일 정리"""
        try:
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            if upload_folder.exists():
                for item in upload_folder.iterdir():
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
        except Exception as e:
            console.print(f"[yellow]임시 파일 정리 오류: {e}[/yellow]")
    
    import atexit
    atexit.register(cleanup_temp_files)
    
    console.print("[bold green]방화벽 정책 검증 웹 애플리케이션 시작[/bold green]")
    console.print("[cyan]로컬 접속: http://127.0.0.1:5000[/cyan]\n")
    
    app.run(host='127.0.0.1', port=5000, debug=True)

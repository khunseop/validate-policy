"""
Flask 웹 애플리케이션: 방화벽 정책 검증

로컬 웹 서버에서 실행하여 파일을 업로드하고 정책을 검증합니다.
방화벽 정책 검증 웹 애플리케이션
"""

from flask import Flask, render_template, request, send_file, jsonify, session
from werkzeug.utils import secure_filename
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
from parse_firewall_policy import (
    parse_policy_file,
    parse_target_file,
    validate_policy_changes
)
from rich.console import Console
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = os.urandom(24)  # 세션 암호화용
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 최대 100MB 파일 업로드
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

console = Console()

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    """파일 확장자 확인"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    """파일 업로드 및 검증 처리"""
    try:
        # 세션 초기화
        session['results'] = None
        session['report_file'] = None
        
        # 파일 확인
        if 'running_file' not in request.files or 'candidate_file' not in request.files:
            return jsonify({'error': 'Running 정책 파일과 Candidate 정책 파일은 필수입니다.'}), 400
        
        running_file = request.files['running_file']
        candidate_file = request.files['candidate_file']
        target_files = request.files.getlist('target_files')
        
        if running_file.filename == '' or candidate_file.filename == '':
            return jsonify({'error': '파일을 선택해주세요.'}), 400
        
        if not (allowed_file(running_file.filename) and allowed_file(candidate_file.filename)):
            return jsonify({'error': 'Excel 파일만 업로드 가능합니다 (.xlsx, .xls)'}), 400
        
        # 파일 저장
        upload_dir = Path(app.config['UPLOAD_FOLDER']) / session.sid
        upload_dir.mkdir(parents=True, exist_ok=True)
        
        running_path = upload_dir / secure_filename(running_file.filename)
        candidate_path = upload_dir / secure_filename(candidate_file.filename)
        
        running_file.save(str(running_path))
        candidate_file.save(str(candidate_path))
        
        # 대상 파일 저장
        target_paths = []
        for target_file in target_files:
            if target_file.filename and allowed_file(target_file.filename):
                target_path = upload_dir / secure_filename(target_file.filename)
                target_file.save(str(target_path))
                target_paths.append(target_path)
        
        # 정책 파일 파싱
        console.print("[cyan]Running 정책 파싱 중...[/cyan]")
        running_policy_data = parse_policy_file(str(running_path))
        console.print(f"[green]✓ {len(running_policy_data)}개 정책 발견[/green]")
        
        console.print("[cyan]Candidate 정책 파싱 중...[/cyan]")
        candidate_policy_data = parse_policy_file(str(candidate_path))
        console.print(f"[green]✓ {len(candidate_policy_data)}개 정책 발견[/green]")
        
        # 대상 정책 목록 로드
        target_policies = []
        if target_paths:
            for target_path in target_paths:
                console.print(f"[cyan]대상 정책 파일 파싱 중: {target_path.name}[/cyan]")
                policies = parse_target_file(str(target_path))
                target_policies.extend(policies)
                console.print(f"[green]✓ {len(policies)}개 정책 발견[/green]")
            
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
        
        # 리포트 저장
        today = datetime.now().strftime("%Y-%m-%d")
        report_filename = f"{today}_validation_report.xlsx"
        report_path = upload_dir / report_filename
        validation_results.to_excel(str(report_path), index=False, engine='openpyxl')
        
        # 결과 요약 생성
        status_counts = validation_results['Status'].value_counts().to_dict()
        target_results = validation_results[validation_results.get('IsTarget', True) == True]
        unexpected_results = validation_results[validation_results.get('IsTarget', True) == False]
        
        summary = {
            'total': len(validation_results),
            'target_total': len(target_results),
            'unexpected_total': len(unexpected_results),
            'deleted': len(target_results[target_results['Status'] == 'DELETED']),
            'disabled': len(target_results[target_results['Status'] == 'DISABLED']),
            'not_disabled': len(target_results[target_results['Status'] == 'NOT_DISABLED']),
            'unexpected_deleted': len(unexpected_results[unexpected_results['Status'] == 'UNEXPECTED_DELETED']),
            'unexpected_disabled': len(unexpected_results[unexpected_results['Status'] == 'UNEXPECTED_DISABLED']),
            'status_counts': {str(k): int(v) for k, v in status_counts.items()}
        }
        
        # 세션에 저장
        session['results'] = validation_results.to_dict('records')
        session['report_file'] = str(report_path)
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
    report_file = session.get('report_file')
    if not report_file or not os.path.exists(report_file):
        return jsonify({'error': '리포트 파일을 찾을 수 없습니다.'}), 404
    
    return send_file(
        report_file,
        as_attachment=True,
        download_name=os.path.basename(report_file),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/results')
def show_results():
    """결과 페이지"""
    results = session.get('results')
    summary = session.get('summary')
    
    if not results or not summary:
        return render_template('index.html', error='검증 결과가 없습니다.')
    
    return render_template('results.html', results=results, summary=summary)


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

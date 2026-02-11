"""
방화벽 정책 검증 자동화 스크립트

이 스크립트는 방화벽 정책 변경이 성공적으로 적용되었는지 자동으로 검증합니다.
- "running" 정책 파일과 "candidate" 정책 파일을 읽습니다
- 대상 정책 이름 목록을 읽어 해당 정책들이 제대로 삭제되었거나 비활성화되었는지 확인합니다
- DRM이 적용된 Excel 파일을 처리하기 위해 xlwings를 사용합니다
"""

import xlwings as xw
import pandas as pd
import os
from pathlib import Path
from typing import List, Optional
from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint

console = Console()


def parse_policy_file(file_path: str) -> pd.DataFrame:
    """
    방화벽 정책 Excel 파일을 파싱하여 'Rulename'과 'Enable' 컬럼을 추출합니다.
    
    성능 최적화: 전체 범위를 한 번에 읽어서 처리합니다.
    빈 셀이 많은 파일에서도 'Rulename'과 'Enable'을 정확히 추출합니다.
    
    Args:
        file_path (str): Excel 파일 경로 (DRM 보호 파일 가능)
    
    Returns:
        pd.DataFrame: 'Rulename'과 'Enable' 컬럼을 가진 DataFrame
                     (중복 제거 및 공백 제거 완료)
    """
    try:
        # xlwings를 사용하여 DRM 보호 파일 열기
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]
            
            # 사용된 범위 가져오기
            if not ws.used_range:
                wb.close()
                return pd.DataFrame(columns=['Rulename', 'Enable'])
            
            # 실제 사용된 마지막 행과 열 가져오기 (제한 없음)
            max_row = ws.used_range.last_cell.row
            max_col = ws.used_range.last_cell.column
            
            # 헤더 행 찾기: 'Rulename'과 'Enable' 컬럼이 있는 행 찾기
            # 첫 50행에서 헤더 찾기 (충분한 범위)
            header_row_idx = None
            rulename_col_idx = None
            enable_col_idx = None
            
            search_rows = min(50, max_row)
            for row_idx in range(1, search_rows + 1):
                for col_idx in range(1, min(max_col + 1, 200)):  # 최대 200열까지 검색
                    cell_value = ws.range((row_idx, col_idx)).value
                    if cell_value:
                        cell_str = str(cell_value).strip().lower()
                        if cell_str == 'rulename' and rulename_col_idx is None:
                            rulename_col_idx = col_idx
                        elif cell_str == 'enable' and enable_col_idx is None:
                            enable_col_idx = col_idx
                
                # 두 컬럼을 모두 찾으면 헤더 행으로 설정
                if rulename_col_idx is not None and enable_col_idx is not None:
                    header_row_idx = row_idx
                    break
            
            if header_row_idx is None or rulename_col_idx is None or enable_col_idx is None:
                wb.close()
                raise ValueError(f"'{file_path}'에서 'Rulename'과 'Enable' 컬럼을 찾을 수 없습니다.")
            
            # 헤더 행 이후부터 마지막 행까지 Rulename과 Enable 컬럼만 읽기
            # 성능 최적화: 필요한 두 컬럼만 직접 읽기
            data_start_row = header_row_idx + 1
            data_end_row = max_row
            
            # 데이터가 있는 경우에만 읽기
            if data_start_row <= data_end_row:
                # Rulename 컬럼 읽기 (헤더 행 다음부터 끝까지)
                rulename_range = ws.range((data_start_row, rulename_col_idx), (data_end_row, rulename_col_idx))
                rulename_values = rulename_range.value
                
                # Enable 컬럼 읽기 (헤더 행 다음부터 끝까지)
                enable_range = ws.range((data_start_row, enable_col_idx), (data_end_row, enable_col_idx))
                enable_values = enable_range.value
            else:
                # 데이터가 없는 경우 빈 리스트
                rulename_values = []
                enable_values = []
            
            wb.close()
        
        # 리스트로 변환 (xlwings 반환값 처리)
        # xlwings는 단일 셀을 읽으면 단일 값, 여러 셀을 읽으면 리스트 또는 2D 배열로 반환
        def normalize_values(values):
            if values is None:
                return []
            elif not isinstance(values, list):
                return [values]
            elif len(values) > 0 and isinstance(values[0], list):
                # 2D 배열인 경우 첫 번째 요소만 추출 (단일 컬럼이므로)
                return [row[0] if row else None for row in values]
            else:
                return values
        
        rulename_values = normalize_values(rulename_values)
        enable_values = normalize_values(enable_values)
        
        # 길이가 다른 경우 짧은 쪽에 맞춤
        min_len = min(len(rulename_values), len(enable_values))
        if len(rulename_values) > min_len:
            rulename_values = rulename_values[:min_len]
        if len(enable_values) > min_len:
            enable_values = enable_values[:min_len]
        
        # DataFrame 생성
        df_filtered = pd.DataFrame({
            'Rulename': rulename_values,
            'Enable': enable_values
        })
        
        # 문자열로 변환하고 공백 제거
        df_filtered['Rulename'] = df_filtered['Rulename'].fillna('').astype(str).str.strip()
        df_filtered['Enable'] = df_filtered['Enable'].fillna('').astype(str).str.strip()
        
        # 빈 행 제거 (두 컬럼이 모두 비어있는 경우)
        df_filtered = df_filtered[
            ~((df_filtered['Rulename'] == '') & (df_filtered['Enable'] == ''))
        ]
        
        # 중복 제거
        df_processed = df_filtered.drop_duplicates().reset_index(drop=True)
        
        return df_processed
    
    except Exception as e:
        console.print(f"[red]파일 파싱 오류 ({file_path}): {e}[/red]")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['Rulename', 'Enable'])


def parse_target_file(file_path: str) -> list:
    """
    대상 정책 파일을 파싱하여 정책 이름 목록을 추출합니다.
    
    - "Rule Name", "Rulename", "Policy Name" 컬럼 모두 지원
    - "작업구분" (Task Type) 컬럼이 있으면 값이 "삭제" (Delete)인 행만 추출
    - Enable 컬럼은 없음
    
    Args:
        file_path (str): 대상 정책 파일 경로 (Excel 파일, DRM 보호 가능)
    
    Returns:
        list: 정책 이름 리스트
    """
    try:
        # xlwings를 사용하여 DRM 보호 파일 열기
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            ws = wb.sheets[0]
            
            # 사용된 범위 가져오기
            if not ws.used_range:
                wb.close()
                return []
            
            # 실제 사용된 마지막 행과 열 가져오기
            max_row = ws.used_range.last_cell.row
            max_col = ws.used_range.last_cell.column
            
            # 헤더 행 찾기: 정책 이름 컬럼과 작업구분 컬럼 찾기
            # 지원하는 컬럼 이름: "Rule Name", "Rulename", "Policy Name"
            # 작업구분 컬럼: "작업구분", "Task Type", "TaskType", "Task"
            header_row_idx = None
            rulename_col_idx = None
            task_type_col_idx = None
            
            search_rows = min(50, max_row)
            for row_idx in range(1, search_rows + 1):
                for col_idx in range(1, min(max_col + 1, 200)):
                    cell_value = ws.range((row_idx, col_idx)).value
                    if cell_value:
                        cell_str = str(cell_value).strip().lower()
                        # 정책 이름 컬럼 찾기
                        if rulename_col_idx is None and cell_str in ['rule name', 'rulename', 'policy name']:
                            rulename_col_idx = col_idx
                        # 작업구분 컬럼 찾기 (한글/영문 모두 지원)
                        if task_type_col_idx is None and cell_str in ['작업구분', 'task type', 'tasktype', 'task']:
                            task_type_col_idx = col_idx
                
                # 정책 이름 컬럼을 찾으면 헤더 행으로 설정
                if rulename_col_idx is not None:
                    header_row_idx = row_idx
                    break
            
            if header_row_idx is None or rulename_col_idx is None:
                wb.close()
                raise ValueError(f"'{file_path}'에서 정책 이름 컬럼('Rule Name', 'Rulename', 또는 'Policy Name')을 찾을 수 없습니다.")
            
            # 헤더 행 이후부터 마지막 행까지 데이터 읽기
            data_start_row = header_row_idx + 1
            data_end_row = max_row
            
            if data_start_row > data_end_row:
                wb.close()
                return []
            
            # 정책 이름 컬럼 읽기
            rulename_range = ws.range((data_start_row, rulename_col_idx), (data_end_row, rulename_col_idx))
            rulename_values = rulename_range.value
            
            # 작업구분 컬럼이 있으면 읽기
            task_type_values = None
            if task_type_col_idx is not None:
                task_type_range = ws.range((data_start_row, task_type_col_idx), (data_end_row, task_type_col_idx))
                task_type_values = task_type_range.value
            
            wb.close()
        
        # 리스트로 변환 (xlwings 반환값 처리)
        def normalize_values(values):
            if values is None:
                return []
            elif not isinstance(values, list):
                return [values]
            elif len(values) > 0 and isinstance(values[0], list):
                return [row[0] if row else None for row in values]
            else:
                return values
        
        rulename_values = normalize_values(rulename_values)
        task_type_values = normalize_values(task_type_values) if task_type_values is not None else None
        
        # 정책 이름 추출
        policies = []
        for idx, rulename in enumerate(rulename_values):
            # 작업구분 컬럼이 있고 값이 "삭제" (Delete)가 아니면 건너뛰기
            if task_type_values is not None and idx < len(task_type_values):
                task_type = str(task_type_values[idx]).strip() if task_type_values[idx] is not None else ''
                task_type_lower = task_type.lower()
                # "삭제" 또는 "delete" 모두 지원
                if task_type_lower not in ['삭제', 'delete']:
                    continue
            
            # 정책 이름이 있으면 추가
            if rulename is not None:
                rulename_str = str(rulename).strip()
                if rulename_str:
                    policies.append(rulename_str)
        
        # 중복 제거
        policies = list(dict.fromkeys(policies))  # 순서 유지하면서 중복 제거
        
        return policies
    
    except Exception as e:
        console.print(f"[red]대상 파일 파싱 오류 ({file_path}): {e}[/red]")
        import traceback
        traceback.print_exc()
        return []


def normalize_enable(value: str) -> str:
    """
    Enable 값을 정규화합니다. Y/N 형식을 처리합니다.
    
    Args:
        value: Enable 값
    
    Returns:
        str: 정규화된 값 ('Y' 또는 'N')
    """
    value_str = str(value).strip().upper()
    if value_str in ['Y', 'YES', 'TRUE', '1', 'ENABLED', 'ENABLE']:
        return 'Y'
    elif value_str in ['N', 'NO', 'FALSE', '0', 'DISABLED', 'DISABLE']:
        return 'N'
    return value_str


def validate_policy_changes(
    running_df: pd.DataFrame,
    candidate_df: pd.DataFrame,
    target_policies: List[str]
) -> pd.DataFrame:
    """
    정책 변경 사항을 검증합니다.
    
    검증 항목:
    1. 대상 정책이 삭제되었는지 확인 (Running에는 있지만 Candidate에는 없음)
    2. 대상 정책이 비활성화되었는지 확인 (Enable 값이 Y → N으로 변경됨)
    3. 대상 외에 삭제되거나 비활성화된 정책 찾기
    4. 덜 삭제/비활성화된 정책 찾기 (대상에는 있지만 실제로는 삭제/비활성화 안됨)
    
    Args:
        running_df (pd.DataFrame): Running 정책 데이터 (Rulename, Enable 컬럼)
        candidate_df (pd.DataFrame): Candidate 정책 데이터 (Rulename, Enable 컬럼)
        target_policies (List[str]): 검증할 대상 정책 이름 리스트
    
    Returns:
        pd.DataFrame: 검증 결과 리포트
                     컬럼: ['Policy', 'Status', 'Running_Enable', 'Candidate_Enable', 'Message']
    """
    results = []
    target_set = set(p.strip() for p in target_policies)
    
    # 1. 대상 정책 검증
    for policy_name in target_policies:
        policy_name = str(policy_name).strip()
        
        # Running 정책에서 찾기
        running_match = running_df[running_df['Rulename'].str.strip() == policy_name]
        # Candidate 정책에서 찾기
        candidate_match = candidate_df[candidate_df['Rulename'].str.strip() == policy_name]
        
        running_enable = None
        candidate_enable = None
        status = ""
        message = ""
        
        if running_match.empty:
            # Running에 없는 경우
            status = "NOT_IN_RUNNING"
            message = "Running 정책에 존재하지 않음"
        elif candidate_match.empty:
            # Running에는 있지만 Candidate에는 없는 경우 (삭제됨)
            running_enable = normalize_enable(running_match.iloc[0]['Enable'])
            status = "DELETED"
            message = "정책이 삭제되었습니다. ✓"
        else:
            # 둘 다 있는 경우 - Enable 상태 확인
            running_enable = normalize_enable(running_match.iloc[0]['Enable'])
            candidate_enable = normalize_enable(candidate_match.iloc[0]['Enable'])
            
            if running_enable == 'Y' and candidate_enable == 'N':
                status = "DISABLED"
                message = "정책이 비활성화되었습니다. ✓"
            elif running_enable == 'N' and candidate_enable == 'Y':
                status = "RE_ENABLED"
                message = "정책이 다시 활성화되었습니다. ⚠"
            elif running_enable == candidate_enable:
                if running_enable == 'Y':
                    status = "NOT_DISABLED"
                    message = "비활성화되지 않았습니다. ⚠"
                else:
                    status = "NO_CHANGE"
                    message = f"변경 없음 (상태: {running_enable})"
            else:
                status = "CHANGED"
                message = f"Enable 상태 변경: {running_enable} -> {candidate_enable}"
        
        results.append({
            'Policy': policy_name,
            'Status': status,
            'Running_Enable': running_enable if running_enable else 'N/A',
            'Candidate_Enable': candidate_enable if candidate_enable else 'N/A',
            'Message': message,
            'IsTarget': True
        })
    
    # 2. 대상 외에 삭제되거나 비활성화된 정책 찾기
    running_policies = set(running_df['Rulename'].str.strip())
    candidate_policies = set(candidate_df['Rulename'].str.strip())
    
    # Running에 있지만 Candidate에 없는 정책 (삭제된 정책)
    deleted_policies = running_policies - candidate_policies - target_set
    
    for policy_name in deleted_policies:
        running_match = running_df[running_df['Rulename'].str.strip() == policy_name]
        if not running_match.empty:
            running_enable = normalize_enable(running_match.iloc[0]['Enable'])
            results.append({
                'Policy': policy_name,
                'Status': 'UNEXPECTED_DELETED',
                'Running_Enable': running_enable,
                'Candidate_Enable': 'N/A',
                'Message': '대상 외 정책이 삭제되었습니다. ⚠',
                'IsTarget': False
            })
    
    # 3. 대상 외에 비활성화된 정책 찾기 (Y → N)
    common_policies = running_policies & candidate_policies - target_set
    
    for policy_name in common_policies:
        running_match = running_df[running_df['Rulename'].str.strip() == policy_name]
        candidate_match = candidate_df[candidate_df['Rulename'].str.strip() == policy_name]
        
        if not running_match.empty and not candidate_match.empty:
            running_enable = normalize_enable(running_match.iloc[0]['Enable'])
            candidate_enable = normalize_enable(candidate_match.iloc[0]['Enable'])
            
            if running_enable == 'Y' and candidate_enable == 'N':
                results.append({
                    'Policy': policy_name,
                    'Status': 'UNEXPECTED_DISABLED',
                    'Running_Enable': running_enable,
                    'Candidate_Enable': candidate_enable,
                    'Message': '대상 외 정책이 비활성화되었습니다. ⚠',
                    'IsTarget': False
                })
    
    return pd.DataFrame(results)


def select_excel_files(current_dir: Path, file_type: str) -> List[str]:
    """
    현재 디렉터리에서 Excel 파일을 선택합니다.
    
    Args:
        current_dir: 현재 디렉터리 경로
        file_type: 파일 타입 설명 ('Running 정책', 'Candidate 정책', '대상 정책')
    
    Returns:
        List[str]: 선택된 파일 경로 리스트
    """
    # Excel 파일 찾기
    excel_files = sorted([f for f in os.listdir(current_dir) 
                         if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')])
    
    if not excel_files:
        console.print(f"[yellow]현재 디렉터리에 Excel 파일이 없습니다.[/yellow]")
        return []
    
    # 파일 선택 테이블 표시
    table = Table(title=f"{file_type} 파일 선택", show_header=True, header_style="bold magenta")
    table.add_column("번호", style="cyan", width=6)
    table.add_column("파일명", style="green")
    
    for idx, filename in enumerate(excel_files, 1):
        table.add_row(str(idx), filename)
    
    console.print(table)
    
    # 여러 파일 선택 (대상 정책의 경우)
    if file_type == "대상 정책":
        console.print(f"\n[bold cyan]{file_type} 파일을 선택하세요 (여러 개 선택 가능, 쉼표로 구분, 예: 1,2,3)[/bold cyan]")
        selection = Prompt.ask("선택", default="")
        
        if not selection.strip():
            return []
        
        try:
            indices = [int(x.strip()) - 1 for x in selection.split(',')]
            selected_files = [excel_files[i] for i in indices if 0 <= i < len(excel_files)]
            return selected_files
        except (ValueError, IndexError):
            console.print("[red]잘못된 선택입니다.[/red]")
            return []
    else:
        # 단일 파일 선택
        console.print(f"\n[bold cyan]{file_type} 파일을 선택하세요[/bold cyan]")
        selection = Prompt.ask("선택 (번호)", default="1")
        
        try:
            idx = int(selection.strip()) - 1
            if 0 <= idx < len(excel_files):
                return [excel_files[idx]]
            else:
                console.print("[red]잘못된 선택입니다.[/red]")
                return []
        except ValueError:
            console.print("[red]잘못된 선택입니다.[/red]")
            return []


def view_report(report_path: Path):
    """
    리포트를 조회합니다.
    
    Args:
        report_path: 리포트 파일 경로
    """
    if not report_path.exists():
        console.print(f"[red]리포트 파일을 찾을 수 없습니다: {report_path}[/red]")
        return
    
    try:
        df = pd.read_excel(report_path)
        
        # 리포트 요약
        console.print("\n" + "="*70)
        console.print("[bold green]검증 결과 요약[/bold green]")
        console.print("="*70)
        
        status_counts = df['Status'].value_counts()
        status_kr = {
            'DELETED': '삭제됨 ✓',
            'DISABLED': '비활성화됨 ✓',
            'NOT_DISABLED': '비활성화 안됨 ⚠',
            'UNEXPECTED_DELETED': '대상 외 삭제됨 ⚠',
            'UNEXPECTED_DISABLED': '대상 외 비활성화됨 ⚠',
            'RE_ENABLED': '재활성화됨 ⚠',
            'NO_CHANGE': '변경 없음',
            'NOT_IN_RUNNING': 'Running에 없음',
            'CHANGED': '변경됨'
        }
        
        summary_table = Table(show_header=True, header_style="bold magenta")
        summary_table.add_column("상태", style="cyan")
        summary_table.add_column("개수", style="green", justify="right")
        
        for status, count in status_counts.items():
            status_str = str(status) if pd.notna(status) else 'UNKNOWN'
            status_name = status_kr.get(status_str, status_str)
            summary_table.add_row(status_name, str(count))
        
        console.print(summary_table)
        
        # 상세 결과 표시
        console.print("\n" + "="*70)
        console.print("[bold green]검증 결과 상세[/bold green]")
        console.print("="*70)
        
        # 대상 정책과 대상 외 정책 분리
        target_results = df[df.get('IsTarget', True) == True]
        unexpected_results = df[df.get('IsTarget', True) == False]
        
        if len(target_results) > 0:
            console.print("\n[bold yellow]대상 정책 검증 결과:[/bold yellow]")
            detail_table = Table(show_header=True, header_style="bold magenta")
            detail_table.add_column("정책명", style="cyan")
            detail_table.add_column("상태", style="green")
            detail_table.add_column("Running", style="yellow")
            detail_table.add_column("Candidate", style="yellow")
            detail_table.add_column("메시지", style="white")
            
            for _, row in target_results.iterrows():
                # 모든 값을 문자열로 변환 (NaN 처리)
                status = str(row['Status']) if pd.notna(row['Status']) else 'UNKNOWN'
                status_name = status_kr.get(status, status)
                policy = str(row['Policy']) if pd.notna(row['Policy']) else 'N/A'
                running_enable = str(row['Running_Enable']) if pd.notna(row['Running_Enable']) else 'N/A'
                candidate_enable = str(row['Candidate_Enable']) if pd.notna(row['Candidate_Enable']) else 'N/A'
                message = str(row['Message']) if pd.notna(row['Message']) else 'N/A'
                
                detail_table.add_row(
                    policy,
                    status_name,
                    running_enable,
                    candidate_enable,
                    message
                )
            
            console.print(detail_table)
        
        if len(unexpected_results) > 0:
            console.print("\n[bold red]대상 외 정책 검증 결과:[/bold red]")
            unexpected_table = Table(show_header=True, header_style="bold magenta")
            unexpected_table.add_column("정책명", style="cyan")
            unexpected_table.add_column("상태", style="red")
            unexpected_table.add_column("Running", style="yellow")
            unexpected_table.add_column("Candidate", style="yellow")
            unexpected_table.add_column("메시지", style="white")
            
            for _, row in unexpected_results.iterrows():
                # 모든 값을 문자열로 변환 (NaN 처리)
                status = str(row['Status']) if pd.notna(row['Status']) else 'UNKNOWN'
                status_name = status_kr.get(status, status)
                policy = str(row['Policy']) if pd.notna(row['Policy']) else 'N/A'
                running_enable = str(row['Running_Enable']) if pd.notna(row['Running_Enable']) else 'N/A'
                candidate_enable = str(row['Candidate_Enable']) if pd.notna(row['Candidate_Enable']) else 'N/A'
                message = str(row['Message']) if pd.notna(row['Message']) else 'N/A'
                
                unexpected_table.add_row(
                    policy,
                    status_name,
                    running_enable,
                    candidate_enable,
                    message
                )
            
            console.print(unexpected_table)
        
    except Exception as e:
        console.print(f"[red]리포트 조회 오류: {e}[/red]")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    current_dir = Path.cwd()
    
    console.print(Panel.fit(
        "[bold cyan]방화벽 정책 검증 자동화 스크립트[/bold cyan]\n"
        "DRM 보호 파일을 처리하고 정책 변경 사항을 검증합니다.",
        border_style="cyan"
    ))
    
    # 1. Running 정책 파일 선택
    console.print("\n[bold]1단계: Running 정책 파일 선택[/bold]")
    running_files = select_excel_files(current_dir, "Running 정책")
    if not running_files:
        console.print("[red]Running 정책 파일이 선택되지 않았습니다.[/red]")
        exit(1)
    
    running_file = running_files[0]
    console.print(f"[green]선택됨: {running_file}[/green]")
    
    # 2. Candidate 정책 파일 선택
    console.print("\n[bold]2단계: Candidate 정책 파일 선택[/bold]")
    candidate_files = select_excel_files(current_dir, "Candidate 정책")
    if not candidate_files:
        console.print("[red]Candidate 정책 파일이 선택되지 않았습니다.[/red]")
        exit(1)
    
    candidate_file = candidate_files[0]
    console.print(f"[green]선택됨: {candidate_file}[/green]")
    
    # 3. 대상 정책 파일 선택 (여러 개 가능)
    console.print("\n[bold]3단계: 대상 정책 파일 선택 (여러 개 선택 가능)[/bold]")
    target_files = select_excel_files(current_dir, "대상 정책")
    if not target_files:
        console.print("[yellow]대상 정책 파일이 선택되지 않았습니다. 계속 진행할까요?[/yellow]")
        if not Confirm.ask("계속 진행"):
            exit(0)
        target_files = []
    
    # 4. 정책 파일 파싱
    console.print("\n" + "="*70)
    console.print("[bold]4단계: 정책 파일 파싱[/bold]")
    console.print("="*70)
    
    console.print(f"\n[cyan]Running 정책 파싱 중...[/cyan]")
    running_policy_data = parse_policy_file(running_file)
    console.print(f"[green]✓ 총 {len(running_policy_data)}개 정책 발견[/green]")
    
    console.print(f"\n[cyan]Candidate 정책 파싱 중...[/cyan]")
    candidate_policy_data = parse_policy_file(candidate_file)
    console.print(f"[green]✓ 총 {len(candidate_policy_data)}개 정책 발견[/green]")
    
    # 5. 대상 정책 목록 로드
    target_policies = []
    if target_files:
        console.print("\n" + "="*70)
        console.print("[bold]5단계: 대상 정책 목록 로드[/bold]")
        console.print("="*70)
        
        for target_file in target_files:
            console.print(f"\n[cyan]대상 정책 파일 파싱 중: {target_file}[/cyan]")
            policies = parse_target_file(target_file)
            target_policies.extend(policies)
            console.print(f"[green]✓ {len(policies)}개 정책 발견[/green]")
        
        # 중복 제거
        target_policies = list(dict.fromkeys(target_policies))
        console.print(f"\n[green]✓ 총 {len(target_policies)}개 고유 대상 정책[/green]")
    
    # 6. 정책 검증
    if len(target_policies) == 0:
        console.print("\n[yellow]⚠ 대상 정책이 없어 검증을 건너뜁니다.[/yellow]")
    elif running_policy_data.empty or candidate_policy_data.empty:
        console.print("\n[yellow]⚠ 정책 데이터가 비어있어 검증을 수행할 수 없습니다.[/yellow]")
    else:
        console.print("\n" + "="*70)
        console.print("[bold]6단계: 정책 검증[/bold]")
        console.print("="*70)
        
        console.print(f"\n[cyan]{len(target_policies)}개 대상 정책 검증 중...[/cyan]")
        validation_results = validate_policy_changes(
            running_policy_data,
            candidate_policy_data,
            target_policies
        )
        
        # 7. 검증 결과 리포트 저장
        console.print("\n" + "="*70)
        console.print("[bold]7단계: 검증 결과 리포트 저장[/bold]")
        console.print("="*70)
        
        validation_report_file = current_dir / "validation_report.xlsx"
        validation_results.to_excel(validation_report_file, index=False)
        console.print(f"\n[green]✓ 검증 결과가 {validation_report_file}에 저장되었습니다.[/green]")
        console.print(f"[green]✓ 총 {len(validation_results)}개 정책 검증 완료[/green]")
        
        # 8. 리포트 조회
        console.print("\n" + "="*70)
        console.print("[bold]8단계: 리포트 조회[/bold]")
        console.print("="*70)
        
        view_report(validation_report_file)
        
        # 리포트 다시 보기 옵션
        console.print("\n")
        if Confirm.ask("[bold cyan]리포트를 다시 조회하시겠습니까?[/bold cyan]"):
            view_report(validation_report_file)
    
    console.print("\n" + "="*70)
    console.print("[bold green]작업 완료![/bold green]")
    console.print("="*70 + "\n")

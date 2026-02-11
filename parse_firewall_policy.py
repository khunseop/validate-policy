"""
방화벽 정책 검증 자동화 스크립트

이 스크립트는 방화벽 정책 변경이 성공적으로 적용되었는지 자동으로 검증합니다.
- "running" 정책 파일과 "candidate" 정책 파일을 읽습니다
- 대상 정책 이름 목록을 읽어 해당 정책들이 제대로 삭제되었거나 비활성화되었는지 확인합니다
- DRM이 적용된 Excel 파일을 처리하기 위해 xlwings를 사용합니다
"""

import xlwings as xw
import pandas as pd


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
        print(f"파일 파싱 오류 ({file_path}): {e}")
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
        print(f"대상 파일 파싱 오류 ({file_path}): {e}")
        import traceback
        traceback.print_exc()
        return []


def load_target_policies(file_path: str) -> list:
    """
    대상 정책 이름 목록을 파일에서 읽어옵니다.
    
    Args:
        file_path (str): 정책 이름 목록이 있는 파일 경로 (텍스트 파일 또는 Excel)
    
    Returns:
        list: 정책 이름 리스트
    """
    try:
        # 텍스트 파일인 경우
        if file_path.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as f:
                policies = [line.strip() for line in f if line.strip()]
            return policies
        
        # Excel 파일인 경우 (첫 번째 컬럼 읽기)
        elif file_path.endswith(('.xlsx', '.xls')):
            with xw.App(visible=False) as app:
                wb = app.books.open(file_path)
                ws = wb.sheets[0]
                
                if not ws.used_range:
                    wb.close()
                    return []
                
                # 첫 번째 컬럼을 한 번에 읽기 (성능 최적화)
                # 행 제한 없이 끝까지 읽기
                max_row = ws.used_range.last_cell.row
                first_col_range = ws.range((1, 1), (max_row, 1))
                values = first_col_range.value
                
                wb.close()
            
            # 리스트로 변환하고 빈 값 제거
            if isinstance(values, list):
                policies = [str(v).strip() for v in values if v]
            else:
                policies = [str(values).strip()] if values else []
            
            return [p for p in policies if p]
        
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {file_path}")
    
    except Exception as e:
        print(f"대상 정책 목록 로드 오류 ({file_path}): {e}")
        return []


def validate_policy_changes(
    running_df: pd.DataFrame,
    candidate_df: pd.DataFrame,
    target_policies: list
) -> pd.DataFrame:
    """
    Running 정책과 Candidate 정책을 비교하여 대상 정책들의 변경 사항을 검증합니다.
    
    검증 항목:
    1. 정책이 삭제되었는지 확인 (Running에는 있지만 Candidate에는 없음)
    2. 정책이 비활성화되었는지 확인 (Enable 값이 False/Disabled로 변경됨)
    
    Args:
        running_df (pd.DataFrame): Running 정책 데이터 (Rulename, Enable 컬럼)
        candidate_df (pd.DataFrame): Candidate 정책 데이터 (Rulename, Enable 컬럼)
        target_policies (list): 검증할 대상 정책 이름 리스트
    
    Returns:
        pd.DataFrame: 검증 결과 리포트
                     컬럼: ['Policy', 'Status', 'Running_Enable', 'Candidate_Enable', 'Message']
    """
    results = []
    
    # Enable 값 정규화 함수 (대소문자 무시)
    def normalize_enable(value: str) -> str:
        value_lower = str(value).strip().lower()
        if value_lower in ['true', 'yes', '1', 'enabled', 'enable']:
            return 'Enabled'
        elif value_lower in ['false', 'no', '0', 'disabled', 'disable']:
            return 'Disabled'
        return str(value).strip()
    
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
            message = "정책이 삭제되었습니다."
        else:
            # 둘 다 있는 경우 - Enable 상태 확인
            running_enable = normalize_enable(running_match.iloc[0]['Enable'])
            candidate_enable = normalize_enable(candidate_match.iloc[0]['Enable'])
            
            if running_enable == 'Enabled' and candidate_enable == 'Disabled':
                status = "DISABLED"
                message = "정책이 비활성화되었습니다."
            elif running_enable == 'Disabled' and candidate_enable == 'Enabled':
                status = "RE_ENABLED"
                message = "정책이 다시 활성화되었습니다."
            elif running_enable == candidate_enable:
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
            'Message': message
        })
    
    return pd.DataFrame(results)


if __name__ == "__main__":
    import os
    import sys
    
    print("="*70)
    print("방화벽 정책 검증 자동화 스크립트")
    print("="*70 + "\n")
    
    # 파일 경로 설정
    running_policy_file = "running_policy.xlsx"
    candidate_policy_file = "candidate_policy.xlsx"
    target_policy_file = "target_policies.xlsx"  # 대상 정책 파일 (작업구분이 "삭제"인 행)
    
    # 출력 파일 경로
    running_output_file = "running_policy_processed.xlsx"
    candidate_output_file = "candidate_policy_processed.xlsx"
    validation_report_file = "validation_report.xlsx"
    
    # ============================================================
    # 1. 정책 파일 파싱
    # ============================================================
    print("="*70)
    print("1단계: 정책 파일 파싱")
    print("="*70)
    
    # Running 정책 파싱
    print(f"\n[1-1] Running 정책 파싱: {running_policy_file}")
    if not os.path.exists(running_policy_file):
        print(f"오류: {running_policy_file} 파일을 찾을 수 없습니다.")
        sys.exit(1)
    
    running_policy_data = parse_policy_file(running_policy_file)
    print(f"  ✓ 총 {len(running_policy_data)}개 정책 발견")
    if not running_policy_data.empty:
        running_policy_data.to_excel(running_output_file, index=False)
        print(f"  ✓ 처리된 데이터가 {running_output_file}에 저장되었습니다.")
    else:
        print("  ⚠ 경고: Running 정책 데이터가 비어있습니다.")
    
    # Candidate 정책 파싱
    print(f"\n[1-2] Candidate 정책 파싱: {candidate_policy_file}")
    if not os.path.exists(candidate_policy_file):
        print(f"오류: {candidate_policy_file} 파일을 찾을 수 없습니다.")
        sys.exit(1)
    
    candidate_policy_data = parse_policy_file(candidate_policy_file)
    print(f"  ✓ 총 {len(candidate_policy_data)}개 정책 발견")
    if not candidate_policy_data.empty:
        candidate_policy_data.to_excel(candidate_output_file, index=False)
        print(f"  ✓ 처리된 데이터가 {candidate_output_file}에 저장되었습니다.")
    else:
        print("  ⚠ 경고: Candidate 정책 데이터가 비어있습니다.")
    
    # ============================================================
    # 2. 대상 정책 목록 로드
    # ============================================================
    print("\n" + "="*70)
    print("2단계: 대상 정책 목록 로드")
    print("="*70)
    
    print(f"\n[2-1] 대상 정책 파일 파싱: {target_policy_file}")
    if not os.path.exists(target_policy_file):
        print(f"  ⚠ 경고: {target_policy_file} 파일을 찾을 수 없습니다.")
        print("  → 대상 정책 파일이 없으면 검증을 건너뜁니다.")
        target_policies = []
    else:
        target_policies = parse_target_file(target_policy_file)
        print(f"  ✓ 총 {len(target_policies)}개 대상 정책 발견")
        if len(target_policies) > 0:
            print(f"  → 대상 정책 목록 (처음 10개): {', '.join(target_policies[:10])}")
            if len(target_policies) > 10:
                print(f"  → ... 외 {len(target_policies) - 10}개")
    
    # ============================================================
    # 3. 정책 검증
    # ============================================================
    print("\n" + "="*70)
    print("3단계: 정책 검증")
    print("="*70)
    
    if len(target_policies) == 0:
        print("\n  ⚠ 대상 정책이 없어 검증을 건너뜁니다.")
        print("  → 대상 정책 파일을 준비하고 다시 실행하세요.")
    elif running_policy_data.empty or candidate_policy_data.empty:
        print("\n  ⚠ 정책 데이터가 비어있어 검증을 수행할 수 없습니다.")
    else:
        print(f"\n[3-1] {len(target_policies)}개 대상 정책 검증 중...")
        validation_results = validate_policy_changes(
            running_policy_data,
            candidate_policy_data,
            target_policies
        )
        
        # 검증 결과 요약
        print("\n[3-2] 검증 결과 요약:")
        status_counts = validation_results['Status'].value_counts()
        for status, count in status_counts.items():
            status_kr = {
                'DELETED': '삭제됨',
                'DISABLED': '비활성화됨',
                'RE_ENABLED': '재활성화됨',
                'NO_CHANGE': '변경 없음',
                'NOT_IN_RUNNING': 'Running에 없음',
                'CHANGED': '변경됨'
            }.get(status, status)
            print(f"  - {status_kr}: {count}개")
        
        # 검증 결과 상세 출력 (처음 20개)
        print("\n[3-3] 검증 결과 상세 (처음 20개):")
        print(validation_results.head(20).to_string(index=False))
        if len(validation_results) > 20:
            print(f"\n  ... 외 {len(validation_results) - 20}개 결과")
        
        # ============================================================
        # 4. 검증 결과 리포트 저장
        # ============================================================
        print("\n" + "="*70)
        print("4단계: 검증 결과 리포트 저장")
        print("="*70)
        
        validation_results.to_excel(validation_report_file, index=False)
        print(f"\n  ✓ 검증 결과가 {validation_report_file}에 저장되었습니다.")
        print(f"  → 총 {len(validation_results)}개 정책 검증 완료")
        
        # 삭제/비활성화 확인된 정책 요약
        deleted_or_disabled = validation_results[
            validation_results['Status'].isin(['DELETED', 'DISABLED'])
        ]
        if len(deleted_or_disabled) > 0:
            print(f"\n  ✓ 삭제 또는 비활성화 확인된 정책: {len(deleted_or_disabled)}개")
        else:
            print("\n  ⚠ 삭제 또는 비활성화 확인된 정책이 없습니다.")
    
    # ============================================================
    # 완료
    # ============================================================
    print("\n" + "="*70)
    print("작업 완료!")
    print("="*70)
    print(f"\n생성된 파일:")
    if os.path.exists(running_output_file):
        print(f"  - {running_output_file}")
    if os.path.exists(candidate_output_file):
        print(f"  - {candidate_output_file}")
    if os.path.exists(validation_report_file):
        print(f"  - {validation_report_file}")
    print()

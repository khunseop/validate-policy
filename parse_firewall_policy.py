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
            
            # 전체 데이터를 한 번에 읽기 (성능 최적화)
            # 최대 1000행, 100열까지 읽기
            max_row = min(ws.used_range.last_cell.row, 1000)
            max_col = min(ws.used_range.last_cell.column, 100)
            
            # 전체 범위를 한 번에 읽어서 pandas DataFrame으로 변환
            data_range = ws.range((1, 1), (max_row, max_col))
            df_raw = data_range.options(pd.DataFrame, header=False, index=False).value
            
            wb.close()
        
        # 헤더 행 찾기: 'Rulename'과 'Enable' 컬럼이 있는 행 찾기
        header_row_idx = None
        rulename_col_idx = None
        enable_col_idx = None
        
        # 첫 20행에서 헤더 찾기
        for idx in range(min(20, len(df_raw))):
            row = df_raw.iloc[idx]
            # 각 셀을 확인하여 'rulename'과 'enable' 찾기
            for col_idx, cell_value in enumerate(row):
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip().lower()
                    if cell_str == 'rulename' and rulename_col_idx is None:
                        rulename_col_idx = col_idx
                    elif cell_str == 'enable' and enable_col_idx is None:
                        enable_col_idx = col_idx
            
            # 두 컬럼을 모두 찾으면 헤더 행으로 설정
            if rulename_col_idx is not None and enable_col_idx is not None:
                header_row_idx = idx
                break
        
        if header_row_idx is None or rulename_col_idx is None or enable_col_idx is None:
            raise ValueError(f"'{file_path}'에서 'Rulename'과 'Enable' 컬럼을 찾을 수 없습니다.")
        
        # 헤더 행 이후의 데이터 추출
        df_data = df_raw.iloc[header_row_idx + 1:].copy()
        
        # 필요한 컬럼만 선택
        df_filtered = pd.DataFrame({
            'Rulename': df_data.iloc[:, rulename_col_idx],
            'Enable': df_data.iloc[:, enable_col_idx]
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
                max_row = min(ws.used_range.last_cell.row, 1000)
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
    # 파일 경로 설정
    running_policy_file = "running_policy.xlsx"
    candidate_policy_file = "candidate_policy.xlsx"
    
    running_output_file = "running_policy_processed.xlsx"
    candidate_output_file = "candidate_policy_processed.xlsx"
    
    # Running 정책 파싱
    print(f"--- Running 정책 파싱: {running_policy_file} ---")
    running_policy_data = parse_policy_file(running_policy_file)
    print(f"총 {len(running_policy_data)}개 정책 발견")
    print(running_policy_data.head(10))
    if not running_policy_data.empty:
        running_policy_data.to_excel(running_output_file, index=False)
        print(f"처리된 running 정책이 {running_output_file}에 저장되었습니다.")
    print("\n" + "="*50 + "\n")
    
    # Candidate 정책 파싱
    print(f"--- Candidate 정책 파싱: {candidate_policy_file} ---")
    candidate_policy_data = parse_policy_file(candidate_policy_file)
    print(f"총 {len(candidate_policy_data)}개 정책 발견")
    print(candidate_policy_data.head(10))
    if not candidate_policy_data.empty:
        candidate_policy_data.to_excel(candidate_output_file, index=False)
        print(f"처리된 candidate 정책이 {candidate_output_file}에 저장되었습니다.")
    print("\n" + "="*50 + "\n")

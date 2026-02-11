"""
정책 파일 파싱 모듈

Excel 파일에서 정책 정보를 추출합니다.
DRM 보호 파일 처리를 위해 xlwings를 사용합니다.
"""

import xlwings as xw
import pandas as pd
from typing import List


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
        raise ValueError(f"파일 파싱 오류 ({file_path}): {e}")


def parse_target_file(file_path: str) -> List[str]:
    """
    대상 정책 파일을 파싱하여 정책 이름 목록을 추출합니다.
    
    - "Rule Name", "Rulename", "Policy Name" 컬럼 모두 지원
    - "작업구분" (Task Type) 컬럼이 있으면 값이 "삭제" (Delete)인 행만 추출 (없으면 모든 행 추출)
    - "제외사유" 컬럼이 있으면 빈 데이터인 행만 추출
    - Enable 컬럼은 없음
    
    Args:
        file_path (str): 대상 정책 파일 경로 (Excel 파일, DRM 보호 가능)
    
    Returns:
        List[str]: 정책 이름 리스트
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
            
            # 헤더 행 찾기: 정책 이름 컬럼, 작업구분 컬럼, 제외사유 컬럼 찾기
            # 지원하는 컬럼 이름: "Rule Name", "Rulename", "Policy Name"
            # 작업구분 컬럼: "작업구분", "Task Type", "TaskType", "Task"
            # 제외사유 컬럼: "제외사유", "Exclusion Reason", "ExclusionReason", "Reason"
            header_row_idx = None
            rulename_col_idx = None
            task_type_col_idx = None
            exclusion_reason_col_idx = None
            
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
                        # 제외사유 컬럼 찾기 (한글/영문 모두 지원)
                        if exclusion_reason_col_idx is None and cell_str in ['제외사유', 'exclusion reason', 'exclusionreason', 'reason', 'exclusion']:
                            exclusion_reason_col_idx = col_idx
                
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
            
            # 제외사유 컬럼이 있으면 읽기
            exclusion_reason_values = None
            if exclusion_reason_col_idx is not None:
                exclusion_reason_range = ws.range((data_start_row, exclusion_reason_col_idx), (data_end_row, exclusion_reason_col_idx))
                exclusion_reason_values = exclusion_reason_range.value
            
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
        exclusion_reason_values = normalize_values(exclusion_reason_values) if exclusion_reason_values is not None else None
        
        # 정책 이름 추출
        policies = []
        for idx, rulename in enumerate(rulename_values):
            # 작업구분 컬럼이 있고 값이 "삭제" (Delete)가 아니면 건너뛰기
            # 작업구분 컬럼이 없으면 모든 행을 추출
            if task_type_values is not None and idx < len(task_type_values):
                task_type = str(task_type_values[idx]).strip() if task_type_values[idx] is not None else ''
                task_type_lower = task_type.lower()
                # "삭제" 또는 "delete" 모두 지원
                if task_type_lower not in ['삭제', 'delete']:
                    continue
            
            # 제외사유 컬럼이 있으면 빈 데이터인 행만 추출
            if exclusion_reason_values is not None and idx < len(exclusion_reason_values):
                exclusion_reason = exclusion_reason_values[idx]
                # None이 아니고 빈 문자열이 아니면 건너뛰기
                if exclusion_reason is not None:
                    exclusion_reason_str = str(exclusion_reason).strip()
                    if exclusion_reason_str:  # 빈 문자열이 아니면 제외
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
        raise ValueError(f"대상 파일 파싱 오류 ({file_path}): {e}")

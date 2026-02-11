"""
벤더별 파서 모듈

벤더별로 다른 정책 파일 포맷을 처리합니다.
"""

import xlwings as xw
import pandas as pd
from typing import List, Optional, Tuple
import re


class PaloaltoParser:
    """Paloalto 정책 파일 파서"""
    
    @staticmethod
    def parse_policy_file(file_path: str) -> pd.DataFrame:
        """
        Paloalto 방화벽 정책 Excel 파일을 파싱합니다.
        
        Args:
            file_path (str): Excel 파일 경로
        
        Returns:
            pd.DataFrame: 'Rulename'과 'Enable' 컬럼을 가진 DataFrame
        """
        try:
            with xw.App(visible=False) as app:
                wb = app.books.open(file_path)
                ws = wb.sheets[0]
                
                if not ws.used_range:
                    wb.close()
                    return pd.DataFrame(columns=['Rulename', 'Enable'])
                
                max_row = ws.used_range.last_cell.row
                max_col = ws.used_range.last_cell.column
                
                # 헤더 행 찾기
                header_row_idx = None
                rulename_col_idx = None
                enable_col_idx = None
                
                search_rows = min(50, max_row)
                for row_idx in range(1, search_rows + 1):
                    for col_idx in range(1, min(max_col + 1, 200)):
                        cell_value = ws.range((row_idx, col_idx)).value
                        if cell_value:
                            cell_str = str(cell_value).strip().lower()
                            if cell_str == 'rulename' and rulename_col_idx is None:
                                rulename_col_idx = col_idx
                            elif cell_str == 'enable' and enable_col_idx is None:
                                enable_col_idx = col_idx
                    
                    if rulename_col_idx is not None and enable_col_idx is not None:
                        header_row_idx = row_idx
                        break
                
                if header_row_idx is None or rulename_col_idx is None or enable_col_idx is None:
                    wb.close()
                    raise ValueError(f"'{file_path}'에서 'Rulename'과 'Enable' 컬럼을 찾을 수 없습니다.")
                
                # 데이터 읽기
                data_start_row = header_row_idx + 1
                data_end_row = max_row
                
                if data_start_row <= data_end_row:
                    rulename_range = ws.range((data_start_row, rulename_col_idx), (data_end_row, rulename_col_idx))
                    rulename_values = rulename_range.value
                    
                    enable_range = ws.range((data_start_row, enable_col_idx), (data_end_row, enable_col_idx))
                    enable_values = enable_range.value
                else:
                    rulename_values = []
                    enable_values = []
                
                wb.close()
            
            # 값 정규화
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
            enable_values = normalize_values(enable_values)
            
            min_len = min(len(rulename_values), len(enable_values))
            if len(rulename_values) > min_len:
                rulename_values = rulename_values[:min_len]
            if len(enable_values) > min_len:
                enable_values = enable_values[:min_len]
            
            # DataFrame 생성
            df = pd.DataFrame({
                'Rulename': rulename_values,
                'Enable': enable_values
            })
            
            df['Rulename'] = df['Rulename'].fillna('').astype(str).str.strip()
            df['Enable'] = df['Enable'].fillna('').astype(str).str.strip()
            
            df = df[~((df['Rulename'] == '') & (df['Enable'] == ''))]
            df = df.drop_duplicates().reset_index(drop=True)
            
            return df
        
        except Exception as e:
            raise ValueError(f"파일 파싱 오류 ({file_path}): {e}")


class SECUIParser:
    """SECUI 정책 파일 파서"""
    
    @staticmethod
    def get_sheets(file_path: str) -> List[str]:
        """
        파일의 모든 시트 이름을 반환합니다.
        
        Args:
            file_path (str): Excel 파일 경로
        
        Returns:
            List[str]: 시트 이름 리스트
        """
        try:
            with xw.App(visible=False) as app:
                wb = app.books.open(file_path)
                sheet_names = [sheet.name for sheet in wb.sheets]
                wb.close()
            return sheet_names
        except Exception as e:
            raise ValueError(f"시트 목록 가져오기 오류 ({file_path}): {e}")
    
    @staticmethod
    def parse_policy_file(file_path: str, sheet_name: str) -> pd.DataFrame:
        """
        SECUI 방화벽 정책 Excel 파일을 파싱합니다.
        
        Args:
            file_path (str): Excel 파일 경로
            sheet_name (str): 시트 이름
        
        Returns:
            pd.DataFrame: 'Rulename' (ID)과 'Enable' 컬럼을 가진 DataFrame
        """
        try:
            with xw.App(visible=False) as app:
                wb = app.books.open(file_path)
                
                if sheet_name not in [s.name for s in wb.sheets]:
                    wb.close()
                    raise ValueError(f"시트 '{sheet_name}'를 찾을 수 없습니다.")
                
                ws = wb.sheets[sheet_name]
                
                if not ws.used_range:
                    wb.close()
                    return pd.DataFrame(columns=['Rulename', 'Enable'])
                
                max_row = ws.used_range.last_cell.row
                max_col = ws.used_range.last_cell.column
                
                # SECUI 포맷: 1-2행 제거, 3-8행에 컬럼명 (병합 셀), 9행부터 데이터
                # 헤더 행 찾기 (3-8행에서 컬럼명 찾기)
                id_col_idx = None
                enable_col_idx = None
                
                # 3-8행에서 컬럼명 찾기 (병합 셀 고려)
                for row_idx in range(3, min(9, max_row + 1)):
                    for col_idx in range(1, min(max_col + 1, 200)):
                        cell_value = ws.range((row_idx, col_idx)).value
                        if cell_value:
                            cell_str = str(cell_value).strip().lower()
                            # ID 컬럼 찾기
                            if id_col_idx is None and cell_str == 'id':
                                id_col_idx = col_idx
                            # Enable 컬럼 찾기
                            if enable_col_idx is None and cell_str == 'enable':
                                enable_col_idx = col_idx
                    
                    # 두 컬럼을 모두 찾으면 종료
                    if id_col_idx is not None and enable_col_idx is not None:
                        break
                
                # ID 컬럼이 없으면 데이터 행에서 숫자만 있는 컬럼 찾기
                if id_col_idx is None:
                    id_col_idx = SECUIParser._find_id_column(ws, 8, max_row, max_col)
                
                if enable_col_idx is None:
                    wb.close()
                    raise ValueError(f"'{file_path}' 시트 '{sheet_name}'에서 'Enable' 컬럼을 찾을 수 없습니다.")
                
                if id_col_idx is None:
                    wb.close()
                    raise ValueError(f"'{file_path}' 시트 '{sheet_name}'에서 ID 컬럼을 찾을 수 없습니다.")
                
                # 데이터 읽기 (9행부터 시작)
                data_start_row = 9
                data_end_row = max_row
                
                if data_start_row <= data_end_row:
                    # ID 컬럼 읽기 (병합 셀 처리)
                    id_values = []
                    enable_values = []
                    
                    # 병합 셀 처리를 위해 마지막 ID 값 저장
                    last_id_value = None
                    
                    for row_idx in range(data_start_row, data_end_row + 1):
                        # ID 값 읽기 (병합 셀의 경우 상위 셀 값 사용)
                        id_cell = ws.range((row_idx, id_col_idx))
                        id_value = id_cell.value
                        
                        # 병합 셀 처리: 값이 None이면 위쪽 셀 값 사용
                        if id_value is None:
                            # 위쪽 행들을 확인하여 값 찾기 (최대 20행까지 확인, 8행까지는 헤더이므로 9행부터)
                            for check_row in range(row_idx - 1, max(data_start_row - 1, row_idx - 20), -1):
                                check_cell = ws.range((check_row, id_col_idx))
                                check_value = check_cell.value
                                if check_value is not None:
                                    check_str = str(check_value).strip()
                                    # 숫자로만 이루어진 값만 사용
                                    if re.match(r'^\d+$', check_str):
                                        id_value = check_value
                                        break
                            
                            # 위에서도 못 찾으면 마지막 ID 값 사용
                            if id_value is None and last_id_value is not None:
                                id_value = last_id_value
                        
                        # Enable 값 읽기
                        enable_cell = ws.range((row_idx, enable_col_idx))
                        enable_value = enable_cell.value
                        
                        # 병합 셀 처리: Enable 값이 None이면 위쪽 셀 값 확인
                        if enable_value is None:
                            for check_row in range(row_idx - 1, max(data_start_row - 1, row_idx - 20), -1):
                                check_cell = ws.range((check_row, enable_col_idx))
                                check_value = check_cell.value
                                if check_value is not None:
                                    enable_value = check_value
                                    break
                        
                        # ID가 숫자로만 이루어져 있는지 확인
                        if id_value is not None:
                            id_str = str(id_value).strip()
                            # 숫자로만 이루어져 있는지 확인
                            if re.match(r'^\d+$', id_str):
                                id_values.append(id_str)
                                enable_values.append(enable_value if enable_value is not None else '')
                                last_id_value = id_str  # 마지막 ID 값 저장
                            else:
                                # 숫자가 아니면 건너뛰기 (헤더나 다른 데이터)
                                continue
                        else:
                            # ID가 없으면 건너뛰기
                            continue
                else:
                    id_values = []
                    enable_values = []
                
                wb.close()
            
            # DataFrame 생성
            df = pd.DataFrame({
                'Rulename': id_values,  # SECUI는 ID를 Rulename으로 사용
                'Enable': enable_values
            })
            
            df['Rulename'] = df['Rulename'].fillna('').astype(str).str.strip()
            df['Enable'] = df['Enable'].fillna('').astype(str).str.strip()
            
            df = df[~((df['Rulename'] == '') & (df['Enable'] == ''))]
            df = df.drop_duplicates().reset_index(drop=True)
            
            return df
        
        except Exception as e:
            raise ValueError(f"파일 파싱 오류 ({file_path}, 시트: {sheet_name}): {e}")
    
    @staticmethod
    def _find_id_column(ws, start_row: int, max_row: int, max_col: int) -> Optional[int]:
        """
        데이터 행에서 숫자로만 이루어진 값이 있는 컬럼을 찾습니다.
        
        Args:
            ws: 워크시트 객체
            start_row: 시작 행 인덱스 (데이터 시작 행)
            max_row: 최대 행
            max_col: 최대 열
        
        Returns:
            Optional[int]: ID 컬럼 인덱스 (없으면 None)
        """
        # 시작 행부터 20행 정도를 확인하여 숫자만 있는 컬럼 찾기
        check_rows = min(20, max_row - start_row + 1)
        
        for col_idx in range(1, min(max_col + 1, 200)):
            numeric_count = 0
            total_count = 0
            
            for row_offset in range(check_rows):
                row_idx = start_row + row_offset
                if row_idx > max_row:
                    break
                    
                cell_value = ws.range((row_idx, col_idx)).value
                
                # 병합 셀 처리: 값이 None이면 위쪽 셀 값 확인
                if cell_value is None:
                    # 위쪽 행들을 확인하여 값 찾기 (최대 10행까지)
                    for check_row in range(row_idx - 1, max(start_row - 1, row_idx - 10), -1):
                        check_cell = ws.range((check_row, col_idx))
                        check_value = check_cell.value
                        if check_value is not None:
                            cell_value = check_value
                            break
                
                if cell_value is not None:
                    total_count += 1
                    cell_str = str(cell_value).strip()
                    # 숫자로만 이루어져 있는지 확인
                    if re.match(r'^\d+$', cell_str):
                        numeric_count += 1
            
            # 충분한 데이터가 있고 대부분의 값이 숫자면 ID 컬럼으로 판단
            if total_count >= 5 and numeric_count >= total_count * 0.7:  # 최소 5개, 70% 이상이 숫자면
                return col_idx
        
        return None

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
        병합된 셀/대용량 시트 대비해 범위를 한 번에 읽고(값만), 병합 셀은 전방 채우기로 처리합니다.
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
                max_col = min(ws.used_range.last_cell.column, 200)

                # 1~2행: 사용 안 함(삭제/무시). 3~8행: 병합된 헤더 영역. 9행~: 데이터.
                # 컬럼 판단: 3행을 먼저 사용(병합 시 값이 3행에 있을 가능성 높음), 없으면 4~8행 스캔
                header_block = ws.range((3, 1), (min(8, max_row), max_col)).value
                if header_block is None:
                    header_block = []
                elif header_block and not isinstance(header_block[0], (list, tuple)):
                    header_block = [header_block]
                id_col_idx = None
                enable_col_idx = None

                def scan_row_for_headers(row, max_c: int) -> None:
                    nonlocal id_col_idx, enable_col_idx
                    row = row if isinstance(row, (list, tuple)) else (row,) if row is not None else ()
                    for c, cell in enumerate(row, 1):
                        if c > max_c:
                            break
                        if cell is None:
                            continue
                        s = str(cell).strip().lower()
                        if id_col_idx is None and s == 'id':
                            id_col_idx = c
                        if enable_col_idx is None and s == 'enable':
                            enable_col_idx = c

                # 3행으로 먼저 컬럼 판단
                if header_block:
                    scan_row_for_headers(header_block[0], max_col)
                for row in (header_block[1:] if len(header_block) > 1 else []):
                    if id_col_idx is not None and enable_col_idx is not None:
                        break
                    scan_row_for_headers(row, max_col)

                if id_col_idx is None and max_row >= 9:
                    data_sample = ws.range((9, 1), (min(28, max_row), max_col)).value
                    if data_sample is not None:
                        id_col_idx = SECUIParser._find_id_column_from_block(data_sample, max_col)
                if enable_col_idx is None:
                    wb.close()
                    raise ValueError(f"'{file_path}' 시트 '{sheet_name}'에서 'Enable' 컬럼을 찾을 수 없습니다.")
                if id_col_idx is None:
                    wb.close()
                    raise ValueError(f"'{file_path}' 시트 '{sheet_name}'에서 ID 컬럼을 찾을 수 없습니다.")

                data_start_row = 9
                if data_start_row > max_row:
                    wb.close()
                    return pd.DataFrame(columns=['Rulename', 'Enable'])

                # 데이터 전체를 한 번에 읽기 (값만 읽어서 병합/수식 부담 감소)
                data_block = ws.range((data_start_row, 1), (max_row, max_col)).value
                wb.close()

            # 블록이 단일 행이면 2차원으로
            if data_block is None:
                data_block = []
            elif not isinstance(data_block[0], (list, tuple)):
                data_block = [data_block]
            id_col_0 = id_col_idx - 1
            enable_col_0 = enable_col_idx - 1
            id_values = []
            enable_values = []
            last_id = None
            last_enable = None
            for row in data_block:
                row = row if isinstance(row, (list, tuple)) else [row]
                if id_col_0 >= len(row):
                    id_val = None
                else:
                    id_val = row[id_col_0]
                if enable_col_0 >= len(row):
                    en_val = None
                else:
                    en_val = row[enable_col_0]
                if id_val is None:
                    id_val = last_id
                else:
                    last_id = id_val
                if en_val is None:
                    en_val = last_enable
                else:
                    last_enable = en_val
                id_str = (id_val if id_val is not None else '').__str__().strip()
                if re.match(r'^\d+$', id_str):
                    id_values.append(id_str)
                    enable_values.append((en_val if en_val is not None else '').__str__().strip())

            df = pd.DataFrame({'Rulename': id_values, 'Enable': enable_values})
            df['Rulename'] = df['Rulename'].fillna('').astype(str).str.strip()
            df['Enable'] = df['Enable'].fillna('').astype(str).str.strip()
            df = df[~((df['Rulename'] == '') & (df['Enable'] == ''))]
            df = df.drop_duplicates().reset_index(drop=True)
            return df
        except Exception as e:
            raise ValueError(f"파일 파싱 오류 ({file_path}, 시트: {sheet_name}): {e}")
    
    @staticmethod
    def _find_id_column_from_block(data_block, max_col: int) -> Optional[int]:
        """데이터 블록(2D 리스트)에서 숫자만 있는 컬럼을 ID 컬럼으로 찾습니다. 반환은 1-based 컬럼 인덱스."""
        if not data_block:
            return None
        if not isinstance(data_block[0], (list, tuple)):
            data_block = [data_block]
        check_rows = data_block[:20]
        for col_0 in range(min(max_col, 200)):
            numeric_count = 0
            total_count = 0
            last_val = None
            for row in check_rows:
                row = row if isinstance(row, (list, tuple)) else [row]
                if col_0 >= len(row):
                    continue
                cell_value = row[col_0] if row[col_0] is not None else last_val
                if cell_value is not None:
                    last_val = cell_value
                    total_count += 1
                    if re.match(r'^\d+$', str(cell_value).strip()):
                        numeric_count += 1
            if total_count >= 5 and numeric_count >= total_count * 0.7:
                return col_0 + 1
        return None

"""
정책 검증 모듈

Running 정책과 Candidate 정책을 비교하여 변경 사항을 검증합니다.
"""

import pandas as pd
from typing import List


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
    정책 변경 사항을 검증합니다. (성능 최적화 버전)
    
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
                     컬럼: ['Policy', 'Status', 'Running_Enable', 'Candidate_Enable', 'Message', 'IsTarget']
    """
    # 성능 최적화: 딕셔너리로 변환하여 O(1) 조회
    running_dict = {}
    for _, row in running_df.iterrows():
        policy_name = str(row['Rulename']).strip()
        if policy_name:
            running_dict[policy_name] = normalize_enable(row['Enable'])
    
    candidate_dict = {}
    for _, row in candidate_df.iterrows():
        policy_name = str(row['Rulename']).strip()
        if policy_name:
            candidate_dict[policy_name] = normalize_enable(row['Enable'])
    
    results = []
    target_set = set(p.strip() for p in target_policies if p.strip())
    
    # 1. 대상 정책 검증
    for policy_name in target_policies:
        policy_name = str(policy_name).strip()
        if not policy_name:
            continue
        
        running_enable = running_dict.get(policy_name)
        candidate_enable = candidate_dict.get(policy_name)
        
        status = ""
        message = ""
        
        if running_enable is None:
            # Running에 없는 경우
            status = "NOT_IN_RUNNING"
            message = "Running 정책에 존재하지 않음"
            running_enable = None
        elif candidate_enable is None:
            # Running에는 있지만 Candidate에는 없는 경우 (삭제됨)
            status = "DELETED"
            message = "정책이 삭제되었습니다. ✓"
        else:
            # 둘 다 있는 경우 - Enable 상태 확인
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
    running_policies_set = set(running_dict.keys())
    candidate_policies_set = set(candidate_dict.keys())
    
    # Running에 있지만 Candidate에 없는 정책 (삭제된 정책)
    deleted_policies = running_policies_set - candidate_policies_set - target_set
    
    for policy_name in deleted_policies:
        running_enable = running_dict[policy_name]
        results.append({
            'Policy': policy_name,
            'Status': 'UNEXPECTED_DELETED',
            'Running_Enable': running_enable,
            'Candidate_Enable': 'N/A',
            'Message': '대상 외 정책이 삭제되었습니다. ⚠',
            'IsTarget': False
        })
    
    # 3. 대상 외에 비활성화된 정책 찾기 (Y → N)
    common_policies = running_policies_set & candidate_policies_set - target_set
    
    for policy_name in common_policies:
        running_enable = running_dict[policy_name]
        candidate_enable = candidate_dict[policy_name]
        
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

"""
유틸리티 함수 모듈
"""

import pandas as pd
from rich.console import Console
from rich.table import Table

console = Console()


def show_summary(validation_results: pd.DataFrame):
    """
    검증 결과 요약을 표시합니다. (CLI용)
    
    Args:
        validation_results: 검증 결과 DataFrame
    """
    if validation_results.empty:
        console.print("[yellow]검증 결과가 없습니다.[/yellow]")
        return
    
    status_counts = validation_results['Status'].value_counts()
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
    
    summary_table = Table(show_header=True, header_style="bold magenta", title="검증 결과 요약")
    summary_table.add_column("상태", style="cyan")
    summary_table.add_column("개수", style="green", justify="right")
    
    for status, count in status_counts.items():
        status_str = str(status) if pd.notna(status) else 'UNKNOWN'
        status_name = status_kr.get(status_str, status_str)
        summary_table.add_row(status_name, str(count))
    
    console.print("\n")
    console.print(summary_table)
    
    # 주요 통계
    target_results = validation_results[validation_results.get('IsTarget', True) == True]
    unexpected_results = validation_results[validation_results.get('IsTarget', True) == False]
    
    deleted_count = len(target_results[target_results['Status'] == 'DELETED'])
    disabled_count = len(target_results[target_results['Status'] == 'DISABLED'])
    not_disabled_count = len(target_results[target_results['Status'] == 'NOT_DISABLED'])
    unexpected_deleted = len(unexpected_results[unexpected_results['Status'] == 'UNEXPECTED_DELETED'])
    unexpected_disabled = len(unexpected_results[unexpected_results['Status'] == 'UNEXPECTED_DISABLED'])
    
    console.print("\n[bold cyan]주요 통계:[/bold cyan]")
    console.print(f"  • 대상 정책 삭제 확인: [green]{deleted_count}개[/green]")
    console.print(f"  • 대상 정책 비활성화 확인: [green]{disabled_count}개[/green]")
    if not_disabled_count > 0:
        console.print(f"  • 비활성화 안됨: [yellow]{not_disabled_count}개[/yellow]")
    if unexpected_deleted > 0:
        console.print(f"  • 대상 외 삭제됨: [red]{unexpected_deleted}개[/red]")
    if unexpected_disabled > 0:
        console.print(f"  • 대상 외 비활성화됨: [red]{unexpected_disabled}개[/red]")


def get_summary_dict(validation_results: pd.DataFrame) -> dict:
    """
    검증 결과 요약을 딕셔너리로 반환합니다. (웹용)
    
    Args:
        validation_results: 검증 결과 DataFrame
    
    Returns:
        dict: 요약 정보 딕셔너리
    """
    if validation_results.empty:
        return {
            'total': 0,
            'target_total': 0,
            'unexpected_total': 0,
            'deleted': 0,
            'disabled': 0,
            'not_disabled': 0,
            'unexpected_deleted': 0,
            'unexpected_disabled': 0,
            'status_counts': {}
        }
    
    status_counts = validation_results['Status'].value_counts().to_dict()
    target_results = validation_results[validation_results.get('IsTarget', True) == True]
    unexpected_results = validation_results[validation_results.get('IsTarget', True) == False]
    
    return {
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

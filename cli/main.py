"""
CLI 메인 진입점

터미널에서 실행하는 명령줄 인터페이스입니다.
"""

import os
import sys
from pathlib import Path
from datetime import datetime
from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.panel import Panel

# 상위 디렉터리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from core import (
    parse_policy_file, 
    parse_target_file, 
    validate_policy_changes, 
    show_summary,
    PaloaltoParser,
    SECUIParser
)

console = Console()


def select_excel_files(current_dir: Path, file_type: str) -> list:
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
    from rich.table import Table
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


def select_vendor() -> str:
    """벤더 선택"""
    from rich.table import Table
    
    table = Table(title="벤더 선택", show_header=True, header_style="bold magenta")
    table.add_column("번호", style="cyan", width=6)
    table.add_column("벤더", style="green")
    table.add_column("설명", style="white")
    
    table.add_row("1", "Paloalto", "Rulename, Enable 컬럼 사용")
    table.add_row("2", "SECUI", "ID(숫자), Enable 컬럼 사용 (작업 전/후 시트)")
    
    console.print(table)
    console.print("\n[bold cyan]벤더를 선택하세요[/bold cyan]")
    selection = Prompt.ask("선택 (번호)", default="1")
    
    if selection == "1":
        return "Paloalto"
    elif selection == "2":
        return "SECUI"
    else:
        console.print("[yellow]잘못된 선택입니다. Paloalto로 설정합니다.[/yellow]")
        return "Paloalto"


def select_secui_sheets(file_path: str) -> tuple:
    """
    SECUI 파일에서 시트 목록을 한 번만 불러와, 작업 전/후 시트를 각각 선택받습니다.
    Returns:
        (running_sheet_name, candidate_sheet_name)
    """
    from rich.table import Table

    try:
        sheets = SECUIParser.get_sheets(file_path)
        if not sheets:
            raise ValueError("시트를 찾을 수 없습니다.")

        table = Table(
            title="SECUI 정책 파일 시트 목록",
            show_header=True,
            header_style="bold magenta",
        )
        table.add_column("번호", style="cyan", width=6)
        table.add_column("시트명", style="green")
        for idx, name in enumerate(sheets, 1):
            table.add_row(str(idx), name)
        console.print(table)

        def ask_sheet(prompt: str, default: str = "1") -> str:
            sel = Prompt.ask(prompt, default=default)
            try:
                idx = int(sel.strip()) - 1
                if 0 <= idx < len(sheets):
                    return sheets[idx]
            except ValueError:
                pass
            raise ValueError("잘못된 선택입니다.")

        console.print("\n[bold cyan]작업 전(Running) 시트를 선택하세요[/bold cyan]")
        running_sheet = ask_sheet("작업 전 시트 (번호)", "1")
        console.print(f"[green]선택됨: {running_sheet}[/green]")

        console.print("\n[bold cyan]작업 후(Candidate) 시트를 선택하세요[/bold cyan]")
        candidate_sheet = ask_sheet("작업 후 시트 (번호)", "2" if len(sheets) > 1 else "1")
        console.print(f"[green]선택됨: {candidate_sheet}[/green]")

        return running_sheet, candidate_sheet
    except Exception as e:
        console.print(f"[red]시트 선택 오류: {e}[/red]")
        raise


def get_sheet_choice(file_path: str, prompt_label: str) -> str:
    """파일에서 시트 목록을 한 번 불러와 하나 선택받습니다."""
    from rich.table import Table
    sheets = SECUIParser.get_sheets(file_path)
    if not sheets:
        raise ValueError("시트를 찾을 수 없습니다.")
    table = Table(title=f"{prompt_label} 시트 선택", show_header=True, header_style="bold magenta")
    table.add_column("번호", style="cyan", width=6)
    table.add_column("시트명", style="green")
    for idx, name in enumerate(sheets, 1):
        table.add_row(str(idx), name)
    console.print(table)
    sel = Prompt.ask(f"{prompt_label} 시트 (번호)", default="1")
    try:
        idx = int(sel.strip()) - 1
        if 0 <= idx < len(sheets):
            return sheets[idx]
    except ValueError:
        pass
    raise ValueError("잘못된 선택입니다.")


def main():
    """CLI 메인 함수"""
    current_dir = Path.cwd()
    
    console.print(Panel.fit(
        "[bold cyan]방화벽 정책 검증 자동화 스크립트[/bold cyan]\n"
        "정책 변경 사항을 검증합니다.",
        border_style="cyan"
    ))
    
    # 0. 벤더 선택
    console.print("\n[bold]0단계: 벤더 선택[/bold]")
    vendor = select_vendor()
    console.print(f"[green]선택됨: {vendor}[/green]")
    
    # 1. Running 정책 파일 선택
    console.print("\n[bold]1단계: Running 정책 파일 선택[/bold]")
    running_files = select_excel_files(current_dir, "Running 정책")
    if not running_files:
        console.print("[red]Running 정책 파일이 선택되지 않았습니다.[/red]")
        sys.exit(1)
    
    running_file = running_files[0]
    console.print(f"[green]선택됨: {running_file}[/green]")
    
    running_sheet = None
    candidate_sheet = None
    if vendor == "SECUI":
        running_file_path = str(current_dir / running_file)
        use_same_file = Confirm.ask(
            "같은 정책 파일 사용? (Running·Candidate 동일 파일, 작업 전/후 시트만 구분)",
            default=True,
        )
        if use_same_file:
            console.print("\n[bold]1단계: SECUI 정책 시트 선택 (작업 전/후)[/bold]")
            running_sheet, candidate_sheet = select_secui_sheets(running_file_path)
            candidate_file = running_file
        else:
            console.print("\n[bold]2단계: Candidate 정책 파일 선택[/bold]")
            candidate_files = select_excel_files(current_dir, "Candidate 정책")
            if not candidate_files:
                console.print("[red]Candidate 정책 파일이 선택되지 않았습니다.[/red]")
                sys.exit(1)
            candidate_file = candidate_files[0]
            console.print(f"[green]선택됨: {candidate_file}[/green]")
            candidate_file_path = str(current_dir / candidate_file)
            console.print("\n[bold]Running 시트 선택[/bold]")
            running_sheet = get_sheet_choice(running_file_path, "Running")
            console.print(f"[green]선택됨: {running_sheet}[/green]")
            console.print("\n[bold]Candidate 시트 선택[/bold]")
            candidate_sheet = get_sheet_choice(candidate_file_path, "Candidate")
            console.print(f"[green]선택됨: {candidate_sheet}[/green]")
    else:
        # 2. Candidate 정책 파일 선택 (Paloalto)
        console.print("\n[bold]2단계: Candidate 정책 파일 선택[/bold]")
        candidate_files = select_excel_files(current_dir, "Candidate 정책")
        if not candidate_files:
            console.print("[red]Candidate 정책 파일이 선택되지 않았습니다.[/red]")
            sys.exit(1)
        
        candidate_file = candidate_files[0]
        console.print(f"[green]선택됨: {candidate_file}[/green]")
    
    # 3. 대상 정책 파일 선택 (여러 개 가능)
    console.print("\n[bold]3단계: 대상 정책 파일 선택 (여러 개 선택 가능)[/bold]")
    target_files = select_excel_files(current_dir, "대상 정책")
    if not target_files:
        console.print("[yellow]대상 정책 파일이 선택되지 않았습니다. 계속 진행할까요?[/yellow]")
        if not Confirm.ask("계속 진행"):
            sys.exit(0)
        target_files = []
    
    # 4. 정책 파일 파싱
    console.print("\n" + "="*70)
    console.print("[bold]4단계: 정책 파일 파싱[/bold]")
    console.print("="*70)
    
    try:
        console.print(f"\n[cyan]Running 정책 파싱 중...[/cyan]")
        if vendor == "SECUI":
            running_policy_data = SECUIParser.parse_policy_file(
                str(current_dir / running_file), 
                running_sheet
            )
        else:
            running_policy_data = PaloaltoParser.parse_policy_file(
                str(current_dir / running_file)
            )
        console.print(f"[green]✓ 총 {len(running_policy_data)}개 정책 발견[/green]")
        
        console.print(f"\n[cyan]Candidate 정책 파싱 중...[/cyan]")
        if vendor == "SECUI":
            candidate_policy_data = SECUIParser.parse_policy_file(
                str(current_dir / candidate_file),
                candidate_sheet
            )
        else:
            candidate_policy_data = PaloaltoParser.parse_policy_file(
                str(current_dir / candidate_file)
            )
        console.print(f"[green]✓ 총 {len(candidate_policy_data)}개 정책 발견[/green]")
    except Exception as e:
        console.print(f"[red]파싱 오류: {e}[/red]")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    # 5. 대상 정책 목록 로드
    target_policies = []
    if target_files:
        console.print("\n" + "="*70)
        console.print("[bold]5단계: 대상 정책 목록 로드[/bold]")
        console.print("="*70)
        
        for target_file in target_files:
            try:
                console.print(f"\n[cyan]대상 정책 파일 파싱 중: {target_file}[/cyan]")
                policies = parse_target_file(str(current_dir / target_file))
                target_policies.extend(policies)
                console.print(f"[green]✓ {len(policies)}개 정책 발견[/green]")
            except Exception as e:
                console.print(f"[yellow]경고: {target_file} 파싱 실패 - {e}[/yellow]")
        
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
        
        try:
            console.print(f"\n[cyan]{len(target_policies)}개 대상 정책 검증 중...[/cyan]")
            validation_results = validate_policy_changes(
                running_policy_data,
                candidate_policy_data,
                target_policies
            )
            
            # 7. 검증 결과 요약 표시
            console.print("\n" + "="*70)
            console.print("[bold]7단계: 검증 결과 요약[/bold]")
            console.print("="*70)
            
            show_summary(validation_results)
            
            # 8. 검증 결과 리포트 저장
            console.print("\n" + "="*70)
            console.print("[bold]8단계: 검증 결과 리포트 저장[/bold]")
            console.print("="*70)
            
            # 날짜+시간으로 파일명 중복 방지
            validation_report_file = current_dir / (datetime.now().strftime("%Y-%m-%d_%H%M%S") + "_validation_report.xlsx")
            validation_results.to_excel(str(validation_report_file), index=False, engine='openpyxl')
            console.print(f"\n[green]✓ 검증 결과가 {validation_report_file.name}에 저장되었습니다.[/green]")
            console.print(f"[green]✓ 총 {len(validation_results)}개 정책 검증 완료[/green]")
        except Exception as e:
            console.print(f"[red]검증 오류: {e}[/red]")
            import traceback
            traceback.print_exc()
            sys.exit(1)
    
    console.print("\n" + "="*70)
    console.print("[bold green]작업 완료![/bold green]")
    console.print("="*70 + "\n")


if __name__ == "__main__":
    main()

"""
mail-check 확장이 보내는 첨부파일명 목록만으로 벤더와 파일 역할(running/candidate/target)을
파일명 키워드로 판별한다. (사내에서 파일명 규칙이 고정되어 있어 하드코딩 방식으로 구현)

판별 규칙:
- 파일명에 'export' 포함           -> SECUI. 같은 파일을 running/candidate로 사용,
                                       시트는 "작업전"(running) / "작업후"(candidate) 고정.
- 파일명이 'running-config'로 시작  -> Paloalto running 파일
- 파일명이 'candidate-config'로 시작 -> Paloalto candidate 파일
- 그 외 전부                        -> 삭제 대상(target) 파일

규칙을 만족하지 못하면(모호하거나 누락) ClassificationError를 던진다.
`stage`는 어느 단계에서 멈췄는지를 나타낸다 ('vendor_detection' | 'target_detection').
"""

from pathlib import Path
from typing import List


class ClassificationError(Exception):
    def __init__(self, message: str, stage: str):
        super().__init__(message)
        self.stage = stage


def _resolve_path(filename: str, download_folder: str) -> Path:
    base = Path.home() / 'Downloads'
    if download_folder:
        base = base / download_folder
    return base / filename


def classify_attachments(attachments: List[str], download_folder: str = '') -> dict:
    secui_files = []
    running_files = []
    candidate_files = []
    target_files = []

    for name in attachments:
        lower = name.lower()
        if 'export' in lower:
            secui_files.append(name)
        elif lower.startswith('running-config'):
            running_files.append(name)
        elif lower.startswith('candidate-config'):
            candidate_files.append(name)
        else:
            target_files.append(name)

    if secui_files and (running_files or candidate_files):
        raise ClassificationError(
            f"SECUI(export 포함: {secui_files})와 Paloalto(running/candidate-config) 파일이 "
            "동시에 감지되었습니다. 정책 파일은 한 벤더의 것만 첨부해주세요.",
            stage='vendor_detection',
        )

    if secui_files:
        if len(secui_files) > 1:
            raise ClassificationError(
                f"export가 포함된 SECUI 파일이 여러 개입니다: {secui_files}",
                stage='vendor_detection',
            )
        secui_path = _resolve_path(secui_files[0], download_folder)
        result = {
            'vendor': 'SECUI',
            'running_path': secui_path,
            'candidate_path': secui_path,
            'running_sheet': '작업전',
            'candidate_sheet': '작업후',
        }
    elif running_files or candidate_files:
        if len(running_files) != 1:
            raise ClassificationError(
                f"running-config로 시작하는 파일을 정확히 1개 찾지 못했습니다 (발견: {running_files}).",
                stage='vendor_detection',
            )
        if len(candidate_files) != 1:
            raise ClassificationError(
                f"candidate-config로 시작하는 파일을 정확히 1개 찾지 못했습니다 (발견: {candidate_files}).",
                stage='vendor_detection',
            )
        result = {
            'vendor': 'Paloalto',
            'running_path': _resolve_path(running_files[0], download_folder),
            'candidate_path': _resolve_path(candidate_files[0], download_folder),
            'running_sheet': None,
            'candidate_sheet': None,
        }
    else:
        raise ClassificationError(
            "정책 파일을 찾지 못했습니다. 파일명에 'export'(SECUI) 또는 "
            "'running-config'/'candidate-config'(Paloalto)가 포함되어야 합니다.",
            stage='vendor_detection',
        )

    if not target_files:
        raise ClassificationError(
            "삭제 대상 파일이 없습니다. 정책/후보 파일 외의 첨부파일이 최소 1개 필요합니다.",
            stage='target_detection',
        )

    result['target_paths'] = [_resolve_path(name, download_folder) for name in target_files]
    return result

"""
core.attachment_classifier.classify_attachments() 테스트

mail-check가 보내는 첨부파일명 목록(vendor/역할 정보 없음)만으로
SECUI/Paloalto 벤더와 running/candidate/target 파일 역할을 하드코딩 키워드로
판별하는 로직을 검증한다.

규칙:
- 파일명에 'export' 포함 -> SECUI (같은 파일을 running/candidate로 사용, 시트 "작업전"/"작업후")
- 파일명이 'running-config'로 시작 -> Paloalto running 파일
- 파일명이 'candidate-config'로 시작 -> Paloalto candidate 파일
- 나머지 전부 -> 삭제 대상(target) 파일
"""

import pytest

from core.attachment_classifier import ClassificationError, classify_attachments


def test_paloalto_classification():
    result = classify_attachments(
        ['running-config_20260101.xlsx', 'candidate-config_20260101.xlsx', '삭제대상.xlsx'],
        'mail-check',
    )
    assert result['vendor'] == 'Paloalto'
    assert result['running_path'].name == 'running-config_20260101.xlsx'
    assert result['candidate_path'].name == 'candidate-config_20260101.xlsx'
    assert [p.name for p in result['target_paths']] == ['삭제대상.xlsx']
    assert 'Downloads/mail-check' in str(result['running_path']).replace('\\', '/')


def test_secui_classification_uses_same_file_for_both_sheets():
    result = classify_attachments(['정책_export.xlsx', '대상.xlsx'], '')
    assert result['vendor'] == 'SECUI'
    assert result['running_path'] == result['candidate_path']
    assert result['running_sheet'] == '작업전'
    assert result['candidate_sheet'] == '작업후'


def test_classification_is_case_insensitive():
    result = classify_attachments(
        ['RUNNING-CONFIG_x.xlsx', 'CANDIDATE-CONFIG_x.xlsx', 't.xlsx'], ''
    )
    assert result['vendor'] == 'Paloalto'


def test_missing_candidate_file_raises_at_vendor_detection():
    with pytest.raises(ClassificationError) as exc:
        classify_attachments(['running-config_x.xlsx', 'target.xlsx'], '')
    assert exc.value.stage == 'vendor_detection'


def test_no_recognizable_vendor_file_raises_at_vendor_detection():
    with pytest.raises(ClassificationError) as exc:
        classify_attachments(['random.xlsx'], '')
    assert exc.value.stage == 'vendor_detection'


def test_export_and_paloalto_files_together_is_ambiguous():
    with pytest.raises(ClassificationError) as exc:
        classify_attachments(
            ['running-config_x.xlsx', 'candidate-config_x.xlsx', 'export_y.xlsx'], ''
        )
    assert exc.value.stage == 'vendor_detection'


def test_no_target_files_raises_at_target_detection():
    with pytest.raises(ClassificationError) as exc:
        classify_attachments(['running-config_x.xlsx', 'candidate-config_x.xlsx'], '')
    assert exc.value.stage == 'target_detection'


def test_duplicate_running_files_is_ambiguous():
    with pytest.raises(ClassificationError) as exc:
        classify_attachments(
            ['running-config_a.xlsx', 'running-config_b.xlsx', 'candidate-config_x.xlsx', 'target.xlsx'],
            '',
        )
    assert exc.value.stage == 'vendor_detection'

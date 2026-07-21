from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))

import review_pdf31_ground_truth as review  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402


def main() -> int:
    state = review.ReviewState(
        schema="pdf31_ground_truth_review_state_v1",
        version="draft",
        source_pdf=str(review.SOURCE),
        source_sha256=review.sha256_file(review.SOURCE),
        reviewer="test",
        last_truth_id="",
        records=review.default_records(),
    )
    counts = review.status_counts(state.records)
    assert len(state.records) == 49
    assert counts.get(review.STATUS_CONFIRMED, 0) == 12
    assert counts.get(review.STATUS_NEEDS_COORD, 0) == 28
    assert counts.get(review.STATUS_RELEASE, 0) == 1
    assert counts.get(review.STATUS_PENDING, 0) == 8
    assert counts.get(review.STATUS_EXCLUDED, 0) == 0
    assert counts.get(review.STATUS_HOLD, 0) == 0
    errors = review.completion_errors(state.records)
    assert errors
    assert any("要座標修正" in error for error in errors)
    assert any("ユーザー確認待ち" in error for error in errors)
    formal_records = [record for record in state.records if record.user_status == review.STATUS_CONFIRMED]
    excluded_records = [record for record in state.records if record.user_status == review.STATUS_EXCLUDED]
    assert len(formal_records) == 12
    assert len(excluded_records) == 0
    test_spinbox_rect_apply()
    test_move_then_approve_preserves_scene_rect()
    print("pdf31_ground_truth_review_tests=passed")
    return 0


def test_spinbox_rect_apply() -> None:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance() or QApplication([])
    original_state_json = review.STATE_JSON
    original_words_for_page = review.GroundTruthReviewDialog.words_for_page
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_state = Path(tmpdir) / "review_state.json"
        review.STATE_JSON = temp_state
        review.GroundTruthReviewDialog.words_for_page = lambda self, page: []
        try:
            state = review.ReviewState(
                schema="pdf31_ground_truth_review_state_v1",
                version="draft",
                source_pdf=str(review.SOURCE),
                source_sha256=review.sha256_file(review.SOURCE),
                reviewer="test",
                last_truth_id="",
                records=review.default_records(),
            )
            dialog = review.GroundTruthReviewDialog(state)
            target = next(record for record in state.records if record.truth_id == "P13-003")
            dialog.select_record(target)
            original_width = target.rect[2] - target.rect[0]
            assert round(original_width, 1) == 82.0
            dialog.w_spin.setValue(70.0)
            dialog.apply_edits()
            assert round(target.rect[2] - target.rect[0], 1) == 70.0
            assert round(dialog.w_spin.value(), 1) == 70.0
            selected_items = [
                item
                for item, record in dialog.items.items()
                if record.truth_id == "P13-003"
            ]
            assert selected_items
            scene_rect = selected_items[0].mapRectToScene(selected_items[0].rect())
            assert round(scene_rect.width() / dialog.scale, 1) == 70.0
            restored = review.load_state(path=temp_state, source_pdf=review.SOURCE)
            restored_target = next(record for record in restored.records if record.truth_id == "P13-003")
            assert round(restored_target.rect[2] - restored_target.rect[0], 1) == 70.0
            restarted = review.GroundTruthReviewDialog(restored)
            restarted.select_record(next(record for record in restored.records if record.truth_id == "P13-003"))
            assert round(restarted.w_spin.value(), 1) == 70.0
            restarted.close()
            dialog.close()
        finally:
            review.STATE_JSON = original_state_json
            review.GroundTruthReviewDialog.words_for_page = original_words_for_page
            app.processEvents()


def test_move_then_approve_preserves_scene_rect() -> None:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance() or QApplication([])
    original_state_json = review.STATE_JSON
    original_words_for_page = review.GroundTruthReviewDialog.words_for_page
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_state = Path(tmpdir) / "review_state.json"
        review.STATE_JSON = temp_state
        review.GroundTruthReviewDialog.words_for_page = lambda self, page: []
        try:
            state = review.ReviewState(
                schema="pdf31_ground_truth_review_state_v1",
                version="draft",
                source_pdf=str(review.SOURCE),
                source_sha256=review.sha256_file(review.SOURCE),
                reviewer="test",
                last_truth_id="",
                records=review.default_records(),
            )
            dialog = review.GroundTruthReviewDialog(state)
            target = next(record for record in state.records if record.truth_id == "P13-003")
            original_rect = list(target.rect)
            dialog.select_record(target)
            target_items = [item for item, record in dialog.items.items() if record.truth_id == "P13-003"]
            assert target_items
            target_items[0].setPos(10.0 * dialog.scale, 5.0 * dialog.scale)
            dialog.approve_current()
            assert target.user_status == review.STATUS_CONFIRMED
            assert round(target.rect[0] - original_rect[0], 1) == 10.0
            assert round(target.rect[1] - original_rect[1], 1) == 5.0
            assert round(target.rect[2] - target.rect[0], 1) == round(original_rect[2] - original_rect[0], 1)
            restored = review.load_state(path=temp_state, source_pdf=review.SOURCE)
            restored_target = next(record for record in restored.records if record.truth_id == "P13-003")
            assert restored_target.user_status == review.STATUS_CONFIRMED
            assert round(restored_target.rect[0] - original_rect[0], 1) == 10.0
            assert round(restored_target.rect[1] - original_rect[1], 1) == 5.0
            restarted = review.GroundTruthReviewDialog(restored)
            restarted.select_record(restored_target)
            assert round(restarted.x_spin.value() - original_rect[0], 1) == 10.0
            assert round(restarted.y_spin.value() - original_rect[1], 1) == 5.0
            restarted.close()
            dialog.close()
        finally:
            review.STATE_JSON = original_state_json
            review.GroundTruthReviewDialog.words_for_page = original_words_for_page
            app.processEvents()


if __name__ == "__main__":
    raise SystemExit(main())

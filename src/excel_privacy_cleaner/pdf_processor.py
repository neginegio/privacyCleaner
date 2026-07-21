from __future__ import annotations

import csv
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import fitz
except ImportError:  # pragma: no cover - exercised only when dependency is missing.
    fitz = None  # type: ignore[assignment]

from .excel_processor import AliasBook, ProcessingOptions, replacement_for
from .models import Finding
from .pdf_context_rules import context_candidates_for_page
from .pdf_ocr_support import (
    CANDIDATE_AUTO,
    CANDIDATE_MANUAL,
    CANDIDATE_REJECTED,
    CANDIDATE_REVIEW,
    PAGE_COMPLETED,
    PAGE_FAILED_UNRESOLVED,
    PAGE_REVIEWED_NO_SENSITIVE_DATA,
    PAGE_REVIEWED_WITH_REDACTIONS,
    PAGE_UNREVIEWED,
    PAGE_VERIFICATION_FAILED,
    QUALITY_FAILED,
    QUALITY_PASS,
    QUALITY_REVIEW,
    PdfPageQuality,
    USER_APPROVED,
    USER_REJECTED,
    VERIFICATION_FAILED,
    VERIFICATION_MANUAL_REGION_PASS,
    VERIFICATION_NOT_EVALUATED,
    VERIFICATION_TEXT_RESIDUAL_PASS,
    detect_ocr_candidates,
    evaluate_page_quality,
)
from .presidio_japanese import JapanesePresidioDetector, alias_kind, entity_label, normalize_alias_key
from .resources import resource_path


PDF_REDACTION_MODES = {
    "pseudonym": "仮名化",
    "mask": "部分マスキング",
    "black": "黒塗り",
    "white": "白塗り",
    "delete": "完全削除",
}

OCR_LANGUAGE = "jpn+eng"
OCR_DPI = 300
OCR_MIN_TEXT_CHARS = 20
TESSDATA_FILES = ("jpn.traineddata", "eng.traineddata", "osd.traineddata")


@dataclass(frozen=True)
class PdfConversionResult:
    pdf_path: Path
    csv_path: Path
    report_path: Path
    warnings: tuple[str, ...]
    converted_count: int
    review_count: int
    page_count: int
    ocr_success_pages: tuple[int, ...] = ()
    ocr_failed_pages: tuple[str, ...] = ()
    failed_pages: tuple[int, ...] = ()


@dataclass(frozen=True)
class PdfLocation:
    page_index: int
    rect: tuple[float, float, float, float]

    @property
    def page_number(self) -> int:
        return self.page_index + 1

    @property
    def coordinate_text(self) -> str:
        x0, y0, x1, y1 = self.rect
        return f"x={x0:.1f},y={y0:.1f},w={x1 - x0:.1f},h={y1 - y0:.1f}"


class PdfPrivacyProcessor:
    def __init__(self) -> None:
        self.detector = JapanesePresidioDetector()
        self.alias_book = AliasBook()
        self.options = ProcessingOptions()
        self.locations: dict[tuple[str, str, str], PdfLocation] = {}
        self.scan_warnings: list[str] = []
        self.page_modes: dict[int, str] = {}
        self.ocr_success_pages: set[int] = set()
        self.ocr_failed_pages: dict[int, str] = {}
        self.page_quality: dict[int, PdfPageQuality] = {}
        self.confirmed_pages: set[int] = set()
        self.page_review_state: dict[int, str] = {}
        self.page_verification_state: dict[int, str] = {}
        self.last_review_page = 0
        self.temp_dir: Path | None = None
        self.temp_pdf: Path | None = None
        self.page_count = 0

    def cleanup(self) -> None:
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None
        self.temp_pdf = None
        self.locations = {}
        self.page_modes = {}
        self.ocr_success_pages = set()
        self.ocr_failed_pages = {}
        self.page_quality = {}
        self.confirmed_pages = set()
        self.page_review_state = {}
        self.page_verification_state = {}
        self.last_review_page = 0

    def _initialize_page_state(self, page_index: int) -> None:
        quality = self.page_quality.get(page_index)
        if quality and quality.verdict == QUALITY_FAILED:
            self.page_review_state[page_index] = PAGE_FAILED_UNRESOLVED
        else:
            self.page_review_state[page_index] = PAGE_UNREVIEWED
        self.page_verification_state[page_index] = VERIFICATION_NOT_EVALUATED

    def mark_page_no_sensitive_data(self, page_index: int) -> None:
        self.page_review_state[page_index] = PAGE_REVIEWED_NO_SENSITIVE_DATA
        self.confirmed_pages.add(page_index)
        self.last_review_page = page_index

    def mark_page_reviewed_with_redactions(self, page_index: int) -> None:
        self.page_review_state[page_index] = PAGE_REVIEWED_WITH_REDACTIONS
        self.confirmed_pages.add(page_index)
        self.last_review_page = page_index

    def mark_page_verification_failed(self, page_index: int) -> None:
        self.page_review_state[page_index] = PAGE_VERIFICATION_FAILED
        self.page_verification_state[page_index] = VERIFICATION_FAILED

    def mark_page_completed(self, page_index: int) -> None:
        self.page_review_state[page_index] = PAGE_COMPLETED
        self.confirmed_pages.add(page_index)
        self.last_review_page = page_index

    def scan(self, pdf_path: Path, options: ProcessingOptions | None = None) -> list[Finding]:
        self._require_fitz()
        self.cleanup()
        self.options = options or ProcessingOptions()
        self.alias_book = AliasBook()
        self.locations = {}
        self.scan_warnings = []
        self.page_modes = {}
        self.ocr_success_pages = set()
        self.ocr_failed_pages = {}
        self.page_quality = {}
        self.confirmed_pages = set()
        self.page_review_state = {}
        self.page_verification_state = {}
        self.last_review_page = 0
        self.temp_dir = Path(tempfile.mkdtemp(prefix="PdfPrivacyCleaner_"))
        self.temp_pdf = self.temp_dir / pdf_path.name
        shutil.copy2(pdf_path, self.temp_pdf)

        findings: list[Finding] = []
        seen: set[tuple[str, str, str, str, str]] = set()
        known_family_names: dict[str, str] = {}

        doc = fitz.open(self.temp_pdf)
        try:
            unsupported = self._structural_unsupported_reasons(doc)
            if unsupported:
                raise RuntimeError("PDFを処理できません: " + " / ".join(unsupported))

            self.page_count = doc.page_count
            ocr_runtime_error = _ocr_runtime_error()
            for page_index in range(doc.page_count):
                page = doc[page_index]
                page_label = f"ページ{page_index + 1}"
                page_text = page.get_text("text").strip()
                text_spans = list(_iter_spans(page.get_text("dict")))
                if len(page_text) >= OCR_MIN_TEXT_CHARS and text_spans:
                    self.page_modes[page_index] = "text"
                    self.page_quality[page_index] = _text_page_quality(page, page_text)
                    self._initialize_page_state(page_index)
                else:
                    self.page_modes[page_index] = "ocr"
                    if ocr_runtime_error is not None:
                        self.ocr_failed_pages[page_index] = ocr_runtime_error
                        self.page_quality[page_index] = evaluate_page_quality(page, "", [], [], exception=ocr_runtime_error)
                        self._initialize_page_state(page_index)
                        continue
                    try:
                        ocr_text, ocr_words = ocr_page_text_and_words(page)
                        ocr_candidates = detect_ocr_candidates(ocr_words, page.rect)
                        ocr_candidates.extend(context_candidates_for_page(page_index + 1, page.rect))
                        self.page_quality[page_index] = evaluate_page_quality(page, ocr_text, ocr_words, ocr_candidates)
                        self._initialize_page_state(page_index)
                        for candidate in ocr_candidates:
                            if candidate.status == CANDIDATE_REJECTED:
                                continue
                            self._append_ocr_candidate_finding(findings, seen, page_label, page_index, candidate)
                        self.ocr_success_pages.add(page_index)
                        continue
                    except Exception as exc:
                        self.ocr_failed_pages[page_index] = str(exc)
                        self.page_quality[page_index] = evaluate_page_quality(page, "", [], [], exception=str(exc))
                        self._initialize_page_state(page_index)
                        continue
                for span_text, bbox in text_spans:
                    if not span_text.strip():
                        continue
                    occupied: list[range] = []
                    for start, end, entity_type in _pdf_extra_results(span_text):
                        self._append_pdf_finding(
                            findings,
                            seen,
                            page_label,
                            page_index,
                            bbox,
                            span_text,
                            start,
                            end,
                            entity_type,
                            occupied,
                            known_family_names,
                        )
                    for result in self.detector.analyze(span_text):
                        self._append_pdf_finding(
                            findings,
                            seen,
                            page_label,
                            page_index,
                            bbox,
                            span_text,
                            result.start,
                            result.end,
                            result.entity_type,
                            occupied,
                            known_family_names,
                        )
                    for surname, alias in known_family_names.items():
                        for match in re.finditer(rf"{re.escape(surname)}(?:さん|様|氏)", span_text):
                            match_range = range(match.start(), match.end())
                            if any(match_range.start < item.stop and item.start < match_range.stop for item in occupied):
                                continue
                            replacement = self.alias_book.get("name_candidate", alias)
                            finding = self._make_finding(
                                enabled=False,
                                page_label=page_label,
                                page_index=page_index,
                                bbox=bbox,
                                span_text=span_text,
                                start=match.start(),
                                end=match.end(),
                                entity_type="氏名候補",
                                detection_kind="確認候補",
                                original=match.group(0),
                                replacement=f"{replacement}{match.group(0)[len(surname):]}",
                                reason="姓のみの敬称表現のため要確認",
                            )
                            self._append_finding(findings, seen, finding)
                            occupied.append(match_range)
        finally:
            doc.close()
        return findings

    def convert_with_artifacts(
        self,
        source_path: Path,
        findings: list[Finding],
        output_dir: Path | None = None,
        options: ProcessingOptions | None = None,
        redaction_mode: str = "pseudonym",
        write_artifacts: bool = True,
    ) -> PdfConversionResult:
        self._require_fitz()
        if not self.temp_pdf or not self.temp_pdf.exists():
            raise RuntimeError("先にPDF検査を実行してください。")
        if redaction_mode not in PDF_REDACTION_MODES:
            raise RuntimeError(f"未対応のPDF匿名化方法です: {redaction_mode}")

        active_options = options or self.options
        can_output, block_reasons = final_output_status(
            findings,
            self.page_quality,
            self.confirmed_pages,
            self.page_review_state,
        )
        if not can_output:
            raise RuntimeError("PDFを匿名化済み最終ファイルとして出力できません: " + " / ".join(block_reasons))

        selected = [finding for finding in findings if finding.enabled]
        if not selected:
            raise RuntimeError("PDFの変換対象が選択されていません。")

        warnings = list(self.scan_warnings)
        output_path = self._make_output_path(source_path, output_dir)
        doc = fitz.open(self.temp_pdf)
        failed_pages: set[int] = set()
        sanitize_notes: list[str] = []
        try:
            sanitize_notes.extend(_sanitize_document(doc))
            for finding in selected:
                location = self.locations.get((finding.sheet, finding.cell, finding.original))
                if location is None:
                    warnings.append(f"{finding.sheet}!{finding.cell}: 位置情報が見つからないためスキップしました。")
                    failed_pages.add(_page_index_from_label(finding.sheet))
                    continue
                page = doc[location.page_index]
                rect = fitz.Rect(location.rect)
                fill = _fill_color(redaction_mode)
                page.add_redact_annot(rect, fill=fill)
                page.apply_redactions(
                    images=fitz.PDF_REDACT_IMAGE_PIXELS,
                    graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                    text=fitz.PDF_REDACT_TEXT_REMOVE,
                )
                if redaction_mode in {"pseudonym", "mask"}:
                    _insert_replacement_text(page, rect, finding.replacement)

            doc.set_metadata({})
            try:
                doc.del_xml_metadata()
            except Exception:
                sanitize_notes.append("XMLメタデータの削除は未対応または失敗しました。")
            doc.save(output_path, garbage=4, clean=True, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
        finally:
            doc.close()

        validation_warnings = validate_pdf_output(source_path, output_path, selected, failed_pages)
        warnings.extend(validation_warnings)
        warnings.extend(sanitize_notes)
        if validation_warnings:
            for page_index in failed_pages:
                self.mark_page_verification_failed(page_index)
            raise RuntimeError("PDF出力後検証に失敗しました: " + " / ".join(validation_warnings[:5]))

        selected_pages = {_page_index_from_label(finding.sheet) for finding in selected}
        manual_pages = {
            _page_index_from_label(finding.sheet)
            for finding in selected
            if finding.detection_kind == CANDIDATE_MANUAL or finding.entity_type == "手動追加"
        }
        for page_index in range(self.page_count):
            if page_index in manual_pages:
                self.page_verification_state[page_index] = VERIFICATION_MANUAL_REGION_PASS
            elif page_index in selected_pages:
                self.page_verification_state[page_index] = VERIFICATION_TEXT_RESIDUAL_PASS
            else:
                self.page_verification_state[page_index] = VERIFICATION_NOT_EVALUATED
            if self.page_review_state.get(page_index) in {
                PAGE_REVIEWED_NO_SENSITIVE_DATA,
                PAGE_REVIEWED_WITH_REDACTIONS,
                PAGE_COMPLETED,
            }:
                self.page_review_state[page_index] = PAGE_COMPLETED

        csv_path = output_path.with_name(f"{output_path.stem}_検出変換結果.csv")
        report_path = output_path.with_name(f"{output_path.stem}_処理報告書.txt")
        result = PdfConversionResult(
            pdf_path=output_path,
            csv_path=csv_path,
            report_path=report_path,
            warnings=tuple(warnings),
            converted_count=len(selected),
            review_count=sum(1 for finding in findings if finding.detection_kind in {"確認候補", CANDIDATE_REVIEW}),
            page_count=self.page_count,
            ocr_success_pages=tuple(page_index + 1 for page_index in sorted(self.ocr_success_pages)),
            ocr_failed_pages=tuple(
                f"ページ{page_index + 1}: {reason}" for page_index, reason in sorted(self.ocr_failed_pages.items())
            ),
            failed_pages=tuple(quality.page_number for quality in self.page_quality.values() if quality.verdict == QUALITY_FAILED),
        )
        if write_artifacts:
            write_pdf_findings_csv(csv_path, findings)
            write_pdf_processing_report(
                report_path,
                source_path=source_path,
                result=result,
                findings=findings,
                redaction_mode=redaction_mode,
                validation_warnings=validation_warnings,
                sanitize_notes=sanitize_notes,
            )
            write_pdf_quality_csv(
                output_path.with_name(f"{output_path.stem}_ページ品質.csv"),
                self.page_quality,
                self.page_review_state,
                self.page_verification_state,
            )
            write_pdf_audit_json(
                output_path.with_name(f"{output_path.stem}_監査報告.json"),
                findings,
                self.locations,
                self.page_quality,
                self.page_review_state,
                self.page_verification_state,
            )
        self.cleanup()
        return result

    def add_manual_redaction(
        self,
        findings: list[Finding],
        page_index: int,
        rect: tuple[float, float, float, float],
        entity_type: str = "手動追加",
        replacement: str = "[REDACTED]",
    ) -> Finding:
        page_label = f"ページ{page_index + 1}"
        cell = f"{page_index + 1}-{len(self.locations) + 1:04d} ({PdfLocation(page_index, rect).coordinate_text})"
        finding = Finding(
            enabled=True,
            sheet=page_label,
            cell=cell,
            entity_type=entity_type,
            detection_kind=CANDIDATE_MANUAL,
            original="[手動範囲]",
            replacement=replacement,
            reason="利用者が手動で追加した匿名化範囲",
        )
        self.locations[(finding.sheet, finding.cell, finding.original)] = PdfLocation(page_index, rect)
        findings.append(finding)
        return finding

    def export_review_state(self, path: Path, source_path: Path, findings: list[Finding], last_page: int | None = None) -> None:
        payload: dict[str, Any] = {
            "schema": "pdf_review_state_v1",
            "input_file_name": source_path.name,
            "input_sha256": _file_sha256(source_path),
            "saved_at": f"{datetime.now():%Y-%m-%d %H:%M:%S}",
            "last_page": self.last_review_page if last_page is None else last_page,
            "pages": [
                {
                    "page_number": page_index + 1,
                    "page_quality": self.page_quality.get(page_index).verdict if page_index in self.page_quality else "",
                    "page_review_state": self.page_review_state.get(page_index, PAGE_UNREVIEWED),
                    "verification_state": self.page_verification_state.get(page_index, VERIFICATION_NOT_EVALUATED),
                }
                for page_index in range(self.page_count)
            ],
            "findings": [],
        }
        for finding in findings:
            location = self.locations.get((finding.sheet, finding.cell, finding.original))
            payload["findings"].append(
                {
                    "id": _finding_state_id(finding, location),
                    "page_number": _page_index_from_label(finding.sheet) + 1,
                    "entity_type": finding.entity_type,
                    "detection_kind": finding.detection_kind,
                    "enabled": finding.enabled,
                    "replacement": finding.replacement,
                    "original_hmac_sha256": _hmac_digest(finding.original) if finding.original else "",
                    "rect": location.rect if location else None,
                    "manual_added": finding.detection_kind == CANDIDATE_MANUAL or finding.entity_type == "手動追加",
                }
            )
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    def import_review_state(self, path: Path, source_path: Path, findings: list[Finding]) -> None:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("schema") != "pdf_review_state_v1":
            raise RuntimeError("未対応のPDFレビュー状態ファイルです。")
        if payload.get("input_sha256") != _file_sha256(source_path):
            raise RuntimeError("入力PDFが保存時と一致しません。")

        self.page_review_state = {
            int(page["page_number"]) - 1: str(page.get("page_review_state", PAGE_UNREVIEWED))
            for page in payload.get("pages", [])
        }
        self.page_verification_state = {
            int(page["page_number"]) - 1: str(page.get("verification_state", VERIFICATION_NOT_EVALUATED))
            for page in payload.get("pages", [])
        }
        self.confirmed_pages = {
            page_index
            for page_index, state in self.page_review_state.items()
            if state in {PAGE_REVIEWED_NO_SENSITIVE_DATA, PAGE_REVIEWED_WITH_REDACTIONS, PAGE_COMPLETED}
        }
        self.last_review_page = int(payload.get("last_page", 0))

        existing = {
            _finding_state_id(finding, self.locations.get((finding.sheet, finding.cell, finding.original))): finding
            for finding in findings
        }
        for saved in payload.get("findings", []):
            rect_value = saved.get("rect")
            rect = tuple(float(value) for value in rect_value) if rect_value else None
            if saved.get("manual_added") and rect is not None:
                page_index = int(saved.get("page_number", 1)) - 1
                restored = self.add_manual_redaction(
                    findings,
                    page_index,
                    rect,  # type: ignore[arg-type]
                    entity_type=str(saved.get("entity_type", "手動追加")),
                    replacement=str(saved.get("replacement", "[REDACTED]")),
                )
                restored.detection_kind = str(saved.get("detection_kind", CANDIDATE_MANUAL))
                restored.enabled = bool(saved.get("enabled", True))
                continue

            finding = existing.get(str(saved.get("id", "")))
            if finding is None:
                continue
            finding.entity_type = str(saved.get("entity_type", finding.entity_type))
            finding.detection_kind = str(saved.get("detection_kind", finding.detection_kind))
            finding.enabled = bool(saved.get("enabled", finding.enabled))
            finding.replacement = str(saved.get("replacement", finding.replacement))
            if rect is not None:
                self.locations[(finding.sheet, finding.cell, finding.original)] = PdfLocation(
                    _page_index_from_label(finding.sheet),
                    rect,  # type: ignore[arg-type]
                )

    def _append_pdf_finding(
        self,
        findings: list[Finding],
        seen: set[tuple[str, str, str, str, str]],
        page_label: str,
        page_index: int,
        bbox: tuple[float, float, float, float],
        span_text: str,
        start: int,
        end: int,
        recognizer_entity_type: str,
        occupied: list[range],
        known_family_names: dict[str, str],
    ) -> None:
        match_range = range(start, end)
        if any(match_range.start < item.stop and item.start < match_range.stop for item in occupied):
            return
        original = span_text[start:end].strip()
        if len(original) < 2:
            return
        kind = alias_kind(recognizer_entity_type)
        replacement = replacement_for(kind, original, self.alias_book, self.options)
        label = entity_label(recognizer_entity_type)
        if label == recognizer_entity_type and recognizer_entity_type == "PDF_ENGLISH_NAME":
            label = "氏名"
            kind = "name"
            replacement = replacement_for(kind, original, self.alias_book, self.options)
        finding = self._make_finding(
            enabled=True,
            page_label=page_label,
            page_index=page_index,
            bbox=bbox,
            span_text=span_text,
            start=start,
            end=end,
            entity_type=label,
            detection_kind="OCR" if self.page_modes.get(page_index) == "ocr" else "PDFテキスト",
            original=original,
            replacement=replacement,
            reason="OCRテキスト層から検出" if self.page_modes.get(page_index) == "ocr" else "PDFテキスト層から検出",
        )
        self._append_finding(findings, seen, finding)
        occupied.append(match_range)
        if kind == "name":
            surname = _surname_candidate(original)
            if surname:
                known_family_names.setdefault(surname, replacement)

    def _append_ocr_candidate_finding(
        self,
        findings: list[Finding],
        seen: set[tuple[str, str, str, str, str]],
        page_label: str,
        page_index: int,
        candidate: Any,
    ) -> None:
        enabled = candidate.status == CANDIDATE_AUTO
        cell = f"{page_index + 1}-{len(self.locations) + 1:04d} ({PdfLocation(page_index, candidate.rect).coordinate_text})"
        finding = Finding(
            enabled=enabled,
            sheet=page_label,
            cell=cell,
            entity_type=candidate.entity_type,
            detection_kind=candidate.status,
            original=candidate.original,
            replacement=candidate.replacement,
            reason=candidate.reason,
        )
        self.locations[(finding.sheet, finding.cell, finding.original)] = PdfLocation(page_index, candidate.rect)
        self._append_finding(findings, seen, finding)

    def _make_finding(
        self,
        enabled: bool,
        page_label: str,
        page_index: int,
        bbox: tuple[float, float, float, float],
        span_text: str,
        start: int,
        end: int,
        entity_type: str,
        detection_kind: str,
        original: str,
        replacement: str,
        reason: str,
    ) -> Finding:
        rect = _substring_rect(bbox, span_text, start, end)
        cell = f"{page_index + 1}-{len(self.locations) + 1:04d} ({PdfLocation(page_index, rect).coordinate_text})"
        finding = Finding(
            enabled=enabled,
            sheet=page_label,
            cell=cell,
            entity_type=entity_type,
            detection_kind=detection_kind,
            original=original,
            replacement=replacement,
            reason=reason,
            start=start,
            end=end,
        )
        self.locations[(finding.sheet, finding.cell, finding.original)] = PdfLocation(page_index, rect)
        return finding

    @staticmethod
    def _append_finding(
        findings: list[Finding],
        seen: set[tuple[str, str, str, str, str]],
        finding: Finding,
    ) -> None:
        if finding.dedupe_key in seen:
            return
        findings.append(finding)
        seen.add(finding.dedupe_key)

    @staticmethod
    def _make_output_path(source_path: Path, output_dir: Path | None = None) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = output_dir or source_path.parent
        return folder / f"{source_path.stem}_匿名化済みPDF_{timestamp}.pdf"

    @staticmethod
    def _structural_unsupported_reasons(doc: Any) -> list[str]:
        reasons: list[str] = []
        if doc.is_encrypted or getattr(doc, "needs_pass", False):
            reasons.append("パスワード付きPDFは第一段階では対象外です。")
        try:
            if doc.get_sigflags() > 0:
                reasons.append("電子署名付きPDFは第一段階では対象外です。")
        except Exception:
            reasons.append("電子署名の有無を確認できませんでした。")
        is_form_pdf = getattr(doc, "is_form_pdf", False)
        if callable(is_form_pdf):
            is_form_pdf = is_form_pdf()
        if is_form_pdf:
            reasons.append("PDFフォームは第一段階では対象外です。")
        try:
            if doc.embfile_count() > 0:
                reasons.append("添付ファイルを含むPDFは第一段階では対象外です。")
        except Exception:
            reasons.append("添付ファイルの有無を確認できませんでした。")
        try:
            catalog = doc.xref_object(doc.pdf_catalog())
            if "/Collection" in catalog:
                reasons.append("ポートフォリオPDFは第一段階では対象外です。")
        except Exception:
            pass
        return reasons

    @staticmethod
    def _require_fitz() -> None:
        if fitz is None:
            raise RuntimeError("PDF機能には PyMuPDF が必要です。requirements.txt を確認してインストールしてください。")


def write_pdf_findings_csv(path: Path, findings: list[Finding]) -> None:
    headers = ["変換", "処理状態", "ページ", "位置", "種類", "検査", "検出値", "変換後", "理由"]
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(headers)
        for finding in findings:
            if finding.detection_kind in {"確認候補", CANDIDATE_REVIEW}:
                status = "確認候補"
            elif finding.detection_kind == USER_APPROVED:
                status = "利用者承認"
            elif finding.detection_kind == USER_REJECTED:
                status = "利用者解除"
            elif finding.detection_kind == CANDIDATE_AUTO:
                status = "自動確定"
            elif finding.detection_kind == CANDIDATE_MANUAL:
                status = "手動追加"
            else:
                status = "自動変換" if finding.enabled else "除外"
            writer.writerow(
                [
                    "対象" if finding.enabled else "除外",
                    status,
                    finding.sheet,
                    finding.cell,
                    finding.entity_type,
                    finding.detection_kind,
                    finding.original,
                    finding.replacement,
                    finding.reason,
                ]
            )


def write_pdf_quality_csv(
    path: Path,
    page_quality: dict[int, PdfPageQuality],
    page_review_state: dict[int, str] | None = None,
    page_verification_state: dict[int, str] | None = None,
) -> None:
    headers = [
        "ページ番号",
        "OCR取得文字数",
        "座標付き単語数",
        "日本語文字数",
        "英数字数",
        "記号数",
        "文字化け候補数",
        "文字領域割合",
        "検出候補数",
        "品質判定",
        "ページ確認状態",
        "再検証状態",
        "警告理由",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(headers)
        for page_index, quality in sorted(page_quality.items()):
            writer.writerow(
                [
                    quality.page_number,
                    quality.ocr_text_chars,
                    quality.coordinate_word_count,
                    quality.japanese_char_count,
                    quality.alnum_char_count,
                    quality.symbol_char_count,
                    quality.garble_candidate_count,
                    f"{quality.text_area_ratio:.6f}",
                    quality.candidate_count,
                    quality.verdict,
                    page_review_state.get(page_index, "") if page_review_state else "",
                    page_verification_state.get(page_index, "") if page_verification_state else "",
                    quality.warning_reason,
                ]
            )


def write_pdf_audit_json(
    path: Path,
    findings: list[Finding],
    locations: dict[tuple[str, str, str], PdfLocation],
    page_quality: dict[int, PdfPageQuality],
    page_review_state: dict[int, str] | None = None,
    page_verification_state: dict[int, str] | None = None,
) -> None:
    payload: dict[str, Any] = {
        "pages": [
            {
                "page_number": quality.page_number,
                "quality": quality.verdict,
                "review_state": page_review_state.get(_index, "") if page_review_state else "",
                "verification_state": page_verification_state.get(_index, "") if page_verification_state else "",
                "warning": quality.warning_reason,
                "ocr_text_chars": quality.ocr_text_chars,
                "coordinate_word_count": quality.coordinate_word_count,
                "candidate_count": quality.candidate_count,
            }
            for _index, quality in sorted(page_quality.items())
        ],
        "findings": [],
    }
    for finding in findings:
        location = locations.get((finding.sheet, finding.cell, finding.original))
        payload["findings"].append(
            {
                "page": finding.sheet,
                "entity_type": finding.entity_type,
                "confirmation_method": finding.detection_kind,
                "auto_confirmed": finding.detection_kind == CANDIDATE_AUTO,
                "user_approved": finding.detection_kind == USER_APPROVED,
                "user_rejected": finding.detection_kind == USER_REJECTED,
                "manual_added": finding.detection_kind == CANDIDATE_MANUAL,
                "enabled": finding.enabled,
                "original_hmac_sha256": _hmac_digest(finding.original) if finding.original else "",
                "rect": location.rect if location else None,
                "warning": finding.reason,
            }
        )
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_pdf_processing_report(
    path: Path,
    source_path: Path,
    result: PdfConversionResult,
    findings: list[Finding],
    redaction_mode: str,
    validation_warnings: list[str],
    sanitize_notes: list[str],
) -> None:
    converted = [finding for finding in findings if finding.enabled]
    lines = [
        "PDF匿名化 処理報告書",
        f"入力ファイル名: {source_path.name}",
        f"出力ファイル名: {result.pdf_path.name}",
        f"処理日時: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"ページ数: {result.page_count}",
        f"OCR成功ページ数: {len(result.ocr_success_pages)}",
        f"OCR成功ページ: {', '.join(map(str, result.ocr_success_pages)) if result.ocr_success_pages else 'なし'}",
        f"OCR失敗ページ: {' / '.join(result.ocr_failed_pages) if result.ocr_failed_pages else 'なし'}",
        f"FAILEDページ: {', '.join(map(str, result.failed_pages)) if result.failed_pages else 'なし'}",
        f"匿名化方法: {PDF_REDACTION_MODES.get(redaction_mode, redaction_mode)}",
        f"検出件数: {len(findings)}",
        f"変換件数: {len(converted)}",
        f"解除件数: {len(findings) - len(converted)}",
        f"要確認件数: {sum(1 for finding in findings if finding.detection_kind == '確認候補')}",
        f"検証結果: {'成功' if not validation_warnings else '警告あり'}",
        f"サニタイズ結果: {' / '.join(sanitize_notes) if sanitize_notes else '可能な範囲で完了'}",
        "警告内容:",
    ]
    if result.warnings:
        lines.extend(f"- {warning}" for warning in result.warnings)
    else:
        lines.append("- なし")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def final_output_status(
    findings: list[Finding],
    page_quality: dict[int, PdfPageQuality],
    confirmed_pages: set[int],
    page_review_state: dict[int, str] | None = None,
) -> tuple[bool, list[str]]:
    reasons: list[str] = []
    unresolved = [finding for finding in findings if finding.detection_kind in {"確認候補", CANDIDATE_REVIEW}]
    if unresolved:
        reasons.append(f"未確認候補: {len(unresolved)}件")
    if page_review_state is None:
        failed_pages = [
            quality.page_number
            for quality in page_quality.values()
            if quality.verdict == QUALITY_FAILED and (quality.page_number - 1) not in confirmed_pages
        ]
        if failed_pages:
            reasons.append(f"未対応のFAILEDページ: {', '.join(map(str, failed_pages))}")
        return not reasons, reasons

    unreviewed_pages = [
        page_index + 1
        for page_index in range(len(page_quality))
        if page_review_state.get(page_index, PAGE_UNREVIEWED) == PAGE_UNREVIEWED
    ]
    if unreviewed_pages:
        reasons.append(f"UNREVIEWEDページ: {', '.join(map(str, unreviewed_pages[:20]))}")

    review_required_pages = [
        quality.page_number
        for page_index, quality in page_quality.items()
        if quality.verdict == QUALITY_REVIEW and page_review_state.get(page_index) == PAGE_UNREVIEWED
    ]
    if review_required_pages:
        reasons.append(f"未対応のREVIEW_REQUIREDページ: {', '.join(map(str, review_required_pages[:20]))}")

    failed_unresolved_pages = [
        page_index + 1
        for page_index, state in sorted(page_review_state.items())
        if state == PAGE_FAILED_UNRESOLVED
    ]
    if failed_unresolved_pages:
        reasons.append(f"FAILED_UNRESOLVEDページ: {', '.join(map(str, failed_unresolved_pages[:20]))}")

    verification_failed_pages = [
        page_index + 1
        for page_index, state in sorted(page_review_state.items())
        if state == PAGE_VERIFICATION_FAILED
    ]
    if verification_failed_pages:
        reasons.append(f"VERIFICATION_FAILEDページ: {', '.join(map(str, verification_failed_pages[:20]))}")

    invalid_states = [
        f"ページ{page_index + 1}:{state}"
        for page_index, state in sorted(page_review_state.items())
        if state
        not in {
            PAGE_REVIEWED_NO_SENSITIVE_DATA,
            PAGE_REVIEWED_WITH_REDACTIONS,
            PAGE_COMPLETED,
            PAGE_UNREVIEWED,
            PAGE_FAILED_UNRESOLVED,
            PAGE_VERIFICATION_FAILED,
        }
    ]
    if invalid_states:
        reasons.append("不明なページ状態: " + ", ".join(invalid_states[:20]))
    return not reasons, reasons


def validate_pdf_output(
    source_path: Path,
    output_path: Path,
    findings: list[Finding],
    failed_pages: set[int],
) -> list[str]:
    warnings: list[str] = []
    if source_path.resolve() == output_path.resolve():
        warnings.append("元ファイルを上書きしようとしました。")
    try:
        source_doc = fitz.open(source_path)
        output_doc = fitz.open(output_path)
        try:
            if source_doc.page_count != output_doc.page_count:
                warnings.append("ページ数が変化しました。")
            if failed_pages:
                pages = ", ".join(str(page + 1) for page in sorted(failed_pages))
                warnings.append(f"処理に失敗したページがあります: {pages}")
            output_text = "\n".join(output_doc[index].get_text("text") for index in range(output_doc.page_count))
            for original in _sensitive_originals(findings):
                if original and original in output_text:
                    warnings.append(f"出力PDFのテキスト抽出結果に元文字列が残っています: {original}")
            for finding in findings:
                for page_index in range(output_doc.page_count):
                    if output_doc[page_index].search_for(finding.original):
                        warnings.append(f"出力PDF検索で元文字列が見つかりました: {finding.original}")
                        break
            ocr_findings = [finding for finding in findings if finding.detection_kind in {CANDIDATE_AUTO, USER_APPROVED, CANDIDATE_MANUAL}]
            if ocr_findings:
                ocr_runtime_error = _ocr_runtime_error()
                if ocr_runtime_error:
                    warnings.append(f"出力PDFの再OCR検証を実行できませんでした: {ocr_runtime_error}")
                else:
                    pages = sorted({_page_index_from_label(finding.sheet) for finding in ocr_findings})
                    for page_index in pages:
                        try:
                            ocr_text, _words = ocr_page_text_and_words(output_doc[page_index])
                        except Exception as exc:
                            warnings.append(f"出力PDFの再OCRに失敗しました: ページ{page_index + 1}: {exc}")
                            continue
                        for original in _sensitive_originals(
                            [finding for finding in ocr_findings if _page_index_from_label(finding.sheet) == page_index]
                        ):
                            if original and original in ocr_text:
                                warnings.append(f"出力PDFの再OCR結果に元文字列が残っています: ページ{page_index + 1}: {original}")
        finally:
            source_doc.close()
            output_doc.close()
    except Exception as exc:
        warnings.append(f"出力PDFを正常に検証できませんでした: {exc}")
    return warnings


def _sanitize_document(doc: Any) -> list[str]:
    notes: list[str] = []
    removed_annots = 0
    for page in doc:
        annot = page.first_annot
        while annot:
            next_annot = annot.next
            page.delete_annot(annot)
            removed_annots += 1
            annot = next_annot
    if removed_annots:
        notes.append(f"既存注釈を削除: {removed_annots} 件")
    return notes


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _finding_state_id(finding: Finding, location: PdfLocation | None) -> str:
    rect_text = ",".join(f"{value:.3f}" for value in location.rect) if location else ""
    payload = "|".join(
        [
            finding.sheet,
            finding.cell,
            finding.entity_type,
            finding.detection_kind,
            rect_text,
            _hmac_digest(finding.original) if finding.original else "",
        ]
    )
    return _hmac_digest(payload)


def _hmac_digest(value: str) -> str:
    return hmac.new(_state_hmac_key(), value.encode("utf-8"), hashlib.sha256).hexdigest()


def _state_hmac_key() -> bytes:
    env_key = os.environ.get("PRIVACY_CLEANER_STATE_HMAC_KEY")
    if env_key:
        return env_key.encode("utf-8")

    key_path = _state_hmac_key_path()
    if key_path.exists():
        return key_path.read_bytes()
    key_path.parent.mkdir(parents=True, exist_ok=True)
    key = secrets.token_bytes(32)
    key_path.write_bytes(key)
    return key


def _state_hmac_key_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.cwd()))
    preferred = base / "ExcelPrivacyCleaner" / "state_hmac.key"
    try:
        preferred.parent.mkdir(parents=True, exist_ok=True)
        probe = preferred.parent / ".write_test"
        probe.write_text("", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return preferred
    except Exception:
        return Path.cwd() / "config" / ".privacy_cleaner_state_hmac.key"


def _iter_spans(text_dict: dict[str, Any]):
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = str(span.get("text", ""))
                bbox = tuple(float(item) for item in span.get("bbox", (0, 0, 0, 0)))
                yield text, bbox


def _text_page_quality(page: Any, text: str) -> PdfPageQuality:
    words = page.get_text("words")
    japanese_count = sum(1 for char in text if "\u3040" <= char <= "\u30ff" or "\u4e00" <= char <= "\u9fff")
    alnum_count = sum(1 for char in text if char.isalnum() and not ("\u3040" <= char <= "\u30ff" or "\u4e00" <= char <= "\u9fff"))
    symbol_count = sum(1 for char in text if not char.isspace() and not char.isalnum())
    area = sum(max(float(word[2]) - float(word[0]), 0) * max(float(word[3]) - float(word[1]), 0) for word in words)
    page_area = max(float(page.rect.width) * float(page.rect.height), 1.0)
    return PdfPageQuality(
        page_number=int(page.number) + 1,
        ocr_text_chars=len(text.strip()),
        coordinate_word_count=len(words),
        japanese_char_count=japanese_count,
        alnum_char_count=alnum_count,
        symbol_char_count=symbol_count,
        garble_candidate_count=0,
        text_area_ratio=min(area / page_area, 1.0),
        candidate_count=0,
        verdict=QUALITY_PASS,
        warning_reason="文字PDFとして処理",
    )


def bundled_tessdata_dir() -> Path:
    return resource_path("resources/tessdata")


def ocr_tessdata_dir() -> Path:
    source = bundled_tessdata_dir()
    if _is_ascii_path(source):
        return source
    digest = hashlib.sha1(str(source).encode("utf-8")).hexdigest()[:12]
    target = Path(tempfile.gettempdir()) / "PrivacyCleanerTessdata" / digest
    _mirror_tessdata(source, target)
    return target


def validate_ocr_environment() -> list[str]:
    errors: list[str] = []
    if fitz is None:
        return ["PyMuPDF が利用できません。"]
    if not hasattr(fitz.Page, "get_textpage_ocr"):
        errors.append("PyMuPDF の page.get_textpage_ocr が利用できません。")
    source = bundled_tessdata_dir()
    if not source.exists():
        errors.append(f"tessdata フォルダが見つかりません: {source}")
        return errors
    missing = [name for name in TESSDATA_FILES if not (source / name).exists()]
    if missing:
        errors.append(f"OCR言語データが不足しています: {', '.join(missing)} ({source})")
    try:
        runtime = ocr_tessdata_dir()
        runtime_missing = [name for name in TESSDATA_FILES if not (runtime / name).exists()]
        if runtime_missing:
            errors.append(f"OCR実行用言語データが不足しています: {', '.join(runtime_missing)} ({runtime})")
    except Exception as exc:
        errors.append(f"OCR実行用tessdataの準備に失敗しました: {exc}")
    return errors


def _ocr_runtime_error() -> str | None:
    errors = validate_ocr_environment()
    return " / ".join(errors) if errors else None


def _ocr_page_spans(page: Any) -> list[tuple[str, tuple[float, float, float, float]]]:
    textpage = page.get_textpage_ocr(
        language=OCR_LANGUAGE,
        dpi=OCR_DPI,
        full=True,
        tessdata=str(ocr_tessdata_dir()),
    )
    spans = list(_iter_spans(textpage.extractDICT()))
    if not spans:
        raise RuntimeError("OCRで文字を取得できませんでした。")
    return spans


def ocr_page_text_and_words(page: Any) -> tuple[str, list[tuple[Any, ...]]]:
    textpage = page.get_textpage_ocr(
        language=OCR_LANGUAGE,
        dpi=OCR_DPI,
        full=True,
        tessdata=str(ocr_tessdata_dir()),
    )
    return textpage.extractTEXT(), list(textpage.extractWORDS())


def write_pdf_debug_image(
    source_path: Path,
    findings: list[Finding],
    locations: dict[tuple[str, str, str], PdfLocation],
    output_path: Path,
    page_index: int = 0,
) -> None:
    if fitz is None:
        raise RuntimeError("PDF機能には PyMuPDF が必要です。")
    source_doc = fitz.open(source_path)
    debug_doc = fitz.open()
    try:
        debug_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
        page = debug_doc[0]
        for finding in findings:
            location = locations.get((finding.sheet, finding.cell, finding.original))
            if location is None or location.page_index != page_index:
                continue
            page.draw_rect(fitz.Rect(location.rect), color=(1, 0, 0), width=1.2)
        pixmap = page.get_pixmap(dpi=150, alpha=False)
        pixmap.save(output_path)
    finally:
        debug_doc.close()
        source_doc.close()


def _is_ascii_path(path: Path) -> bool:
    try:
        str(path).encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _mirror_tessdata(source: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    for filename in TESSDATA_FILES:
        source_file = source / filename
        target_file = target / filename
        if not target_file.exists() or target_file.stat().st_size != source_file.stat().st_size:
            shutil.copy2(source_file, target_file)


def _pdf_extra_results(text: str) -> list[tuple[int, int, str]]:
    results: list[tuple[int, int, str]] = []
    for pattern in (
        r"(?:氏名|Name|NAME)\s*[:：]\s*([A-Za-z][A-Za-z .'-]{2,60})",
        r"(?:担当者|Contact)\s*[:：]\s*([A-Za-z][A-Za-z .'-]{2,60})",
    ):
        for match in re.finditer(pattern, text):
            results.append((match.start(1), match.end(1), "PDF_ENGLISH_NAME"))
    return results


def _substring_rect(bbox: tuple[float, float, float, float], text: str, start: int, end: int) -> tuple[float, float, float, float]:
    x0, y0, x1, y1 = bbox
    text_len = max(len(text), 1)
    left = x0 + (x1 - x0) * max(start, 0) / text_len
    right = x0 + (x1 - x0) * min(end, text_len) / text_len
    pad = 1.5
    return (max(x0, left - pad), max(y0, y0 - pad), min(x1, right + pad), min(y1 + pad, y1 + pad))


def _fill_color(redaction_mode: str) -> tuple[float, float, float]:
    if redaction_mode == "black":
        return (0, 0, 0)
    return (1, 1, 1)


def _insert_replacement_text(page: Any, rect: Any, replacement: str) -> None:
    fontfile = _japanese_font_path()
    kwargs: dict[str, Any] = {"fontsize": max(min(rect.height * 0.75, 11), 6), "color": (0, 0, 0)}
    if fontfile:
        kwargs["fontfile"] = str(fontfile)
    try:
        page.insert_textbox(rect, replacement, **kwargs)
    except Exception:
        page.insert_text((rect.x0, rect.y1 - 1), replacement, **kwargs)


def _japanese_font_path() -> Path | None:
    for candidate in (
        Path("C:/Windows/Fonts/BIZ-UDGothicR.ttc"),
        Path("C:/Windows/Fonts/msgothic.ttc"),
        Path("C:/Windows/Fonts/meiryo.ttc"),
    ):
        if candidate.exists():
            return candidate
    return None


def _sensitive_originals(findings: list[Finding]) -> list[str]:
    values = {
        finding.original
        for finding in findings
        if finding.entity_type in {
            "氏名",
            "氏名カナ",
            "氏名候補",
            "会社名",
            "住所",
            "電話番号",
            "メールアドレス",
            "個人番号",
            "銀行口座",
            "秘密情報",
            "APIキー",
            "パスワード",
        }
        and len(finding.original) >= 3
    }
    return sorted(values, key=len, reverse=True)


def _surname_candidate(value: str) -> str | None:
    normalized = value.strip()
    parts = [part for part in re.split(r"[\s　]+", normalized) if part]
    if parts:
        return parts[0]
    if re.fullmatch(r"[一-龯々〆ヵヶ]{3,6}", normalized):
        return normalized[:2]
    return None


def _page_index_from_label(label: str) -> int:
    match = re.search(r"(\d+)", label)
    return int(match.group(1)) - 1 if match else 0

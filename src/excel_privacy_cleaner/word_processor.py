from __future__ import annotations

import hashlib
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document


WORD_SUPPORTED_EXTENSION = ".docx"
WORD_MACRO_EXTENSION = ".docm"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class UnsupportedWordDocumentError(RuntimeError):
    pass


@dataclass(frozen=True)
class WordTextRun:
    story_type: str
    section_index: int | None
    container_type: str
    paragraph_index: int
    run_index: int
    char_start: int
    char_end: int
    text: str
    paragraph_text: str
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    cell_paragraph_index: int | None = None
    header_footer_type: str | None = None
    part_name: str = ""
    element_path: str = ""
    hyperlink_url: str | None = None
    hyperlink_rel_id: str | None = None
    is_linked_to_previous: bool = False

    @property
    def location_id(self) -> str:
        values = [
            self.story_type,
            str(self.section_index),
            self.header_footer_type or "",
            self.container_type,
            str(self.table_index),
            str(self.row_index),
            str(self.cell_index),
            str(self.paragraph_index),
            str(self.cell_paragraph_index),
            str(self.run_index),
            str(self.char_start),
            str(self.char_end),
            self.part_name,
        ]
        return ":".join(values)


@dataclass(frozen=True)
class WordParagraphText:
    story_type: str
    section_index: int | None
    container_type: str
    paragraph_index: int
    text: str
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    cell_paragraph_index: int | None = None
    header_footer_type: str | None = None
    part_name: str = ""
    element_path: str = ""
    is_linked_to_previous: bool = False


@dataclass(frozen=True)
class WordHyperlink:
    story_type: str
    section_index: int | None
    container_type: str
    paragraph_index: int
    text: str
    url: str
    rel_id: str
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    cell_paragraph_index: int | None = None
    header_footer_type: str | None = None
    part_name: str = ""
    element_path: str = ""


@dataclass(frozen=True)
class WordDocumentProperties:
    author: str
    last_modified_by: str
    title: str
    subject: str
    keywords: str
    category: str
    comments: str
    created: datetime | None
    modified: datetime | None
    content_status: str
    identifier: str
    language: str
    revision: int | None
    version: str


@dataclass(frozen=True)
class UnsupportedWordFeature:
    feature_type: str
    part_name: str
    count: int
    detail: str = ""


@dataclass(frozen=True)
class WordStructureInventory:
    source_path: Path
    source_sha256: str
    paragraphs: tuple[WordParagraphText, ...]
    runs: tuple[WordTextRun, ...]
    hyperlinks: tuple[WordHyperlink, ...]
    document_properties: WordDocumentProperties
    unsupported_features: tuple[UnsupportedWordFeature, ...]
    macro_parts: tuple[str, ...]


def extract_word_structure(path: Path) -> WordStructureInventory:
    validate_docx_input(path)
    source_sha256 = _file_sha256(path)
    unsupported_features, macro_parts = inspect_docx_package(path)
    if macro_parts:
        raise UnsupportedWordDocumentError("マクロを含むWord文書はv1対象外です: " + ", ".join(macro_parts))

    document = Document(path)
    paragraphs: list[WordParagraphText] = []
    runs: list[WordTextRun] = []
    hyperlinks: list[WordHyperlink] = []

    for paragraph_index, paragraph in enumerate(document.paragraphs):
        _append_paragraph(
            paragraphs,
            runs,
            hyperlinks,
            paragraph,
            story_type="body",
            section_index=None,
            container_type="paragraph",
            paragraph_index=paragraph_index,
            element_path=f"body/p[{paragraph_index}]",
        )

    for table_index, table in enumerate(document.tables):
        _append_table(
            paragraphs,
            runs,
            hyperlinks,
            table,
            story_type="body",
            section_index=None,
            header_footer_type=None,
            table_index=table_index,
            part_name=_part_name(table),
            element_path_prefix=f"body/tbl[{table_index}]",
        )

    for section_index, section in enumerate(document.sections):
        for header_footer_type, story_type, story in (
            ("default", "header", section.header),
            ("first_page", "header", section.first_page_header),
            ("even_page", "header", section.even_page_header),
            ("default", "footer", section.footer),
            ("first_page", "footer", section.first_page_footer),
            ("even_page", "footer", section.even_page_footer),
        ):
            linked = bool(story.is_linked_to_previous)
            part_name = _part_name(story)
            story_prefix = f"{story_type}[{section_index}]/{header_footer_type}"
            for paragraph_index, paragraph in enumerate(story.paragraphs):
                _append_paragraph(
                    paragraphs,
                    runs,
                    hyperlinks,
                    paragraph,
                    story_type=story_type,
                    section_index=section_index,
                    container_type="paragraph",
                    paragraph_index=paragraph_index,
                    header_footer_type=header_footer_type,
                    part_name=part_name,
                    element_path=f"{story_prefix}/p[{paragraph_index}]",
                    is_linked_to_previous=linked,
                )
            for table_index, table in enumerate(story.tables):
                _append_table(
                    paragraphs,
                    runs,
                    hyperlinks,
                    table,
                    story_type=story_type,
                    section_index=section_index,
                    header_footer_type=header_footer_type,
                    table_index=table_index,
                    part_name=part_name or _part_name(table),
                    element_path_prefix=f"{story_prefix}/tbl[{table_index}]",
                    is_linked_to_previous=linked,
                )

    return WordStructureInventory(
        source_path=path,
        source_sha256=source_sha256,
        paragraphs=tuple(paragraphs),
        runs=tuple(runs),
        hyperlinks=tuple(hyperlinks),
        document_properties=_document_properties(document),
        unsupported_features=unsupported_features,
        macro_parts=macro_parts,
    )


def validate_docx_input(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix == WORD_MACRO_EXTENSION:
        raise UnsupportedWordDocumentError(".docmはWord匿名化v1の対象外です。")
    if suffix != WORD_SUPPORTED_EXTENSION:
        raise UnsupportedWordDocumentError(f"Word匿名化v1は.docxのみ対応です: {path.name}")
    if not path.exists():
        raise FileNotFoundError(path)


def inspect_docx_package(path: Path) -> tuple[tuple[UnsupportedWordFeature, ...], tuple[str, ...]]:
    features: dict[tuple[str, str], int] = {}
    details: dict[tuple[str, str], str] = {}
    macro_parts: list[str] = []

    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            for name in names:
                lowered = name.lower()
                if lowered.endswith("vbaproject.bin") or "vbaproject" in lowered:
                    macro_parts.append(name)
                if name.startswith("word/embeddings/"):
                    _count_feature(features, details, "embedded_object", name)
                if name.startswith("customXml/"):
                    _count_feature(features, details, "custom_xml", name)
                if name.startswith("word/comments") and name.endswith(".xml"):
                    _count_feature(features, details, "comments", name)
                if name == "word/footnotes.xml":
                    _count_feature(features, details, "footnotes", name)
                if name == "word/endnotes.xml":
                    _count_feature(features, details, "endnotes", name)
                if not name.endswith((".xml", ".rels")):
                    continue
                data = archive.read(name)
                text = data.decode("utf-8", errors="ignore")
                _inspect_xml_text(name, text, features, details)
    except zipfile.BadZipFile as exc:
        raise UnsupportedWordDocumentError(f"DOCX ZIP構造を読み取れませんでした: {exc}") from exc

    unsupported = tuple(
        UnsupportedWordFeature(feature_type=feature_type, part_name=part_name, count=count, detail=details.get((feature_type, part_name), ""))
        for (feature_type, part_name), count in sorted(features.items())
    )
    return unsupported, tuple(sorted(macro_parts))


def _append_table(
    paragraphs: list[WordParagraphText],
    runs: list[WordTextRun],
    hyperlinks: list[WordHyperlink],
    table: Any,
    *,
    story_type: str,
    section_index: int | None,
    header_footer_type: str | None,
    table_index: int,
    part_name: str,
    element_path_prefix: str,
    is_linked_to_previous: bool = False,
) -> None:
    for row_index, row in enumerate(table.rows):
        for cell_index, cell in enumerate(row.cells):
            for cell_paragraph_index, paragraph in enumerate(cell.paragraphs):
                _append_paragraph(
                    paragraphs,
                    runs,
                    hyperlinks,
                    paragraph,
                    story_type=story_type,
                    section_index=section_index,
                    container_type="table_cell",
                    paragraph_index=cell_paragraph_index,
                    table_index=table_index,
                    row_index=row_index,
                    cell_index=cell_index,
                    cell_paragraph_index=cell_paragraph_index,
                    header_footer_type=header_footer_type,
                    part_name=part_name,
                    element_path=f"{element_path_prefix}/tr[{row_index}]/tc[{cell_index}]/p[{cell_paragraph_index}]",
                    is_linked_to_previous=is_linked_to_previous,
                )


def _append_paragraph(
    paragraphs: list[WordParagraphText],
    runs: list[WordTextRun],
    hyperlinks: list[WordHyperlink],
    paragraph: Any,
    *,
    story_type: str,
    section_index: int | None,
    container_type: str,
    paragraph_index: int,
    table_index: int | None = None,
    row_index: int | None = None,
    cell_index: int | None = None,
    cell_paragraph_index: int | None = None,
    header_footer_type: str | None = None,
    part_name: str = "",
    element_path: str,
    is_linked_to_previous: bool = False,
) -> None:
    part_name = part_name or _part_name(paragraph)
    extracted_runs = _paragraph_runs_with_offsets(paragraph)
    paragraph_text = "".join(item["text"] for item in extracted_runs)
    paragraphs.append(
        WordParagraphText(
            story_type=story_type,
            section_index=section_index,
            container_type=container_type,
            paragraph_index=paragraph_index,
            text=paragraph_text,
            table_index=table_index,
            row_index=row_index,
            cell_index=cell_index,
            cell_paragraph_index=cell_paragraph_index,
            header_footer_type=header_footer_type,
            part_name=part_name,
            element_path=element_path,
            is_linked_to_previous=is_linked_to_previous,
        )
    )
    for run_index, item in enumerate(extracted_runs):
        text = str(item["text"])
        run_path = f"{element_path}/r[{run_index}]"
        run = WordTextRun(
            story_type=story_type,
            section_index=section_index,
            container_type=container_type,
            paragraph_index=paragraph_index,
            table_index=table_index,
            row_index=row_index,
            cell_index=cell_index,
            cell_paragraph_index=cell_paragraph_index,
            run_index=run_index,
            char_start=int(item["char_start"]),
            char_end=int(item["char_end"]),
            text=text,
            paragraph_text=paragraph_text,
            header_footer_type=header_footer_type,
            part_name=part_name,
            element_path=run_path,
            hyperlink_url=item.get("hyperlink_url"),
            hyperlink_rel_id=item.get("hyperlink_rel_id"),
            is_linked_to_previous=is_linked_to_previous,
        )
        runs.append(run)
    for item in extracted_runs:
        rel_id = item.get("hyperlink_rel_id")
        url = item.get("hyperlink_url")
        if rel_id and url:
            hyperlinks.append(
                WordHyperlink(
                    story_type=story_type,
                    section_index=section_index,
                    container_type=container_type,
                    paragraph_index=paragraph_index,
                    table_index=table_index,
                    row_index=row_index,
                    cell_index=cell_index,
                    cell_paragraph_index=cell_paragraph_index,
                    header_footer_type=header_footer_type,
                    text=str(item["text"]),
                    url=str(url),
                    rel_id=str(rel_id),
                    part_name=part_name,
                    element_path=element_path,
                )
            )


def _paragraph_runs_with_offsets(paragraph: Any) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    cursor = 0
    for child in paragraph._p.iterchildren():
        tag = _local_name(child.tag)
        if tag == "r":
            text = _run_text(child)
            if text:
                items.append({"text": text, "char_start": cursor, "char_end": cursor + len(text)})
                cursor += len(text)
        elif tag == "hyperlink":
            rel_id = child.get(f"{{{R_NS}}}id")
            url = _hyperlink_target(paragraph, rel_id) if rel_id else None
            for run in child.iterchildren():
                if _local_name(run.tag) != "r":
                    continue
                text = _run_text(run)
                if text:
                    items.append(
                        {
                            "text": text,
                            "char_start": cursor,
                            "char_end": cursor + len(text),
                            "hyperlink_url": url,
                            "hyperlink_rel_id": rel_id,
                        }
                    )
                    cursor += len(text)
    return items


def _run_text(run_element: Any) -> str:
    parts: list[str] = []
    for child in run_element.iterchildren():
        tag = _local_name(child.tag)
        if tag == "t":
            parts.append(child.text or "")
        elif tag == "tab":
            parts.append("\t")
        elif tag in {"br", "cr"}:
            parts.append("\n")
    return "".join(parts)


def _hyperlink_target(paragraph: Any, rel_id: str | None) -> str | None:
    if not rel_id:
        return None
    part = getattr(paragraph, "part", None)
    if part is None:
        parent = getattr(paragraph, "_parent", None)
        part = getattr(parent, "part", None)
    if part is None:
        return None
    relationship = part.rels.get(rel_id)
    return getattr(relationship, "target_ref", None) if relationship is not None else None


def _document_properties(document: Any) -> WordDocumentProperties:
    props = document.core_properties
    return WordDocumentProperties(
        author=str(props.author or ""),
        last_modified_by=str(props.last_modified_by or ""),
        title=str(props.title or ""),
        subject=str(props.subject or ""),
        keywords=str(props.keywords or ""),
        category=str(props.category or ""),
        comments=str(props.comments or ""),
        created=props.created,
        modified=props.modified,
        content_status=str(props.content_status or ""),
        identifier=str(props.identifier or ""),
        language=str(props.language or ""),
        revision=props.revision,
        version=str(props.version or ""),
    )


def _inspect_xml_text(
    part_name: str,
    text: str,
    features: dict[tuple[str, str], int],
    details: dict[tuple[str, str], str],
) -> None:
    checks = (
        ("textbox", ("txbxContent", "v:textbox", "wps:txbx")),
        ("drawing_or_shape", ("<w:drawing", "<v:shape", "<wps:wsp")),
        ("tracked_changes", ("<w:ins", "<w:del", "<w:moveFrom", "<w:moveTo")),
        ("complex_field_code", ("<w:fldChar", "<w:instrText")),
    )
    for feature_type, markers in checks:
        count = sum(text.count(marker) for marker in markers)
        if count:
            _count_feature(features, details, feature_type, part_name, count=count)
    if part_name.endswith(".rels") and 'TargetMode="External"' in text:
        _count_feature(features, details, "external_link", part_name, count=text.count('TargetMode="External"'))
    if part_name.endswith(".xml") and _may_hold_text(part_name, text):
        _count_feature(features, details, "text_holding_xml_part", part_name)


def _may_hold_text(part_name: str, text: str) -> bool:
    supported_prefixes = (
        "word/document.xml",
        "word/header",
        "word/footer",
        "docProps/",
    )
    if part_name.startswith(supported_prefixes):
        return False
    return "<w:t" in text or "<dc:" in text or "<cp:" in text


def _count_feature(
    features: dict[tuple[str, str], int],
    details: dict[tuple[str, str], str],
    feature_type: str,
    part_name: str,
    *,
    count: int = 1,
    detail: str = "",
) -> None:
    key = (feature_type, part_name)
    features[key] = features.get(key, 0) + count
    if detail:
        details[key] = detail


def _part_name(item: Any) -> str:
    part = getattr(item, "part", None)
    if part is None:
        parent = getattr(item, "_parent", None)
        part = getattr(parent, "part", None)
    partname = getattr(part, "partname", "")
    return str(partname)


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

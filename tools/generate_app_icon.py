from __future__ import annotations

import struct
import sys
from pathlib import Path

from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QApplication


ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = ROOT / "assets"
PNG_PATH = ASSET_DIR / "app_icon.png"
ICO_PATH = ASSET_DIR / "app_icon.ico"


def rounded_rect_path(rect: QRectF, radius: float) -> QPainterPath:
    path = QPainterPath()
    path.addRoundedRect(rect, radius, radius)
    return path


def draw_icon(size: int) -> QImage:
    scale = size / 256
    image = QImage(size, size, QImage.Format_ARGB32)
    image.fill(Qt.transparent)

    painter = QPainter(image)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.TextAntialiasing, True)

    def rect(x: float, y: float, w: float, h: float) -> QRectF:
        return QRectF(x * scale, y * scale, w * scale, h * scale)

    def pen(color: str, width: float) -> QPen:
        return QPen(QColor(color), width * scale, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)

    # Bright neutral tile: independent of Excel/Word while remaining visible.
    bg = rounded_rect_path(rect(12, 12, 232, 232), 40 * scale)
    painter.fillPath(bg, QColor("#fff7ed"))
    painter.setPen(pen("#fb7185", 7))
    painter.drawPath(bg)

    # Oversized redaction marker body.
    marker = rounded_rect_path(rect(36, 82, 184, 84), 26 * scale)
    painter.fillPath(marker, QColor("#dc2626"))

    cap = rounded_rect_path(rect(34, 74, 58, 100), 18 * scale)
    painter.fillPath(cap, QColor("#991b1b"))

    tip = QPainterPath()
    tip.moveTo(QPointF(215 * scale, 88 * scale))
    tip.lineTo(QPointF(244 * scale, 124 * scale))
    tip.lineTo(QPointF(215 * scale, 160 * scale))
    tip.closeSubpath()
    painter.fillPath(tip, QColor("#7f1d1d"))

    # Gloss and edge lines keep the mark readable at taskbar size.
    painter.setPen(Qt.NoPen)
    painter.fillPath(rounded_rect_path(rect(63, 93, 126, 10), 5 * scale), QColor("#fca5a5"))
    painter.fillPath(rounded_rect_path(rect(62, 144, 132, 9), 4 * scale), QColor("#b91c1c"))

    # Redaction bars behind the marker communicate "masked text" without document branding.
    painter.setPen(pen("#111827", 14))
    painter.drawLine(QPointF(51 * scale, 194 * scale), QPointF(207 * scale, 194 * scale))
    painter.drawLine(QPointF(77 * scale, 218 * scale), QPointF(179 * scale, 218 * scale))

    painter.end()
    return image


def image_to_png_bytes(image: QImage) -> bytes:
    byte_array = QByteArray()
    buffer = QBuffer(byte_array)
    buffer.open(QIODevice.WriteOnly)
    image.save(buffer, "PNG")
    buffer.close()
    return bytes(byte_array)


def write_ico(path: Path, sizes: list[int]) -> None:
    png_entries = [(size, image_to_png_bytes(draw_icon(size))) for size in sizes]
    header_size = 6 + (16 * len(png_entries))
    offset = header_size

    directory = bytearray()
    data = bytearray()
    for size, png in png_entries:
        width = 0 if size == 256 else size
        height = 0 if size == 256 else size
        directory += struct.pack("<BBBBHHII", width, height, 0, 0, 1, 32, len(png), offset)
        data += png
        offset += len(png)

    path.write_bytes(struct.pack("<HHH", 0, 1, len(png_entries)) + directory + data)


def main() -> int:
    QApplication.instance() or QApplication(sys.argv)
    ASSET_DIR.mkdir(exist_ok=True)
    draw_icon(1024).save(str(PNG_PATH), "PNG")
    write_ico(ICO_PATH, [16, 24, 32, 48, 64, 128, 256])
    print(f"Generated {PNG_PATH}")
    print(f"Generated {ICO_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

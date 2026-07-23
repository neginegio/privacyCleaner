from __future__ import annotations

import cgi
import html
import secrets
import shutil
import tempfile
import threading
import urllib.parse
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from .excel_processor import ExcelPrivacyProcessor
from .models import Finding


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


@dataclass
class Download:
    filename: str
    content: bytes


class WebState:
    def __init__(self) -> None:
        self.processor = ExcelPrivacyProcessor()
        self.source_path: Path | None = None
        self.findings: list[Finding] = []
        self.history: list[str] = []
        self.downloads: dict[str, Download] = {}

    def reset_input(self) -> None:
        self.processor.cleanup()
        self.source_path = None
        self.findings = []


STATE = WebState()


class ExcelPrivacyWebHandler(BaseHTTPRequestHandler):
    server_version = "ExcelPrivacyCleaner/0.1"

    def do_GET(self) -> None:
        route = urllib.parse.urlparse(self.path).path
        if route == "/":
            self._send_html(render_page())
            return
        if route.startswith("/download/"):
            self._download(route.removeprefix("/download/"))
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        route = urllib.parse.urlparse(self.path).path
        if route == "/scan":
            self._scan()
            return
        if route == "/convert":
            self._convert()
            return
        if route == "/clear-history":
            STATE.history.clear()
            self._send_html(render_page(status="履歴を消去しました。"))
            return
        if route == "/shutdown":
            self._send_html(render_shutdown_page())
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: object) -> None:
        return

    def _scan(self) -> None:
        upload_dir = Path(tempfile.mkdtemp(prefix="ExcelPrivacyCleanerUpload_"))
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                },
            )
            file_item = form["excel"] if "excel" in form else None
            if file_item is None or not getattr(file_item, "filename", ""):
                self._send_html(render_page(error="Excel ファイルを選択してください。"))
                return

            original_name = Path(file_item.filename).name
            suffix = Path(original_name).suffix.lower()
            if suffix not in SUPPORTED_EXTENSIONS:
                self._send_html(render_page(error="対応形式は .xlsx / .xlsm です。"))
                return

            uploaded_path = upload_dir / original_name
            with uploaded_path.open("wb") as output:
                shutil.copyfileobj(file_item.file, output)

            STATE.reset_input()
            STATE.source_path = uploaded_path
            STATE.findings = STATE.processor.scan(uploaded_path)
            self._send_html(render_page(status=f"検査完了: {len(STATE.findings)} 件を検出しました。"))
        except Exception as exc:
            STATE.processor.cleanup()
            self._send_html(render_page(error=f"検査エラー: {html.escape(str(exc))}"))
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)

    def _convert(self) -> None:
        if STATE.source_path is None:
            self._send_html(render_page(error="先に Excel ファイルを検査してください。"))
            return

        length = int(self.headers.get("Content-Length", "0") or "0")
        params = urllib.parse.parse_qs(self.rfile.read(length).decode("utf-8"))
        for index, finding in enumerate(STATE.findings):
            finding.enabled = f"enabled_{index}" in params
            replacement = params.get(f"replacement_{index}", [finding.replacement])[0].strip()
            if replacement:
                finding.replacement = replacement

        output_dir = Path(tempfile.mkdtemp(prefix="ExcelPrivacyCleanerOutput_"))
        try:
            output_path = STATE.processor.convert(STATE.source_path, STATE.findings, output_dir=output_dir)
            content = output_path.read_bytes()
            token = secrets.token_urlsafe(16)
            STATE.downloads[token] = Download(filename=output_path.name, content=content)
            converted_count = sum(1 for finding in STATE.findings if finding.enabled)
            STATE.history.insert(
                0,
                f"{datetime.now():%Y/%m/%d %H:%M:%S}  {converted_count} 件変換  {output_path.name}  一時ファイル削除済み",
            )
            self._send_html(render_page(status="変換が完了しました。", download_token=token))
        except Exception as exc:
            self._send_html(render_page(error=f"変換エラー: {html.escape(str(exc))}"))
        finally:
            shutil.rmtree(output_dir, ignore_errors=True)

    def _download(self, token: str) -> None:
        download = STATE.downloads.get(token)
        if download is None:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        encoded_name = urllib.parse.quote(download.filename)
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_name}")
        self.send_header("Content-Length", str(len(download.content)))
        self.end_headers()
        self.wfile.write(download.content)

    def _send_html(self, body: str) -> None:
        content = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)


def render_page(status: str = "", error: str = "", download_token: str | None = None) -> str:
    rows = []
    for index, finding in enumerate(STATE.findings):
        checked = "checked" if finding.enabled else ""
        rows.append(
            "<tr>"
            f"<td><input type='checkbox' name='enabled_{index}' {checked}></td>"
            f"<td>{escape(finding.sheet)}</td>"
            f"<td>{escape(finding.cell)}</td>"
            f"<td>{escape(finding.entity_type)}</td>"
            f"<td>{escape(finding.detection_kind)}</td>"
            f"<td>{escape(finding.original)}</td>"
            f"<td><input name='replacement_{index}' value='{escape_attr(finding.replacement)}'></td>"
            f"<td>{escape(finding.reason)}</td>"
            "</tr>"
        )

    history = "".join(f"<li>{escape(item)}</li>" for item in STATE.history[:20]) or "<li>まだ変換履歴はありません。</li>"
    source = escape(STATE.source_path.name) if STATE.source_path else "未選択"
    message = ""
    if status:
        message = f"<div class='status'>{escape(status)}</div>"
    if error:
        message = f"<div class='error'>{error}</div>"
    download = ""
    if download_token:
        download = f"<p><a class='button primary' href='/download/{download_token}'>匿名化済み Excel をダウンロード</a></p>"

    results = ""
    if STATE.findings:
        results = (
            "<form method='post' action='/convert'>"
            "<div class='toolbar'>"
            "<button type='button' onclick='setChecks(true)'>すべて変換</button>"
            "<button type='button' onclick='setChecks(false)'>すべて除外</button>"
            "<button class='primary' type='submit'>確認済みを変換</button>"
            "</div>"
            "<div class='table-wrap'><table><thead><tr>"
            "<th>変換</th><th>シート</th><th>セル</th><th>種類</th><th>検査</th><th>検出値</th><th>変換後</th><th>理由</th>"
            "</tr></thead><tbody>"
            + "".join(rows)
            + "</tbody></table></div></form>"
        )

    return f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>hoso Privacy Cleaner</title>
  <style>
    body {{ font-family: "Meiryo UI", "Yu Gothic UI", sans-serif; margin: 0; color: #1f2937; background: #f6f7f9; }}
    main {{ max-width: 1180px; margin: 0 auto; padding: 24px; }}
    h1 {{ font-size: 24px; margin: 0 0 8px; }}
    p {{ margin: 6px 0 16px; }}
    section {{ background: #fff; border: 1px solid #d8dee8; border-radius: 8px; padding: 16px; margin-bottom: 16px; }}
    input[type=file], input[type=text], input:not([type]) {{ font: inherit; }}
    input[name^=replacement_] {{ width: 120px; box-sizing: border-box; }}
    button, .button {{ border: 1px solid #aab4c2; background: #fff; color: #1f2937; padding: 8px 12px; border-radius: 6px; cursor: pointer; text-decoration: none; display: inline-block; }}
    .primary {{ background: #0f766e; border-color: #0f766e; color: white; }}
    .toolbar {{ display: flex; gap: 8px; margin-bottom: 12px; }}
    .status {{ background: #ecfdf5; border: 1px solid #a7f3d0; color: #065f46; padding: 10px; border-radius: 6px; margin-bottom: 12px; }}
    .error {{ background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; padding: 10px; border-radius: 6px; margin-bottom: 12px; }}
    .table-wrap {{ overflow: auto; max-height: 430px; border: 1px solid #d8dee8; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    th, td {{ border-bottom: 1px solid #e5e7eb; padding: 8px; text-align: left; vertical-align: top; white-space: nowrap; }}
    th {{ position: sticky; top: 0; background: #eef2f7; z-index: 1; }}
    td:nth-child(6), td:nth-child(8) {{ white-space: normal; min-width: 180px; }}
    .meta {{ color: #4b5563; }}
    ul {{ margin-bottom: 0; }}
  </style>
  <script>
    function setChecks(value) {{
      document.querySelectorAll("input[type=checkbox][name^=enabled_]").forEach(function(box) {{ box.checked = value; }});
    }}
  </script>
</head>
<body>
<main>
  <h1>hoso Privacy Cleaner</h1>
  <p class="meta">処理は localhost 上のこの PC 内だけで行います。外部クラウドには送信しません。原本は上書きしません。</p>
  {message}
  {download}
  <section>
    <form method="post" action="/scan" enctype="multipart/form-data">
      <p>現在のファイル: {source}</p>
      <input type="file" name="excel" accept=".xlsx,.xlsm" required>
      <button class="primary" type="submit">検査開始</button>
    </form>
  </section>
  <section>
    <h2>検出結果</h2>
    {results or "<p>Excel を投入すると検出結果がここに表示されます。</p>"}
  </section>
  <section>
    <h2>変換履歴</h2>
    <ul>{history}</ul>
    <form method="post" action="/clear-history" style="margin-top: 12px;"><button type="submit">履歴消去</button></form>
  </section>
  <section>
    <form method="post" action="/shutdown"><button type="submit">アプリを終了</button></form>
  </section>
</main>
</body>
</html>"""


def render_shutdown_page() -> str:
    return """<!doctype html><html lang="ja"><meta charset="utf-8"><title>終了</title><body><p>アプリを終了しました。このタブは閉じてかまいません。</p></body></html>"""


def escape(value: object) -> str:
    return html.escape(str(value), quote=False)


def escape_attr(value: object) -> str:
    return html.escape(str(value), quote=True)


def run(host: str = "127.0.0.1", port: int = 0) -> None:
    server = ThreadingHTTPServer((host, port), ExcelPrivacyWebHandler)
    actual_port = server.server_address[1]
    url = f"http://{host}:{actual_port}/"
    webbrowser.open(url)
    try:
        server.serve_forever()
    finally:
        STATE.processor.cleanup()
        server.server_close()

# -*- coding: utf-8 -*-
"""MIMO 工资评分生成器 Final V1.1
零 FastAPI 依赖版：新电脑只需要 Python，本包自带 Excel 依赖 vendor/openpyxl。
"""
import os
import sys
import json
import tempfile
import traceback
import webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote
from email.parser import BytesParser
from email.policy import default

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VENDOR_DIR = os.path.join(BASE_DIR, "vendor")
if os.path.isdir(VENDOR_DIR) and VENDOR_DIR not in sys.path:
    sys.path.insert(0, VENDOR_DIR)

from wage_core import build_preview_rows, generate_workbook, VERSION

FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
TMP_DIR = os.path.join(BASE_DIR, "runtime_uploads")
os.makedirs(TMP_DIR, exist_ok=True)


def _json_bytes(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def _safe_name(name, default="网约工工资明细_生成版.xlsx"):
    name = (name or default).replace("/", "_").replace("\\", "_").strip() or default
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


class Handler(BaseHTTPRequestHandler):
    server_version = "MIMO-WageScore/" + VERSION

    def log_message(self, fmt, *args):
        sys.stdout.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _send(self, status=200, body=b"", content_type="text/plain; charset=utf-8", headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = unquote(self.path.split("?", 1)[0])
        if path == "/api/health":
            return self._send(200, _json_bytes({"ok": True, "version": VERSION, "app": "MIMO 工资评分生成器", "mode": "html_stdlib"}), "application/json; charset=utf-8")
        if path == "/":
            path = "/index.html"
        if path.startswith("/static/"):
            rel = path[len("/static/"):]
        else:
            rel = path.lstrip("/")
        safe = os.path.normpath(rel)
        full = os.path.join(FRONTEND_DIR, safe)
        if not full.startswith(FRONTEND_DIR) or not os.path.exists(full) or os.path.isdir(full):
            return self._send(404, "文件不存在".encode("utf-8"), "text/plain; charset=utf-8")
        ext = os.path.splitext(full)[1].lower()
        ctype = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".png": "image/png",
        }.get(ext, "application/octet-stream")
        with open(full, "rb") as f:
            data = f.read()
        return self._send(200, data, ctype)

    def _parse_multipart(self):
        length = int(self.headers.get("Content-Length", "0"))
        ctype = self.headers.get("Content-Type", "")
        body = self.rfile.read(length)
        raw = ("Content-Type: " + ctype + "\r\nMIME-Version: 1.0\r\n\r\n").encode("utf-8") + body
        msg = BytesParser(policy=default).parsebytes(raw)
        fields = {}
        files = {}
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            data = part.get_payload(decode=True)
            if not name:
                continue
            if filename:
                files[name] = {"filename": filename, "data": data}
            else:
                fields[name] = data.decode("utf-8", errors="ignore") if data else ""
        return fields, files

    def do_POST(self):
        try:
            if self.path not in ["/api/wage-score/preview", "/api/wage-score/generate"]:
                return self._send(404, _json_bytes({"ok": False, "error": "接口不存在"}), "application/json; charset=utf-8")
            fields, files = self._parse_multipart()
            if "score_file" not in files:
                raise ValueError("请上传米墨评分 Excel。")
            strategy = fields.get("strategy") or "human_like"
            output_name = _safe_name(fields.get("output_name") or "网约工工资明细_生成版.xlsx")

            score_path = None
            settlement_path = None
            output_path = None
            try:
                score_fd, score_path = tempfile.mkstemp(prefix="score_", suffix=".xlsx", dir=TMP_DIR)
                with os.fdopen(score_fd, "wb") as f:
                    f.write(files["score_file"]["data"])
                # 兼容新旧字段名：settlement_file / process_file
                settlement_data = None
                if "settlement_file" in files:
                    settlement_data = files["settlement_file"]["data"]
                elif "process_file" in files:
                    settlement_data = files["process_file"]["data"]
                if settlement_data:
                    proc_fd, settlement_path = tempfile.mkstemp(prefix="settlement_", suffix=".xlsx", dir=TMP_DIR)
                    with os.fdopen(proc_fd, "wb") as f:
                        f.write(settlement_data)

                if self.path == "/api/wage-score/preview":
                    data = build_preview_rows(score_path, settlement_path, strategy=strategy)
                    data = {k: v for k, v in data.items() if k not in ("people", "settlement")}
                    data["ok"] = True
                    data["message"] = "预览展示全部导出数据；生成 Excel 会按同一批数据处理。"
                    return self._send(200, _json_bytes(data), "application/json; charset=utf-8")

                output_path = os.path.join(TMP_DIR, output_name)
                generate_workbook(score_path, output_path, settlement_path, strategy=strategy)
                with open(output_path, "rb") as f:
                    output = f.read()
                quoted = ''.join('%%%02X' % b for b in output_name.encode('utf-8'))
                headers = {"Content-Disposition": f"attachment; filename*=UTF-8''{quoted}"}
                return self._send(200, output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers)
            finally:
                for p in (score_path, settlement_path, output_path):
                    if p and os.path.exists(p):
                        try:
                            os.remove(p)
                        except Exception:
                            pass
        except Exception as e:
            traceback.print_exc()
            return self._send(400, _json_bytes({"ok": False, "error": str(e)}), "application/json; charset=utf-8")


def run():
    for port in range(5190, 5200):
        try:
            httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            url = f"http://127.0.0.1:{port}"
            print("=" * 60)
            print("MIMO 工资评分生成器", VERSION)
            print("已启动：", url)
            print("关闭软件：回到这个窗口，按 Ctrl + C")
            print("=" * 60)
            try:
                webbrowser.open(url)
            except Exception:
                pass
            httpd.serve_forever()
            return
        except OSError:
            continue
    print("5190-5199 端口都被占用，请关闭旧版工资评分生成器后重试。")


if __name__ == "__main__":
    run()

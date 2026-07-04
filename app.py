#!/usr/bin/env python3
from __future__ import annotations

import html
import subprocess
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

from fit_to_excel import (
    DEFAULT_OUTPUT_DIR,
    DROPDOWN_CONFIG_PATH,
    WORKBOOK_VERSION_NAME,
    create_workbook,
    load_dropdown_options,
)


ROOT = Path(__file__).resolve().parent
FIT_DIR = ROOT / "FIT"
HOST = "127.0.0.1"
PORT = 8765
DEFAULT_FIT_LIST_LIMIT = 30


def all_fit_files():
    FIT_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(FIT_DIR.glob("*.fit"), key=lambda path: path.stat().st_mtime, reverse=True)


def fit_files(selected_fit=""):
    files = all_fit_files()
    total_count = len(files)
    limit = DEFAULT_FIT_LIST_LIMIT
    limited = files[:limit]
    if selected_fit and not any(path.name == selected_fit for path in limited):
        selected_path = FIT_DIR / selected_fit
        if selected_path.exists() and selected_path.suffix.lower() == ".fit":
            limited.insert(0, selected_path)
    return limited, total_count, limit


def parse_number(value):
    value = (value or "").strip()
    if not value:
        return ""
    try:
        return float(value)
    except ValueError:
        return value


def first_value(form, key, default=""):
    return (form.get(key, [default])[0] or "").strip()


def content_disposition_params(value):
    params = {}
    for part in value.split(";"):
        part = part.strip()
        if "=" not in part:
            continue
        key, raw = part.split("=", 1)
        params[key.lower()] = raw.strip().strip('"')
    return params


def parse_multipart(body, content_type):
    marker = "boundary="
    if marker not in content_type:
        return {}, {}
    boundary = content_type.split(marker, 1)[1].split(";", 1)[0].strip().strip('"').encode()
    delimiter = b"--" + boundary
    form = {}
    files = {}

    for part in body.split(delimiter):
        if not part or part in (b"--\r\n", b"--"):
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        if part.endswith(b"--\r\n"):
            part = part[:-4]
        elif part.endswith(b"--"):
            part = part[:-2]
        if part.endswith(b"\r\n"):
            part = part[:-2]
        if b"\r\n\r\n" not in part:
            continue

        raw_headers, data = part.split(b"\r\n\r\n", 1)
        headers = {}
        for line in raw_headers.decode("utf-8", "replace").split("\r\n"):
            if ":" in line:
                key, value = line.split(":", 1)
                headers[key.lower()] = value.strip()
        disposition = headers.get("content-disposition", "")
        params = content_disposition_params(disposition)
        name = params.get("name")
        if not name:
            continue
        filename = params.get("filename")
        if filename:
            files[name] = {"filename": Path(filename).name, "content": data}
        else:
            form.setdefault(name, []).append(data.decode("utf-8", "replace"))
    return form, files


def parse_post_data(headers, body):
    content_type = headers.get("Content-Type", "")
    if content_type.startswith("multipart/form-data"):
        return parse_multipart(body, content_type)
    return parse_qs(body.decode("utf-8")), {}


def save_uploaded_fit(upload):
    if not upload:
        return None
    filename = upload.get("filename") or ""
    content = upload.get("content") or b""
    if not filename or not content:
        return None
    if Path(filename).suffix.lower() != ".fit":
        raise ValueError("上傳檔案必須是 .fit。")
    FIT_DIR.mkdir(parents=True, exist_ok=True)
    target = FIT_DIR / Path(filename).name
    target.write_bytes(content)
    return target


def build_metadata(form):
    return {
        "shoe": first_value(form, "shoe"),
        "weather_temp": parse_number(first_value(form, "weather_temp")),
        "humidity": parse_number(first_value(form, "humidity")),
        "wind_direction": first_value(form, "wind_direction"),
        "wind_speed": first_value(form, "wind_speed"),
        "workout_type": first_value(form, "workout_type"),
        "training_focus": first_value(form, "training_focus"),
        "rpe": first_value(form, "rpe"),
        "fueling": first_value(form, "fueling"),
        "max_hr": parse_number(first_value(form, "max_hr")),
        "critical_power": parse_number(first_value(form, "critical_power")),
        "training_effect_aerobic": parse_number(first_value(form, "training_effect_aerobic")),
        "training_effect_anaerobic": parse_number(first_value(form, "training_effect_anaerobic")),
        "training_load": parse_number(first_value(form, "training_load")),
        "recovery_time_hr": parse_number(first_value(form, "recovery_time_hr")),
        "notes": first_value(form, "notes"),
    }


def option_tags(options, selected=""):
    tags = ['<option value="">自動 / 留空</option>']
    for option in options:
        value = html.escape(str(option), quote=True)
        is_selected = " selected" if str(option) == selected else ""
        tags.append(f'<option value="{value}"{is_selected}>{html.escape(str(option))}</option>')
    return "\n".join(tags)


def input_field(label, name, value="", input_type="text", placeholder=""):
    return f"""
      <label>
        <span>{html.escape(label)}</span>
        <input type="{input_type}" name="{html.escape(name)}" value="{html.escape(str(value or ''), quote=True)}" placeholder="{html.escape(placeholder, quote=True)}">
      </label>
    """


def render_page(message="", error="", selected_fit=""):
    dropdown_options = load_dropdown_options(DROPDOWN_CONFIG_PATH)
    files, total_count, list_limit = fit_files(selected_fit)
    fit_options = []
    for path in files:
        value = html.escape(path.name, quote=True)
        is_selected = " selected" if path.name == selected_fit else ""
        fit_options.append(f'<option value="{value}"{is_selected}>{html.escape(path.name)}</option>')
    if not fit_options:
        fit_options.append('<option value="">FIT 資料夾目前沒有 .fit 檔</option>')

    status = ""
    if message:
        status = f'<section class="status ok">{message}</section>'
    if error:
        status = f'<section class="status error">{html.escape(error)}</section>'

    list_note = f"FIT 資料夾共有 {total_count} 個檔案，目前清單只顯示最近 {min(total_count, list_limit)} 個。"

    return f"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>跑步分析資料轉檔</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #1f2933;
      --muted: #65758b;
      --line: #d9e2ec;
      --accent: #1f6f8b;
      --accent-dark: #14506a;
      --surface: #ffffff;
      --page: #f4f7fa;
      --error: #b42318;
      --ok: #166534;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", sans-serif;
      background: var(--page);
      color: var(--ink);
    }}
    main {{
      width: min(1040px, calc(100vw - 32px));
      margin: 32px auto;
    }}
    h1 {{
      margin: 0 0 6px;
      font-size: 28px;
      letter-spacing: 0;
    }}
    .subtitle {{
      margin: 0 0 24px;
      color: var(--muted);
      font-size: 15px;
    }}
    form {{
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 22px;
      box-shadow: 0 8px 24px rgba(31, 41, 51, 0.07);
    }}
    fieldset {{
      border: 0;
      padding: 0;
      margin: 0 0 24px;
    }}
    legend {{
      padding: 0;
      margin: 0 0 14px;
      font-size: 17px;
      font-weight: 700;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
    }}
    label {{
      display: flex;
      flex-direction: column;
      gap: 6px;
      min-width: 0;
    }}
    label.wide {{ grid-column: span 3; }}
    span {{
      font-size: 13px;
      color: var(--muted);
    }}
    input, select, textarea {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 11px;
      font: inherit;
      color: var(--ink);
      background: #fff;
    }}
    textarea {{
      min-height: 76px;
      resize: vertical;
    }}
    .inline {{
      display: flex;
      align-items: center;
      gap: 10px;
      color: var(--ink);
      margin-top: 2px;
    }}
    .inline input {{
      width: auto;
    }}
    .actions {{
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 4px;
    }}
    button, .button {{
      appearance: none;
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: #fff;
      padding: 11px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-block;
    }}
    button:hover, .button:hover {{
      background: var(--accent-dark);
    }}
    .secondary {{
      background: #e7eef5;
      color: var(--ink);
    }}
    .secondary:hover {{
      background: #d7e3ee;
    }}
    .status {{
      border-radius: 8px;
      padding: 14px 16px;
      margin: 0 0 18px;
      border: 1px solid var(--line);
      background: #fff;
      line-height: 1.55;
    }}
    .ok {{ color: var(--ok); }}
    .error {{ color: var(--error); }}
    .note {{
      color: var(--muted);
      font-size: 13px;
      margin: 8px 0 0;
    }}
    code {{
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      font-size: 13px;
      color: var(--ink);
    }}
    @media (max-width: 760px) {{
      main {{ width: min(100vw - 20px, 1040px); margin: 18px auto; }}
      form {{ padding: 16px; }}
      .grid {{ grid-template-columns: 1fr; }}
      label.wide {{ grid-column: span 1; }}
      .actions {{ flex-direction: column; align-items: stretch; }}
      button, .button {{ text-align: center; }}
    }}
  </style>
</head>
<body>
  <main>
    <h1>跑步分析資料轉檔</h1>
    <p class="subtitle">選擇 Garmin FIT 檔，產生固定格式 Excel。最大心率、Critical Power、Training Effect 與天氣會盡量自動帶入。</p>
    {status}
    <form method="post" action="/convert" enctype="multipart/form-data">
      <fieldset>
        <legend>檔案</legend>
        <div class="grid">
          <label class="wide">
            <span>從電腦選擇 FIT 檔</span>
            <input type="file" name="upload_fit" accept=".fit">
          </label>
          <label class="wide">
            <span>或使用 FIT 資料夾裡的檔案</span>
            <select name="fit_file">
              {"".join(fit_options)}
            </select>
            <p class="note">{html.escape(list_note)}</p>
          </label>
          <label class="wide">
            <span>輸出檔名，可留空</span>
            <input name="output_name" placeholder="{html.escape(WORKBOOK_VERSION_NAME)}_活動檔名.xlsx">
          </label>
          <label class="inline wide">
            <input type="checkbox" name="fetch_weather" value="1" checked>
            <span>自動抓 Open-Meteo 歷史天氣</span>
          </label>
        </div>
      </fieldset>

      <fieldset>
        <legend>活動資訊</legend>
        <div class="grid">
          <label>
            <span>鞋款</span>
            <select name="shoe">{option_tags(dropdown_options["shoes"])}</select>
          </label>
          <label>
            <span>課表類型</span>
            <select name="workout_type">{option_tags(dropdown_options["workout_types"])}</select>
          </label>
          <label>
            <span>訓練目的</span>
            <select name="training_focus">{option_tags(dropdown_options["training_focus"])}</select>
          </label>
          <label>
            <span>Garmin 主觀感受</span>
            <select name="rpe">{option_tags(dropdown_options["garmin_rpe"])}</select>
          </label>
          {input_field("最大心率", "max_hr", input_type="number", placeholder="自動")}
          {input_field("Critical Power(W)", "critical_power", input_type="number", placeholder="自動")}
          {input_field("氣溫(°C)", "weather_temp", input_type="number", placeholder="自動")}
          {input_field("濕度(%)", "humidity", input_type="number", placeholder="自動")}
          {input_field("風向", "wind_direction", placeholder="自動")}
          {input_field("風速", "wind_speed", placeholder="自動")}
          {input_field("Recovery Time (hr)", "recovery_time_hr", input_type="number")}
          {input_field("Training Load", "training_load", input_type="number", placeholder="自動")}
          <label class="wide">
            <span>補給紀錄</span>
            <textarea name="fueling"></textarea>
          </label>
          <label class="wide">
            <span>備註</span>
            <textarea name="notes"></textarea>
          </label>
        </div>
      </fieldset>

      <div class="actions">
        <button type="submit">轉成 Excel</button>
        <a class="button secondary" href="/">重新整理</a>
      </div>
    </form>
  </main>
</body>
</html>"""


class AppHandler(BaseHTTPRequestHandler):
    def send_html(self, content, status=200):
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        if parsed.path == "/open":
            output = Path(first_value(query, "path"))
            if output.exists() and output.is_file():
                subprocess.run(["open", str(output)], check=False)
                self.send_html(render_page(message=f"已要求系統開啟 <code>{html.escape(str(output))}</code>"))
            else:
                self.send_html(render_page(error="找不到輸出檔。"), status=404)
            return
        self.send_html(render_page())

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/convert":
            self.send_html(render_page(error="不支援的操作。"), status=404)
            return

        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        try:
            form, files = parse_post_data(self.headers, body)
            fit_path = save_uploaded_fit(files.get("upload_fit"))
        except ValueError as error:
            self.send_html(render_page(error=str(error)), status=400)
            return

        fit_name = first_value(form, "fit_file")
        if fit_path is None:
            fit_path = FIT_DIR / fit_name
        if not fit_path.exists() or fit_path.suffix.lower() != ".fit":
            self.send_html(render_page(error="請選擇一個 .fit 檔，或從 FIT 資料夾清單選擇有效檔案。"), status=400)
            return
        fit_name = fit_path.name

        output_name = first_value(form, "output_name")
        output_path = DEFAULT_OUTPUT_DIR / output_name if output_name else DEFAULT_OUTPUT_DIR / f"{WORKBOOK_VERSION_NAME}_{fit_path.stem}.xlsx"
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        metadata = build_metadata(form)
        fetch_weather = first_value(form, "fetch_weather") == "1"

        try:
            saved = create_workbook(fit_path, output_path, metadata=metadata, fetch_weather=fetch_weather)
        except Exception as error:
            self.send_html(render_page(error=f"轉檔失敗：{error}", selected_fit=fit_name), status=500)
            return

        open_link = "/open?" + urlencode({"path": str(saved)})
        message = (
            f"轉檔完成：<code>{html.escape(str(saved))}</code><br>"
            f'<a class="button" href="{html.escape(open_link, quote=True)}">開啟 Excel</a>'
        )
        self.send_html(render_page(message=message, selected_fit=fit_name))

    def log_message(self, format, *args):
        return


def open_browser_later(url):
    timer = threading.Timer(0.6, lambda: webbrowser.open(url))
    timer.daemon = True
    timer.start()


def main():
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    url = f"http://{HOST}:{PORT}"
    print(f"Running Analytics app: {url}")
    open_browser_later(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()

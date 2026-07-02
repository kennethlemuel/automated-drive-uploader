#!/usr/bin/env python3
import json
import os
import shutil
import sys
from pathlib import Path

import webview

from excel_drive_links import auth_drive, convert, upload_image, upload_workbook


APP_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def app_data_dir():
    if not getattr(sys, "frozen", False):
        return APP_DIR
    root = Path(os.getenv("APPDATA", Path.home() / "AppData/Roaming")) if sys.platform == "win32" else Path.home() / "Library/Application Support"
    return root / "Drive Uploader"  # Preserve existing OAuth tokens after the app rename.


DATA_DIR = app_data_dir()
DATA_DIR.mkdir(parents=True, exist_ok=True)


class Api:
    def __init__(self):
        self.window = None

    def initial_state(self):
        credentials = DATA_DIR / "credentials.json"
        token = DATA_DIR / ".google-drive-token.json"
        return {
            "credentials": str(credentials) if credentials.exists() else "",
            "tokenReady": token.exists(),
        }

    def choose_credentials(self):
        path = self.pick_file("Credentials JSON", ("JSON files (*.json)",))
        if not path:
            return ""
        destination = DATA_DIR / "credentials.json"
        if Path(path).resolve() != destination.resolve():
            shutil.copy2(path, destination)
        return str(destination)

    def choose_excel(self):
        path = self.pick_file("Excel workbook", ("Excel files (*.xlsx)",))
        if not path:
            return {"path": "", "output": ""}
        workbook = Path(path)
        return {"path": path, "output": str(workbook.with_name(f"{workbook.stem}-with-links.xlsx"))}

    def choose_screenshots(self):
        return dialog_path(self.window.create_file_dialog(webview.FileDialog.FOLDER))

    def choose_output(self, current=""):
        directory = str(Path(current).parent) if current else str(Path.home() / "Documents")
        name = Path(current).name if current else "workbook-with-links.xlsx"
        path = dialog_path(self.window.create_file_dialog(webview.FileDialog.SAVE, directory=directory, save_filename=name, file_types=("Excel files (*.xlsx)",)))
        if not path or Path(path).is_dir():
            return current or ""
        return path if path.lower().endswith(".xlsx") else f"{path}.xlsx"

    def run_upload(self, payload):
        input_path = Path(payload.get("excel") or "")
        output_path = Path(payload.get("output") or input_path.with_name(f"{input_path.stem}-with-links.xlsx"))
        credentials = Path(payload.get("credentials") or "")
        image_root = payload.get("screenshots") or None
        folder_id = payload.get("folderId") or None
        public = bool(payload.get("publicLinks"))
        dry_run = bool(payload.get("dryRun"))
        upload_sheet = bool(payload.get("uploadSheet"))

        if not input_path.exists() or input_path.suffix.lower() != ".xlsx":
            return {"ok": False, "error": "Choose a valid .xlsx workbook."}
        if not dry_run and not credentials.exists():
            return {"ok": False, "error": "Choose credentials.json or enable dry run."}

        logs = [f"Input: {input_path}", f"Output: {output_path}"]
        try:
            self.progress(3, "Preparing upload")
            if dry_run:
                upload = lambda _data, name: f"dry-run://{name}"
                service = None
            else:
                logs.append("Authorizing Google Drive...")
                self.progress(8, "Authorizing Google Drive")
                service = auth_drive(credentials, DATA_DIR / ".google-drive-token.json")
                upload = lambda data, name: upload_image(service, data, name, folder_id, public)

            def on_progress(done, total, name):
                percent = 10 + int((done / max(total, 1)) * 80)
                self.progress(percent, f"Uploading {name}")

            uploaded, missing = convert(input_path, output_path, upload, image_root, on_progress)
            logs.append(f"Image uploads: {uploaded}")
            sheet_url = ""
            if upload_sheet and not dry_run:
                logs.append("Uploading finished workbook as Google Sheet...")
                self.progress(94, "Publishing Google Sheet")
                sheet_url = upload_workbook(service, output_path, folder_id, public, as_sheets=True)
                logs.append(sheet_url)
            if missing:
                logs.append(f"Missing local image files: {len(missing)}")
                logs.extend(missing[:8])
            if uploaded == 0:
                logs.append("No embedded images or existing local image paths were found.")

            return {
                "ok": True,
                "output": str(output_path),
                "workbookName": output_path.stem,
                "uploaded": uploaded,
                "missing": len(missing),
                "sheetUrl": sheet_url,
                "logs": logs,
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc), "logs": logs}

    def pick_file(self, _title, file_types):
        return dialog_path(self.window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=False, file_types=file_types))

    def progress(self, percent, message):
        if self.window:
            self.window.evaluate_js(f"window.uploadProgress({json.dumps({'percent': percent, 'message': message})})")


def dialog_path(result):
    if not result:
        return ""
    if isinstance(result, (str, Path)):
        return str(result)
    return str(result[0]) if result else ""


def main():
    api = Api()
    html = (APP_DIR / "web" / "index.html").read_text(encoding="utf-8")
    window = webview.create_window("Efficient Audit Uploader", html=html, js_api=api, width=550, height=850, min_size=(550, 850))
    api.window = window
    webview.start(debug=False)


if __name__ == "__main__":
    main()

from __future__ import annotations

from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

from excel_export import build_workbook

try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    Credentials = None
    build = None
    MediaFileUpload = None
    MediaIoBaseDownload = None


APP_TITLE = "Khảo sát du lịch Meiwa"
APP_TITLE_JP = "Meiwa社員旅行アンケート"
APP_VERSION = "v1.6.0"
APP_AUTHOR = "Nguyen Duy Hoa"

DEPARTMENTS = ["GA", "CR", "CS", "CD", "PT", "QA", "MOLD"]
EXPORT_FILE_NAME = "Form khao sat Du lich Cong ty meiwa nam 2026.xlsx"
EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

JOIN_LABELS = {
    "": "Chọn xác nhận / 参加確認を選択",
    "Co": "Có / 参加",
    "Khong": "Không / 不参加",
}
DESTINATION_LABELS = {
    "": "Chọn địa điểm / 行き先を選択",
    "Nha Trang": "Nha Trang / ニャチャン",
    "Da Lat": "Đà Lạt / ダラット",
}


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🧳",
    layout="wide",
)


def normalize_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_spaces(value: object) -> str:
    return " ".join(normalize_text(value).split())


def normalize_name(value: object) -> str:
    raw = normalize_spaces(value)
    if not raw:
        return ""
    return " ".join(part[:1].upper() + part[1:].lower() for part in raw.split(" "))


def normalize_department(value: object) -> str:
    raw = normalize_text(value).upper().replace(".", "").replace(" ", "")
    return raw if raw in DEPARTMENTS else raw


def normalize_join_value(value: object) -> str:
    raw = normalize_text(value).lower()
    if raw in {"co", "có", "yes", "y", "1", "true", "v", "x"}:
        return "Co"
    if raw in {"khong", "không", "no", "n", "0", "false"}:
        return "Khong"
    return ""


def normalize_destination_value(value: object) -> str:
    raw = normalize_text(value).lower()
    if raw in {"nha trang", "nhatrang"}:
        return "Nha Trang"
    if raw in {"da lat", "đà lạt", "dalat"}:
        return "Da Lat"
    return ""


def is_joining(value: object) -> bool:
    return normalize_join_value(value) == "Co"


def clean_record(record: dict[str, Any]) -> dict[str, str]:
    join_value = normalize_join_value(record.get("tham_gia"))
    destination = normalize_destination_value(record.get("dia_diem")) if join_value == "Co" else ""
    return {
        "msnv": normalize_spaces(record.get("msnv")),
        "ho_ten": normalize_name(record.get("ho_ten")),
        "bo_phan": normalize_department(record.get("bo_phan")),
        "cong_doan": normalize_spaces(record.get("cong_doan")),
        "tham_gia": join_value,
        "dia_diem": destination,
    }


def init_state() -> None:
    if "bootstrapped" not in st.session_state:
        st.session_state.bootstrapped = False
    if "records" not in st.session_state:
        st.session_state.records = []
    if "last_saved_path" not in st.session_state:
        st.session_state.last_saved_path = ""
    if "last_saved_at" not in st.session_state:
        st.session_state.last_saved_at = ""
    if "last_drive_url" not in st.session_state:
        st.session_state.last_drive_url = ""
    if "save_message" not in st.session_state:
        st.session_state.save_message = ""
    if "save_message_kind" not in st.session_state:
        st.session_state.save_message_kind = "info"


def get_export_path() -> Path:
    data_dir = os.getenv("DATA_DIR", "")
    export_dir = Path(data_dir) if data_dir else Path(__file__).with_name("exports")
    export_dir.mkdir(parents=True, exist_ok=True)
    return export_dir / EXPORT_FILE_NAME


def cleanup_old_export_files(export_dir: Path, keep_name: str) -> None:
    for file_path in export_dir.glob("*.xlsx"):
        if file_path.name != keep_name:
            try:
                file_path.unlink()
            except OSError:
                pass


def build_export_bytes(records: list[dict[str, str]]) -> bytes:
    workbook = build_workbook(records)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def save_export_file(records: list[dict[str, str]]) -> Path:
    export_path = get_export_path()
    cleanup_old_export_files(export_path.parent, EXPORT_FILE_NAME)
    workbook = build_workbook(records)
    workbook.save(export_path)
    return export_path


def load_existing_records() -> list[dict[str, str]]:
    export_path = get_export_path()
    if not export_path.exists():
        return []

    workbook = load_workbook(export_path, data_only=False)
    sheet = workbook.active
    loaded: list[dict[str, str]] = []

    for row_num in range(9, sheet.max_row + 1):
        msnv = normalize_text(sheet.cell(row_num, 2).value)
        if msnv.upper() == "TOTAL":
            break

        record = clean_record(
            {
                "msnv": msnv,
                "ho_ten": sheet.cell(row_num, 3).value,
                "bo_phan": sheet.cell(row_num, 4).value,
                "cong_doan": sheet.cell(row_num, 5).value,
                "tham_gia": "Co" if normalize_text(sheet.cell(row_num, 6).value).lower() == "v" else "Khong",
                "dia_diem": (
                    "Nha Trang"
                    if normalize_text(sheet.cell(row_num, 8).value).lower() == "v"
                    else "Da Lat"
                    if normalize_text(sheet.cell(row_num, 9).value).lower() == "v"
                    else ""
                ),
            }
        )
        if any(record.values()):
            loaded.append(record)

    return loaded


def read_google_config_from_env() -> tuple[dict[str, str], dict[str, Any]]:
    drive_settings = {
        "folder_id": os.getenv("GOOGLE_DRIVE_FOLDER_ID", ""),
        "shared_url": os.getenv("GOOGLE_DRIVE_SHARED_URL", ""),
    }
    service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account: dict[str, Any] = {}
    if service_account_json:
        try:
            service_account = json.loads(service_account_json)
        except json.JSONDecodeError:
            service_account = {}
    return drive_settings, service_account


def get_google_drive_settings() -> dict[str, str]:
    drive_settings: dict[str, Any] = {}
    service_account: dict[str, Any] = {}

    try:
        if "google_drive" in st.secrets:
            drive_settings = dict(st.secrets["google_drive"])
        if "google_service_account" in st.secrets:
            service_account = dict(st.secrets["google_service_account"])
    except Exception:
        drive_settings, service_account = read_google_config_from_env()
    else:
        if not drive_settings and not service_account:
            drive_settings, service_account = read_google_config_from_env()

    folder_id = normalize_text(drive_settings.get("folder_id", ""))
    shared_url = normalize_text(drive_settings.get("shared_url", ""))
    enabled = bool(folder_id and service_account)
    return {
        "folder_id": folder_id,
        "shared_url": shared_url,
        "enabled": "1" if enabled else "0",
    }


@st.cache_resource(show_spinner=False)
def get_google_drive_service():
    if Credentials is None or build is None:
        return None

    service_account_info: dict[str, Any] | None = None
    try:
        if "google_service_account" in st.secrets:
            service_account_info = dict(st.secrets["google_service_account"])
    except Exception:
        service_account_info = None

    if not service_account_info:
        _, env_service_account = read_google_config_from_env()
        service_account_info = env_service_account or None

    if not service_account_info:
        return None

    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=credentials)


def sync_export_to_google_drive(local_path: Path) -> dict[str, str]:
    settings = get_google_drive_settings()
    if settings["enabled"] != "1":
        return {"status": "not_configured", "url": settings["shared_url"]}

    service = get_google_drive_service()
    if service is None or MediaFileUpload is None:
        return {"status": "missing_library", "url": settings["shared_url"]}

    file_name = local_path.name.replace("'", "\\'")
    query = (
        f"name = '{file_name}' and "
        f"'{settings['folder_id']}' in parents and trashed = false"
    )
    existing = (
        service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, webViewLink, webContentLink)",
            pageSize=1,
        )
        .execute()
        .get("files", [])
    )

    media = MediaFileUpload(str(local_path), mimetype=EXCEL_MIME_TYPE, resumable=False)

    if existing:
        file_id = existing[0]["id"]
        updated = (
            service.files()
            .update(fileId=file_id, media_body=media, fields="id, webViewLink, webContentLink")
            .execute()
        )
        return {
            "status": "updated",
            "url": updated.get("webViewLink") or updated.get("webContentLink") or settings["shared_url"],
        }

    created = (
        service.files()
        .create(
            body={"name": local_path.name, "parents": [settings["folder_id"]]},
            media_body=media,
            fields="id, webViewLink, webContentLink",
        )
        .execute()
    )
    return {
        "status": "created",
        "url": created.get("webViewLink") or created.get("webContentLink") or settings["shared_url"],
    }


def ensure_local_export_from_google_drive() -> dict[str, str]:
    export_path = get_export_path()
    if export_path.exists():
        return {"status": "local_exists", "path": str(export_path)}

    settings = get_google_drive_settings()
    if settings["enabled"] != "1":
        return {"status": "not_configured", "path": str(export_path)}

    service = get_google_drive_service()
    if service is None or MediaIoBaseDownload is None:
        return {"status": "missing_library", "path": str(export_path)}

    file_name = export_path.name.replace("'", "\\'")
    query = (
        f"name = '{file_name}' and "
        f"'{settings['folder_id']}' in parents and trashed = false"
    )
    existing = (
        service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id, name)",
            pageSize=1,
        )
        .execute()
        .get("files", [])
    )
    if not existing:
        return {"status": "drive_file_missing", "path": str(export_path)}

    request = service.files().get_media(fileId=existing[0]["id"])
    export_path.parent.mkdir(parents=True, exist_ok=True)
    with export_path.open("wb") as file_obj:
        downloader = MediaIoBaseDownload(file_obj, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    return {"status": "downloaded", "path": str(export_path)}


def bootstrap_records() -> None:
    if st.session_state.bootstrapped:
        return

    settings = get_google_drive_settings()
    st.session_state.last_drive_url = settings["shared_url"]
    ensure_local_export_from_google_drive()
    st.session_state.records = dedupe_records(load_existing_records())
    st.session_state.bootstrapped = True


def dedupe_records(records: list[dict[str, str]]) -> list[dict[str, str]]:
    ordered: dict[str, dict[str, str]] = {}
    for item in records:
        clean = clean_record(item)
        key = clean["msnv"] or f"__row__{len(ordered)}"
        ordered[key] = clean
    return list(ordered.values())


def upsert_record(records: list[dict[str, str]], record: dict[str, str]) -> list[dict[str, str]]:
    record = clean_record(record)
    key = record["msnv"]
    if not key:
        return records

    updated = list(records)
    for index, existing in enumerate(updated):
        if normalize_text(existing.get("msnv")) == key:
            updated[index] = record
            return updated
    updated.append(record)
    return updated


def build_records_frame(records: list[dict[str, str]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(
            columns=[
                "MSNV",
                "Họ tên",
                "Bộ phận",
                "Công đoạn",
                "Tham gia",
                "Địa điểm",
            ]
        )

    rows = []
    for item in records:
        rows.append(
            {
                "MSNV": item["msnv"],
                "Họ tên": item["ho_ten"],
                "Bộ phận": item["bo_phan"],
                "Công đoạn": item["cong_doan"],
                "Tham gia": "Có" if item["tham_gia"] == "Co" else "Không",
                "Địa điểm": "Nha Trang" if item["dia_diem"] == "Nha Trang" else "Đà Lạt" if item["dia_diem"] == "Da Lat" else "",
            }
        )
    return pd.DataFrame(rows)


def dashboard_metrics(records: list[dict[str, str]]) -> dict[str, int]:
    total = len(records)
    joined = sum(1 for item in records if item["tham_gia"] == "Co")
    not_joined = sum(1 for item in records if item["tham_gia"] == "Khong")
    nha_trang = sum(1 for item in records if item["tham_gia"] == "Co" and item["dia_diem"] == "Nha Trang")
    da_lat = sum(1 for item in records if item["tham_gia"] == "Co" and item["dia_diem"] == "Da Lat")
    return {
        "total": total,
        "joined": joined,
        "not_joined": not_joined,
        "nha_trang": nha_trang,
        "da_lat": da_lat,
    }


def build_overview_table(records: list[dict[str, str]]) -> pd.DataFrame:
    stats = dashboard_metrics(records)
    total = stats["total"]
    joined = stats["joined"]
    rows = [
        {
            "Hạng mục": "Tổng phiếu",
            "日本語": "総票数",
            "Số lượng": total,
            "Tỷ lệ": "100.0%" if total else "0.0%",
        },
        {
            "Hạng mục": "Tham gia",
            "日本語": "参加",
            "Số lượng": joined,
            "Tỷ lệ": f"{(joined / total * 100):.1f}%" if total else "0.0%",
        },
        {
            "Hạng mục": "Không tham gia",
            "日本語": "不参加",
            "Số lượng": stats["not_joined"],
            "Tỷ lệ": f"{(stats['not_joined'] / total * 100):.1f}%" if total else "0.0%",
        },
        {
            "Hạng mục": "Nha Trang",
            "日本語": "ニャチャン",
            "Số lượng": stats["nha_trang"],
            "Tỷ lệ": f"{(stats['nha_trang'] / joined * 100):.1f}%" if joined else "0.0%",
        },
        {
            "Hạng mục": "Đà Lạt",
            "日本語": "ダラット",
            "Số lượng": stats["da_lat"],
            "Tỷ lệ": f"{(stats['da_lat'] / joined * 100):.1f}%" if joined else "0.0%",
        },
    ]
    return pd.DataFrame(rows)


def build_overview_chart_frame(records: list[dict[str, str]]) -> pd.DataFrame:
    stats = dashboard_metrics(records)
    return pd.DataFrame(
        [
            {"Nhóm": "Tham gia", "Số lượng": stats["joined"]},
            {"Nhóm": "Không tham gia", "Số lượng": stats["not_joined"]},
            {"Nhóm": "Nha Trang", "Số lượng": stats["nha_trang"]},
            {"Nhóm": "Đà Lạt", "Số lượng": stats["da_lat"]},
        ]
    ).set_index("Nhóm")


def build_department_summary(records: list[dict[str, str]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    totals = {
        "Bộ phận": "TỔNG CỘNG",
        "日本語": "合計",
        "Tổng phiếu": 0,
        "Tham gia": 0,
        "Không tham gia": 0,
        "Nha Trang": 0,
        "Đà Lạt": 0,
    }
    extras: dict[str, dict[str, Any]] = {}

    for department in DEPARTMENTS:
        dept_records = [item for item in records if item["bo_phan"] == department]
        if not dept_records:
            continue
        stats = dashboard_metrics(dept_records)
        row = {
            "Bộ phận": department,
            "日本語": f"{department} 部門",
            "Tổng phiếu": stats["total"],
            "Tham gia": stats["joined"],
            "Không tham gia": stats["not_joined"],
            "Nha Trang": stats["nha_trang"],
            "Đà Lạt": stats["da_lat"],
        }
        rows.append(row)
        for key in ["Tổng phiếu", "Tham gia", "Không tham gia", "Nha Trang", "Đà Lạt"]:
            totals[key] += row[key]

    for item in records:
        department = item["bo_phan"]
        if department in DEPARTMENTS:
            continue

        bucket = department or "Khác/Chưa rõ"
        if bucket not in extras:
            extras[bucket] = {
                "Bộ phận": bucket,
                "日本語": "未分類",
                "Tổng phiếu": 0,
                "Tham gia": 0,
                "Không tham gia": 0,
                "Nha Trang": 0,
                "Đà Lạt": 0,
            }

        extras[bucket]["Tổng phiếu"] += 1
        if item["tham_gia"] == "Co":
            extras[bucket]["Tham gia"] += 1
            if item["dia_diem"] == "Nha Trang":
                extras[bucket]["Nha Trang"] += 1
            elif item["dia_diem"] == "Da Lat":
                extras[bucket]["Đà Lạt"] += 1
        else:
            extras[bucket]["Không tham gia"] += 1

    for extra in extras.values():
        rows.append(extra)
        for key in ["Tổng phiếu", "Tham gia", "Không tham gia", "Nha Trang", "Đà Lạt"]:
            totals[key] += extra[key]

    if not rows:
        return pd.DataFrame(
            columns=["Bộ phận", "日本語", "Tổng phiếu", "Tham gia", "Không tham gia", "Nha Trang", "Đà Lạt"]
        )

    rows.append(totals)
    return pd.DataFrame(rows)


def build_department_chart_frame(records: list[dict[str, str]]) -> pd.DataFrame:
    summary = build_department_summary(records)
    if summary.empty:
        return pd.DataFrame(columns=["Tham gia", "Không tham gia", "Nha Trang", "Đà Lạt"])
    detail = summary[summary["Bộ phận"] != "TỔNG CỘNG"].copy()
    detail = detail.set_index("Bộ phận")[["Tham gia", "Không tham gia", "Nha Trang", "Đà Lạt"]]
    return detail


def render_close_warning(should_warn: bool) -> None:
    components.html(
        f"""
        <script>
        window.onbeforeunload = {str(should_warn).lower()}
          ? function(event) {{
              event.preventDefault();
              event.returnValue = "";
              return "";
            }}
          : null;
        </script>
        """,
        height=0,
    )


def show_save_message() -> None:
    message = st.session_state.save_message
    if not message:
        return

    kind = st.session_state.save_message_kind
    if kind == "success":
        st.success(message)
    elif kind == "warning":
        st.warning(message)
    else:
        st.info(message)


def set_save_message(message: str, kind: str = "info") -> None:
    st.session_state.save_message = message
    st.session_state.save_message_kind = kind


def render_hero() -> None:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at top left, rgba(255, 244, 231, 0.85), transparent 35%),
                radial-gradient(circle at top right, rgba(233, 239, 255, 0.95), transparent 35%),
                #fcfcfd;
        }}
        .block-container {{
            padding-top: 1.25rem;
            padding-bottom: 3rem;
            max-width: 1180px;
        }}
        .hero-card {{
            display: flex;
            gap: 1rem;
            justify-content: space-between;
            align-items: flex-start;
            padding: 1.35rem 1.4rem;
            border: 1px solid rgba(35, 43, 62, 0.10);
            border-radius: 28px;
            background: linear-gradient(135deg, rgba(255, 248, 241, 0.92), rgba(239, 241, 255, 0.94));
            box-shadow: 0 18px 40px rgba(31, 41, 55, 0.08);
            margin-bottom: 1rem;
        }}
        .hero-copy {{
            flex: 1 1 auto;
            min-width: 0;
        }}
        .hero-title-vn {{
            margin: 0;
            color: #2b3140;
            font-size: 2.15rem;
            font-weight: 800;
            line-height: 1.12;
            letter-spacing: -0.02em;
            word-break: keep-all;
        }}
        .hero-title-jp {{
            margin: 0.2rem 0 0;
            color: #58637a;
            font-size: 1.18rem;
            font-weight: 700;
            line-height: 1.2;
        }}
        .hero-desc-vn, .hero-desc-jp {{
            margin: 0.7rem 0 0;
            color: #5c677d;
            line-height: 1.45;
        }}
        .hero-desc-vn {{
            font-size: 0.98rem;
        }}
        .hero-desc-jp {{
            font-size: 0.84rem;
        }}
        .hero-meta {{
            flex: 0 0 240px;
            border-radius: 20px;
            padding: 1rem 1rem 0.9rem;
            background: rgba(255, 255, 255, 0.74);
            border: 1px solid rgba(35, 43, 62, 0.08);
        }}
        .meta-label {{
            font-size: 0.75rem;
            color: #8a93a7;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            margin-bottom: 0.18rem;
        }}
        .meta-value {{
            font-size: 1rem;
            color: #1f2937;
            font-weight: 700;
            margin-bottom: 0.8rem;
        }}
        .section-title {{
            margin-top: 0.25rem;
            margin-bottom: 0.05rem;
            color: #263043;
            font-weight: 800;
            font-size: 1.25rem;
        }}
        .section-subtitle {{
            margin-top: 0;
            margin-bottom: 0.8rem;
            color: #6b7280;
            font-size: 0.84rem;
        }}
        .save-panel {{
            margin-top: 1rem;
            border-radius: 22px;
            padding: 1rem 1rem 1.05rem;
            border: 1px solid rgba(208, 213, 221, 0.9);
            background: rgba(255, 255, 255, 0.88);
        }}
        .save-panel-title {{
            color: #b42318;
            font-size: 1.05rem;
            font-weight: 800;
            margin: 0 0 0.2rem;
        }}
        .save-panel-sub {{
            color: #b54708;
            font-size: 0.9rem;
            margin: 0;
        }}
        div[data-testid="stFormSubmitButton"] > button {{
            min-height: 3.25rem;
            border-radius: 16px;
            border: 2px solid #ef4444;
            background: linear-gradient(180deg, #ff6868, #ef4444);
            color: white;
            font-weight: 800;
            letter-spacing: 0.02em;
            box-shadow: 0 0 0 4px rgba(239, 68, 68, 0.12);
        }}
        div[data-testid="stFormSubmitButton"] > button:hover {{
            border-color: #dc2626;
            background: linear-gradient(180deg, #ff5a5a, #dc2626);
        }}
        div[data-testid="stDownloadButton"] > button {{
            min-height: 3.1rem;
            border-radius: 16px;
            border: 2px solid #0f766e;
            background: linear-gradient(180deg, #18b7a5, #0f766e);
            color: white;
            font-weight: 800;
            letter-spacing: 0.02em;
            box-shadow: 0 0 0 4px rgba(15, 118, 110, 0.12);
        }}
        div[data-testid="stDownloadButton"] > button:hover {{
            border-color: #115e59;
            background: linear-gradient(180deg, #14a394, #115e59);
            color: white;
        }}
        .support-links {{
            margin-top: 0.65rem;
            font-size: 0.97rem;
            font-weight: 700;
        }}
        .support-links a {{
            color: #1d4ed8;
            text-decoration: underline;
        }}
        @media (max-width: 820px) {{
            .hero-card {{
                flex-direction: column;
                padding: 1.1rem 1rem;
                border-radius: 22px;
            }}
            .hero-title-vn {{
                font-size: 0.98rem;
                line-height: 1.2;
            }}
            .hero-title-jp {{
                font-size: 0.72rem;
            }}
            .hero-desc-vn {{
                font-size: 0.86rem;
            }}
            .hero-desc-jp {{
                font-size: 0.72rem;
                line-height: 1.38;
            }}
            .hero-meta {{
                width: 100%;
                flex-basis: auto;
            }}
            .section-title {{
                font-size: 1.08rem;
            }}
            .section-subtitle {{
                font-size: 0.78rem;
            }}
        }}
        </style>
        <div class="hero-card">
            <div class="hero-copy">
                <h1 class="hero-title-vn">{APP_TITLE}</h1>
                <p class="hero-title-jp">{APP_TITLE_JP}</p>
                <p class="hero-desc-vn">
                    Mở app là nhập ngay. Nếu chọn không tham gia thì phần địa điểm sẽ tự ẩn.
                </p>
                <p class="hero-desc-jp">
                    アプリを開いたらすぐ入力できます。不参加を選ぶと行き先は自動で非表示になります。
                </p>
            </div>
            <div class="hero-meta">
                <div class="meta-label">Người lập</div>
                <div class="meta-value">{APP_AUTHOR}</div>
                <div class="meta-label">Phiên bản</div>
                <div class="meta-value">{APP_VERSION}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_section_title(vn: str, jp: str) -> None:
    st.markdown(
        f"""
        <div class="section-title">{vn}</div>
        <div class="section-subtitle">{jp}</div>
        """,
        unsafe_allow_html=True,
    )


def render_form() -> None:
    render_section_title("Phiếu khảo sát", "アンケート入力")

    with st.form("survey_form", clear_on_submit=True):
        msnv = st.text_input("MSNV", placeholder="Ví dụ: G06071209")
        name = st.text_input("Họ tên", placeholder="Ví dụ: Nguyễn Duy Hoà")
        department = st.selectbox("Bộ phận", DEPARTMENTS, index=3)
        process = st.text_input("Công đoạn", placeholder="Ví dụ: 56-0")
        join_value = st.selectbox(
            "Bạn có tham gia chuyến đi không?",
            options=["", "Co", "Khong"],
            format_func=lambda value: JOIN_LABELS[value],
        )

        destination_value = ""
        if join_value == "Co":
            destination_value = st.selectbox(
                "Chọn 1 địa điểm",
                options=["", "Nha Trang", "Da Lat"],
                format_func=lambda value: DESTINATION_LABELS[value],
            )
        else:
            st.caption("Nếu chọn Không thì không cần chọn địa điểm. / 不参加の場合は行き先の選択は不要です。")

        submitted = st.form_submit_button("LƯU KẾT QUẢ KHẢO SÁT / アンケート結果を保存")

    form_dirty = any(
        [
            normalize_text(msnv),
            normalize_text(name),
            normalize_text(process),
            normalize_text(join_value),
            normalize_text(destination_value),
        ]
    )
    render_close_warning(form_dirty)

    if not submitted:
        return

    clean_msnv = normalize_spaces(msnv)
    clean_name = normalize_name(name)
    clean_process = normalize_spaces(process)

    if not clean_msnv:
        set_save_message("Vui lòng nhập MSNV trước khi lưu. / 保存前に社員番号を入力してください。", "warning")
        return
    if not clean_name:
        set_save_message("Vui lòng nhập họ tên trước khi lưu. / 保存前に氏名を入力してください。", "warning")
        return
    if not join_value:
        set_save_message("Vui lòng chọn tham gia hoặc không tham gia. / 参加または不参加を選択してください。", "warning")
        return
    if join_value == "Co" and not destination_value:
        set_save_message("Nếu chọn Có thì cần chọn địa điểm. / 参加する場合は行き先を選択してください。", "warning")
        return

    record = clean_record(
        {
            "msnv": clean_msnv,
            "ho_ten": clean_name,
            "bo_phan": department,
            "cong_doan": clean_process,
            "tham_gia": join_value,
            "dia_diem": destination_value,
        }
    )
    st.session_state.records = upsert_record(st.session_state.records, record)

    saved_path = save_export_file(st.session_state.records)
    drive_result = sync_export_to_google_drive(saved_path)

    st.session_state.last_saved_path = str(saved_path)
    st.session_state.last_saved_at = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    if drive_result.get("url"):
        st.session_state.last_drive_url = drive_result["url"]

    drive_text = ""
    if drive_result["status"] in {"created", "updated"}:
        drive_text = " Đã cập nhật Google Drive. / Google Driveにも更新しました。"
    elif drive_result["status"] == "not_configured":
        drive_text = " Google Drive chưa cấu hình. / Google Driveは未設定です。"

    set_save_message(
        f"Đã lưu phiếu và cập nhật file tổng lúc {st.session_state.last_saved_at}.{drive_text}",
        "success",
    )
    st.rerun()


def render_records_table(records: list[dict[str, str]]) -> None:
    render_section_title("Danh sách đã nhập", "入力済み一覧")
    frame = build_records_frame(records)
    st.dataframe(frame, use_container_width=True, hide_index=True)


def render_download_panel(records: list[dict[str, str]]) -> None:
    settings = get_google_drive_settings()
    drive_url = st.session_state.last_drive_url or settings["shared_url"]

    st.markdown(
        """
        <div class="save-panel">
            <p class="save-panel-title">TỰ ĐỘNG LƯU FILE TỔNG</p>
            <p class="save-panel-sub">集計ファイル自動保存</p>
            <p class="save-panel-sub">
                Mỗi lần bấm nút lưu, app sẽ cập nhật ngay vào 1 file tổng duy nhất.
                保存ボタンを押すたびに、1つの集計ファイルへ自動反映されます。
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    export_bytes = build_export_bytes(records)
    st.download_button(
        "TẢI FILE TỔNG / 集計ファイルをダウンロード",
        data=export_bytes,
        file_name=EXPORT_FILE_NAME,
        mime=EXCEL_MIME_TYPE,
        use_container_width=True,
    )

    if drive_url:
        st.markdown(
            f"""
            <div class="support-links">
                <a href="{drive_url}" target="_blank">
                    MỞ FILE TỔNG TRÊN GOOGLE DRIVE / Google Driveで集計ファイルを開く
                </a>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_dashboard(records: list[dict[str, str]]) -> None:
    render_section_title("Dashboard tổng quan", "全体ダッシュボード")
    stats = dashboard_metrics(records)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Tổng phiếu", stats["total"])
    c2.metric("Tham gia", stats["joined"])
    c3.metric("Không tham gia", stats["not_joined"])
    c4.metric("Nha Trang", stats["nha_trang"])
    c5.metric("Đà Lạt", stats["da_lat"])

    overview_table = build_overview_table(records)
    overview_chart = build_overview_chart_frame(records)

    left, right = st.columns([1.12, 1])
    with left:
        render_section_title("Bảng tổng toàn công ty", "全社集計表")
        st.dataframe(overview_table, use_container_width=True, hide_index=True)
    with right:
        render_section_title("Biểu đồ tổng toàn công ty", "全社集計グラフ")
        st.caption(
            f"Tổng phiếu: {stats['total']} | Tham gia: {stats['joined']} | Không tham gia: {stats['not_joined']}"
        )
        st.bar_chart(overview_chart, use_container_width=True)

    department_summary = build_department_summary(records)
    department_chart = build_department_chart_frame(records)

    render_section_title("Chi tiết từng bộ phận", "部門別集計")
    if department_summary.empty:
        st.info("Chưa có dữ liệu để hiển thị dashboard. / ダッシュボード表示用のデータはまだありません。")
        return

    total_row = department_summary.iloc[-1]
    st.caption(
        "Tổng cộng bộ phận phải khớp với tổng quan: "
        f"{int(total_row['Tổng phiếu'])} phiếu, "
        f"{int(total_row['Tham gia'])} tham gia, "
        f"{int(total_row['Không tham gia'])} không tham gia."
    )
    left, right = st.columns([1.15, 1])
    with left:
        st.dataframe(department_summary, use_container_width=True, hide_index=True)
    with right:
        st.caption(
            "Biểu đồ theo bộ phận: tham gia, không tham gia, Nha Trang, Đà Lạt."
        )
        st.bar_chart(department_chart, use_container_width=True)


def render_app() -> None:
    init_state()
    bootstrap_records()
    render_hero()
    show_save_message()

    records = st.session_state.records

    render_form()

    if records:
        render_records_table(records)
        render_download_panel(records)
        render_dashboard(records)
    else:
        st.info("Chưa có dữ liệu khảo sát. Hãy nhập phiếu đầu tiên. / まだ回答がありません。最初の1件を入力してください。")


render_app()

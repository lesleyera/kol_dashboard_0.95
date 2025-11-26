import streamlit as st
import pandas as pd
import numpy as np
import datetime
import altair as alt
import calendar
import base64
import os
import streamlit.components.v1 as components
from streamlit_calendar import calendar as st_calendar

# -----------------------------------------------------------------
# 1. Page Config & CSS
# -----------------------------------------------------------------
st.set_page_config(
    page_title="MEDIT KOL Dashboard",
    page_icon="🟦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

def local_css():
    st.markdown(
        """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Roboto', sans-serif;
            color: #111111;
            background-color: #F5F6FA;
        }

        .block-container {
            padding-top: 2.6rem !important;
            padding-bottom: 2rem !important;
        }

        /* 헤더 영역 */
        .app-header {
            background: linear-gradient(135deg, #2D5AF5 0%, #00A9E0 40%, #3EC9A7 100%);
            border-radius: 20px;
            padding: 22px 26px 20px 26px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.12);
            display: flex;
            align-items: center;
            justify-content: space-between;
            color: #FFFFFF; /* Header 텍스트 전체 색상 */
        }
        .app-header-left {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }
        .app-title {
            font-size: 1.9rem;
            font-weight: 900; 
            letter-spacing: 0.02em;
            color: #FFFFFF !important; /* 수정: 제목 색상 흰색 강제 */
        }
        .app-subtitle {
            font-size: 0.95rem;
            opacity: 0.9;
        }
        .app-tag {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 0.8rem;
            padding: 4px 10px;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.18);
            backdrop-filter: blur(6px);
            border: 1px solid rgba(255, 255, 255, 0.25);
        }
        .app-tag-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #3EC9A7;
            box-shadow: 0 0 6px rgba(62, 201, 167, 0.9);
        }

        /* 데이터프레임 헤더 */
        thead tr th {
            background-color: #2D5AF5 !important; 
            color: #FFFFFF !important; 
            font-size: 13px !important;
            font-weight: 600 !important;
        }

        /* FullCalendar 스크롤 제거 (달력 칸 높이 고정에 필수) */
        .fc-scroller {
            overflow: hidden !important;
        }

        /* 섹션 대제목 스타일 */
        h3 {
            font-family: 'Roboto', sans-serif;
            font-weight: 900 !important; 
            font-size: 1.5rem !important;
            color: #111111;
            margin-top: 2.4rem !important;
            letter-spacing: -0.02em;
        }

        /* 달력 제목 스타일 */
        .fc-toolbar-title {
            font-family: 'Roboto', sans-serif !important; /* 수정: 폰트 통일 */
            font-weight: 900 !important;
            font-size: 1.15rem !important; /* 수정: 크기 축소 */
            color: #111111 !important; 
            text-transform: capitalize; 
        }
        
        /* KPI 텍스트 */
        .kpi-label {
            font-size: 0.9rem;
            color: #6B7280;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 4px;
        }
        .kpi-value {
            font-size: 2.3rem;
            font-weight: 900;
            color: #2D5AF5;
            line-height: 1.1;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

local_css()

# -----------------------------------------------------------------
# 2. Constants & Settings
# -----------------------------------------------------------------
FILE_SETTINGS = {
    "FILE_PATH": "KOL_consolidated_251123(V).xlsx",
    "MASTER_TAB": "kol_master",
    "CONTRACT_TAB": "contract_tasks",
    "ACTIVITY_TAB": "activity_log",
}

# --- Google Maps API Key 설정 (하드코딩 방식) ---
# 이 값을 유효한 키로 변경해야 지도가 정상 작동합니다.
GOOGLE_MAPS_API_KEY = "AIzaSyAVIHGVbAa47uwyQvo0OKW7Hu7M1DVrpYI" 

COLOR_PRIMARY = "#2D5AF5"
COLOR_DANGER = "#DC2626"
COLOR_WARNING = "#F59E0B"

MONTH_NAME_MAP = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}
MONTH_NAME_TO_NUM = {v: k for k, v in MONTH_NAME_MAP.items()}

TASK_COLOR_MAP = {
    "Lecture": "#1D4ED8",
    "Case Report": "#0EA5E9",
    "SNS Posting": "#EC4899",
    "Article": "#F97316",
    "Webinar": "#22C55E",
    "Testimonial": "#6366F1",
}

# -----------------------------------------------------------------
# 3. Helper Functions
# -----------------------------------------------------------------
def find_col(df, candidates):
    cols = list(df.columns)
    norm = {c: c.strip().lower().replace(" ", "_").replace("(", "").replace(")", "").replace("/", "_") for c in cols}
    normalized_candidates = [c.strip().lower().replace(" ", "_").replace("(", "").replace(")", "").replace("/", "_") for c in candidates]
    for original, n in norm.items():
        for nc in normalized_candidates:
            if nc in n: 
                return original
    return None

def normalize_status(val: str) -> str:
    if pd.isna(val):
        return "Planned"
    s = str(val).strip().lower().replace("_", " ").replace("-", " ")
    if s in ["planned", "plan"]:
        return "Planned"
    if any(x in s for x in ["on progress", "in progress", "ongoing", "progress", "doing"]):
        return "On Progress"
    if any(x in s for x in ["done", "finished", "complete", "completed", "end"]):
        return "Done"
    return s.title()

def delayed_to_bool(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s in ["1", "y", "yes", "true", "delayed", "delay", "o"]

def warning_to_bool(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return "warning" in s

def highlight_critical_rows(row):
    style = ''
    status_val = ""
    
    if "Warning/Delayed" in row.index:
        status_val = str(row["Warning/Delayed"]).lower()
    elif "Delayed" in row.index:
        val = str(row["Delayed"]).lower()
        if val in ["true", "delayed", "yes", "y", "1"]:
            status_val = "delayed"
        elif "warning" in val:
            status_val = "warning"
            
    if "delayed" in status_val:
        style = 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;' # Red
    elif "warning" in status_val:
        style = 'background-color: #FEF3C7; color: #92400E; font-weight: bold;' # Orange

    return [style] * len(row)

def create_warning_delayed_col(row):
    if row.get("Delayed_flag") is True:
        return "Delayed"
    elif row.get("Warning_flag") is True:
        return "Warning"
    
    if "Contract_End" in row and pd.notna(row["Contract_End"]):
        try:
            today_ts = pd.Timestamp(datetime.date.today())
            contract_end_ts = pd.to_datetime(row["Contract_End"])
            diff = contract_end_ts - today_ts
            if 0 <= diff.days <= 30:
                return "Warning"
        except:
            pass
    return "-"

def kpi_text(label: str, value: str, color: str = COLOR_PRIMARY):
    st.markdown(
        f"""
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color:{color};">{value}</div>
        """,
        unsafe_allow_html=True,
    )

def show_pdf(file_path):
    """PDF 파일을 다운로드 링크로 제공 (Edge/보안 차단 회피)"""
    if not os.path.exists(file_path):
        st.caption(f"No PDF found: {os.path.basename(file_path)}")
        return

    try:
        with open(file_path, "rb") as f:
            base64_pdf = base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")
        return
    
    href = f"""
    <a href="data:application/pdf;base64,{base64_pdf}" download="{os.path.basename(file_path)}" target="_blank" 
        style="text-decoration: none; padding: 10px 15px; background-color: #2D5AF5; color: white; border-radius: 5px; 
        font-weight: 500; display: inline-block;">
        ⬇️ {os.path.basename(file_path)} 프로필 다운로드 (클릭)
    </a>
    """

    st.markdown("---")
    st.markdown("#### 프로필 PDF 파일 접근", unsafe_allow_html=True)
    st.markdown(href, unsafe_allow_html=True)
    st.caption("⚠️ 브라우저 보안 정책으로 인해 파일을 다운로드하여 열람해 주세요.")


def render_task_legend():
    legend_html = '<div style="display: flex; gap: 12px; margin-bottom: 14px; flex-wrap: wrap;">'
    for task, color in TASK_COLOR_MAP.items():
        legend_html += f'<div style="display: flex; align-items: center; gap: 5px;"><div style="width: 12px; height: 12px; background-color: {color}; border-radius: 3px;"></div><span style="font-size: 0.8rem; color: #374151; font-weight: 500;">{task}</span></div>'
    legend_html += f'<div style="display: flex; align-items: center; gap: 5px; margin-left: 8px; border-left: 1px solid #ddd; padding: left: 8px;"><div style="width: 12px; height: 12px; background-color: {COLOR_DANGER}; border-radius: 3px;"></div><span style="font-size: 0.8rem; color: {COLOR_DANGER}; font-weight: 700;">Delayed</span></div>'
    legend_html += '</div>'
    st.markdown(legend_html, unsafe_allow_html=True)

# [Google Maps API 기반 지도 렌더링 함수 - 하드코딩된 상수 사용]
def render_google_map(df_master, area_filter=None):
    # 전역 상수로 정의된 GOOGLE_MAPS_API_KEY 사용
    api_key = GOOGLE_MAPS_API_KEY
    if api_key == "YOUR_GOOGLE_MAPS_API_KEY_HERE" or not api_key:
        # 이 시점에서 키가 설정되지 않았다면 오류 메시지를 반환합니다.
        return "<div style='padding:16px; color:red;'>🚨 Map Error: Google Maps API Key is not configured.</div>"

    lat_col = find_col(df_master, ["lat", "latitude", "Latitude"])
    lon_col = find_col(df_master, ["lon", "longitude", "Longitude"])
    if lat_col is None or lon_col is None:
        return "<div style='padding:8px; font-size:14px; color:#4B5563;'>No location (lat/lon) data in kol_master.</div>"

    df_map = df_master.dropna(subset=[lat_col, lon_col]).copy()
    if area_filter is not None:
        df_map = df_map[df_map["Area"] == area_filter]

    if df_map.empty:
        map_center_lat = 37.5665
        map_center_lng = 126.9780
        markers_json = "[]"
    else:
        df_map["Latitude_Raw"] = pd.to_numeric(df_map[lat_col], errors="coerce")
        df_map["Longitude_Raw"] = pd.to_numeric(df_map[lon_col], errors="coerce")
        df_map = df_map.dropna(subset=["Latitude_Raw", "Longitude_Raw"])

        # 🚨🚨🚨 좌표 순서 교환 로직: KML/GIS (경도, 위도) 순서가 웹 (위도, 경도) 순서로 바뀌었을 때 적용 🚨🚨🚨
        # 만약 로컬에서 순서가 맞는데 Streamlit에서 안 맞는다면, 데이터 로딩 후 순서가 꼬였을 가능성 높음.
        # 강제로 Latitude(Y)와 Longitude(X)를 교환하여 웹 표준 (Latitude, Longitude)을 맞춥니다.
        
        # 교환 전 원본 값을 복사
        lat_values_raw = df_map["Latitude_Raw"].copy()
        lon_values_raw = df_map["Longitude_Raw"].copy()
        
        # Longitude 컬럼에 Latitude 값을 할당 (경도 자리에 위도 값을 넣음)
        df_map["Longitude"] = lat_values_raw
        
        # Latitude 컬럼에 Longitude 값을 할당 (위도 자리에 경도 값을 넣음)
        df_map["Latitude"] = lon_values_raw
        
        # 이 시점에서 df_map["Latitude"]는 경도 값, df_map["Longitude"]는 위도 값을 담고 있습니다.
        # 그러나 JavaScript는 {lat: 위도, lng: 경도}를 기대합니다.
        # 즉, df_map['Latitude']를 lng에, df_map['Longitude']를 lat에 넣어야 합니다.

        # (위도, 경도) 순서를 맞추기 위해, 여기서 다시 한번 순서를 교환합니다. 
        # 이는 Google Sheet의 컬럼 이름(Latitude, Longitude)과 파이썬 변수 이름('lat', 'lon')이
        # 서로 엇갈려 사용되었을 때 발생하는 흔한 오류를 해결하기 위한 안전 장치입니다.
        
        # 최종적으로 JavaScript에 전달할 {lat: 위도, lng: 경도} 순서를 맞추기 위해 변수를 준비합니다.
        # df_map["Latitude"] (현재 경도값) -> lng (경도)
        # df_map["Longitude"] (현재 위도값) -> lat (위도)

        map_center_lat = df_map["Longitude"].mean()
        map_center_lng = df_map["Latitude"].mean()

        markers_list = []
        for _, row in df_map.iterrows():
            name = row.get("Name", "Unknown")
            hospital = row.get("Hospital", "")
            country = row.get("Country", "")
            area = row.get("Area", "")

            if pd.isna(name): name = "Unknown"
            if pd.isna(hospital): hospital = ""
            if pd.isna(country): country = ""
            if pd.isna(area): area = ""

            info_content = f"<b>{name}</b><br>"
            if isinstance(hospital, str) and hospital.strip():
                info_content += f"{hospital}<br>"
            extra = " · ".join(
                [x for x in [country, area] if isinstance(x, str) and x.strip()]
            )
            if extra:
                info_content += extra

            markers_list.append(
                {
                    "name": name,
                    "lat": float(row["Longitude"]), # 위도 값
                    "lng": float(row["Latitude"]),  # 경도 값
                    "info": info_content,
                }
            )
        import json as _json
        markers_json = _json.dumps(markers_list)

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            #map {{
                height: 100%;
                width: 100%;
                border-radius: 12px;
            }}
            html, body {{
                height: 100%;
                margin: 0;
                padding: 0;
                font-family: 'Roboto', sans-serif;
            }}
        </style>
    </head>
    <body>
        <div id="map"></div>
        <script>
            function initMap() {{
                const map = new google.maps.Map(document.getElementById("map"), {{
                    zoom: 4,
                    center: {{ lat: {map_center_lat}, lng: {map_center_lng} }},
                    mapTypeControl: false,
                    streetViewControl: false
                }});

                const markersData = {markers_json};
                const infoWindow = new google.maps.InfoWindow();

                markersData.forEach((data) => {{
                    const marker = new google.maps.Marker({{
                        position: {{ lat: data.lat, lng: data.lng }},
                        map: map,
                        title: data.name
                    }});

                    marker.addListener("click", () => {{
                        infoWindow.setContent(data.info);
                        infoWindow.open(map, marker);
                    }});
                }});
            }}
        </script>
        <script
            src="https://maps.googleapis.com/maps/api/js?key={api_key}&callback=initMap"
            async defer>
        </script>
    </body>
    </html>
    """
    return html_code

@st.cache_data(ttl=600)
def load_data(file_path, master_tab, contract_tab, activity_tab):
    try:
        # 데이터 파일 로드 (로컬 경로에 파일이 있어야 함)
        df_master_raw = pd.read_excel(file_path, sheet_name=master_tab, engine="openpyxl")
        df_contract = pd.read_excel(file_path, sheet_name=contract_tab, engine="openpyxl")
        df_act = pd.read_excel(file_path, sheet_name=activity_tab, engine="openpyxl")

        df_act = df_act.drop_duplicates()

        # --- Master ---
        col_kol_id_m = find_col(df_master_raw, ["KOL_ID", "kol_id"])
        col_name_m = find_col(df_master_raw, ["Name"])
        col_area_m = find_col(df_master_raw, ["Area"])
        col_country_m = find_col(df_master_raw, ["Country"])

        if col_name_m is None:
            st.error("kol_master 시트에서 Name 컬럼을 찾을 수 없습니다.")
            return None, None, None, None

        master_name_count = df_master_raw[col_name_m].dropna().astype(str).nunique()
        df_master = df_master_raw.copy()

        if col_kol_id_m is not None:
            df_master[col_kol_id_m] = pd.to_numeric(df_master[col_kol_id_m], errors="coerce")

        # 엑셀 시트에서 'Latitude' 컬럼을 lat_col로, 'Longitude' 컬럼을 lon_col로 찾음
        lat_col = find_col(df_master, ["lat", "latitude"])
        lon_col = find_col(df_master, ["lon", "longitude"])
        
        # 로드된 컬럼 이름이 'Latitude'/'Longitude'라고 가정하고, 
        # 이후 로직에서 'Latitude' 컬럼이 lat에, 'Longitude' 컬럼이 lon에 매핑됨
        if lat_col is not None and lon_col is not None:
            df_master[lat_col] = pd.to_numeric(df_master[lat_col], errors="coerce")
            df_master[lon_col] = pd.to_numeric(df_master[lon_col], errors="coerce")

        rename_map = {
            col_name_m: "Name",
            col_area_m: "Area" if col_area_m else "Area",
            col_country_m: "Country" if col_country_m else "Country",
        }
        if col_kol_id_m is not None:
            rename_map[col_kol_id_m] = "KOL_ID"

        df_master = df_master.rename(columns=rename_map)

        # --- Contract ---
        col_kol_id_c = find_col(df_contract, ["KOL_ID", "kol_id"])
        col_name_c = find_col(df_contract, ["Name"])
        col_task_c = find_col(df_contract, ["Task"])
        col_cstart = find_col(df_contract, ["Contract_Start", "Contract Start"])
        col_cend = find_col(df_contract, ["Contract_End", "Contract End"])
        col_freq = find_col(df_contract, ["Freq.", "Freq", "Frequency"])

        if col_task_c is None:
            st.error("contract_tasks 시트에서 Task 컬럼을 찾을 수 없습니다.")
            return None, None, None, None

        if col_kol_id_c is not None:
            df_contract[col_kol_id_c] = pd.to_numeric(df_contract[col_kol_id_c], errors="coerce")
        if col_cstart is not None:
            df_contract[col_cstart] = pd.to_datetime(df_contract[col_cstart], errors="coerce")
        if col_cend is not None:
            df_contract[col_cend] = pd.to_datetime(df_contract[col_cend], errors="coerce")
        if col_freq is not None:
            df_contract[col_freq] = pd.to_numeric(df_contract[col_freq], errors="coerce")

        rename_c = {
            col_task_c: "Task",
            col_name_c: "Name" if col_name_c else "Name",
            col_cstart: "Contract_Start" if col_cstart else "Contract_Start",
            col_cend: "Contract_End" if col_cend else "Contract_End",
            col_freq: "Freq" if col_freq else "Freq",
        }
        if col_kol_id_c is not None:
            rename_c[col_kol_id_c] = "KOL_ID"

        df_contract = df_contract.rename(columns=rename_c)

        # --- Activity ---
        col_kol_id_a = find_col(df_act, ["KOL_ID", "kol_id"])
        col_name_a = find_col(df_act, ["Name"])
        col_date = find_col(df_act, ["Date"])
        col_task_a = find_col(df_act, ["Task"])
        col_activity_a = find_col(df_act, ["Activity", "Topic", "Title", "Details"])
        col_status = find_col(df_act, ["Status", "Status(Planned/On Progress/Done)", "Current Status", "State"])
        col_delayed = find_col(df_act, ["Delayed", "Delay"])
        col_area_a = find_col(df_act, ["Area"])
        col_country_a = find_col(df_act, ["Country"])

        if col_date is None:
            st.error("activity_log 시트에서 Date 컬럼을 찾을 수 없습니다.")
            return None, None, None, None

        if col_kol_id_a is not None:
            df_act[col_kol_id_a] = pd.to_numeric(df_act[col_kol_id_a], errors="coerce")

        df_act[col_date] = pd.to_datetime(df_act[col_date], errors="coerce")
        df_act = df_act.dropna(subset=[col_date])

        if col_status is not None:
            df_act["Status_norm"] = df_act[col_status].apply(normalize_status)
        else:
            df_act["Status_norm"] = "Planned"

        if col_delayed is not None:
            df_act["Delayed_flag"] = df_act[col_delayed].apply(delayed_to_bool)
            df_act["Warning_flag"] = df_act[col_delayed].apply(warning_to_bool)
        else:
            df_act["Delayed_flag"] = False
            df_act["Warning_flag"] = False

        if col_area_a is None or col_country_a is None:
            if "KOL_ID" in df_master.columns and col_kol_id_a is not None:
                df_act = df_act.merge(
                    df_master[["KOL_ID", "Area", "Country"]],
                    left_on=col_kol_id_a,
                    right_on="KOL_ID",
                    how="left",
                    suffixes=("", "_m"),
                )
                col_area_a = "Area"
                col_country_a = "Country"

        rename_a = {
            col_name_a: "Name",
            col_date: "Date",
            col_task_a: "Task",
            col_status: "Status",
            col_delayed: "Delayed",
            col_area_a: "Area",
            col_country_a: "Country",
        }
        if col_kol_id_a is not None:
            rename_a[col_kol_id_a] = "KOL_ID"
        
        if col_activity_a is not None:
            rename_a[col_activity_a] = "Activity"

        df_act = df_act.rename(columns=rename_a)
        
        if "Activity" not in df_act.columns:
            df_act["Activity"] = ""

        return df_master, df_contract, df_act, master_name_count

    except Exception as e:
        st.error(f"Data Load Error: {e}")
        return None, None, None, None

def build_task_status(df_master, df_contract, df_act_filtered):
    if df_contract is None or df_contract.empty:
        return pd.DataFrame()

    base_cols = ["Name", "Task"]
    if "KOL_ID" in df_contract.columns:
        base_cols.append("KOL_ID")
    if "Contract_End" in df_contract.columns:
        base_cols.append("Contract_End")

    base = (
        df_contract[base_cols]
        .dropna(subset=["Name", "Task"])
        .drop_duplicates()
        .copy()
    )

    if "Name" in df_master.columns:
        base = base.merge(
            df_master[["Name", "Area", "Country"]].drop_duplicates(subset=["Name"]),
            on="Name",
            how="left",
        )

    if df_act_filtered is not None and not df_act_filtered.empty:
        grp_keys = ["Name", "Task"]
        if "KOL_ID" in df_act_filtered.columns and "KOL_ID" in base.columns:
            grp_keys = ["KOL_ID", "Name", "Task"]

        grp = (
            df_act_filtered.groupby(grp_keys, dropna=False)
            .agg(
                Latest_Date=("Date", "max"),
                Latest_Status=(
                    "Status_norm",
                    lambda s: s.dropna().iloc[-1] if len(s.dropna()) > 0 else None,
                ),
                Any_Delayed=("Delayed_flag", "any"),
                Any_Warning=("Warning_flag", "any"),
                Activity_Count=("Task", "size"),
            )
            .reset_index()
        )
    else:
        grp = pd.DataFrame(
            columns=["Name", "Task", "Latest_Date", "Latest_Status", "Any_Delayed", "Any_Warning", "Activity_Count"]
        )

    if "KOL_ID" in base.columns and "KOL_ID" in grp.columns:
        merge_keys = ["KOL_ID", "Name", "Task"]
    else:
        merge_keys = ["Name", "Task"]

    merged = base.merge(
        grp,
        on=merge_keys,
        how="left",
    )

    merged["Activity_Count"] = merged["Activity_Count"].fillna(0).astype(int)
    merged["Any_Delayed"] = merged["Any_Delayed"].fillna(False)
    merged["Any_Warning"] = merged["Any_Warning"].fillna(False)

    def summarize(row):
        if row["Any_Delayed"]:
            return "Delayed"
        if row["Any_Warning"]:
            return "Warning"
        st_val = row["Latest_Status"]
        if st_val == "Done":
            return "Done"
        if st_val == "On Progress":
            return "On Progress"
        if st_val == "Planned" or pd.isna(st_val):
            return "Planned"
        return "Planned"

    merged["Summary_Status"] = merged.apply(summarize, axis=1)
    return merged

def render_kol_detail(kol_name: str, df_master: pd.DataFrame, df_activity: pd.DataFrame, df_contract: pd.DataFrame):
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### KOL Detail")

    info = df_master[df_master["Name"] == kol_name].head(1)

    # --- New Logic: Find and format contract period ---
    contract_info = df_contract[df_contract["Name"] == kol_name].copy()
    contract_period_str = ""
    if not contract_info.empty and "Contract_Start" in contract_info.columns and "Contract_End" in contract_info.columns:
        # Get the earliest start and latest end date for all contracts of this KOL
        start_date = contract_info["Contract_Start"].min()
        end_date = contract_info["Contract_End"].max()

        # Format dates if valid
        start_str = start_date.strftime("%Y.%m.%d") if pd.notna(start_date) and isinstance(start_date, pd.Timestamp) else "?"
        end_str = end_date.strftime("%Y.%m.%d") if pd.notna(end_date) and isinstance(end_date, pd.Timestamp) else "?"
        
        # Only show period if at least one date is valid
        if start_str != '?' or end_str != '?':
            contract_period_str = f" ({start_str} ~ {end_str})"
    # ---------------------------------------------------

    if not info.empty:
        r = info.iloc[0]
        # Append the contract period to the name line
        line = f"**{r.get('Name', kol_name)}{contract_period_str}**" 
        extra_parts = []
        country = r.get("Country", "")
        area = r.get("Area", "")
        if isinstance(country, str) and country.strip():
            extra_parts.append(country)
        if isinstance(area, str) and area.strip():
            extra_parts.append(area)
        if extra_parts:
            line += "  \n" + " · ".join(extra_parts)
        st.markdown(line)
    else:
        st.markdown(f"**{kol_name}{contract_period_str}**") # KOL detail이 없을 때도 기간은 표시

    log = df_activity[df_activity["Name"] == kol_name].copy()
    if log.empty:
        st.caption("No activity records for this KOL.")
    else:
        log = log.sort_values("Date", ascending=False)
        log["Date"] = log["Date"].dt.strftime("%Y-%m-%d")
        
        log["Warning/Delayed"] = log.apply(create_warning_delayed_col, axis=1)
        
        cols = [
            "Date", "Task", "Activity", "Status_norm", "Warning/Delayed", "Area",
        ]
        cols = [c for c in cols if c in log.columns]
        log_disp = log[cols].rename(
            columns={"Status_norm": "Status"}
        )
        st.markdown("**Activity Log**")
        st.dataframe(
            log_disp.style.apply(highlight_critical_rows, axis=1), 
            use_container_width=True, hide_index=True
        )

    st.markdown("---")
    st.markdown("**Profile (PDF)**")
    pdf_filename = f"{kol_name}.pdf"
    pdf_path = os.path.join("profiles", pdf_filename)
    show_pdf(pdf_path)

# -----------------------------------------------------------------
# 5. Load Data
# -----------------------------------------------------------------
df_master, df_contract, df_activity, master_name_count = load_data(
    FILE_SETTINGS["FILE_PATH"],
    FILE_SETTINGS["MASTER_TAB"],
    FILE_SETTINGS["CONTRACT_TAB"],
    FILE_SETTINGS["ACTIVITY_TAB"],
)

if df_master is None or df_activity is None:
    st.stop()

df_activity = df_activity.copy()
df_activity["Year"] = df_activity["Date"].dt.year
df_activity["Month_Num"] = df_activity["Date"].dt.month
df_activity["Month_Name"] = df_activity["Month_Num"].map(MONTH_NAME_MAP)

available_years = sorted(df_activity["Year"].dropna().unique().tolist())
today = datetime.date.today()
default_year = today.year if today.year in available_years else (max(available_years) if available_years else today.year)

available_month_nums = sorted(df_activity["Month_Num"].dropna().unique().tolist())
available_month_names = [MONTH_NAME_MAP[m] for m in available_month_nums] if available_month_nums else list(MONTH_NAME_MAP.values())

# -----------------------------------------------------------------
# 6. Top Controls
# -----------------------------------------------------------------
c_page, c_year, c_month, c_area = st.columns([1.4, 0.8, 0.8, 1.0])

with c_page:
    page = st.selectbox("Board", ["Executive Board", "Regional Board", "Admin Board"])

with c_year:
    selected_year = st.selectbox("Year", options=available_years, index=available_years.index(default_year))

with c_month:
    month_options = ["All"] + available_month_names
    default_month_name = MONTH_NAME_MAP.get(today.month, available_month_names[0] if available_month_names else "Jan")
    if default_month_name not in month_options:
        default_month_name = "All"
    selected_month_name = st.selectbox("Month", options=month_options, index=month_options.index(default_month_name))

with c_area:
    area_options = ["All"] + sorted(df_master["Area"].dropna().unique().tolist())
    selected_area = st.selectbox("Area (for Regional Board)", options=area_options, index=0)

calc_initial_date = today
if selected_month_name != "All":
    _m = MONTH_NAME_TO_NUM[selected_month_name]
    calc_initial_date = datetime.date(selected_year, _m, 1)
else:
    if selected_year != today.year:
        calc_initial_date = datetime.date(selected_year, 1, 1)
    else:
        calc_initial_date = today

calc_initial_date_str = calc_initial_date.strftime("%Y-%m-%d")

mask = df_activity["Year"] == selected_year
if selected_month_name != "All":
    month_num_selected = MONTH_NAME_TO_NUM[selected_month_name]
    mask &= df_activity["Month_Num"] == month_num_selected
else:
    month_num_selected = None

df_activity_filtered = df_activity[mask].copy()
TODAY = today

df_task_status_all = build_task_status(df_master, df_contract, df_activity)
df_task_status_filtered = build_task_status(df_master, df_contract, df_activity_filtered)
total_kols_master = master_name_count if master_name_count is not None else df_master["Name"].dropna().astype(str).nunique()

# -----------------------------------------------------------------
# 7. Executive Board
# -----------------------------------------------------------------
if page == "Executive Board":
    st.markdown(
        f"""
    <div class="app-header">
        <div class="app-header-left">
            <div class="app-title">MEDIT KOL Performance Cockpit</div>
            <div class="app-subtitle">
                Global KOL contract & activity overview · As of <b>{TODAY.strftime('%Y-%m-%d')}</b>
            </div>
            <div class="app-tag">
                <div class="app-tag-dot"></div>
                Schedule · KPIs · KOL map
            </div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown("### Schedule")
    
    if selected_month_name == "All":
        st.info("Please select a specific month to view the Daily Schedule.")
        clicked_kol_exec = None
    else:
        render_task_legend()
        events = []
        for _, row in df_activity_filtered.iterrows():
            date_str = row["Date"].strftime("%Y-%m-%d")
            task_label = row["Task"] if pd.notna(row["Task"]) else ""
            status = row["Status_norm"]
            delayed_flag = bool(row["Delayed_flag"])

            if delayed_flag:
                color = COLOR_DANGER
            else:
                color = TASK_COLOR_MAP.get(str(task_label).strip(), COLOR_PRIMARY)

            title = f"{row['Name']} · {task_label}".strip(" ·")
            if delayed_flag:
                title = "[Delayed] " + title

            area_val = row.get("Area", "")
            if pd.isna(area_val): area_val = ""
            
            events.append(
                {
                    "title": title,
                    "start": date_str,
                    "end": date_str,
                    "allDay": True,
                    "backgroundColor": color,
                    "borderColor": color,
                    "extendedProps": {
                        "status": status,
                        "delayed": delayed_flag,
                        "task": task_label,
                        "kol_name": row["Name"],
                        "area": area_val,
                    },
                }
            )

        cal_state = st_calendar(
            events=events,
            options={
                "initialDate": calc_initial_date_str,
                "headerToolbar": {"left": "", "center": "title", "right": ""},
                "height": 760,
                "contentHeight": 720,
                "fixedWeekCount": True,  # 달력 주 개수 고정
                "dayMaxEvents": 3,      # 하루 최대 이벤트 개수 제한
            },
            key=f"exec_calendar_{selected_year}_{selected_month_name}",
        )
        
        clicked_kol_exec = None
        if cal_state and cal_state.get("callback") == "eventClick":
            ev = cal_state.get("eventClick", {}).get("event", {})
            if isinstance(ev, dict):
                ext = ev.get("extendedProps", {}) or {}
                clicked_kol_exec = ext.get("kol_name")
                if not clicked_kol_exec:
                    title = ev.get("title", "")
                    title = title.replace("[Delayed]", "").strip()
                    clicked_kol_exec = title.split("·")[0].strip()

    if clicked_kol_exec:
        render_kol_detail(clicked_kol_exec, df_master, df_activity, df_contract)

    st.markdown("### KPIs")
    total_planned_tasks = df_task_status_all.shape[0] 
    onprogress_tasks = df_activity_filtered[df_activity_filtered["Status_norm"] == "On Progress"].shape[0]
    done_tasks = df_activity_filtered[df_activity_filtered["Status_norm"] == "Done"].shape[0]
    delayed_tasks = df_activity_filtered[df_activity_filtered["Delayed_flag"] == True].shape[0]
    warning_tasks = df_activity_filtered[df_activity_filtered["Warning_flag"] == True].shape[0]

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: kpi_text("Total KOLs", f"{total_kols_master}")
    with k2: kpi_text("Planned (Tasks)", f"{total_planned_tasks}")
    with k3: kpi_text("On Progress (Acts)", f"{onprogress_tasks}")
    with k4: kpi_text("Done (Acts)", f"{done_tasks}")
    with k5: kpi_text("Delayed (Acts)", f"{delayed_tasks}", color=COLOR_DANGER)
    with k6: kpi_text("Warning", f"{warning_tasks}", color=COLOR_WARNING)

    st.markdown("### Active & Delayed Tasks")
    
    status_df = df_activity_filtered[
        (df_activity_filtered["Status_norm"] == "On Progress") | 
        (df_activity_filtered["Delayed_flag"] == True) |
        (df_activity_filtered["Warning_flag"] == True)
    ].copy()
    
    if status_df.empty:
        st.caption("No active, delayed, or warning tasks for selected period.")
    else:
        status_df["Warning/Delayed"] = status_df.apply(create_warning_delayed_col, axis=1)
        status_cols = ["Date", "Name", "Task", "Activity", "Status_norm", "Warning/Delayed", "Area", "Country"]
        status_disp = status_df[status_cols].rename(columns={"Status_norm": "Status"})
        status_disp["Date"] = status_disp["Date"].dt.strftime("%Y-%m-%d")
        status_disp = status_disp.sort_values(by=["Warning/Delayed", "Date"], ascending=[False, True])
        
        st.dataframe(
            status_disp.style.apply(highlight_critical_rows, axis=1),
            use_container_width=True, 
            hide_index=True
        )

    st.markdown("### KOL Location Map")
    map_html = render_google_map(df_master) 
    components.html(map_html, height=540)

# -----------------------------------------------------------------
# 8. Regional Board
# -----------------------------------------------------------------
elif page == "Regional Board":
    if selected_area == "All":
        st.warning("Please select a specific Area at the top for the Regional Board.")
        st.stop()

    st.markdown(
        f"""
    <div class="app-header">
        <div class="app-header-left">
            <div class="app-title">Regional Performance Cockpit</div>
            <div class="app-subtitle">
                Area: <b>{selected_area}</b> · As of <b>{TODAY.strftime('%Y-%m-%d')}</b>
            </div>
            <div class="app-tag">
                <div class="app-tag-dot"></div>
                Schedule · KPIs · Regional map
            </div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    df_activity_region = df_activity_filtered[df_activity_filtered["Area"] == selected_area].copy()
    df_master_region = df_master[df_master["Area"] == selected_area].copy()
    
    clicked_kol_region = None # NameError 방지를 위해 초기화
    
    st.markdown("### Schedule")
    
    if selected_month_name == "All":
        st.info("Please select a specific month to view the Daily Schedule.")
        clicked_kol_region = None
    else:
        render_task_legend()
        events = []
        for _, row in df_activity_region.iterrows():
            date_str = row["Date"].strftime("%Y-%m-%d")
            task_label = row["Task"] if pd.notna(row["Task"]) else ""
            status = row["Status_norm"]
            delayed_flag = bool(row["Delayed_flag"])

            if delayed_flag:
                color = COLOR_DANGER
            else:
                color = TASK_COLOR_MAP.get(str(task_label).strip(), COLOR_PRIMARY)

            title = f"{row['Name']} · {task_label}".strip(" ·")
            if delayed_flag:
                title = "[Delayed] " + title

            area_val = row.get("Area", "")
            if pd.isna(area_val): area_val = ""

            events.append(
                {
                    "title": title,
                    "start": date_str,
                    "end": date_str,
                    "allDay": True,
                    "backgroundColor": color,
                    "borderColor": color,
                    "extendedProps": {
                        "status": status,
                        "delayed": delayed_flag,
                        "task": task_label,
                        "kol_name": row["Name"],
                        "area": area_val,
                    },
                }
            )

        cal_state_r = st_calendar(
            events=events,
            options={
                "initialDate": calc_initial_date_str,
                "headerToolbar": {"left": "", "center": "title", "right": ""},
                "height": 760,
                "contentHeight": 720,
                "fixedWeekCount": True,  # 달력 주 개수 고정
                "dayMaxEvents": 3,      # 하루 최대 이벤트 개수 제한
            },
            key=f"region_calendar_{selected_area}_{selected_year}_{selected_month_name}",
        )

        if cal_state_r and cal_state_r.get("callback") == "eventClick":
            ev = cal_state_r.get("eventClick", {}).get("event", {})
            if isinstance(ev, dict):
                ext = ev.get("extendedProps", {}) or {}
                clicked_kol_region = ext.get("kol_name")
                if not clicked_kol_region:
                    title = ev.get("title", "")
                    title = title.replace("[Delayed]", "").strip()
                    clicked_kol_region = title.split("·")[0].strip()

    if clicked_kol_region:
        render_kol_detail(clicked_kol_region, df_master_region, df_activity_region, df_contract)

    st.markdown("### KPIs (Region)")
    total_kols_region = df_master_region["Name"].dropna().astype(str).nunique()
    total_planned_region = df_task_status_all[df_task_status_all["Area"] == selected_area].shape[0]
    onprogress_region = df_activity_region[df_activity_region["Status_norm"] == "On Progress"].shape[0]
    done_region = df_activity_region[df_activity_region["Status_norm"] == "Done"].shape[0]
    delayed_region = df_activity_region[df_activity_region["Delayed_flag"] == True].shape[0]
    warning_region = df_activity_region[df_activity_region["Warning_flag"] == True].shape[0]

    r1, r2, r3, r4, r5, r6 = st.columns(6)
    with r1: kpi_text("KOLs in Region", f"{total_kols_region}")
    with r2: kpi_text("Planned (Tasks)", f"{total_planned_region}")
    with r3: kpi_text("On Progress (Acts)", f"{onprogress_region}")
    with r4: kpi_text("Done (Acts)", f"{done_region}")
    with r5: kpi_text("Delayed (Acts)", f"{delayed_region}", color=COLOR_DANGER)
    with r6: kpi_text("Warning", f"{warning_region}", color=COLOR_WARNING)

    st.markdown("### Active & Delayed Tasks (Region)")
    status_df_r = df_activity_region[
        (df_activity_region["Status_norm"] == "On Progress") | 
        (df_activity_region["Delayed_flag"] == True) |
        (df_activity_region["Warning_flag"] == True)
    ].copy()
    
    if status_df_r.empty:
        st.caption("No active, delayed, or warning tasks in this region.")
    else:
        status_df_r["Warning/Delayed"] = status_df_r.apply(create_warning_delayed_col, axis=1)
        status_cols = ["Date", "Name", "Task", "Activity", "Status_norm", "Warning/Delayed"]
        status_disp_r = status_df_r[status_cols].rename(columns={"Status_norm": "Status"})
        status_disp_r["Date"] = status_disp_r["Date"].dt.strftime("%Y-%m-%d")
        status_disp_r = status_disp_r.sort_values(by=["Warning/Delayed", "Date"], ascending=[False, True])
        
        st.dataframe(
            status_disp_r.style.apply(highlight_critical_rows, axis=1),
            use_container_width=True, 
            hide_index=True
        )

    st.markdown("### KOL Location Map")
    map_html_r = render_google_map(df_master, area_filter=selected_area)
    components.html(map_html_r, height=540)

# -----------------------------------------------------------------
# 9. Admin Board
# -----------------------------------------------------------------
else:  # Admin Board
    st.markdown(
        f"""
    <div class="app-header">
        <div class="app-header-left">
            <div class="app-title">Admin · KOL Data Control Center</div>
            <div class="app-subtitle">
                Raw activity logs, task summary, delayed list for operations.
            </div>
            <div class="app-tag">
                <div class="app-tag-dot"></div>
                Activity_log view · Contract-task summary
            </div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown("### Activity Log Viewer")

    col_a1, col_a2, col_a3 = st.columns(3)
    with col_a1:
        f_area = st.selectbox("Area", ["All"] + sorted(df_activity["Area"].dropna().unique().tolist()))
    with col_a2:
        f_status = st.selectbox("Status (normalized)", ["All", "Planned", "On Progress", "Done"])
    with col_a3:
        f_critical = st.selectbox("Warning/Delayed Flag", ["All", "Delayed only", "Warning only", "Non-critical"])

    df_log = df_activity.copy()
    if f_area != "All":
        df_log = df_log[df_log["Area"] == f_area]
    if f_status != "All":
        df_log = df_log[df_log["Status_norm"] == f_status]
    
    if f_critical == "Delayed only":
        df_log = df_log[df_log["Delayed_flag"]]
    elif f_critical == "Warning only":
        df_log = df_log[df_log["Warning_flag"]]
    elif f_critical == "Non-critical":
        df_log = df_log[~(df_log["Delayed_flag"] | df_log["Warning_flag"])]

    if df_log.empty:
        st.caption("No activity-log rows for selected filters.")
    else:
        df_log["Warning/Delayed"] = df_log.apply(create_warning_delayed_col, axis=1)
        
        cols_to_show = [
            "Date", "Name", "KOL_ID", "Area", "Country", 
            "Task", "Activity", 
            "Status_norm", "Warning/Delayed"
        ]
        cols_to_show = [c for c in cols_to_show if c in df_log.columns or c == "KOL_ID" and "KOL_ID" in df_log]
        
        df_log_disp = df_log[cols_to_show].copy()
        df_log_disp["Date"] = df_log_disp["Date"].dt.strftime("%Y-%m-%d")
        df_log_disp = df_log_disp.rename(
            columns={"Status_norm": "Status"}
        )
        
        if "KOL_ID" in df_log_disp.columns:
            df_log_disp["KOL_ID"] = pd.to_numeric(df_log_disp["KOL_ID"], errors='coerce').astype('Int64')
        
        st.dataframe(
            df_log_disp.style.apply(highlight_critical_rows, axis=1), 
            use_container_width=True, hide_index=True
        )

    st.markdown("---")
    st.markdown("### KOL Detail Info")
    
    all_kol_names = sorted(df_master["Name"].dropna().unique().tolist())
    selected_prof = st.selectbox("Select KOL", ["-"] + all_kol_names)
    
    if selected_prof != "-":
        st.markdown("#### Contract Progress Rates")
        subset_c = df_contract[df_contract["Name"] == selected_prof].copy()
        subset_a = df_activity[df_activity["Name"] == selected_prof].copy()
        
        if not subset_c.empty:
            act_counts = subset_a.groupby("Task").size().reset_index(name="Actual_Count")
            progress_df = subset_c.merge(act_counts, on="Task", how="left")
            
            if "Freq" in progress_df.columns:
                progress_df["Actual_Count"] = pd.to_numeric(progress_df["Actual_Count"], errors='coerce').fillna(0).astype(int)
                progress_df["Freq"] = pd.to_numeric(progress_df["Freq"], errors="coerce").fillna(0)
                
                progress_df["Progress (%)"] = np.where(
                    progress_df["Freq"] > 0,
                    (progress_df["Actual_Count"] / progress_df["Freq"] * 100).round(1),
                    0
                )
                
                if "Contract_Start" in progress_df.columns:
                    progress_df["Contract_Start"] = progress_df["Contract_Start"].dt.strftime("%Y-%m-%d")
                if "Contract_End" in progress_df.columns:
                    progress_df["Contract_End"] = progress_df["Contract_End"].dt.strftime("%Y-%m-%d")
                    
                disp_cols = ["Task", "Contract_Start", "Contract_End", "Freq", "Actual_Count", "Progress (%)"]
                disp_cols = [c for c in disp_cols if c in progress_df.columns]
                
                st.dataframe(
                    progress_df[disp_cols].style.format({"Progress (%)": "{:.1f}%"}),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("경고: 'Freq' (빈도수) 컬럼을 찾을 수 없어 진척률을 계산할 수 없습니다.")
        else:
            st.info("선택된 KOL에 대한 계약 정보가 없습니다.")
        
        # 수정된 함수 호출: df_contract 추가
        render_kol_detail(selected_prof, df_master, df_activity, df_contract)
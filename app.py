from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from html import escape
from io import BytesIO
import hmac
import os
import re
from typing import Any
from urllib.parse import urlencode, urljoin

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth


ODATA_ENDPOINT = "https://online.marorka.com/Odata/v1/ODataService.svc/ReportData"
MAX_ODATA_PAGES = 100

EXCLUDED_REPORT_TYPES = [
    "Intake Report",
    "Fuel Change Report",
]

SOURCE_INDEX = [
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "LapTime",
]

REPORT_GROUP_KEYS = [
    "ShipName",
    "EndDateTimeGMT",
]

SOURCE_COLUMNS = [
    *SOURCE_INDEX,
    "ValueDescription",
    "ReportedValue",
]

RENAME_COLUMNS = {
    "ShipName": "Ship Name",
    "ReportType": "Report Type",
    "StartDateTimeGMT": "Start Date & Time GMT",
    "EndDateTimeGMT": "End Date & Time GMT",
    "LapTime": "Lap Time",
    "Total DG Power [kW] (kW)": "Total Daily Average Power [kW]",
}

VALUE_ALIASES = {
    "Total DG Power [kW] (kW)": [
        "Total DG Power [kW] (kW)",
        "Total DG Power [kW]",
        "Total Daily Average Power [kW]",
    ],
    "DG1 Running Hours [hh:mm]": [
        "DG1 Running Hours [hh:mm]",
        "DG 1 Running Hours [hh:mm]",
        "DG1 Running Hours",
        "DG 1 Running Hours",
    ],
    "DG2 Running Hours [hh:mm]": [
        "DG2 Running Hours [hh:mm]",
        "DG 2 Running Hours [hh:mm]",
        "DG2 Running Hours",
        "DG 2 Running Hours",
    ],
    "DG3 Running Hours [hh:mm]": [
        "DG3 Running Hours [hh:mm]",
        "DG 3 Running Hours [hh:mm]",
        "DG3 Running Hours",
        "DG 3 Running Hours",
    ],
    "DG4 Running Hours [hh:mm]": [
        "DG4 Running Hours [hh:mm]",
        "DG 4 Running Hours [hh:mm]",
        "DG4 Running Hours",
        "DG 4 Running Hours",
    ],
    "Shaft Generator Running Hours [hh:mm]": [
        "Shaft Generator Running Hours [hh:mm]",
        "Shaft Generator Running Hours",
    ],
    "Reefer Energy [kWh]": [
        "Reefer Energy [kWh]",
    ],
    "Reefer Power [kW]": [
        "Reefer Power [kW]",
    ],
    "Total Number Reefer Units (20 and 40ft)": [
        "Total Number Reefer Units (20 and 40ft)",
        "Total Number Reefer Units (20 and 40 ft)",
        "Total Number Reefer Units",
        "Total Number of Reefer Units (20 and 40ft)",
        "Total Number of Reefer Units (20 and 40 ft)",
    ],
}

DATETIME_COLUMNS = [
    "Start Date & Time GMT",
    "End Date & Time GMT",
]

GENERATOR_COLUMNS = [
    "DG1 Running Hours [hh:mm]",
    "DG2 Running Hours [hh:mm]",
    "DG3 Running Hours [hh:mm]",
    "DG4 Running Hours [hh:mm]",
]

REEFER_UNITS_COLUMN = "Total Number Reefer Units (20 and 40ft)"

DAILY_COALESCE_COLUMNS = [
    "Total Daily Average Power [kW]",
    "Reefer Energy [kWh]",
    "Reefer Daily Average Power [kW]",
    "Reefer Power [kW]",
    REEFER_UNITS_COLUMN,
    *GENERATOR_COLUMNS,
]

DISPLAY_COLUMNS = [
    "Ship Name",
    "Report Type",
    "Start Date & Time GMT",
    "End Date & Time GMT",
    "Lap Time",
    "Total Daily Average Power [kW]",
    "DG1 Running Hours [hh:mm]",
    "DG2 Running Hours [hh:mm]",
    "DG3 Running Hours [hh:mm]",
    "DG4 Running Hours [hh:mm]",
    "Shaft Generator Running Hours [hh:mm]",
    "Reefer Energy [kWh]",
    "Reefer Daily Average Power [kW]",
    "Reefer Power [kW]",
    REEFER_UNITS_COLUMN,
    "Average Power per Reefer unit [kW]",
]


st.set_page_config(
    page_title="Reefer Dashboard",
    layout="wide",
)


def apply_custom_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --bg: #0B1018;
            --panel: #111827;
            --panel-soft: #162033;
            --border: rgba(148, 163, 184, 0.20);
            --text-soft: #9CA3AF;
            --cyan: #00D1FF;
            --green: #00FFA3;
        }

        html, body,
        .stApp,
        div[data-testid="stAppViewContainer"],
        div[data-testid="stAppViewContainer"] > section,
        main,
        div[data-testid="stMain"],
        div[data-testid="stMain"] > div {
            background:
                radial-gradient(circle at top left, rgba(0, 209, 255, 0.14), transparent 34rem),
                radial-gradient(circle at top right, rgba(0, 255, 163, 0.09), transparent 30rem),
                var(--bg) !important;
            background-color: var(--bg) !important;
        }

        header[data-testid="stHeader"],
        header[data-testid="stHeader"] > div,
        div[data-testid="stToolbar"],
        div[data-testid="stDecoration"] {
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: 0 !important;
            box-shadow: none !important;
        }

        /* Remove the separate dark top band above the dashboard body. */
        div[data-testid="stAppViewContainer"] > .main,
        div[data-testid="stAppViewContainer"] .main,
        section.main,
        .main .block-container {
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
        }

        .block-container {
            padding-top: 3.5rem;
            padding-bottom: 3rem;
            max-width: 1280px;
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #101827 0%, #0B1018 100%);
            border-right: 1px solid var(--border);
        }

        section[data-testid="stSidebar"] label {
            color: #E5E7EB !important;
            font-weight: 700 !important;
        }

        div[data-baseweb="select"] > div {
            background-color: #0F172A !important;
            border: 1px solid rgba(0, 209, 255, 0.22) !important;
            border-radius: 14px !important;
        }

        .dashboard-hero {
            padding: 1.8rem 2rem;
            border: 1px solid var(--border);
            border-radius: 24px;
            background: linear-gradient(135deg, rgba(17, 24, 39, 0.96), rgba(15, 23, 42, 0.78));
            box-shadow: 0 24px 70px rgba(0,0,0,0.36);
            margin-bottom: 1.4rem;
        }

        .eyebrow {
            color: var(--cyan);
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-size: 0.78rem;
            font-weight: 800;
            margin-bottom: 0.35rem;
        }

        .dashboard-title {
            font-size: clamp(2.4rem, 5vw, 4.8rem);
            line-height: 1.02;
            font-weight: 900;
            color: #F8FAFC;
            margin: 0;
        }

        .dashboard-subtitle {
            color: var(--text-soft);
            font-size: 1rem;
            margin-top: 0.8rem;
        }

        .api-load-caption {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            margin: -0.55rem 0 1.15rem 0.15rem;
            padding: 0.35rem 0.65rem;
            border: 1px solid rgba(148, 163, 184, 0.18);
            border-radius: 999px;
            background: rgba(15, 23, 42, 0.55);
            color: #9CA3AF;
            font-size: 0.78rem;
            font-weight: 650;
        }

        .api-load-caption span {
            color: #AEB8C7;
        }

        .section-title {
            font-size: 1.35rem;
            font-weight: 850;
            color: #F8FAFC;
            margin: 1.6rem 0 0.75rem 0;
        }

        div[data-testid="stMetric"] {
            background: linear-gradient(180deg, rgba(22, 32, 51, 0.98), rgba(17, 24, 39, 0.98));
            border: 1px solid var(--border);
            border-radius: 20px;
            padding: 1.05rem 1.1rem;
            box-shadow: 0 14px 34px rgba(0,0,0,0.28);
            min-height: 132px;
        }

        div[data-testid="stMetric"]:hover {
            border-color: rgba(0, 209, 255, 0.55);
            transform: translateY(-1px);
            transition: all 160ms ease;
        }

        div[data-testid="stMetricLabel"] p {
            color: #AEB8C7 !important;
            font-weight: 750 !important;
            font-size: 0.88rem !important;
        }

        div[data-testid="stMetricValue"] {
            color: #F8FAFC !important;
            font-size: clamp(1.35rem, 1.9vw, 2rem) !important;
            line-height: 1.12 !important;
            font-weight: 850 !important;
            white-space: normal !important;
            overflow-wrap: anywhere !important;
        }

        .remark-card {
            border: 1px solid rgba(0, 209, 255, 0.28);
            background: linear-gradient(135deg, rgba(0, 209, 255, 0.11), rgba(0, 255, 163, 0.07));
            border-radius: 20px;
            padding: 1.2rem 1.4rem;
            box-shadow: 0 16px 40px rgba(0,0,0,0.26);
            margin-bottom: 1.2rem;
        }

        .remark-label {
            color: var(--cyan);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 850;
            margin-bottom: 0.4rem;
        }

        .remark-text {
            color: #F8FAFC;
            font-size: 1.08rem;
            font-weight: 650;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border);
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 14px 36px rgba(0,0,0,0.24);
        }

        .stDownloadButton button, .stButton button {
            border-radius: 14px !important;
            border: 1px solid rgba(0, 209, 255, 0.35) !important;
            background: linear-gradient(135deg, rgba(0, 209, 255, 0.95), rgba(0, 255, 163, 0.80)) !important;
            color: #061018 !important;
            font-weight: 850 !important;
        }

        hr {
            border-color: rgba(148, 163, 184, 0.16);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


class MarorkaConfigError(RuntimeError):
    pass


def read_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, os.getenv(name, default))
    except Exception:
        value = os.getenv(name, default)
    return str(value).strip() if value is not None else default


def app_timezone() -> ZoneInfo:
    """Return dashboard display timezone. Defaults to Greece local time."""
    timezone_name = read_secret("APP_TIMEZONE", "Europe/Athens")
    try:
        return ZoneInfo(timezone_name)
    except Exception:
        return ZoneInfo("Europe/Athens")


def current_local_api_load_time() -> str:
    """Timestamp shown in the UI for the latest API/cache load."""
    return datetime.now(app_timezone()).strftime("%Y-%m-%d %H:%M:%S %Z")


def get_query_param(name: str, default: str = "") -> str:
    """Read one query parameter value, compatible with newer and older Streamlit versions."""
    try:
        value = st.query_params.get(name, default)
    except Exception:
        try:
            value = st.experimental_get_query_params().get(name, [default])
        except Exception:
            value = default

    if isinstance(value, list):
        value = value[0] if value else default

    return str(value) if value is not None else default


def is_warmup_request() -> bool:
    return get_query_param("warmup", "0") == "1"


def warmup_token_is_valid() -> bool:
    expected_token = read_secret("WARMUP_TOKEN")
    provided_token = get_query_param("token", "")

    if not expected_token:
        return False

    return hmac.compare_digest(provided_token, expected_token)


def require_dashboard_password() -> None:
    dashboard_password = read_secret("DASHBOARD_PASSWORD")
    if not dashboard_password:
        return

    if st.session_state.get("dashboard_authenticated"):
        return

    apply_custom_css()
    st.markdown(
        """
        <div class="dashboard-hero">
            <div class="eyebrow">Secure access</div>
            <h1 class="dashboard-title">Reefer Dashboard</h1>
            <div class="dashboard-subtitle">Enter your dashboard password to continue.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    entered_password = st.text_input("Password", type="password")

    if st.button("Sign in", type="primary"):
        if hmac.compare_digest(entered_password, dashboard_password):
            st.session_state["dashboard_authenticated"] = True
            st.rerun()
        st.error("Invalid password.")

    st.stop()


def build_odata_url(start_date: date) -> str:
    start_text = start_date.strftime("%Y-%m-%d")
    params = {
        "$filter": f"StartDateTimeGMT gt DateTime'{start_text}'",
        "$select": ",".join(SOURCE_COLUMNS),
    }
    return f"{ODATA_ENDPOINT}?{urlencode(params)}"


def request_auth(username: str, password: str, auth_method: str) -> Any:
    method = auth_method.lower()
    if method == "basic":
        return HTTPBasicAuth(username, password)
    if method == "digest":
        return HTTPDigestAuth(username, password)
    if method == "bearer":
        return None
    if method in {"none", "anonymous", ""}:
        return None
    raise MarorkaConfigError(
        "Unsupported MARORKA_AUTH_METHOD. Use basic, digest, bearer, or none."
    )


def request_headers(token: str, auth_method: str) -> dict[str, str]:
    headers = {
        "Accept": "application/json",
    }
    if auth_method.lower() == "bearer":
        if not token:
            raise MarorkaConfigError("MARORKA_TOKEN is required for bearer auth.")
        headers["Authorization"] = f"Bearer {token}"
    return headers


def extract_odata_page(payload: Any) -> tuple[list[dict[str, Any]], str | None]:
    if isinstance(payload, list):
        return payload, None

    if not isinstance(payload, dict):
        raise ValueError("Could not parse OData response payload.")

    rows = payload.get("value")
    next_link = payload.get("@odata.nextLink") or payload.get("odata.nextLink")

    if rows is None and isinstance(payload.get("d"), dict):
        data = payload["d"]
        rows = data.get("results")
        next_link = next_link or data.get("__next")

    if rows is None:
        raise ValueError("Could not find OData rows in the API response.")

    return rows, next_link


def rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    if "__metadata" in df.columns:
        df = df.drop(columns=["__metadata"])
    return df


@st.cache_data(ttl=21600, show_spinner=False)
def fetch_report_data(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    days_back: int,
) -> pd.DataFrame:
    start_date = datetime.now(timezone.utc).date() - timedelta(days=days_back)
    next_url = build_odata_url(start_date)
    all_rows: list[dict[str, Any]] = []
    seen_urls: set[str] = set()

    for _ in range(MAX_ODATA_PAGES):
        if next_url in seen_urls:
            break
        seen_urls.add(next_url)

        response = requests.get(
            next_url,
            auth=request_auth(username, password, auth_method),
            headers=request_headers(token, auth_method),
            timeout=60,
        )
        response.raise_for_status()

        page_rows, next_link = extract_odata_page(response.json())
        all_rows.extend(page_rows)

        if not next_link:
            break
        next_url = urljoin(next_url, next_link)

    return rows_to_dataframe(all_rows)


def first_non_null(series: pd.Series) -> Any:
    values = series.dropna()
    if values.empty:
        return pd.NA
    return values.iloc[0]


def last_non_null(series: pd.Series) -> Any:
    values = series.dropna()
    if values.empty:
        return pd.NA
    return values.iloc[-1]


def normalize_value_description(value: Any) -> str:
    text = str(value).lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def wanted_value_keys() -> set[str]:
    return {
        normalize_value_description(alias)
        for aliases in VALUE_ALIASES.values()
        for alias in aliases
    }


def drop_incomplete_initial_cycles(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df.empty or REEFER_UNITS_COLUMN not in report_df.columns:
        return report_df

    trimmed_vessels = []
    sort_columns = ["End Date & Time GMT", "Start Date & Time GMT"]

    for _, vessel_df in report_df.groupby("Ship Name", sort=False):
        vessel_df = vessel_df.sort_values(sort_columns)
        report_type = vessel_df["Report Type"].astype("string")
        has_departure = report_type.str.contains(
            "Departure Report",
            case=False,
            na=False,
        )
        has_units = vessel_df[REEFER_UNITS_COLUMN].notna()
        start_candidates = vessel_df.loc[has_departure & has_units].index

        if len(start_candidates) == 0:
            start_candidates = vessel_df.loc[has_units].index

        if len(start_candidates) > 0:
            start_position = vessel_df.index.get_loc(start_candidates[0])
            vessel_df = vessel_df.iloc[start_position:]

        trimmed_vessels.append(vessel_df)

    if not trimmed_vessels:
        return report_df

    return pd.concat(trimmed_vessels).sort_values(
        ["Ship Name", "End Date & Time GMT", "Start Date & Time GMT"]
    )


@st.cache_data(ttl=21600, show_spinner=False)
def transform_report_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    missing_columns = sorted(set(SOURCE_COLUMNS).difference(raw_df.columns))
    if missing_columns:
        raise ValueError(f"Missing expected API columns: {', '.join(missing_columns)}")

    df = raw_df.copy()
    df["_value_key"] = df["ValueDescription"].map(normalize_value_description)
    df = df[
        df["ValueDescription"].notna()
        & df["_value_key"].isin(wanted_value_keys())
        & ~df["ReportType"].isin(EXCLUDED_REPORT_TYPES)
    ].copy()

    if df.empty:
        return pd.DataFrame(columns=DISPLAY_COLUMNS)

    df["_source_order"] = range(len(df))
    df["StartDateTimeGMT"] = parse_datetime_series(df["StartDateTimeGMT"])
    df["EndDateTimeGMT"] = parse_datetime_series(df["EndDateTimeGMT"])
    df["LapTime"] = parse_numeric_series(df["LapTime"])
    df["ParsedValue"] = parse_numeric_series(df["ReportedValue"])

    report_df = build_report_rows(df)
    report_df = report_df.rename(columns=RENAME_COLUMNS)
    report_df = report_df.sort_values(
        ["Ship Name", "End Date & Time GMT", "Start Date & Time GMT"]
    )
    report_df = drop_incomplete_initial_cycles(report_df)

    report_df[REEFER_UNITS_COLUMN] = report_df.groupby("Ship Name", sort=False)[
        REEFER_UNITS_COLUMN
    ].ffill()
    report_df["Reefer Daily Average Power [kW]"] = report_df.apply(
        lambda row: calculate_reefer_daily_average_power(
            row.get("Reefer Energy [kWh]"),
            row.get("Lap Time"),
        ),
        axis=1,
    )
    report_df["Average Power per Reefer unit [kW]"] = (
        report_df["Reefer Daily Average Power [kW]"] / report_df[REEFER_UNITS_COLUMN]
    ).where(report_df[REEFER_UNITS_COLUMN].gt(0))
    report_df["Generators in Use"] = report_df.apply(format_generators_in_use, axis=1)
    report_df["Remarks"] = report_df.apply(format_basic_remark, axis=1)
    report_df["Remark Sequence"] = build_remark_sequences(report_df)

    return report_df




@st.cache_data(ttl=21600, show_spinner=False)
def load_dashboard_data(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    days_back: int,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Load and transform API data once, keeping the API load timestamp in cache."""
    raw_df = fetch_report_data(
        username=username,
        password=password,
        token=token,
        auth_method=auth_method,
        days_back=days_back,
    )
    df = transform_report_data(raw_df)
    load_meta = {
        "last_api_load_local": current_local_api_load_time(),
        "api_rows": int(len(raw_df)),
        "dashboard_rows": int(len(df)),
    }
    return raw_df, df, load_meta


def load_dashboard_data_fresh(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    days_back: int,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Load a fresh API snapshot without clearing the active cache first.

    This is used by manual refresh and warmup force-refresh so the currently
    displayed data remains available if the new API request fails.
    """
    raw_loader = getattr(fetch_report_data, "__wrapped__", fetch_report_data)
    transform_loader = getattr(transform_report_data, "__wrapped__", transform_report_data)

    raw_df = raw_loader(
        username=username,
        password=password,
        token=token,
        auth_method=auth_method,
        days_back=days_back,
    )
    df = transform_loader(raw_df)
    load_meta = {
        "last_api_load_local": current_local_api_load_time(),
        "api_rows": int(len(raw_df)),
        "dashboard_rows": int(len(df)),
        "fresh_load": True,
    }
    return raw_df, df, load_meta


def set_active_dashboard_data(
    raw_df: pd.DataFrame,
    df: pd.DataFrame,
    load_meta: dict[str, Any],
) -> None:
    """Store the active dataset for this browser session after a successful load."""
    st.session_state["reefer_raw_df"] = raw_df
    st.session_state["reefer_df"] = df
    st.session_state["reefer_load_meta"] = load_meta


def get_active_dashboard_data() -> tuple[pd.DataFrame | None, pd.DataFrame | None, dict[str, Any] | None]:
    return (
        st.session_state.get("reefer_raw_df"),
        st.session_state.get("reefer_df"),
        st.session_state.get("reefer_load_meta"),
    )


def build_report_rows(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    grouped = df.sort_values("_source_order").groupby(
        REPORT_GROUP_KEYS,
        sort=False,
        dropna=False,
    )

    for keys, group in grouped:
        key_values = keys if isinstance(keys, tuple) else (keys,)
        record = dict(zip(REPORT_GROUP_KEYS, key_values, strict=True))
        record["ReportType"] = last_non_null(group["ReportType"])
        record["StartDateTimeGMT"] = last_non_null(group["StartDateTimeGMT"])
        record["LapTime"] = last_non_null(group["LapTime"])

        for column, aliases in VALUE_ALIASES.items():
            alias_keys = {normalize_value_description(alias) for alias in aliases}
            matching_values = group.loc[
                group["_value_key"].isin(alias_keys),
                "ParsedValue",
            ]
            record[column] = last_non_null(matching_values)

        records.append(record)

    return pd.DataFrame(records)


def calculate_reefer_daily_average_power(
    reefer_energy: Any,
    lap_time: Any,
) -> Any:
    if pd.isna(reefer_energy) or pd.isna(lap_time):
        return pd.NA

    try:
        energy = float(reefer_energy)
        hours = float(lap_time)
    except (TypeError, ValueError):
        return pd.NA

    if hours <= 0:
        return pd.NA

    return energy / hours


def parse_datetime_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", utc=True)
    missing_mask = parsed.isna()

    if missing_mask.any():
        date_text = series.astype("string")
        dotnet_millis = date_text.str.extract(r"/Date\((-?\d+)").iloc[:, 0]
        dotnet_parsed = pd.to_datetime(
            pd.to_numeric(dotnet_millis, errors="coerce"),
            errors="coerce",
            unit="ms",
            utc=True,
        )
        parsed = parsed.mask(missing_mask, dotnet_parsed)

    return parsed


def parse_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.map(parse_numeric_value), errors="coerce")


def parse_numeric_value(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return pd.NA

    duration_match = re.fullmatch(r"(-?\d+):([0-5]?\d)(?::([0-5]?\d))?", text)
    if duration_match:
        hours = int(duration_match.group(1))
        sign = -1 if hours < 0 else 1
        minutes = int(duration_match.group(2))
        seconds = int(duration_match.group(3) or 0)
        return sign * (abs(hours) + minutes / 60 + seconds / 3600)

    numeric_text = text.replace(" ", "")
    if re.fullmatch(r"-?\d+,\d+", numeric_text):
        numeric_text = numeric_text.replace(",", ".")
    else:
        numeric_text = numeric_text.replace(",", "")

    numeric_text = re.sub(r"[^0-9.\-]", "", numeric_text)
    if numeric_text in {"", "-", ".", "-."}:
        return pd.NA

    try:
        return float(numeric_text)
    except ValueError:
        return pd.NA


def format_generators_in_use(row: pd.Series) -> Any:
    generators = []
    for generator_number in range(1, 5):
        column = f"DG{generator_number} Running Hours [hh:mm]"
        value = row.get(column)
        if pd.notna(value) and float(value) > 0:
            generators.append(f"No.{generator_number}")

    if not generators:
        return pd.NA

    return f"{len(generators)} ({', '.join(generators)})"


def format_basic_remark(row: pd.Series) -> str:
    report_type = row.get("Report Type")
    lap_time = row.get("Lap Time")
    if pd.isna(lap_time):
        return str(report_type)
    return f"{report_type}, {float(lap_time):g} Hrs"


def format_report_transition(previous_report_type: Any, row: pd.Series) -> str:
    lap_time = row.get("Lap Time")
    lap_text = "Unknown Hrs" if pd.isna(lap_time) else f"{float(lap_time):g} Hrs"
    return f"From {previous_report_type} to {row['Report Type']}, {lap_text}"


def build_remark_sequences(df: pd.DataFrame) -> pd.Series:
    sequences = pd.Series(index=df.index, dtype="object")
    for _, vessel_df in df.groupby("Ship Name", sort=False):
        vessel_df = vessel_df.sort_values(
            ["End Date & Time GMT", "Start Date & Time GMT"]
        )
        previous_report_type = vessel_df["Report Type"].shift(1)

        for idx, row in vessel_df.iterrows():
            if pd.isna(previous_report_type.loc[idx]):
                sequences.loc[idx] = format_basic_remark(row)
            else:
                sequences.loc[idx] = format_report_transition(
                    previous_report_type.loc[idx],
                    row,
                )
    return sequences


def select_report_row(
    vessel_df: pd.DataFrame,
    selected_report_time: pd.Timestamp,
) -> pd.Series:
    sorted_vessel_df = vessel_df.sort_values(
        ["End Date & Time GMT", "Start Date & Time GMT"]
    )
    rows_for_time = sorted_vessel_df[
        sorted_vessel_df["End Date & Time GMT"] == selected_report_time
    ]
    if rows_for_time.empty:
        rows_for_time = sorted_vessel_df

    selected_index = rows_for_time.index[-1]
    selected_position = sorted_vessel_df.index.get_loc(selected_index)
    selected_row = sorted_vessel_df.loc[selected_index].copy()

    for column in DAILY_COALESCE_COLUMNS:
        if column not in rows_for_time.columns:
            continue

        values = rows_for_time[column].dropna()
        if values.empty:
            continue

        if column in GENERATOR_COLUMNS:
            selected_row[column] = values.max()
        elif pd.isna(selected_row.get(column)):
            selected_row[column] = values.iloc[-1]

    reefer_units = selected_row.get(REEFER_UNITS_COLUMN)
    selected_row["Reefer Daily Average Power [kW]"] = (
        calculate_reefer_daily_average_power(
            selected_row.get("Reefer Energy [kWh]"),
            selected_row.get("Lap Time"),
        )
    )

    reefer_daily_average_power = selected_row.get("Reefer Daily Average Power [kW]")
    if (
        pd.notna(reefer_units)
        and float(reefer_units) > 0
        and pd.notna(reefer_daily_average_power)
    ):
        selected_row["Average Power per Reefer unit [kW]"] = (
            float(reefer_daily_average_power) / float(reefer_units)
        )

    selected_row["Generators in Use"] = format_generators_in_use(selected_row)
    if selected_position > 0:
        previous_row = sorted_vessel_df.iloc[selected_position - 1]
        selected_row["Remark Sequence"] = format_report_transition(
            previous_row["Report Type"],
            selected_row,
        )
    else:
        selected_row["Remark Sequence"] = format_basic_remark(selected_row)

    return selected_row


def format_metric_value(value: Any, suffix: str = "", decimals: int = 1) -> str:
    if pd.isna(value):
        return "-"
    if isinstance(value, str):
        return value
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)

    if decimals == 0:
        return f"{number:,.0f}{suffix}"
    return f"{number:,.{decimals}f}{suffix}"


def show_kpis(selected_row: pd.Series) -> None:
    first_row = st.columns(3)
    second_row = st.columns(3)
    third_row = st.columns(3)

    first_row[0].metric(
        "Total Daily Average Power",
        format_metric_value(selected_row.get("Total Daily Average Power [kW]"), " kW"),
    )
    first_row[1].metric(
        "Reefer Energy",
        format_metric_value(selected_row.get("Reefer Energy [kWh]"), " kWh"),
    )
    first_row[2].metric(
        "Lap Time",
        format_metric_value(selected_row.get("Lap Time"), " Hrs"),
    )

    second_row[0].metric(
        "Reefer Daily Average Power",
        format_metric_value(selected_row.get("Reefer Daily Average Power [kW]"), " kW"),
    )
    second_row[1].metric(
        "Reefer Power",
        format_metric_value(selected_row.get("Reefer Power [kW]"), " kW"),
    )
    second_row[2].metric(
        "Total Number Reefer Units",
        format_metric_value(selected_row.get(REEFER_UNITS_COLUMN), " Reefers", 0),
    )

    third_row[0].metric(
        "Average Power / Reefer Unit",
        format_metric_value(
            selected_row.get("Average Power per Reefer unit [kW]"), " kW", 2
        ),
    )
    third_row[1].metric(
        "Generators in Use",
        format_metric_value(selected_row.get("Generators in Use")),
    )
    third_row[2].metric("Report Type", selected_row.get("Report Type", "-"))


def available_vessels(df: pd.DataFrame) -> list[str]:
    return sorted(df["Ship Name"].dropna().unique().tolist())


def available_report_times_for_vessel(df: pd.DataFrame, vessel: str) -> list[pd.Timestamp]:
    vessel_df = df[df["Ship Name"] == vessel]
    return sorted(vessel_df["End Date & Time GMT"].dropna().unique().tolist(), reverse=True)


def format_report_time_option(vessel_df: pd.DataFrame, report_time: pd.Timestamp) -> str:
    return report_time.strftime("%d-%m-%Y %H:%M")


def display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    columns = [column for column in DISPLAY_COLUMNS if column in df.columns]
    display_df = df[columns].copy()
    for column in DATETIME_COLUMNS:
        if column in display_df.columns:
            display_df[column] = display_df[column].dt.strftime("%Y-%m-%d %H:%M")
    return display_df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reefer History")
        worksheet = writer.sheets["Reefer History"]
        for column_cells in worksheet.columns:
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in column_cells
            )
            adjusted_width = min(max(max_length + 2, 12), 45)
            worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width
    return output.getvalue()


def sidebar_controls(df: pd.DataFrame) -> tuple[str, pd.Timestamp]:
    vessels = available_vessels(df)
    if not vessels:
        st.stop()

    selected_vessel = st.sidebar.selectbox("Ship Name", vessels)
    available_report_times = available_report_times_for_vessel(df, selected_vessel)

    if not available_report_times:
        st.sidebar.warning("No report times available for the selected vessel.")
        st.stop()

    selected_report_time = st.sidebar.selectbox(
        "End Date & Time GMT",
        available_report_times,
        index=0,
        format_func=lambda value: format_report_time_option(df, value),
    )

    return selected_vessel, selected_report_time


def render_header(selected_vessel: str, selected_report_time: pd.Timestamp) -> None:
    st.markdown(
        f"""
        <div class="dashboard-hero">
            <div class="eyebrow">Fleet reefer monitoring</div>
            <h1 class="dashboard-title">Reefer Dashboard</h1>
            <div class="dashboard-subtitle">
                {escape(selected_vessel)} | selected report {selected_report_time.strftime('%d-%m-%Y %H:%M')} GMT
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_remark_card(text: Any) -> None:
    remark = "-" if pd.isna(text) else str(text)
    st.markdown(
        f"""
        <div class="remark-card">
            <div class="remark-label">Remark sequence</div>
            <div class="remark-text">{escape(remark)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_api_load_caption(load_meta: dict[str, Any]) -> None:
    """Render one small, discrete last-load indicator below the dashboard hero."""
    last_load = load_meta.get("last_api_load_local", "-")
    last_load_raw = str(last_load).replace(" EEST", "").replace(" EET", "")

    try:
        last_load_display = pd.to_datetime(last_load_raw).strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        last_load_display = str(last_load)

    st.markdown(
        f"""
        <div class="api-load-caption">
            Last API load: <span>{escape(last_load_display)} LT</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def run_warmup_if_requested() -> None:
    """Warm up API/cache via a secret-token URL without showing the normal dashboard UI."""
    if not is_warmup_request():
        return

    if not warmup_token_is_valid():
        st.error("Invalid or missing warmup token.")
        st.stop()

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")
    days_back = int(read_secret("MARORKA_DAYS_BACK", "10"))

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.error("Warmup failed: MARORKA_USERNAME and MARORKA_PASSWORD are required.")
        st.stop()

    force_refresh = get_query_param("force", "0") == "1"

    try:
        with st.spinner("Warming up API..."):
            if force_refresh:
                # Load fresh data first. Only after a successful API + transform cycle
                # do we clear/reseed Streamlit cache, so a failed warmup does not
                # remove the last good cached dataset.
                raw_df, df, load_meta = load_dashboard_data_fresh(
                    username=username,
                    password=password,
                    token=token,
                    auth_method=auth_method,
                    days_back=days_back,
                )
                fetch_report_data.clear()
                transform_report_data.clear()
                load_dashboard_data.clear()
                # Reseed Streamlit's shared function cache only after the fresh
                # request above has succeeded. This keeps the old cache intact
                # during API outages and makes the next normal user load fast.
                raw_df, df, load_meta = load_dashboard_data(
                    username=username,
                    password=password,
                    token=token,
                    auth_method=auth_method,
                    days_back=days_back,
                )
            else:
                raw_df, df, load_meta = load_dashboard_data(
                    username=username,
                    password=password,
                    token=token,
                    auth_method=auth_method,
                    days_back=days_back,
                )
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        st.error(f"Warmup failed: Marorka API request failed with status {status}.")
        st.stop()
    except (MarorkaConfigError, ValueError, requests.RequestException) as exc:
        st.error(f"Warmup failed: {exc}")
        st.stop()

    st.success("Warmup OK.")
    st.write(
        {
            "last_api_load_local": load_meta.get("last_api_load_local"),
            "api_rows": int(len(raw_df)),
            "dashboard_rows": int(len(df)),
            "force_refresh": force_refresh,
        }
    )
    st.stop()


def main() -> None:
    run_warmup_if_requested()
    require_dashboard_password()
    apply_custom_css()

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")
    days_back = int(read_secret("MARORKA_DAYS_BACK", "10"))

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.info("Add MARORKA_USERNAME and MARORKA_PASSWORD to .streamlit/secrets.toml.")
        st.stop()

    refresh_requested = st.sidebar.button("Refresh API data")
    if refresh_requested:
        st.session_state["confirm_api_refresh"] = True

    if st.session_state.get("confirm_api_refresh"):
        load_meta = st.session_state.get("reefer_load_meta") or {}
        last_load = load_meta.get("last_api_load_local", "-")
        last_load_raw = str(last_load).replace(" EEST", "").replace(" EET", "")

        try:
            last_load_display = pd.to_datetime(last_load_raw).strftime("%d-%m-%Y %H:%M:%S")
        except Exception:
            last_load_display = str(last_load)

        st.sidebar.warning(
            f"Refresh will call the API and may take a while.\n\n"
            f"Last updated data was on: {last_load_display} LT"
        )

        col1, col2 = st.sidebar.columns(2)

        if col1.button("Confirm"):
            try:
                with st.spinner("Refreshing Marorka report data..."):
                    raw_df, df, load_meta = load_dashboard_data_fresh(
                        username=username,
                        password=password,
                        token=token,
                        auth_method=auth_method,
                        days_back=days_back,
                    )
            except requests.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else "unknown"
                st.sidebar.error(f"Refresh failed with status {status}. Existing data was kept.")
            except (MarorkaConfigError, ValueError, requests.RequestException) as exc:
                st.sidebar.error(f"Refresh failed. Existing data was kept. {exc}")
            else:
                set_active_dashboard_data(raw_df, df, load_meta)
                st.session_state["confirm_api_refresh"] = False
                st.rerun()

        if col2.button("Cancel"):
            st.session_state["confirm_api_refresh"] = False
            st.rerun()

    try:
        raw_df, df, load_meta = get_active_dashboard_data()
        if raw_df is None or df is None or load_meta is None:
            with st.spinner("Loading Marorka report data..."):
                raw_df, df, load_meta = load_dashboard_data(
                    username=username,
                    password=password,
                    token=token,
                    auth_method=auth_method,
                    days_back=days_back,
                )
                set_active_dashboard_data(raw_df, df, load_meta)
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        st.error(f"Marorka API request failed with status {status}.")
        st.caption("If credentials are correct, try MARORKA_AUTH_METHOD = 'digest'.")
        st.stop()
    except (MarorkaConfigError, ValueError, requests.RequestException) as exc:
        st.error(str(exc))
        st.stop()

    if df.empty:
        st.warning("No matching Marorka report values were returned for the date window.")
        st.stop()

    selected_vessel, selected_report_time = sidebar_controls(df)

    vessel_df = df[df["Ship Name"] == selected_vessel].sort_values("End Date & Time GMT")
    selected_row = select_report_row(vessel_df, selected_report_time)

    render_header(selected_vessel, selected_report_time)
    render_api_load_caption(load_meta)
    show_kpis(selected_row)
    render_remark_card(selected_row.get("Remark Sequence", selected_row.get("Remarks", "-")))

    st.markdown('<div class="section-title">Historical reports</div>', unsafe_allow_html=True)
    history_df = vessel_df.sort_values("End Date & Time GMT", ascending=False)
    export_df = display_dataframe(history_df)
    st.dataframe(
        export_df,
        use_container_width=True,
        hide_index=True,
    )

    excel_data = to_excel_bytes(export_df)
    st.download_button(
        "Download vessel history as Excel",
        excel_data,
        file_name=f"{selected_vessel.replace(' ', '_')}_reefer_history.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()

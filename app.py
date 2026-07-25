from __future__ import annotations

import html
import json
import re
import time
import uuid
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

from backend import backend_enabled, post_backend
from reporting import build_version_csv, build_versions_pdf, encode_attachment

st.set_page_config(
    page_title="Coal-to-Clean Jurisdictional Readiness Index 2026",
    page_icon=Path(__file__).parent / "assets" / "favicon.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_FILE = Path(__file__).parent / "data" / "Index_Model.xlsx"
SHEET_NAME = "Streamlit"
OVERLAY_SHEET_NAME = "OVERLAY"

PILLAR_SOURCES = [
    "PILLAR 1 score\nEnergy system conditions",
    "PILLAR 2 score\nPolicy and transition commitment",
    "PILLAR 3 score\nGovernance and institutional capacity",
    "PILLAR 4 score\nCarbon market maturity",
    "PILLAR 5 score\nMacro-financial conditions",
    "PILLAR 6 score\nJust transition and social credibility",
]

PILLAR_FULL = {
    PILLAR_SOURCES[0]: "Pillar 1: Energy system conditions",
    PILLAR_SOURCES[1]: "Pillar 2: Policy and transition commitment",
    PILLAR_SOURCES[2]: "Pillar 3: Governance and institutional capacity",
    PILLAR_SOURCES[3]: "Pillar 4: Carbon market maturity",
    PILLAR_SOURCES[4]: "Pillar 5: Macro-financial conditions",
    PILLAR_SOURCES[5]: "Pillar 6: Just transition and social credibility",
}

PILLAR_TABLE = {
    PILLAR_SOURCES[0]: "Pillar: Energy system conditions",
    PILLAR_SOURCES[1]: "Pillar: Policy and transition commitment",
    PILLAR_SOURCES[2]: "Pillar: Governance and institutional capacity",
    PILLAR_SOURCES[3]: "Pillar: Carbon market maturity",
    PILLAR_SOURCES[4]: "Pillar: Macro-financial conditions",
    PILLAR_SOURCES[5]: "Pillar: Just transition and social credibility",
}

PILLAR_SHORT = [
    "Energy system conditions",
    "Policy and transition commitment",
    "Governance and institutional capacity",
    "Carbon market maturity",
    "Macro-financial conditions",
    "Just transition and social credibility",
]

PILLAR_DESCRIPTIONS = [
    "Signals whether the power system can replace coal while maintaining reliable electricity supply.",
    "Signals whether announced transition policies, coal commitments and regulatory measures are likely to be sustained and implemented.",
    "Signals whether institutions, regulation and public administration can support credible project delivery.",
    "Signals experience in developing carbon projects, issuing credits, transferring units and reaching buyers.",
    "Signals whether financing, currency and payment conditions support investable and bankable projects.",
    "Signals whether worker, community and social-protection risks are recognised and managed during coal phase-down.",
]

BASE_WEIGHTS = [0.20, 0.20, 0.15, 0.20, 0.15, 0.10]

RECOMMENDATIONS = {
    PILLAR_SOURCES[0]: "Test coal-retirement feasibility, replacement capacity and grid-readiness assumptions at project level.",
    PILLAR_SOURCES[1]: "Test whether policy commitments translate into durable approvals, implementation and credible retirement pathways.",
    PILLAR_SOURCES[2]: "Strengthen regulatory, institutional and counterparty safeguards before committing capital or delivery obligations.",
    PILLAR_SOURCES[3]: "Verify project-development capability, issuance delivery, transfer infrastructure and evidence of buyer use.",
    PILLAR_SOURCES[4]: "Assess currency, financing and payment risks, including suitable risk-mitigation and capital structures.",
    PILLAR_SOURCES[5]: "Examine worker, community and social-protection arrangements within the transition plan.",
}

INDUSTRY_MULTIPLIERS = {
    "Agriculture": [0.95, 1.00, 1.00, 1.10, 1.00, 1.10],
    "Aviation": [0.95, 1.05, 1.00, 1.25, 1.05, 0.95],
    "Automotive": [1.05, 1.05, 1.00, 1.00, 1.10, 1.00],
    "Banking and financial services": [1.00, 1.05, 1.10, 1.10, 1.25, 0.90],
    "Carbon markets and climate services": [0.90, 1.10, 1.05, 1.35, 1.00, 0.90],
    "Cement and building materials": [1.15, 1.10, 1.00, 0.90, 1.15, 0.90],
    "Chemicals": [1.10, 1.05, 1.00, 0.95, 1.15, 0.95],
    "Construction": [1.10, 1.05, 1.00, 0.95, 1.10, 1.00],
    "Consumer goods": [0.95, 1.00, 1.00, 1.15, 1.00, 1.05],
    "Data centres": [1.10, 1.00, 1.00, 1.10, 1.00, 0.95],
    "Development finance and multilateral institutions": [1.00, 1.15, 1.20, 1.05, 1.20, 1.10],
    "Food and beverage": [0.95, 1.00, 1.00, 1.10, 1.00, 1.10],
    "Healthcare": [0.95, 1.00, 1.00, 1.05, 1.00, 1.05],
    "Hospitality and tourism": [0.95, 1.00, 1.00, 1.15, 1.00, 1.05],
    "Infrastructure": [1.10, 1.05, 1.00, 0.95, 1.10, 1.00],
    "Insurance": [0.95, 1.05, 1.10, 1.05, 1.25, 1.00],
    "Manufacturing": [1.10, 1.05, 1.00, 0.95, 1.15, 0.95],
    "Metals and steel": [1.15, 1.05, 1.00, 0.95, 1.15, 0.95],
    "Mining": [1.15, 1.05, 1.00, 0.95, 1.15, 0.95],
    "Mobility and fleet services": [1.05, 1.00, 1.00, 1.00, 1.10, 1.00],
    "Oil and gas": [1.15, 1.05, 1.00, 0.95, 1.10, 1.00],
    "Pharmaceuticals": [0.95, 1.00, 1.00, 1.05, 1.00, 1.05],
    "Power generation": [1.20, 1.10, 1.00, 0.90, 1.00, 1.10],
    "Professional services": [1.00, 1.05, 1.10, 1.05, 1.05, 1.00],
    "Public transport and rail": [1.10, 1.05, 1.00, 0.95, 1.10, 1.05],
    "Pulp and paper": [1.10, 1.05, 1.00, 0.95, 1.10, 1.00],
    "Real estate": [1.05, 1.05, 1.00, 0.95, 1.10, 1.00],
    "Retail": [0.95, 1.00, 1.00, 1.10, 1.00, 1.05],
    "Road freight and logistics": [1.00, 1.05, 1.00, 1.05, 1.10, 0.95],
    "Semiconductors and electronics": [1.05, 1.00, 1.00, 1.05, 1.10, 0.95],
    "Shipping": [1.00, 1.05, 1.00, 1.15, 1.10, 0.95],
    "Technology": [1.05, 1.00, 1.00, 1.10, 1.00, 0.95],
    "Telecommunications": [1.05, 1.00, 1.00, 1.05, 1.00, 1.00],
    "Textiles and apparel": [0.95, 1.00, 1.00, 1.10, 1.00, 1.05],
    "Utilities and energy networks": [1.15, 1.10, 1.00, 0.90, 1.00, 1.10],
    "Waste management": [1.00, 1.00, 1.00, 1.10, 1.00, 1.05],
    "Other": [1.00, 1.00, 1.00, 1.00, 1.00, 1.00],
}

ROLE_OPTIONS = [
    "Corporate buyer or end user",
    "Market intermediary or adviser",
    "Investor or lender",
    "Project developer or originator",
    "Government or regulator",
    "Standard, registry or assurance provider",
    "Civil society, research or academia",
    "Asset owner or operator",
    "Other",
]

SURVEY_ITEMS = [
    {
        "key": "reliable_power",
        "label": "Reliable power during the transition",
        "description": "Keeping electricity dependable while coal is replaced by renewable or other lower-emissions power.",
        "pillars": [0],
    },
    {
        "key": "government_follow_through",
        "label": "Government follow-through",
        "description": "The host government implements and maintains its announced energy-transition policies and regulations.",
        "pillars": [1],
    },
    {
        "key": "project_delivery",
        "label": "Ability to deliver the project",
        "description": "Coal-to-clean projects can move from planning through construction, operation and credit issuance.",
        "pillars": [2],
    },
    {
        "key": "carbon_market_track_record",
        "label": "Carbon-market track record",
        "description": "The host country has credible experience developing projects, issuing credits and reaching buyers.",
        "pillars": [3],
    },
    {
        "key": "stable_finance",
        "label": "Stable finance and payments",
        "description": "The host country offers workable financing, currency and payment conditions for project delivery.",
        "pillars": [4],
    },
    {
        "key": "workers_communities",
        "label": "Protection for workers and communities",
        "description": "The transition addresses impacts on workers, communities and local livelihoods affected by coal phase-down.",
        "pillars": [5],
    },
    {
        "key": "time_to_delivery",
        "label": "Time to delivery",
        "description": "Projects can reach operation and issue credits within a predictable timeframe.",
        "pillars": [1, 2],
    },
    {
        "key": "credit_price",
        "label": "Credit price",
        "description": "The expected carbon-credit price is competitive enough to influence your decision.",
        "pillars": [4],
    },
]

REQUIRED_COLUMNS = [
    "Country",
    "ISO3",
    "Eligible",
    "Base index rank",
    "Base index score",
    *PILLAR_SOURCES,
]

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');

:root {
    --main-surface: rgba(236, 242, 249, .55);
    --sidebar: #9fbedf;
    --ink: rgba(0, 0, 0, .90);
    --ink-muted: rgba(0, 0, 0, .62);
    --accent: #2b5688;
    --accent-mid: #79a6d2;
    --accent-soft: #8bb0da;
    --accent-pale: #c4d7ed;
    --card: rgba(255, 255, 255, .94);
    --card-border: rgba(43, 86, 136, .13);
    --slider-red: #cf4f57;
    --warm-accent: #fff5cc;
    --warm-accent-strong: #ffe680;
}


html { scroll-behavior: smooth; }
body, .stApp { font-family: 'Montserrat', sans-serif; }
.stApp { background: var(--main-surface); color: var(--ink); }
.stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp label, .stApp button, .stApp input, .stApp textarea,
.stApp [data-testid="stMarkdownContainer"] { font-family: 'Montserrat', sans-serif !important; }
header[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }
.block-container { max-width: 1460px; padding-top: 2.2rem; padding-bottom: 4rem; }

[data-testid="stSidebar"] {
    background: var(--sidebar);
    border-right: 1px solid rgba(43, 86, 136, .10);
}
[data-testid="stSidebar"] > div:first-child { padding-top: 2.8rem; }
[data-testid="stSidebarNav"] { display: none; }
[data-testid="stSidebar"] [data-testid="stIconMaterial"],
[data-testid="stSidebar"] span[class*="material-symbols"] {
    font-family: 'Material Symbols Rounded', 'Material Symbols Outlined' !important;
}
.sidebar-nav a,
.sidebar-methodology a {
    display: block;
    color: #ffffff !important;
    text-decoration: none !important;
    font-size: 14px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .03em;
    transition: opacity .15s ease, transform .15s ease;
}
.sidebar-nav a { margin: 0 0 20px 0; }
.sidebar-nav a:nth-child(4) { margin-top: 24px; }
.sidebar-nav a:hover,
.sidebar-methodology a:hover { color: #ffffff !important; opacity: .78; transform: translateX(3px); }
.sidebar-methodology { position: fixed; bottom: 64px; width: 235px; }


[data-testid="stSidebar"] .stButton > button {
    background: transparent !important; border: 0 !important; box-shadow: none !important;
    color: #ffffff !important; padding: 0 !important; min-height: 0 !important; height: auto !important;
    justify-content: flex-start !important; font-size: 14px !important; font-weight: 300 !important;
    text-transform: uppercase; letter-spacing: .03em; border-radius: 0 !important;
}
[data-testid="stSidebar"] .stButton > button:hover { color:#ffffff !important; opacity:.78; transform:translateX(3px); }
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0 !important; }
[data-testid="stSidebar"] [data-testid="stButton"] { margin-bottom: 10px; }
.st-key-sidebar_priorities_gap {
    height: 30px;
    margin-top: 50px;
    border-top: 5px solid #FFF5CCAA;
    padding: 10px;
    width: 50%;
}
.st-key-sidebar_methodology_button { position:fixed; bottom:64px; width:235px; }

.st-key-top_nav [data-testid="stHorizontalBlock"] { align-items:flex-start; }
.st-key-top_nav [data-testid="column"] { min-width:0 !important; }
.st-key-top_nav .stButton > button {
    position:relative; background:transparent !important; border:0 !important; box-shadow:none !important;
    border-radius:0 !important; padding:0 0 15px 0 !important; min-height:0 !important; height:auto !important;
    color:var(--ink-muted) !important; font-size:12px !important; font-weight:600 !important;
    letter-spacing:.04em; text-transform:uppercase; white-space:nowrap;
}
.st-key-top_nav .stButton > button:hover { color:var(--accent) !important; box-shadow:inset 0 -5px 0 var(--warm-accent) !important; }
[data-testid="stSidebar"] .stButton > button *,
.st-key-top_nav .stButton > button * {
    font-weight: 700 !important;
}
.st-key-nav_overall_active .stButton > button,
.st-key-nav_market_active .stButton > button,
.st-key-nav_priorities_active .stButton > button {
    color:var(--accent) !important; box-shadow:inset 0 -5px 0 var(--warm-accent) !important;
}
.st-key-period_slot { height:55px; padding-top:10px; }
.st-key-period_slot [data-testid="stHorizontalBlock"] { align-items:flex-start; }
.st-key-period_slot .stButton > button {
    border:1px solid rgba(43,86,136,.25) !important; border-radius:999px !important;
    background:rgba(255,255,255,.48) !important; color:var(--accent) !important;
    padding:7px 13px !important; min-height:0 !important; height:auto !important;
    font-size:12px !important; font-weight:600 !important; white-space:nowrap;
}
.st-key-period_slot .stButton > button:hover,
.st-key-period_q1_active .stButton > button,
.st-key-period_q2_active .stButton > button { background:var(--accent-pale) !important; }
.st-key-period_clear .stButton > button { border:0 !important; background:transparent !important; text-decoration:underline !important; }

.main-title-row { display:flex; align-items:baseline; gap:8px; position:relative; width:100%; }
.main-title {
    color: var(--ink);
    font-weight: 600;
    font-size: clamp(31px, 2.45vw, 39px);
    line-height: 1.12;
    letter-spacing: -.025em;
    margin: 0;
}
.main-subtitle { font-size: 17px; color: var(--ink-muted); margin: 10px 0 30px 0; }
.info-details { display:inline-block; position:relative; }
.info-details summary {
    list-style:none; cursor:pointer; color:var(--accent-mid); font-size:19px; line-height:1;
    user-select:none; outline:none;
}
.info-details summary::-webkit-details-marker { display:none; }
.info-details summary:hover { color:var(--accent); }
.info-panel {
    position:absolute; right:0; top:30px; z-index:1000; width:min(520px, 82vw);
    padding:18px 20px; background:#ffffff; border:1px solid rgba(43,86,136,.18);
    border-radius:12px; box-shadow:0 18px 50px rgba(43,86,136,.18);
    font-size:15px !important; line-height:1.58; color:var(--ink);
}
.info-panel p { margin:0 0 10px 0; font-size:15px !important; }
.info-panel a { color:var(--ink) !important; font-weight:700; text-decoration:underline !important; text-decoration-color:var(--warm-accent-strong) !important; text-decoration-thickness:2px !important; text-underline-offset:3px; }
.info-panel p:last-child { margin-bottom:0; }

.top-nav {
    display: flex; flex-wrap: wrap; align-items: flex-start; gap: 42px; margin: 0 0 4px 0;
}
.top-nav a {
    position: relative; display: inline-block; padding-bottom: 15px; color: var(--ink-muted) !important;
    text-decoration: none !important; font-size: 12px; font-weight: 600; letter-spacing: .04em;
    text-transform: uppercase; cursor: pointer;
}
.top-nav a::after {
    content: ''; position: absolute; left: 0; bottom: 0; width: 100%; height: 5px; border-radius: 99px;
    background: var(--warm-accent); transform: scaleX(0); transform-origin: left; transition: transform .16s ease;
}
.top-nav a:hover, .top-nav a.active { color: var(--accent) !important; }
.top-nav a:hover::after, .top-nav a.active::after { transform: scaleX(1); }
.period-slot { height: 55px; display: flex; align-items: flex-start; padding-top:10px; }
.period-nav { display: flex; gap: 10px; margin-left: 120px; align-items: center; }
.period-nav a {
    display: inline-block; padding: 7px 13px; border: 1px solid rgba(43, 86, 136, .25);
    border-radius: 999px; color: var(--accent) !important; text-decoration: none !important;
    font-size: 12px; font-weight: 600; background: rgba(255, 255, 255, .48);
}
.period-nav a:hover, .period-nav a.active { background: var(--accent-pale); }
.period-nav a.clear { border: 0; background: transparent; text-decoration: underline !important; }
.period-placeholder { visibility: hidden; height: 55px; }

.section-title { color: var(--ink); font-weight: 600; font-size: 26px; line-height: 1.25; margin: 0 0 12px 0; }
.section-kicker { color: var(--ink-muted); font-size: 11px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; margin: -2px 0 12px 0; }
.section-gap { height: 92px; }
.anchor { display: block; position: relative; top: -18px; visibility: hidden; }
.st-key-chart_header_slot { height: auto; min-height: 0; padding-top: 12px; padding-bottom: 0; overflow: visible; }
.st-key-chart_header_slot .section-title { margin-bottom: 5px; }
.st-key-chart_header_slot .base-chart-title { margin-bottom: 10px !important; }
.st-key-chart_header_slot [data-testid="stToggle"] { margin-top: -2px !important; margin-bottom: -18px !important; }
.st-key-ranking_chart_slot { margin-top: -4px !important; }
.st-key-ranking_chart_slot [data-testid="stPlotlyChart"] { margin-top: 0 !important; padding-top: 0 !important; }

.st-key-ranking_section, .st-key-table_section, .st-key-assessment_section, .st-key-priorities_section, .st-key-methodology_section {
    background: var(--card); border: 1px solid var(--card-border); border-radius: 16px; padding: 28px;
    box-shadow: 0 8px 24px rgba(43, 86, 136, .05);
}
.st-key-assessment_section { padding-bottom: 50px; }
.st-key-priorities_section [data-testid="stForm"] { padding: 50px !important; }
.st-key-priorities_section [data-testid="stFormSubmitButton"] { margin-top: 34px; }

.stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {
    border-radius: 8px !important; border-color: rgba(43, 86, 136, .30) !important; font-weight: 600 !important;
}
.stButton > button:hover, .stFormSubmitButton > button:hover, .stDownloadButton > button:hover {
    border-color: var(--accent) !important; color: var(--accent) !important;
}
[data-testid="stToggle"] label { color: var(--ink) !important; }
[data-testid="stToggle"]:hover div[role="checkbox"] { background: var(--accent-pale) !important; }
[data-testid="stToggle"] div[role="checkbox"][aria-checked="true"] { background: var(--accent) !important; }

[data-testid="stElementToolbar"] { opacity: 1 !important; background: transparent !important; }
[data-testid="stElementToolbar"] button { color: var(--accent) !important; background: transparent !important; opacity: 1 !important; }
[data-testid="stElementToolbar"] button:hover { background: var(--warm-accent) !important; }
.modebar-container, .modebar { opacity: 1 !important; background: transparent !important; }
.modebar-group { background: transparent !important; }
.modebar-btn { background: transparent !important; border-radius: 6px !important; }
.modebar-btn:hover { background: var(--warm-accent) !important; }
.modebar-btn svg { width: 19px !important; height: 19px !important; fill: var(--accent) !important; }
.modebar-btn:hover svg { fill: var(--accent) !important; }

.priority-form-space { height: 38px; }
.priority-label { color: var(--ink); font-size: 18px; font-weight: 600; margin: 18px 0 3px 0; }
.priority-description { color: var(--ink-muted); font-size: 13px; line-height: 1.5; margin: 0 0 4px 0; max-width: 900px; }
.priority-scale-note {
    display: flex; justify-content: space-between; color: var(--ink-muted); font-size: 10px; font-weight: 600;
    letter-spacing: .07em; text-transform: uppercase; margin: 10px 0 2px 0;
}
[data-testid="stSlider"] { margin-bottom: 1rem; }
[data-testid="stSlider"] [data-baseweb="slider"] > div > div {
    min-height: 11px !important; border-radius: 99px !important; background: var(--slider-red) !important;
}
[data-testid="stSlider"] [role="slider"] {
    width: 27px !important; height: 27px !important; background: var(--slider-red) !important;
    border: 3px solid #ffffff !important; box-shadow: 0 0 0 2px rgba(207, 79, 87, .25) !important;
}

.metric-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin: 10px 0 34px 0; }
.metric-card { background: rgba(121, 166, 210, .16); border: 1px solid rgba(43, 86, 136, .13); border-radius: 10px; padding: 14px 16px; }
.metric-label { color: var(--ink-muted); font-size: 12px; font-weight: 600; margin-bottom: 5px; }
.metric-value { color: var(--accent); font-size: 28px; font-weight: 600; line-height: 1.1; }
.assessment-heading, .priority-heading { color: var(--ink); font-size: 17px; font-weight: 600; margin: 0 0 4px 0; }
.assessment-copy { color: var(--ink); font-size: 14px; line-height: 1.68; }
.assessment-copy strong { color: var(--accent); font-weight: 700; }
.priority-heading { margin-top: 30px; margin-bottom: 12px; }
.priority-card {
    padding: 16px 18px; border-left: 5px solid var(--accent-mid); background: rgba(121, 166, 210, .13);
    border-radius: 0 9px 9px 0; color: var(--ink); font-size: 14px; line-height: 1.55; margin: 0;
}
.priority-card p { margin: 0; }
.priority-card p + p { margin-top: 12px; }
.st-key-country_picker { width: 100%; max-width: 100%; }
.st-key-country_picker div[data-baseweb="select"],
.st-key-industry_picker div[data-baseweb="select"],
.st-key-role_picker div[data-baseweb="select"] { font-size: 16px; }
.st-key-country_picker div[data-baseweb="select"] > div,
.st-key-industry_picker div[data-baseweb="select"] > div,
.st-key-role_picker div[data-baseweb="select"] > div { background: var(--warm-accent) !important; }
.st-key-assessment_bottom_row [data-testid="stHorizontalBlock"] { align-items: stretch; }
.st-key-assessment_bottom_row [data-testid="column"] { display: flex; flex-direction: column; }
.st-key-priority_panel, .st-key-pillar_explanation_panel { height: 100%; display: flex; flex-direction: column; justify-content: flex-end; }
.pillar-legend-box {
    margin-top: 0; padding: 18px 20px; border: 1px solid var(--warm-accent-strong); border-radius: 10px;
    display: grid; grid-template-columns: 1fr; gap: 15px; font-size: 12px; line-height: 1.4; color: var(--ink);
}
.pillar-definition strong { display: block; margin-bottom: 5px; }
.pillar-definition p { margin: 0; color: var(--ink-muted); line-height: 1.5; }
.assessment-footnote { margin-top:28px; padding-top:14px; border-top:1px solid rgba(43,86,136,.12); color:var(--ink-muted); font-size:11px; line-height:1.45; }

.action-progress-wrap {
    min-height: 46px;
    margin: 18px 0 10px 0;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
    color: var(--ink-muted);
    font-size: 13px;
    font-weight: 500;
}
.action-progress-spinner {
    width: 24px;
    height: 24px;
    border: 2px solid rgba(43, 86, 136, .16);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: priority-spin .75s linear infinite;
}
.st-key-custom_actions {
    position: sticky;
    bottom: 12px;
    z-index: 200;
    margin-top: 6px;
    padding: 18px 0 6px 0;
    background: rgba(255,255,255,.98);
    border: 0;
    border-top: 1px solid var(--warm-accent-strong);
    box-shadow: none;
    border-radius: 0;
}
.st-key-custom_actions .stButton > button,
.st-key-custom_actions .stDownloadButton > button { width: 100% !important; min-width: 132px !important; }
.action-warning {
    max-width: 620px; margin: 0 auto 8px auto; padding: 9px 12px; border-radius: 7px;
    background: #fffae6; color: var(--ink); font-size: 12px; line-height: 1.45; text-align: center;
}
.action-message {
    max-width: 760px;
    margin: 20px auto;
    padding: 0;
    border: 0 !important;
    border-radius: 0;
    background: transparent !important;
    color: var(--accent) !important;
    font-size: 12px;
    font-weight: 600;
    line-height: 1.45;
    text-align: center;
}
.action-message.success,
.action-message.error,
.action-message.warning {
    border: 0 !important;
    background: transparent !important;
    color: var(--accent) !important;
}
.saved-line { color: var(--ink-muted); font-size: 12px; margin: 12px 0 0 0; text-align: center; }
.form-note, .backend-note { color: var(--ink-muted); font-size: 12px; }
.backend-note { padding: 10px 12px; background: rgba(121, 166, 210, .14); border-radius: 7px; }
.gate-copy { font-size:14px; line-height:1.55; color:var(--ink); margin-bottom:10px; }
.gate-fine-print { font-size:10px; line-height:1.45; color:var(--ink-muted); margin-top:10px; margin-bottom:24px; }
.gate-progress-wrap {
    min-height: 180px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 14px;
    color: var(--accent);
    font-size: 13px;
    font-weight: 600;
    text-align: center;
}
.gate-progress-spinner {
    width: 28px;
    height: 28px;
    border: 2px solid rgba(43, 86, 136, .16);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: priority-spin .75s linear infinite;
}
[data-testid="stDialog"], [data-testid="stDialog"] * { font-family:'Montserrat', sans-serif !important; }
[data-testid="stDialog"] [data-testid="stFormSubmitButton"] { margin-top: 18px; }
.priority-processing {
    min-height: 260px; display: flex; flex-direction: column; align-items: center; justify-content: center;
    color: var(--ink-muted); font-size: 13px; gap: 14px;
}
.priority-processing-spinner {
    width: 28px; height: 28px; border: 2px solid rgba(43,86,136,.18); border-top-color: var(--accent);
    border-radius: 50%; animation: priority-spin .8s linear infinite;
}
@keyframes priority-spin { to { transform: rotate(360deg); } }
.methodology-copy { max-width: 1020px; font-size: 15px; line-height: 1.72; color: var(--ink); }
.methodology-copy p { margin: 0 0 16px 0; }
.methodology-copy ul { margin: 0 0 18px 22px; padding: 0; }
.methodology-copy li { margin: 0 0 6px 0; }
.methodology-subhead { font-size: 19px; font-weight: 500; line-height: 1.35; margin: 30px 0 12px 0; color: var(--ink); }

@media (max-width: 900px) {
    .main-title { font-size: 31px; }
    .main-subtitle { font-size: 16px; }
    .top-nav { gap: 24px; }
    .period-nav { margin-left: 0; }
    .sidebar-methodology { position: static; width: auto; margin-top: 46px; }
    .st-key-ranking_section, .st-key-table_section, .st-key-assessment_section, .st-key-priorities_section, .st-key-methodology_section { padding: 20px 18px; }
    .pillar-legend-box { grid-template-columns:1fr; }
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_index_data(file_path: str) -> pd.DataFrame:
    """Read cached Excel values without changing the workbook."""
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    if SHEET_NAME not in workbook.sheetnames:
        workbook.close()
        raise ValueError(f"The workbook does not contain a '{SHEET_NAME}' sheet.")

    worksheet = workbook[SHEET_NAME]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        workbook.close()
        raise ValueError("The Streamlit sheet is empty.")

    headers = list(rows[0])
    while headers and headers[-1] is None:
        headers.pop()

    data_rows: list[list[Any]] = []
    for row in rows[1:]:
        trimmed = list(row[: len(headers)])
        if any(value not in (None, "") for value in trimmed):
            data_rows.append(trimmed)

    frame = pd.DataFrame(data_rows, columns=headers)
    frame = frame[frame["Country"].notna() & frame["ISO3"].notna()].copy()

    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        workbook.close()
        raise ValueError("Missing required columns: " + ", ".join(missing))

    numeric_columns = ["Base index rank", "Base index score", *PILLAR_SOURCES]
    for column in numeric_columns:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    if frame["Base index score"].isna().any():
        workbook.close()
        raise ValueError(
            "Some formula results are unavailable. Open the workbook in Excel, allow recalculation, save it, and replace data/Index_Model.xlsx."
        )

    frame = frame[frame["Eligible"].astype(str).str.strip().str.lower().isin({"yes", "true", "1"})]

    if "Screening" in workbook.sheetnames:
        screening = workbook["Screening"]
        screening_rows = list(screening.iter_rows(values_only=True))
        if screening_rows:
            screening_headers = list(screening_rows[0])
            screening_data = pd.DataFrame(screening_rows[1:], columns=screening_headers)
            if "ISO3" in screening_data.columns and "Operating coal capacity (MW)" in screening_data.columns:
                screening_data = screening_data[["ISO3", "Operating coal capacity (MW)"]].copy()
                screening_data["Operating coal capacity (MW)"] = pd.to_numeric(
                    screening_data["Operating coal capacity (MW)"], errors="coerce"
                )
                frame = frame.merge(screening_data, on="ISO3", how="left")
    workbook.close()

    if "Operating coal capacity (MW)" not in frame.columns:
        frame["Operating coal capacity (MW)"] = pd.NA

    return frame.sort_values("Base index rank").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def load_market_overlay(file_path: str) -> pd.DataFrame:
    """Read the final Q1 and Q2 country overlay results from the OVERLAY sheet."""
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    if OVERLAY_SHEET_NAME not in workbook.sheetnames:
        workbook.close()
        raise ValueError(f"The workbook does not contain an '{OVERLAY_SHEET_NAME}' sheet.")

    worksheet = workbook[OVERLAY_SHEET_NAME]
    rows = list(worksheet.iter_rows(values_only=True))
    header_index = next(
        (
            index
            for index, row in enumerate(rows)
            if len(row) >= 2 and str(row[0]).strip() == "Country" and str(row[1]).strip() == "ISO3"
        ),
        None,
    )
    if header_index is None:
        workbook.close()
        raise ValueError("The OVERLAY sheet does not contain the final country-results table.")

    headers = list(rows[header_index])
    while headers and headers[-1] is None:
        headers.pop()

    data_rows: list[list[Any]] = []
    for row in rows[header_index + 1 :]:
        trimmed = list(row[: len(headers)])
        if not trimmed or trimmed[0] in (None, ""):
            break
        data_rows.append(trimmed)
    workbook.close()

    frame = pd.DataFrame(data_rows, columns=headers)
    required = [
        "Country",
        "ISO3",
        "Q1 overlay score",
        "Q1 rank",
        "Q1 rank change",
        "Q2 overlay score",
        "Q2 rank",
        "Q2 rank change",
    ]
    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError("Missing OVERLAY columns: " + ", ".join(missing))

    numeric_columns = [
        "Q1 overlay score",
        "Q1 rank",
        "Q1 score change",
        "Q1 rank change",
        "Q2 overlay score",
        "Q2 rank",
        "Q2 rank change",
    ]
    for column in numeric_columns:
        if column in frame.columns:
            frame[column] = pd.to_numeric(frame[column], errors="coerce")

    if frame[["Q1 overlay score", "Q1 rank", "Q2 overlay score", "Q2 rank"]].isna().any().any():
        raise ValueError("Some final Q1 or Q2 overlay results are unavailable in the OVERLAY sheet.")
    if frame["ISO3"].duplicated().any():
        raise ValueError("The OVERLAY country-results table contains duplicate ISO3 codes.")

    return frame.reset_index(drop=True)


def market_results(index_data: pd.DataFrame, overlay_data: pd.DataFrame, period: str) -> pd.DataFrame:
    score_column = "Q1 overlay score" if period == "q1" else "Q2 overlay score"
    rank_column = "Q1 rank" if period == "q1" else "Q2 rank"
    rank_change_column = "Q1 rank change" if period == "q1" else "Q2 rank change"
    score_change_column = "Q1 score change" if period == "q1" else None

    result = base_results(index_data)
    overlay_columns = ["ISO3", score_column, rank_column, rank_change_column]
    if score_change_column and score_change_column in overlay_data.columns:
        overlay_columns.append(score_change_column)
    result = result.merge(overlay_data[overlay_columns], on="ISO3", how="left", validate="one_to_one")

    if result[[score_column, rank_column]].isna().any().any():
        missing = result.loc[result[score_column].isna() | result[rank_column].isna(), "ISO3"].tolist()
        raise ValueError("Missing market-overlay results for: " + ", ".join(missing))

    result["Overall score"] = result[score_column].astype(float)
    result["Rank"] = result[rank_column].astype(int)
    result["Rank change"] = result[rank_change_column].astype(int)
    if score_change_column and score_change_column in result.columns:
        result["Score change"] = result[score_change_column].astype(float)
    else:
        result["Score change"] = result["Overall score"] - result["Base score"]

    return result.sort_values(["Rank", "Country"]).reset_index(drop=True)


SINGAPORE_TZ = ZoneInfo("Asia/Singapore")

def singapore_now() -> str:
    return datetime.now(SINGAPORE_TZ).replace(microsecond=0).isoformat()


def initialise_state() -> None:
    defaults = {
        "session_id": str(uuid.uuid4()),
        "session_started_at": singapore_now(),
        "previews": [],
        "saved_preview_ids": [],
        "current_preview_id": None,
        "selected_country": "Chile",
        "contact": {"name": "", "organisation": "", "email": ""},
        "profile_complete": False,
        "gate_action": None,
        "gate_version_id": None,
        "session_logged": False,
        "scroll_to_top": False,
        "scroll_target": None,
        "notice": None,
        "action_notice": None,
        "action_notice_type": "warning",
        "logging_notice": None,
        "profile_industry": "Select your industry",
        "profile_role": "Select your role",
        "pending_priority_submission": None,
        "action_in_progress": None,
        "action_version_id": None,
        "cached_pdf": None,
        "cached_pdf_version_id": None,
        "profile_capture_pending": False,
        "profile_capture_reason": None,
        "gate_processing": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def query_value(key: str, default: str) -> str:
    value = st.query_params.get(key, default)
    if isinstance(value, list):
        value = value[0] if value else default
    return str(value)


def set_view(view: str) -> None:
    st.query_params["view"] = view
    if view != "market" and "period" in st.query_params:
        del st.query_params["period"]


def reset_priority_survey() -> None:
    """Clear the current customised result while retaining the user's profile for this session."""
    st.session_state.previews = []
    st.session_state.saved_preview_ids = []
    st.session_state.current_preview_id = None
    st.session_state.gate_action = None
    st.session_state.gate_version_id = None
    st.session_state.action_notice = None
    st.session_state.action_notice_type = "warning"
    st.session_state.logging_notice = None
    st.session_state.action_in_progress = None
    st.session_state.action_version_id = None
    st.session_state.cached_pdf = None
    st.session_state.cached_pdf_version_id = None
    st.session_state.pending_priority_submission = None
    st.session_state.profile_capture_pending = False
    st.session_state.profile_capture_reason = None
    st.session_state.gate_processing = False

    for item in SURVEY_ITEMS:
        st.session_state[f"priority_slider_{item['key']}"] = 3

    # Preserve the session's selected industry and role.
    st.session_state.industry_widget = st.session_state.profile_industry
    st.session_state.role_widget = st.session_state.profile_role

    if "result" in st.query_params:
        del st.query_params["result"]


def navigate_to(view: str, period: str | None = None, scroll_target: str = "top") -> None:
    st.query_params["view"] = view
    if period:
        st.query_params["period"] = period
    elif "period" in st.query_params:
        del st.query_params["period"]
    if "result" in st.query_params:
        del st.query_params["result"]
    st.session_state.scroll_target = scroll_target
    st.session_state.scroll_to_top = scroll_target == "top"
    st.rerun()


def render_sidebar(view: str) -> None:
    with st.sidebar:
        if st.button("Readiness Index", key="sidebar_readiness", type="tertiary"):
            navigate_to("overall", scroll_target="top")
        if st.button("Pillar Scores", key="sidebar_pillars", type="tertiary"):
            navigate_to("overall", scroll_target="pillar-scores")
        if st.button("Country Assessment", key="sidebar_assessment", type="tertiary"):
            navigate_to("overall", scroll_target="country-assessment")
        with st.container(key="sidebar_priorities_gap"):
            st.markdown("", unsafe_allow_html=True)
        if st.button("Add Your Priorities", key="sidebar_priorities", type="tertiary"):
            if view != "priorities":
                reset_priority_survey()
            navigate_to("priorities", scroll_target="top")
        with st.container(key="sidebar_methodology_button"):
            if st.button("Methodology", key="sidebar_methodology", type="tertiary"):
                navigate_to("methodology", scroll_target="top")


def render_header() -> None:
    st.markdown('<span id="top" class="anchor"></span>', unsafe_allow_html=True)
    header_html = """
    <div class="main-title-row">
      <div class="main-title">Coal-to-Clean Jurisdictional Readiness Index 2026</div>
      <details class="info-details" id="index-info-details">
        <summary aria-label="About the index" title="About the index">ⓘ</summary>
        <div class="info-panel">
          <p><strong>About the index</strong></p>
          <p>The Coal-to-Clean Jurisdictional Readiness Index compares coal-dependent markets on their readiness to support credible coal-to-clean energy transition opportunities.</p>
          <p>The index combines six pillars covering energy system conditions, policy commitment, governance capacity, carbon market maturity, macro-financial conditions, and just transition credibility. Scores are calculated from normalised indicators and weighted to produce an overall jurisdictional readiness score.</p>
          <p>The base index primarily reflects 2024 data, using the latest available year only where 2024 values were unavailable.</p>
          <p>The index is a jurisdiction-level screening tool. It does not assess the viability of individual power plants, projects or investments.</p>
          <p style="margin-top:16px;"><strong>About the market confidence overlay</strong></p>
          <p>The market-confidence overlay adds more recent evidence to the underlying index. It adjusts selected indicator weights where market developments suggest that certain readiness factors have become more or less important.</p>
          <p>The overlay does not replace the base index or alter its underlying data. It provides a separate, time-specific view for Q1 2026 and Q2 2026.</p>
          <p>Survey-based user weightings are calculated separately and are not combined with the main market-confidence overlay.</p>
          <p style="margin-top:18px;">For more information, please see <a href="?view=methodology#top" class="methodology-link">Methodology</a>.</p>
        </div>
      </details>
    </div>
    <script>
      (() => {
        const details = document.getElementById('index-info-details');
        if (!details) return;
        document.addEventListener('click', (event) => {
          if (details.open && !details.contains(event.target)) details.removeAttribute('open');
        });
        document.addEventListener('keydown', (event) => {
          if (event.key === 'Escape') details.removeAttribute('open');
        });
        const methodologyLink = details.querySelector('.methodology-link');
        if (methodologyLink) {
          methodologyLink.addEventListener('click', (event) => {
            event.preventDefault();
            details.removeAttribute('open');
            const methodologyButton = Array.from(document.querySelectorAll('button')).find(
              (button) => button.textContent.trim().toLowerCase() === 'methodology'
            );
            if (methodologyButton) methodologyButton.click();
            else window.location.href = '?view=methodology#top';
          });
        }
      })();
    </script>
    """
    st.html(header_html, unsafe_allow_javascript=True)
    st.markdown(
        '<div class="main-subtitle">Assessing energy transition opportunities across coal-dependent markets</div>',
        unsafe_allow_html=True,
    )

def render_top_navigation(view: str, market_period: str | None = None) -> None:
    with st.container(key="top_nav"):
        nav_columns = st.columns([1.2, 2.0, 2.1, 3.7], gap="medium")
        nav_items = [
            ("overall", "Main Index"),
            ("market", "Market Confidence"),
            ("priorities", "Add Your Priorities"),
        ]
        for column, (target, label) in zip(nav_columns[:3], nav_items):
            state = "active" if view == target else "inactive"
            with column:
                with st.container(key=f"nav_{target}_{state}"):
                    if st.button(label, key=f"top_nav_{target}", type="tertiary"):
                        if target == "priorities" and view != "priorities":
                            reset_priority_survey()
                        navigate_to(target, scroll_target="top")

    if view == "market":
        period = market_period or "q1"
        with st.container(key="period_slot"):
            # The empty first column matches the Main Index navigation column,
            # so Q1 begins directly beneath Market Confidence.
            period_columns = st.columns(
                [1.2, 0.9, 0.9, 1.5, 4.5],
                gap="medium",
            )
            with period_columns[1]:
                with st.container(key=f"period_q1_{'active' if period == 'q1' else 'inactive'}"):
                    if st.button("Q1 2026", key="period_q1", type="tertiary"):
                        navigate_to("market", period="q1", scroll_target="top")
            with period_columns[2]:
                with st.container(key=f"period_q2_{'active' if period == 'q2' else 'inactive'}"):
                    if st.button("Q2 2026", key="period_q2", type="tertiary"):
                        navigate_to("market", period="q2", scroll_target="top")
            with period_columns[3]:
                with st.container(key="period_clear"):
                    if st.button("Clear adjustment", key="period_clear_button", type="tertiary"):
                        navigate_to("overall", scroll_target="top")
    else:
        st.markdown('<div class="period-placeholder">Placeholder</div>', unsafe_allow_html=True)


def base_results(index_data: pd.DataFrame) -> pd.DataFrame:
    result = index_data[
        ["Country", "ISO3", "Base index rank", "Base index score", "Operating coal capacity (MW)", *PILLAR_SOURCES]
    ].copy()
    result = result.rename(columns={"Base index rank": "Base rank", "Base index score": "Base score"})
    result["Rank"] = result["Base rank"].astype(int)
    result["Overall score"] = result["Base score"].astype(float)
    result["Rank change"] = 0
    result["Score change"] = 0.0
    return result.sort_values("Rank").reset_index(drop=True)


def adjusted_weights(industry: str, responses: dict[str, int]) -> list[float]:
    slider_multipliers = {1: 0.60, 2: 0.80, 3: 1.00, 4: 1.20, 5: 1.40}
    pillar_responses: list[list[int]] = [[] for _ in range(6)]
    for item in SURVEY_ITEMS:
        for pillar_index in item["pillars"]:
            pillar_responses[pillar_index].append(responses[item["key"]])
    response_values = [sum(values) / len(values) if values else 3 for values in pillar_responses]
    industry_values = INDUSTRY_MULTIPLIERS[industry]
    pre_normalised = [
        base * industry_multiplier * slider_multipliers[int(round(response_value))]
        for base, industry_multiplier, response_value in zip(BASE_WEIGHTS, industry_values, response_values)
    ]
    total = sum(pre_normalised)
    return [value / total for value in pre_normalised]


def calculate_custom_results(index_data: pd.DataFrame, weights: list[float]) -> pd.DataFrame:
    result = index_data[
        ["Country", "ISO3", "Base index rank", "Base index score", "Operating coal capacity (MW)", *PILLAR_SOURCES]
    ].copy()
    result = result.rename(columns={"Base index rank": "Base rank", "Base index score": "Base score"})
    result["Overall score"] = sum(result[column] * weight for column, weight in zip(PILLAR_SOURCES, weights))
    result["Rank"] = result["Overall score"].rank(method="min", ascending=False).astype(int)
    result["Rank change"] = result["Base rank"].astype(int) - result["Rank"]
    result["Score change"] = result["Overall score"] - result["Base score"]
    return result.sort_values(["Rank", "Country"]).reset_index(drop=True)


def ranking_figure(results: pd.DataFrame, show_all: bool, mode: str) -> go.Figure:
    displayed = results.sort_values("Rank").copy()
    if not show_all:
        displayed = displayed.head(10)
    displayed = displayed.sort_values("Rank", ascending=False)

    colour = "#2b5688"
    if mode == "market":
        colour = "#79a6d2"
    elif mode == "custom":
        colour = "#8bb0da"

    overlay_mode = mode in {"market", "custom"}
    customdata = []
    for _, row in displayed.iterrows():
        capacity = row.get("Operating coal capacity (MW)")
        capacity_text = "Not available" if pd.isna(capacity) else f"{float(capacity):,.0f} MW"
        rank_change = int(row.get("Rank change", 0))
        change_text = "No change" if rank_change == 0 else f"{rank_change:+d} position{'s' if abs(rank_change) != 1 else ''}"
        customdata.append([int(row["Rank"]), change_text, capacity_text, int(row.get("Base rank", row["Rank"]))])

    score_text = displayed["Overall score"].map(lambda value: f"{float(value):.2f}")
    hover_template = (
        "<b>%{y}</b><br>Adjusted rank: %{customdata[0]}<br>"
        "Position change: %{customdata[1]}<br>Operating coal capacity: %{customdata[2]}<extra></extra>"
        if overlay_mode
        else "<b>%{y}</b><br>Rank: %{customdata[0]}<br>Operating coal capacity: %{customdata[2]}<extra></extra>"
    )

    figure = go.Figure()
    figure.add_trace(
        go.Bar(
            x=displayed["Overall score"],
            y=displayed["Country"],
            orientation="h",
            marker_color=colour,
            text=score_text,
            textposition="inside" if overlay_mode else "outside",
            insidetextanchor="start" if overlay_mode else "end",
            textfont=dict(
                family="Montserrat",
                size=12,
                color="#fff5cc" if overlay_mode else "rgba(0,0,0,.86)",
            ),
            cliponaxis=False,
            customdata=customdata,
            hovertemplate=hover_template,
            name="Adjusted score" if overlay_mode else "Overall score",
        )
    )

    if overlay_mode:
        connector_x: list[float | None] = []
        connector_y: list[str | None] = []
        for _, row in displayed.iterrows():
            connector_x.extend([float(row["Base score"]), float(row["Overall score"]), None])
            connector_y.extend([str(row["Country"]), str(row["Country"]), None])
        figure.add_trace(
            go.Scatter(
                x=connector_x,
                y=connector_y,
                mode="lines",
                line=dict(color="rgba(43,86,136,.48)", width=2),
                hoverinfo="skip",
                showlegend=False,
            )
        )
        figure.add_trace(
            go.Scatter(
                x=displayed["Base score"],
                y=displayed["Country"],
                mode="markers",
                customdata=customdata,
                marker=dict(symbol="diamond", size=13, color="#2b5688", line=dict(width=1, color="#ffffff")),
                hovertemplate="<b>%{y}</b><br>Base rank: %{customdata[3]}<br>Base index score: %{x:.2f}<extra></extra>",
                name="Base index",
            )
        )

    maximum = max(100.0, float(displayed["Overall score"].max()) + 8)
    row_height = 42
    chart_height = (len(displayed) * row_height) + 66
    figure.update_layout(
        autosize=False,
        height=chart_height,
        margin=dict(l=280, r=78, t=0, b=66, autoexpand=False),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Montserrat", color="rgba(0,0,0,.90)"),
        bargap=0.22,
        showlegend=overlay_mode,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1.0,
            xanchor="right",
            x=0.99,
            font=dict(size=11),
            traceorder="normal",
            bgcolor="rgba(0,0,0,0)",
        ),
        hoverlabel=dict(
            bgcolor="#76a7d5",
            bordercolor="#ffffff",
            font=dict(family="Montserrat", size=12, color="#ffffff"),
        ),
        dragmode=False,
        uniformtext=dict(minsize=10, mode="show"),
    )
    figure.update_xaxes(
        title=dict(text="OVERALL SCORE", font=dict(size=14, family="Montserrat"), standoff=18),
        range=[0, maximum],
        fixedrange=True,
        gridcolor="rgba(43,86,136,.11)",
        zeroline=False,
        tickfont=dict(size=11),
    )
    figure.update_yaxes(
        fixedrange=True,
        domain=[0.0, 1.0],
        range=[-0.5, len(displayed) - 0.5],
        tickfont=dict(size=13, family="Montserrat", color="rgba(0,0,0,.88)"),
        automargin=False,
        ticklabelposition="outside",
        ticklabelstandoff=28,
        categoryorder="array",
        categoryarray=displayed["Country"].tolist(),
    )
    return figure

def pillar_figure(data_row: pd.Series, weights: list[float] | None = None) -> go.Figure:
    values = [float(data_row[column]) for column in PILLAR_SOURCES]
    labels = [f"Pillar {index}" for index in range(1, 7)]
    text = [f"{value:.1f}" for value in values]
    figure = go.Figure(
        go.Bar(
            x=labels,
            y=values,
            marker_color="#2b5688" if weights is None else "#8bb0da",
            width=0.38,
            text=text,
            textposition="outside",
            textfont=dict(family="Montserrat", size=12),
            hoverinfo="skip",
        )
    )
    figure.update_layout(
        height=430,
        margin=dict(l=72, r=20, t=20, b=78),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Montserrat", color="rgba(0,0,0,.90)"),
        dragmode=False,
        showlegend=False,
    )
    figure.update_yaxes(
        title=dict(text="PILLAR SCORE", font=dict(size=14, family="Montserrat"), standoff=16),
        range=[0, 105],
        fixedrange=True,
        gridcolor="rgba(43,86,136,.11)",
        zeroline=False,
        tickfont=dict(size=11),
    )
    figure.update_xaxes(
        fixedrange=True,
        tickfont=dict(size=12, family="Montserrat"),
        ticklabelstandoff=16,
    )
    return figure

def descriptor_for_country(
    country: str,
    current_row: pd.Series,
    results: pd.DataFrame,
    index_data: pd.DataFrame,
    custom: bool,
) -> str:
    ordered = results.sort_values("Rank").reset_index(drop=True)
    rank = int(current_row["Rank"])
    score = float(current_row["Overall score"])
    total = len(ordered)
    median_score = float(ordered["Overall score"].median())
    median_gap = score - median_score

    sentences: list[str] = []
    if rank == 1 and total > 1:
        second = ordered.iloc[1]
        lead = score - float(second["Overall score"])
        sentences.append(
            f"{country} leads the index by {lead:.1f} points over {second['Country']} and scores {abs(median_gap):.1f} points above the jurisdiction median."
        )
    elif rank == total and total > 1:
        above = ordered.iloc[-2]
        gap = float(above["Overall score"]) - score
        sentences.append(
            f"{country} trails the next-ranked jurisdiction by {gap:.1f} points and sits {abs(median_gap):.1f} points below the median."
        )
    else:
        above = ordered.iloc[rank - 2] if rank > 1 else None
        below = ordered.iloc[rank] if rank < total else None
        neighbours = []
        if above is not None:
            neighbours.append((float(above["Overall score"]) - score, str(above["Country"])))
        if below is not None:
            neighbours.append((score - float(below["Overall score"]), str(below["Country"])))
        nearest_gap, nearest_country = min(neighbours, key=lambda item: item[0])
        direction = "above" if median_gap >= 0 else "below"
        sentences.append(
            f"{country} sits {abs(median_gap):.1f} points {direction} the median and is only {nearest_gap:.1f} points from {nearest_country}."
        )

    data_row = index_data.loc[index_data["Country"] == country].iloc[0]
    pillar_scores = pd.Series({column: float(data_row[column]) for column in PILLAR_SOURCES})
    pillar_medians = index_data[PILLAR_SOURCES].median()
    above_count = int((pillar_scores > pillar_medians).sum())
    score_range = float(pillar_scores.max() - pillar_scores.min())

    if score_range <= 20:
        sentences.append(
            f"Its readiness profile is comparatively balanced, with scores above the peer median in {above_count} of six pillars."
        )
    else:
        sentences.append(
            f"Its readiness profile is uneven, with scores above the peer median in {above_count} of six pillars and a {score_range:.1f}-point spread across pillars."
        )

    sorted_scores = pillar_scores.sort_values(ascending=False)
    strongest_gap = float(sorted_scores.iloc[0] - sorted_scores.iloc[1])
    weakest_gap = float(sorted_scores.iloc[-2] - sorted_scores.iloc[-1])
    if strongest_gap >= 10:
        source = sorted_scores.index[0]
        sentences.append(
            f"Its clearest differentiator is **{PILLAR_FULL[source]}**, which is {strongest_gap:.1f} points ahead of its next-best pillar."
        )
    elif weakest_gap >= 10:
        source = sorted_scores.index[-1]
        sentences.append(
            f"Its clearest constraint is **{PILLAR_FULL[source]}**, which is {weakest_gap:.1f} points behind its next-weakest pillar."
        )

    if custom:
        rank_change = int(current_row["Rank change"])
        score_change = float(current_row["Score change"])
        if rank_change > 0:
            sentences.append(
                f"Under the selected priorities, it rises {rank_change} position{'s' if rank_change != 1 else ''} and its score changes by {score_change:+.1f} points."
            )
        elif rank_change < 0:
            sentences.append(
                f"Under the selected priorities, it falls {abs(rank_change)} position{'s' if rank_change != -1 else ''} and its score changes by {score_change:+.1f} points."
            )
        else:
            sentences.append(
                f"Under the selected priorities, its rank is unchanged and its score changes by {score_change:+.1f} points."
            )

    return " ".join(sentences)


def bold_markdown_to_html(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)


def priority_considerations(country: str, index_data: pd.DataFrame) -> list[tuple[str, str]]:
    row = index_data.loc[index_data["Country"] == country].iloc[0]
    scores = pd.Series({column: float(row[column]) for column in PILLAR_SOURCES}).sort_values()
    selected = list(scores.index[:2])
    return [(PILLAR_FULL[column], RECOMMENDATIONS[column]) for column in selected]


def current_result_query(view: str, market_period: str | None, custom_version: dict[str, Any] | None) -> str:
    parts = [f"view={view}"]
    if market_period:
        parts.append(f"period={market_period}")
    if custom_version:
        parts.append(f"result=preview_{custom_version['id']}")
    return "?" + "&".join(parts)


def render_html_table(
    table: pd.DataFrame,
    current_query: str,
    height: int = 560,
) -> None:
    rows = []
    for _, row in table.iterrows():
        cells = []
        for column in table.columns:
            value = row[column]
            if pd.isna(value):
                rendered = ""
            elif isinstance(value, (float, int)) and column not in {"Rank", "Change"}:
                rendered = f"{float(value):.2f}"
            elif column in {"Rank", "Change"}:
                rendered = str(int(value)) if float(value).is_integer() else str(value)
            else:
                rendered = html.escape(str(value))
            class_name = "rank-cell" if column == "Rank" else ""
            cells.append(f'<td class="{class_name}" data-col="{html.escape(column)}">{rendered}</td>')
        country = quote(str(row["Country"]))
        target = f"{current_query}&country={country}#country-assessment"
        rows.append(f'<tr onclick="window.location.href=\'{target}\'">{"".join(cells)}</tr>')

    headers = "".join(
        f'<th data-col="{html.escape(column)}">{html.escape(column)}</th>' for column in table.columns
    )
    column_options = "".join(
        f'<label><input type="checkbox" checked data-toggle="{html.escape(column)}"><span>{html.escape(column)}</span></label>'
        for column in table.columns
    )

    col_widths = []
    for column in table.columns:
        if column == "Rank":
            width = 62
        elif column == "Country":
            width = 200
        elif column == "Overall Score":
            width = 100
        elif column == "Change":
            width = 105
        else:
            width = 145
        col_widths.append(f'<col data-col="{html.escape(column)}" style="width:{width}px">')

    component_html = f"""
    <div class="readiness-table-component" id="table-component">
      <style>
        .readiness-table-component {{ height:{height}px; display:flex; flex-direction:column; font-family:Montserrat,sans-serif; color:rgba(0,0,0,.90); }}
        .readiness-table-component * {{ box-sizing:border-box; }}
        .readiness-table-toolbar {{ height:44px; flex:0 0 44px; display:flex; align-items:center; justify-content:flex-end; gap:8px; background:transparent; position:relative; overflow:visible; z-index:100; }}
        .readiness-tool {{ position:relative; }}
        .readiness-tool summary {{ width:38px; height:34px; border:1px solid rgba(43,86,136,.42); border-radius:7px; background:#ffffff; display:flex; align-items:center; justify-content:center; cursor:pointer; list-style:none; color:#2b5688 !important; opacity:1; }}
        .readiness-tool summary::-webkit-details-marker {{ display:none; }}
        .readiness-tool summary:hover {{ background:#fff5cc; }}
        .columns-glyph {{ display:block; color:#2b5688; font-size:22px; line-height:1; font-family:Arial,sans-serif; font-weight:700; transform:translateY(-1px); }}
        .readiness-column-menu {{ position:absolute; right:0; top:39px; z-index:2000; width:min(760px, 92vw); background:white; border:1px solid rgba(43,86,136,.16); border-radius:10px; padding:14px 16px; box-shadow:0 12px 30px rgba(43,86,136,.18); display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); column-gap:24px; }}
        .readiness-column-menu label {{ display:flex; align-items:center; gap:10px; font-size:12px; line-height:1.25; margin:8px 0; white-space:nowrap; }}
        .readiness-column-menu input[type="checkbox"] {{ width:16px; height:16px; accent-color:#ffe680; flex:0 0 auto; }}
        .readiness-table-shell {{ flex:1; min-height:0; border:1px solid rgba(43,86,136,.14); border-radius:10px; overflow:hidden; background:white; }}
        .readiness-table-scroll {{ height:100%; overflow:auto; }}
        .readiness-table-component table {{ border-collapse:collapse; width:100%; min-width:1232px; table-layout:fixed; }}
        .readiness-table-component thead th {{ position:sticky; top:0; z-index:3; height:76px; padding:10px 11px; text-align:left; vertical-align:middle; white-space:normal; line-height:1.25; font-size:13px; font-weight:600; background:#fffae6; border-right:1px solid rgba(43,86,136,.10); border-bottom:1px solid rgba(43,86,136,.14); }}
        .readiness-table-component tbody td {{ height:39px; padding:8px 11px; font-size:13px; font-weight:500; border-right:1px solid rgba(43,86,136,.08); border-bottom:1px solid rgba(43,86,136,.08); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
        .readiness-table-component tbody tr {{ cursor:pointer; }}
        .readiness-table-component tbody tr:nth-child(even) {{ background:#fffdf3; }}
        .readiness-table-component tbody tr:hover {{ background:#fff5cc; }}
        .readiness-table-component .rank-cell {{ text-align:center; }}
        .readiness-tooltip {{ position:absolute; bottom:42px; right:50%; transform:translateX(50%); padding:5px 8px; background:rgba(0,0,0,.78); color:white; border-radius:5px; font-size:10px; white-space:nowrap; opacity:0; pointer-events:none; transition:opacity .12s ease; }}
        .readiness-tool:hover .readiness-tooltip {{ opacity:1; }}
      </style>
      <div class="readiness-table-toolbar">
        <details class="readiness-tool" id="column-picker">
          <summary aria-label="Choose columns"><span class="columns-glyph" aria-hidden="true">▥</span></summary>
          <span class="readiness-tooltip">Choose columns</span>
          <div class="readiness-column-menu">{column_options}</div>
        </details>
      </div>
      <div class="readiness-table-shell">
        <div class="readiness-table-scroll">
          <table>
            <colgroup>{''.join(col_widths)}</colgroup>
            <thead><tr>{headers}</tr></thead>
            <tbody>{''.join(rows)}</tbody>
          </table>
        </div>
      </div>
    </div>
    <script>
      document.querySelectorAll('#table-component [data-toggle]').forEach(function(box) {{
        box.addEventListener('change', function() {{
          const col = this.getAttribute('data-toggle');
          document.querySelectorAll('#table-component [data-col="' + CSS.escape(col) + '"]').forEach(function(cell) {{
            cell.style.display = box.checked ? '' : 'none';
          }});
        }});
      }});
      document.addEventListener('click', function(event) {{
        const picker = document.getElementById('column-picker');
        if (picker && picker.open && !picker.contains(event.target)) picker.removeAttribute('open');
      }});
    </script>
    """
    st.html(component_html, unsafe_allow_javascript=True)

def render_chart(
    figure: go.Figure,
    filename: str,
    key: str,
) -> None:
    st.plotly_chart(
        figure,
        width="stretch",
        theme=None,
        key=key,
        config={
            "displayModeBar": True,
            "displaylogo": False,
            "scrollZoom": False,
            "responsive": True,
            "showTips": True,
            "modeBarButtonsToRemove": [
                "zoom2d",
                "pan2d",
                "select2d",
                "lasso2d",
                "zoomIn2d",
                "zoomOut2d",
                "autoScale2d",
                "resetScale2d",
            ],
            "toImageButtonOptions": {"format": "png", "filename": filename, "scale": 2},
        },
    )

@st.fragment
def render_ranking_chart_fragment(
    results: pd.DataFrame,
    view: str,
    mode: str,
    custom_id: str | None,
    market_period: str | None,
) -> None:
    show_all_key = f"show_all_{view}_{custom_id or market_period or 'base'}"
    if show_all_key not in st.session_state:
        st.session_state[show_all_key] = False
    show_all = bool(st.session_state[show_all_key])

    if view == "market":
        label = "Q1 2026" if market_period == "q1" else "Q2 2026"
        chart_title = f"{label}: Market Confidence-Adjusted Country Ranking"
        title_class = "section-title"
    elif mode == "custom":
        chart_title = "Your Priority-Adjusted Country Ranking"
        title_class = "section-title"
    else:
        chart_title = "2026 Country Ranking" if show_all else "Top 10 Country Ranking"
        title_class = "section-title base-chart-title"

    with st.container(key="chart_header_slot"):
        st.markdown(f'<div class="{title_class}">{chart_title}</div>', unsafe_allow_html=True)
        show_all = st.toggle("Show all countries", key=show_all_key)

    with st.container(key="ranking_chart_slot"):
        render_chart(
            ranking_figure(results, show_all=show_all, mode=mode),
            "readiness_index_country_ranking",
            f"ranking_chart_{view}_{custom_id or market_period or 'base'}",
        )


def get_visible_previews() -> list[dict[str, Any]]:
    saved_ids = set(st.session_state.saved_preview_ids)
    visible = [preview for preview in st.session_state.previews if preview["id"] in saved_ids]
    current_id = st.session_state.current_preview_id
    if current_id and current_id not in saved_ids:
        current = next((preview for preview in st.session_state.previews if preview["id"] == current_id), None)
        if current:
            visible.append(current)
    return sorted(visible, key=lambda item: item["number"])


def resolve_custom_version() -> dict[str, Any] | None:
    result_param = query_value("result", "")
    if not result_param.startswith("preview_"):
        return None
    preview_id = result_param.replace("preview_", "", 1)
    match = next((preview for preview in st.session_state.previews if preview["id"] == preview_id), None)
    if match:
        st.session_state.current_preview_id = match["id"]
    return match

def render_result_selector(custom_version: dict[str, Any] | None) -> dict[str, Any] | None:
    visible = get_visible_previews()
    if not visible:
        return custom_version
    labels = ["Main index"] + [f"Customised view {preview['number']}" for preview in visible]
    current_label = "Main index"
    if custom_version:
        current_label = f"Customised view {custom_version['number']}"
    if "result_view_selector" not in st.session_state or st.session_state.result_view_selector not in labels:
        st.session_state.result_view_selector = current_label
    selected = st.segmented_control(
        "Result view",
        labels,
        key="result_view_selector",
        label_visibility="collapsed",
    )
    if not selected or selected == current_label:
        return custom_version
    if selected == "Main index":
        st.session_state.current_preview_id = None
        if "result" in st.query_params:
            del st.query_params["result"]
        st.rerun()
    number = int(selected.split()[-1])
    chosen = next(preview for preview in visible if preview["number"] == number)
    st.session_state.current_preview_id = chosen["id"]
    st.query_params["result"] = f"preview_{chosen['id']}"
    st.rerun()
    return custom_version


def display_results(
    index_data: pd.DataFrame,
    results: pd.DataFrame,
    view: str,
    custom_version: dict[str, Any] | None = None,
    market_period: str | None = None,
) -> None:
    custom = custom_version is not None
    mode = "custom" if custom else ("market" if view == "market" else "base")

    with st.container(key="ranking_section"):
        render_top_navigation(view, market_period)
        if custom_version:
            results = custom_version["results"]

        render_ranking_chart_fragment(
            results=results,
            view=view,
            mode=mode,
            custom_id=custom_version["id"] if custom_version else None,
            market_period=market_period,
        )

        if custom:
            render_custom_follow_up(index_data, custom_version)

    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.container(key="table_section"):
        st.markdown('<span id="pillar-scores" class="anchor"></span>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Jurisdiction Rankings and Pillar Scores</div>', unsafe_allow_html=True)

        table = results.sort_values("Rank").copy()
        table_columns = ["Rank", "Country", "Overall score"]
        if custom:
            table_columns.append("Rank change")
        table_columns.extend(PILLAR_SOURCES)
        table = table[table_columns].rename(columns=PILLAR_TABLE)
        table = table.rename(columns={"Overall score": "Overall Score", "Rank change": "Change"})
        query = current_result_query(view, market_period, custom_version)
        render_html_table(table, query)

    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.container(key="assessment_section"):
        st.markdown('<span id="country-assessment" class="anchor"></span>', unsafe_allow_html=True)
        st.markdown('<div class="section-title" style="margin-bottom:20px;">Country Assessment</div>', unsafe_allow_html=True)

        countries = sorted(results["Country"].tolist())
        query_country = query_value("country", "Chile")
        preferred = query_country if query_country in countries else st.session_state.selected_country
        if preferred not in countries:
            preferred = "Chile" if "Chile" in countries else countries[0]
        st.session_state.selected_country = preferred
        if "country_selector" not in st.session_state or st.session_state.country_selector not in countries:
            st.session_state.country_selector = preferred

        left, right = st.columns([1.05, 1.85], gap="large")
        with left:
            with st.container(key="country_picker"):
                selected_country = st.selectbox(
                    "Search jurisdiction",
                    countries,
                    key="country_selector",
                    width="stretch",
                )
            st.session_state.selected_country = selected_country
            current_row = results.loc[results["Country"] == selected_country].iloc[0]
            st.markdown(
                f"""
                <div class="metric-grid">
                  <div class="metric-card"><div class="metric-label">Overall rank</div><div class="metric-value">{int(current_row['Rank'])} of {len(results)}</div></div>
                  <div class="metric-card"><div class="metric-label">Overall score</div><div class="metric-value">{float(current_row['Overall score']):.2f}</div></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            assessment = descriptor_for_country(selected_country, current_row, results, index_data, custom)
            st.markdown('<div class="assessment-heading">Assessment</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="assessment-copy">{bold_markdown_to_html(assessment)}</div>', unsafe_allow_html=True)

        with right:
            data_row = index_data.loc[index_data["Country"] == selected_country].iloc[0]
            weights = custom_version["weights"] if custom else None
            render_chart(
                pillar_figure(data_row, weights),
                f"{selected_country}_pillar_scores",
                f"pillar_chart_{selected_country}_{custom_version['id'] if custom_version else 'base'}",
            )

        with st.container(key="assessment_bottom_row"):
            bottom_left, bottom_right = st.columns([1.05, 1.85], gap="large")
            with bottom_left:
                with st.container(key="priority_panel"):
                    st.markdown('<div class="priority-heading">Priority considerations</div>', unsafe_allow_html=True)
                    priority_paragraphs = "".join(
                        f'<p><strong>{html.escape(label)}</strong> – {html.escape(recommendation)}</p>'
                        for label, recommendation in priority_considerations(selected_country, index_data)
                    )
                    st.markdown(
                        f'<div class="priority-card">{priority_paragraphs}</div>',
                        unsafe_allow_html=True,
                    )

            with bottom_right:
                with st.container(key="pillar_explanation_panel"):
                    legend_items = "".join(
                        (
                            '<div class="pillar-definition">'
                            f'<strong>{html.escape(PILLAR_FULL[source])}</strong>'
                            f'<p>{html.escape(description)}</p>'
                            '</div>'
                        )
                        for source, description in zip(PILLAR_SOURCES, PILLAR_DESCRIPTIONS)
                    )
                    st.markdown(f'<div class="pillar-legend-box">{legend_items}</div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="assessment-footnote">This commentary is generated from relative index and pillar scores. It is not a substitute for transaction-specific due diligence.</div>',
            unsafe_allow_html=True,
        )

def report_versions_for_export(current_version: dict[str, Any]) -> list[dict[str, Any]]:
    if current_version["id"] in st.session_state.saved_preview_ids:
        return [current_version]
    return []


def save_preview(current_version: dict[str, Any]) -> None:
    st.session_state.saved_preview_ids = [current_version["id"]]


def begin_custom_action(action: str, current_version: dict[str, Any]) -> None:
    st.session_state.action_notice = None
    st.session_state.action_in_progress = action
    st.session_state.action_version_id = current_version["id"]


def render_action_progress(label: str) -> None:
    st.markdown(
        (
            '<div class="action-progress-wrap">'
            '<div class="action-progress-spinner" aria-hidden="true"></div>'
            f'<div>{html.escape(label)}</div>'
            '</div>'
        ),
        unsafe_allow_html=True,
    )


def perform_custom_action(index_data: pd.DataFrame, current_version: dict[str, Any]) -> None:
    """Complete a queued save or email action after the progress UI is visible."""
    action = st.session_state.action_in_progress
    version_id = st.session_state.action_version_id
    if not action or version_id != current_version["id"]:
        return

    # Give the browser a brief moment to paint the spinner and disabled buttons.
    time.sleep(0.18)

    if st.session_state.profile_capture_pending:
        contact = st.session_state.contact
        profile_ok, profile_message = post_backend(
            "log_profile",
            {
                "session_id": st.session_state.session_id,
                "captured_at": singapore_now(),
                "name": contact["name"],
                "organisation": contact["organisation"],
                "email": contact["email"],
                "reason": st.session_state.profile_capture_reason or action,
                "industry": st.session_state.profile_industry,
                "role": st.session_state.profile_role,
                "started_at": st.session_state.session_started_at,
            },
            timeout=12,
        )
        st.session_state.session_logged = st.session_state.session_logged or profile_ok
        st.session_state.logging_notice = None if profile_ok else (
            "Your details were saved for this session, but Google Sheets logging did not complete. "
            + profile_message
        )
        st.session_state.profile_capture_pending = False
        st.session_state.profile_capture_reason = None

    if action == "save":
        save_preview(current_version)
        try:
            pdf_content = build_versions_pdf(
                base_results(index_data),
                report_versions_for_export(current_version),
            )
            st.session_state.cached_pdf = pdf_content
            st.session_state.cached_pdf_version_id = current_version["id"]
            st.session_state.action_notice = "Saved. You can now download or email this view."
            st.session_state.action_notice_type = "success"
        except Exception as exc:
            st.session_state.cached_pdf = None
            st.session_state.cached_pdf_version_id = None
            st.session_state.action_notice = (
                "The view was saved, but the PDF could not be prepared yet. "
                + str(exc)
            )
            st.session_state.action_notice_type = "error"

    elif action == "email":
        ok, message = send_results_email(index_data, current_version)
        st.session_state.action_notice = message
        st.session_state.action_notice_type = "success" if ok else "error"

    st.session_state.action_in_progress = None
    st.session_state.action_version_id = None
    st.rerun()


def render_custom_follow_up(index_data: pd.DataFrame, current_version: dict[str, Any]) -> None:
    in_progress = (
        st.session_state.action_in_progress in {"save", "email"}
        and st.session_state.action_version_id == current_version["id"]
    )

    if in_progress:
        progress_label = (
            "Saving your view..."
            if st.session_state.action_in_progress == "save"
            else "Emailing your results..."
        )
        render_action_progress(progress_label)

    with st.container(key="custom_actions"):
        if st.session_state.logging_notice:
            st.markdown(
                f'<div class="action-message error">{html.escape(st.session_state.logging_notice)}</div>',
                unsafe_allow_html=True,
            )
        if st.session_state.action_notice:
            notice_type = st.session_state.action_notice_type or "warning"
            st.markdown(
                f'<div class="action-message {html.escape(notice_type)}">{html.escape(st.session_state.action_notice)}</div>',
                unsafe_allow_html=True,
            )

        left_space, save_col, gap_one, download_col, gap_two, email_col, right_space = st.columns(
            [2.4, 1.5, 0.5, 1.5, 0.5, 1.5, 2.4],
            gap="small",
        )
        del left_space, gap_one, gap_two, right_space

        export_versions = report_versions_for_export(current_version)
        cached_pdf_ready = (
            st.session_state.cached_pdf is not None
            and st.session_state.cached_pdf_version_id == current_version["id"]
        )

        with save_col:
            if in_progress:
                save_label = "Saved" if current_version["id"] in st.session_state.saved_preview_ids else "Save"
                st.button(
                    save_label,
                    disabled=True,
                    key=f"processing_save_{current_version['id']}",
                    width="stretch",
                )
            elif current_version["id"] in st.session_state.saved_preview_ids:
                st.button("Saved", disabled=True, key=f"saved_{current_version['id']}", width="stretch")
            elif st.button("Save", key=f"save_{current_version['id']}", width="stretch"):
                if st.session_state.profile_complete:
                    begin_custom_action("save", current_version)
                else:
                    st.session_state.action_notice = None
                    st.session_state.gate_action = "save"
                    st.session_state.gate_version_id = current_version["id"]
                st.rerun()

        with download_col:
            if in_progress:
                st.button(
                    "Download PDF",
                    disabled=True,
                    key=f"processing_download_{current_version['id']}",
                    width="stretch",
                )
            elif st.session_state.profile_complete and export_versions and cached_pdf_ready:
                st.download_button(
                    "Download PDF",
                    data=st.session_state.cached_pdf,
                    file_name="Coal-to-Clean Jurisdictional Readiness Index 2026.pdf",
                    mime="application/pdf",
                    on_click="ignore",
                    key=f"download_{current_version['id']}",
                    width="stretch",
                )
            elif st.button("Download PDF", key=f"gate_download_{current_version['id']}", width="stretch"):
                if not export_versions:
                    st.session_state.action_notice = "Save this customised view before downloading the PDF."
                    st.session_state.action_notice_type = "warning"
                elif not cached_pdf_ready:
                    begin_custom_action("save", current_version)
                else:
                    st.session_state.action_notice = None
                    st.session_state.gate_action = "download"
                    st.session_state.gate_version_id = current_version["id"]
                st.rerun()

        with email_col:
            if in_progress:
                st.button(
                    "Email results",
                    disabled=True,
                    key=f"processing_email_{current_version['id']}",
                    width="stretch",
                )
            elif st.button("Email results", key=f"email_{current_version['id']}", width="stretch"):
                if not export_versions:
                    st.session_state.action_notice = "Save this customised view before emailing results."
                    st.session_state.action_notice_type = "warning"
                elif st.session_state.profile_complete:
                    begin_custom_action("email", current_version)
                else:
                    st.session_state.action_notice = None
                    st.session_state.gate_action = "email"
                    st.session_state.gate_version_id = current_version["id"]
                st.rerun()

    if in_progress:
        perform_custom_action(index_data, current_version)


def valid_email(value: str) -> bool:
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value.strip()))


def send_results_email(index_data: pd.DataFrame, current_version: dict[str, Any]) -> tuple[bool, str]:
    if not backend_enabled():
        return False, "Email delivery is not connected yet. Please download the PDF instead."

    contact = st.session_state.contact
    selected_versions = report_versions_for_export(current_version)
    pdf_content = build_versions_pdf(base_results(index_data), selected_versions)
    attachments = [
        encode_attachment("Coal-to-Clean Jurisdictional Readiness Index 2026.pdf", "application/pdf", pdf_content)
    ]

    ok, message = post_backend(
        "send_email",
        {
            "session_id": st.session_state.session_id,
            "requested_at": singapore_now(),
            "name": contact["name"],
            "organisation": contact["organisation"],
            "email": contact["email"],
            "versions_requested": [f"Customised view {version['number']}" for version in selected_versions],
            "attachments": attachments,
        },
    )
    if ok:
        return True, "Your results have been emailed!"
    if "Unauthorised" in message or "unauthorised" in message:
        return False, "The Google Sheets connection needs to be reauthorised before email delivery can work."
    return False, "Email delivery was not completed. Please download the PDF instead"

def lookup_preview(preview_id: str | None) -> dict[str, Any] | None:
    if not preview_id:
        return None
    return next((preview for preview in st.session_state.previews if preview["id"] == preview_id), None)

@st.dialog("A small request", width="small")
def research_gate_dialog(index_data: pd.DataFrame) -> None:
    action = st.session_state.gate_action
    current_version = lookup_preview(st.session_state.gate_version_id) or resolve_custom_version()

    if st.session_state.gate_processing:
        progress_copy = {
            "save": "Preparing your customised readiness index...",
            "download": "Saving your details and preparing your download...",
            "email": "Saving your details and preparing your email...",
        }.get(action, "Saving your details...")
        st.markdown(
            (
                '<div class="gate-progress-wrap">'
                '<div class="gate-progress-spinner" aria-hidden="true"></div>'
                f'<div>{html.escape(progress_copy)}</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )
        time.sleep(0.18)

        if st.session_state.profile_capture_pending:
            contact = st.session_state.contact
            profile_ok, profile_message = post_backend(
                "log_profile",
                {
                    "session_id": st.session_state.session_id,
                    "captured_at": singapore_now(),
                    "name": contact["name"],
                    "organisation": contact["organisation"],
                    "email": contact["email"],
                    "reason": st.session_state.profile_capture_reason or action,
                    "industry": st.session_state.profile_industry,
                    "role": st.session_state.profile_role,
                    "started_at": st.session_state.session_started_at,
                },
                timeout=12,
            )
            st.session_state.session_logged = st.session_state.session_logged or profile_ok
            st.session_state.logging_notice = None if profile_ok else (
                "Your details were saved for this session, but database logging did not complete. "
                + profile_message
            )
            st.session_state.profile_capture_pending = False
            st.session_state.profile_capture_reason = None

        st.session_state.gate_processing = False
        st.session_state.gate_action = None
        st.session_state.gate_version_id = None

        if current_version is None:
            st.session_state.action_notice = "No customised view is available for this action."
            st.session_state.action_notice_type = "error"
            st.rerun()

        if action == "save":
            begin_custom_action("save", current_version)
        elif action == "email":
            begin_custom_action("email", current_version)
        elif action == "download":
            st.session_state.action_notice = "Thanks! Please click Download PDF to continue."
            st.session_state.action_notice_type = "success"
        st.rerun()

    action_copy = {
        "save": "Please provide your details to save your priority-adjusted view",
        "download": "Please provide your details to download your results as a PDF",
        "email": "Please provide your details to email the results to yourself",
    }.get(action, "Please provide your details to continue")

    st.markdown(f'<div class="gate-copy">{action_copy}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="gate-copy">This helps us understand how different market participants use the index.</div>',
        unsafe_allow_html=True,
    )

    if not st.session_state.profile_complete:
        contact = st.session_state.contact
        with st.form("research_profile_form"):
            name = st.text_input("Name", value=contact.get("name", ""))
            organisation = st.text_input("Organisation", value=contact.get("organisation", ""))
            email = st.text_input("Email", value=contact.get("email", ""))
            st.markdown(
                '<div class="gate-fine-print">Your details will be used only to extend this research and to fulfil your request. They will not be used for marketing or shared with third parties.</div>',
                unsafe_allow_html=True,
            )
            submitted = st.form_submit_button("Continue")

        if not submitted:
            return
        if not name.strip() or not organisation.strip() or not valid_email(email):
            st.error("Complete all fields / Enter a valid email address")
            return

        st.session_state.contact = {
            "name": name.strip(),
            "organisation": organisation.strip(),
            "email": email.strip(),
        }
        st.session_state.profile_complete = True
        st.session_state.profile_capture_pending = True
        st.session_state.profile_capture_reason = action
        st.session_state.gate_processing = True
        st.rerun()

    if current_version is None:
        st.error("No customised view is available for this action")
        return

    # This path is used only if the profile was already completed before the dialog opened.
    st.session_state.gate_processing = True
    st.rerun()


def render_overall(index_data: pd.DataFrame) -> None:
    custom_version = resolve_custom_version()
    results = custom_version["results"] if custom_version else base_results(index_data)
    display_results(index_data, results, view="overall", custom_version=custom_version)

def render_market(index_data: pd.DataFrame, overlay_data: pd.DataFrame) -> None:
    period = query_value("period", "q1")
    if period not in {"q1", "q2"}:
        period = "q1"
    display_results(
        index_data,
        market_results(index_data, overlay_data, period),
        view="market",
        market_period=period,
    )


def process_priority_submission(index_data: pd.DataFrame, submission: dict[str, Any]) -> None:
    industry = submission["industry"]
    role = submission["role"]
    responses = submission["responses"]

    st.session_state.profile_industry = industry
    st.session_state.profile_role = role

    weights = adjusted_weights(industry, responses)
    results = calculate_custom_results(index_data, weights)
    preview_id = uuid.uuid4().hex[:10]
    submitted_at = singapore_now()
    version = {
        "id": preview_id,
        "number": 1,
        "submitted_at": submitted_at,
        "industry": industry,
        "role": role,
        "responses": responses,
        "weights": weights,
        "results": results,
    }
    st.session_state.previews = [version]
    st.session_state.current_preview_id = preview_id

    rankings_payload = [
        {
            "ISO3": str(row["ISO3"]),
            "Rank": int(row["Rank"]),
            "Overall score": round(float(row["Overall score"]), 6),
            "Rank change": int(row["Rank change"]),
        }
        for _, row in results.sort_values("Rank").iterrows()
    ]
    logging_ok, logging_message = post_backend(
        "log_submission",
        {
            "session": {
                "session_id": st.session_state.session_id,
                "started_at": st.session_state.session_started_at,
                "industry": industry,
                "role": role,
            },
            "version": {
                "session_id": st.session_state.session_id,
                "version": 1,
                "submitted_at": submitted_at,
                "survey_responses": {"industry": industry, "role": role, **responses},
                "adjusted_weights": {PILLAR_SHORT[i]: weights[i] for i in range(6)},
                "resulting_rankings": rankings_payload,
            },
        },
        timeout=15,
    )
    st.session_state.session_logged = logging_ok
    st.session_state.logging_notice = None if logging_ok else (
        "Your ranking was created, but usage logging did not complete. " + logging_message
    )

    st.session_state.pending_priority_submission = None
    st.query_params["view"] = "overall"
    st.query_params["result"] = f"preview_{preview_id}"
    st.session_state.scroll_target = "top"
    st.session_state.scroll_to_top = True
    st.rerun()


def render_priorities(index_data: pd.DataFrame) -> None:
    existing = st.session_state.previews[0] if st.session_state.previews else None
    if existing is not None:
        st.query_params["view"] = "overall"
        st.query_params["result"] = f"preview_{existing['id']}"
        st.session_state.scroll_target = "top"
        st.session_state.scroll_to_top = True
        st.rerun()

    industries = sorted((name for name in INDUSTRY_MULTIPLIERS if name != "Other"), key=str.casefold)
    industry_options = ["Select your industry", *industries, "Other"]
    role_options = ["Select your role", *ROLE_OPTIONS]

    if st.session_state.profile_industry not in industry_options:
        st.session_state.profile_industry = "Select your industry"
    if st.session_state.profile_role not in role_options:
        st.session_state.profile_role = "Select your role"
    if "industry_widget" not in st.session_state or st.session_state.industry_widget not in industry_options:
        st.session_state.industry_widget = st.session_state.profile_industry
    if "role_widget" not in st.session_state or st.session_state.role_widget not in role_options:
        st.session_state.role_widget = st.session_state.profile_role

    with st.container(key="priorities_section"):
        render_top_navigation("priorities")
        st.markdown('<div class="section-title">Set Your Market Priorities</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="assessment-copy" style="max-width:980px; font-size:16px;">
            <p>The published index highlights where coal-to-clean transition opportunities appear most credible, executable and investable across jurisdictions. Different organisations, however, weigh policy certainty, market maturity, financial conditions and social safeguards differently.</p>
            <p>Adjusting the weights lets you apply your own decision criteria to the base index. This reveals which jurisdictions rise or fall under your priorities, helping focus further due diligence, market entry and engagement.</p>
            </div>
            <div class="priority-form-space"></div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("priority_form"):
            col1, col2 = st.columns(2, gap="large")
            with col1:
                with st.container(key="industry_picker"):
                    industry = st.selectbox("Your industry", industry_options, key="industry_widget")
            with col2:
                with st.container(key="role_picker"):
                    role = st.selectbox("Your role", role_options, key="role_widget")
            st.markdown("**When assessing a coal-to-clean opportunity, how important are the following?**")
            st.markdown(
                '<div class="priority-scale-note"><span>1 = Less important</span><span>5 = More important</span></div>',
                unsafe_allow_html=True,
            )

            responses: dict[str, int] = {}
            for item in SURVEY_ITEMS:
                st.markdown(f'<div class="priority-label">{html.escape(item["label"])}</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="priority-description">{html.escape(item["description"])}</div>',
                    unsafe_allow_html=True,
                )
                slider_key = f"priority_slider_{item['key']}"
                if slider_key not in st.session_state:
                    st.session_state[slider_key] = 3
                responses[item["key"]] = st.slider(
                    item["label"], min_value=1, max_value=5, step=1, key=slider_key, label_visibility="collapsed"
                )

            applied = st.form_submit_button("Apply priorities")

        if not applied:
            return
        if industry == "Select your industry" or role == "Select your role":
            st.error("Select both your industry and role before applying your priorities")
            return

        st.session_state.profile_industry = industry
        st.session_state.profile_role = role
        submission = {
            "industry": industry,
            "role": role,
            "responses": dict(responses),
        }
        with st.spinner("Updating your ranking...", show_time=False):
            time.sleep(0.7)
            process_priority_submission(index_data, submission)


def render_methodology() -> None:
    with st.container(key="methodology_section"):
        render_top_navigation("methodology")
        st.markdown('<div class="section-title">Methodology</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="methodology-copy">
              <div class="methodology-subhead">Purpose and scope</div>
              <p>The Coal-to-Clean Transition Jurisdiction Readiness Index is designed as an initial screening tool for identifying jurisdictions with relatively favourable conditions for credible coal-transition opportunities. It does not assess individual power plants or predict whether a specific transaction will succeed. Instead, it compares the national-level energy, policy, institutional, carbon-market, financial and social conditions that could enable or constrain project development.</p>
              <p>The index uses the latest sufficiently comparable public data available, generally representing 2024 or 2025 conditions. Recent market developments are incorporated separately through quarterly market-salience overlays for Q1 and Q2 2026. This separates relatively structural country readiness from faster-moving changes in market priorities.</p>

              <div class="methodology-subhead">Country universe and materiality screening</div>
              <p>The index's universe comprises jurisdictions with operating coal-fired power generation capacity recorded in Global Energy Monitor’s January 2026 Global Coal Plant Tracker. Jurisdictions without recorded operating coal capacity were excluded because they fall outside the intended coal-transition opportunity set.</p>
              <p>A central materiality screen was then applied using three dimensions: operating coal capacity, coal dependence in the electricity mix and estimated lifetime emissions associated with the operating fleet. A jurisdiction passed the central screen where it met at least two of the three median-based thresholds. This produced a final comparative universe of 36 jurisdictions.</p>
              <p>The screen measures the materiality of the coal-transition opportunity rather than readiness. Coal capacity was therefore not rewarded again within the readiness score, avoiding double counting between opportunity scale and enabling conditions.</p>

              <div class="methodology-subhead">Index structure</div>
              <p>The index comprises six pillars:</p>
              <ul>
                <li>Energy system readiness</li>
                <li>Policy and transition commitment</li>
                <li>Governance and institutional capacity</li>
                <li>Carbon market maturity</li>
                <li>Financial and market viability</li>
                <li>Just transition and social credibility</li>
              </ul>
              <p>The indicators combine current conditions with recent trajectory measures. Current-condition indicators capture relatively structural characteristics, such as renewable electricity penetration, governance quality and carbon-market experience. Trajectory indicators capture the direction of change, including recent trends in coal share, renewable electricity and grid emissions intensity.</p>
              <p>Indicators were selected according to six principles: relevance to coal-transition execution; coverage across the eligible country universe; comparability between jurisdictions; limited conceptual overlap; data quality and institutional credibility; and reproducibility using publicly available information. Where two indicators appeared to measure substantially the same underlying condition, one was removed, archived or assigned no weight.</p>
              <p>A detailed indicator register, including definitions, data years, sources and scoring treatment, is provided in the appendix and supporting workbooks.</p>

              <div class="methodology-subhead">Data treatment and scoring</div>
              <p>Raw indicators were first aligned so that higher values consistently represented stronger readiness. Measures of risk or transition pressure, such as coal construction, inflation or exchange-rate instability, were reverse-scored.</p>
              <p>Indicators were converted to a common 0–100 scale. Percentile bounds were used where necessary to reduce the influence of extreme outliers, while logarithmic transformation was applied to highly skewed carbon-market activity measures such as project counts, credit issuance and retirements. Country scores were then aggregated into pillar scores and an overall index score.</p>
              <p>Missing observations were not automatically treated as zero. Where defensible estimates were required, they were produced using documented comparator or secondary-source evidence and explicitly flagged. The sensitivity testing assessed whether these estimates materially affected the results.</p>

              <div class="methodology-subhead">Pillar weights</div>
              <p>The six pillars were assigned overall weights of 20% for energy-system readiness, 20% for policy and transition commitment, 15% for governance and institutional capacity, 20% for carbon-market maturity, 15% for financial and market viability and 10% for just-transition and social credibility. The weights reflect their relative relevance to coal-transition execution, the strength of the available cross-country evidence and the need to avoid duplicating similar risks across pillars.</p>

              <div class="methodology-subhead">Indicator weights</div>
              <p>Indicators within each pillar were weighted according to their direct relationship with the pillar’s due-diligence question, data quality and distinctiveness from other measures. Detailed indicator weights are retained in the analytical model rather than reproduced in the report. Alternative weighting scenarios were tested to determine whether the final rankings depended excessively on the selected configuration.</p>

              <div class="methodology-subhead">Market-salience overlay</div>
              <p>The quarterly overlay assesses which readiness factors were receiving increased attention from investors, buyers, policymakers and other market participants during Q1 and Q2 2026. It does not assign positive or negative media scores to individual countries. Instead, it changes the relative emphasis placed on existing indicators, with the same adjusted weights applied across all jurisdictions.</p>
              <p>Nexis was used as the main systematic discovery corpus, supported by targeted Google searches for primary-source verification. A pilot collection from QCIntel was excluded from the formal quarterly comparison because comparable access could not be completed after the subscription ended earlier than expected.</p>
              <p>Duplicate reporting of the same development was clustered as one evidence event. Each event could support more than one construct where it contained genuinely distinct claims. Evidence was assessed according to its relevance to coal transition, evidence type, strength and source independence. Construct-level evidence was converted into multipliers of 1.00, 1.05, 1.10 or 1.20.</p>
              <p>Each scored construct was assigned to its closest primary index indicator. Cross-cutting or project-specific issues without a sufficiently direct country-level indicator were retained as contextual findings rather than forced into the scoring model. Adjusted indicator weights were renormalised to 100% before calculating Q1 and Q2 overlay scores.</p>

              <div class="methodology-subhead">Validation and sensitivity testing</div>
              <p>The index was tested under alternative pillar and indicator weighting scenarios. Rank correlations across these scenarios ranged from approximately 0.93 to 0.99, indicating that the broad ranking pattern was stable. The maximum observed inter-pillar correlation was 0.764, below the 0.800 threshold used to flag potentially excessive overlap.</p>
              <p>The market overlays produced similarly limited changes. Spearman rank correlations between the base index and the Q1 and Q2 results were approximately 0.998, and no jurisdiction moved by more than two ranking positions. The overlay therefore adds current market context without overturning the structural findings of the base index.</p>

              <div class="methodology-subhead">AI-assisted analysis and quality control</div>
              <p>Generative AI was used to support data extraction, evidence coding, formula development, consistency checks and iterative analytical testing. The researcher designed and refined the prompts, defined the classification framework, reviewed source documents, corrected coding decisions and retained responsibility for the final methodology and interpretation. AI outputs were treated as provisional until checked against the underlying data or evidence source.</p>

              <div class="methodology-subhead">Limitations</div>
              <p>The index is a national-level screening tool, whereas coal-transition transactions ultimately depend on plant-level economics, ownership, contracts, grid access and community consent. Some international datasets are published with a time lag, and English-language source availability is uneven across jurisdictions.</p>
              <p>Several relevant issues could not be measured consistently, including plant-specific community support, grid-connection constraints and transaction-level bankability. These are identified as areas for subsequent project due diligence rather than represented through weak or inconsistent proxies. The results should therefore support prioritisation and comparison, not be interpreted as investment, legal or carbon-credit integrity advice.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


initialise_state()

try:
    index_data = load_index_data(str(DATA_FILE))
    overlay_data = load_market_overlay(str(DATA_FILE))
except Exception as exc:
    st.error(str(exc))
    st.stop()

view = query_value("view", "overall")
if view not in {"overall", "market", "priorities", "methodology"}:
    view = "overall"

render_sidebar(view)
render_header()

if view == "market":
    render_market(index_data, overlay_data)
elif view == "priorities":
    render_priorities(index_data)
elif view == "methodology":
    render_methodology()
else:
    render_overall(index_data)

if st.session_state.scroll_to_top or st.session_state.scroll_target:
    target_id = st.session_state.scroll_target or "top"
    components.html(
        f"""
        <script>
        (() => {{
          const win = window.parent;
          const doc = win.document;
          const targetId = {json.dumps(target_id)};
          const scrollTarget = () => {{
            try {{
              if (doc.activeElement && typeof doc.activeElement.blur === 'function') {{
                doc.activeElement.blur();
              }}
              const candidates = [
                doc.querySelector('section[data-testid="stMain"]'),
                doc.querySelector('[data-testid="stMain"]'),
                doc.querySelector('[data-testid="stAppViewContainer"]'),
                doc.scrollingElement,
                doc.documentElement,
                doc.body
              ].filter(Boolean);
              if (targetId === 'top') {{
                candidates.forEach((target) => {{
                  target.scrollTop = 0;
                  if (typeof target.scrollTo === 'function') target.scrollTo(0, 0);
                }});
                win.scrollTo(0, 0);
              }}
              const anchor = doc.getElementById(targetId);
              if (anchor && typeof anchor.scrollIntoView === 'function') {{
                anchor.scrollIntoView({{block: 'start', inline: 'nearest'}});
              }}
              const url = new URL(win.location.href);
              url.hash = targetId;
              win.history.replaceState({{}}, '', url.toString());
            }} catch (error) {{
              win.scrollTo(0, 0);
            }}
          }};
          [0, 50, 200, 500, 1000, 1800, 2800].forEach((delay) => win.setTimeout(scrollTarget, delay));
        }})();
        </script>
        """,
        height=0,
        width=0,
    )
    st.session_state.scroll_to_top = False
    st.session_state.scroll_target = None

if st.session_state.gate_action:
    research_gate_dialog(index_data)

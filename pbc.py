"""
pbc.py - the OAG Preliminary Information Request tracker UI.

Entry point used by app.py:
    render_page()   -> renders the whole tracker
"""
import base64
import html as _html
import inspect
import os
import re
from datetime import date, datetime, timedelta, timezone
from functools import lru_cache

import streamlit as st
from streamlit_quill import st_quill

import pbc_db as pdb
import pbc_export as pex
from pbc_parser import html_to_text, parse_request_docx, sanitize_html

HERE = os.path.dirname(os.path.abspath(__file__))
SEED_PATH = os.path.join(HERE, "pbc_seed.json")
LOGO_PATH = os.path.join(HERE, "logo.jpg")

# Streamlit renamed use_container_width -> width="stretch"; support both old and new versions.
STRETCH = {"width": "stretch"} if "width" in inspect.signature(st.button).parameters else {"use_container_width": True}

# Each request row is its own fragment: clicking a button inside a row re-runs only that row,
# not the whole page (big speed-up). Falls back gracefully on very old Streamlit versions.
_fragment = getattr(st, "fragment", None) or (lambda f: f)

KIGALI = timezone(timedelta(hours=2))  # Rwanda: UTC+2, no daylight saving


def today():
    return datetime.now(KIGALI).date()


# ---------------------------------------------------------------------------
# Rich-text editor toolbars
# ---------------------------------------------------------------------------
ADV_TOOLBAR = [
    [{"header": [1, 2, 3, False]}, {"size": ["small", False, "large", "huge"]}],
    ["bold", "italic", "underline", "strike"],
    [{"color": []}, {"background": []}],
    [{"list": "ordered"}, {"list": "bullet"}, {"indent": "-1"}, {"indent": "+1"}],
    [{"align": []}],
    [{"script": "sub"}, {"script": "super"}],
    ["blockquote", "link", "clean"],
]
COMPACT_TOOLBAR = [["bold", "italic", "underline"], [{"list": "ordered"}, {"list": "bullet"}], ["link", "clean"]]


# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
PAGE_CSS = """
<style>
.pbc-wrap, .pbc-wrap *{font-family:'Inter','Segoe UI','Source Sans Pro',system-ui,-apple-system,sans-serif;}
@keyframes pbcFade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}

header{visibility:hidden;}
#MainMenu{visibility:hidden;}

/* ---------- hero (compact) ---------- */
.block-container,[data-testid="stMainBlockContainer"]{padding-top:.2rem!important;padding-left:.9rem!important;padding-right:.9rem!important;max-width:100%!important;}

.pbc-delivery{margin-bottom:8px;}
.pbc-hero{position:relative;overflow:hidden;display:flex;align-items:center;gap:14px;flex-wrap:wrap;
  background:linear-gradient(120deg,#12294f 0%,#1e3c72 42%,#2f5fb3 100%);color:#fff;border-radius:14px;
  padding:10px 18px;margin:6px 0 8px;box-shadow:0 6px 16px rgba(30,60,114,.22);}
.pbc-logo{width:42px;height:42px;border-radius:10px;background:#fff;display:flex;align-items:center;justify-content:center;
  overflow:hidden;flex:0 0 auto;z-index:1;}
.pbc-logo img{width:100%;height:100%;object-fit:contain;}
.pbc-hero-text{flex:1 1 300px;z-index:1;min-width:240px;}
.pbc-eyebrow{font-size:9.5px;letter-spacing:.12em;font-weight:700;opacity:.7;text-transform:uppercase;}
.pbc-h1{font-size:18px;font-weight:800;line-height:1.2;margin:1px 0 1px;color:#fff;}
.pbc-sub{font-size:11.5px;opacity:.82;color:#fff;}
.pbc-hero-right{display:flex;align-items:center;gap:10px;z-index:1;}
.pbc-ring{--p:0;width:50px;height:50px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  background:conic-gradient(#4ade80 calc(var(--p)*1%),rgba(255,255,255,.22) 0);}
.pbc-ring>div{width:40px;height:40px;border-radius:50%;background:#1b3868;display:flex;align-items:center;justify-content:center;}
.pbc-ring b{font-size:12px;font-weight:800;color:#fff;}
.pbc-ring span{display:none;}
.pbc-due{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.28);border-radius:10px;
  padding:5px 12px;text-align:center;line-height:1.15;}
.pbc-due b{display:inline;font-size:18px;font-weight:800;color:#fff;margin-right:4px;}
.pbc-due span{font-size:10.5px;opacity:.85;text-transform:uppercase;letter-spacing:.05em;color:#fff;}
.pbc-due div{font-size:10px!important;margin-top:1px!important;}
.pbc-due.late{background:rgba(239,68,68,.28);border-color:rgba(254,202,202,.6);}
.pbc-due.soon{background:rgba(245,158,11,.28);border-color:rgba(253,230,138,.6);}

/* ---------- KPI cards ---------- */
.pbc-kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin-bottom:10px;}
@media (max-width:1150px){.pbc-kpis{grid-template-columns:repeat(3,minmax(0,1fr));}}
@media (max-width:640px){.pbc-kpis{grid-template-columns:repeat(2,minmax(0,1fr));}}
.pbc-kpi{background:#fff;border:1px solid #e6ebf4;border-radius:16px;padding:8px 14px;position:relative;overflow:hidden;
  box-shadow:0 2px 8px rgba(15,23,42,.05);transition:transform .2s,box-shadow .2s;}
.pbc-kpi:hover{transform:translateY(-3px);box-shadow:0 10px 22px rgba(15,23,42,.10);}
.pbc-kpi:before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--c);}
.pbc-kpi .ic{position:absolute;right:10px;top:8px;font-size:17px;opacity:.9;}
.pbc-kpi .lb{font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#64748b;}
.pbc-kpi .vl{font-size:22px;font-weight:800;color:#0f172a;line-height:1.15;margin-top:2px;}
.pbc-kpi .sb{font-size:11px;color:#64748b;font-weight:600;min-height:0;}

/* ---------- section headers ---------- */
.pbc-sec{display:flex;align-items:center;gap:14px;margin:16px 0 8px;padding:8px 14px;border-radius:12px;
  background:linear-gradient(90deg,#eaf0fa,#f6f9ff);border:1px solid #dbe5f5;}
.pbc-sec .num{min-width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,#1e3c72,#2f5fb3);color:#fff;
  font-weight:800;display:flex;align-items:center;justify-content:center;font-size:15px;padding:0 6px;}
.pbc-sec .tt{font-size:17px;font-weight:800;color:#12294f;flex:1;}
.pbc-sec .mt{font-size:12px;color:#475569;font-weight:600;text-align:right;white-space:nowrap;}
.pbc-sec .bar{width:140px;height:8px;border-radius:99px;background:#d6e0f2;overflow:hidden;flex:0 0 auto;}
.pbc-sec .bar span{display:block;height:100%;background:linear-gradient(90deg,#34d399,#22c55e);border-radius:99px;}
.pbc-sub-h{margin:12px 0 6px 6px;font-size:12px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:#64748b;}
.pbc-sub-h:before{content:"▸ ";color:#2f5fb3;}

/* ---------- request cards (st.container with key) ---------- */
div[class*="st-key-pbc-card-"]{background:#fff;border:1px solid #e6ebf4;border-left:6px solid #94a3b8;border-radius:14px;
  padding:14px 16px 16px;margin-bottom:8px;box-shadow:0 1px 4px rgba(15,23,42,.05);transition:box-shadow .2s,transform .2s;}
div[class*="st-key-pbc-card-"]:hover{box-shadow:0 10px 24px rgba(15,23,42,.10);transform:translateY(-1px);}
div[class*="st-key-pbc-card-pending-"]{border-left-color:#94a3b8;}
div[class*="st-key-pbc-card-progress-"]{border-left-color:#f59e0b;}
div[class*="st-key-pbc-card-partial-"]{border-left-color:#3b82f6;}
div[class*="st-key-pbc-card-submitted-"]{border-left-color:#22c55e;background:linear-gradient(90deg,#f4fdf7,#fff 40%);}
div[class*="st-key-pbc-card-overdue-"]{border-left-color:#ef4444;background:linear-gradient(90deg,#fff5f5,#fff 40%);}
div[class*="st-key-pbc-card-na-"]{border-left-color:#cbd5e1;opacity:.78;}
.pbc-ref{display:inline-block;min-width:46px;text-align:center;padding:6px 10px;border-radius:10px;background:#eef2ff;color:#3730a3;
  font-weight:800;font-size:13px;letter-spacing:.02em;}
.pbc-desc{color:#1f2937;font-size:14.5px;line-height:1.55;word-break:break-word;}
.pbc-desc p{margin:0 0 .35rem;}
.pbc-desc ul,.pbc-desc ol{margin:.1rem 0 .4rem 1.2rem;padding:0;}
.pbc-desc li{margin:.1rem 0;}
.pbc-desc a{color:#2563eb;text-decoration:none;border-bottom:1px dotted #93b4f5;}
.pbc-desc blockquote{margin:.3rem 0;padding:.1rem .8rem;border-left:3px solid #cbd5e1;color:#475569;}
.pbc-meta{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;}
.chip{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600;
  border:1px solid transparent;max-width:100%;line-height:1.5;}
.chip.dir{background:#eef2ff;color:#3730a3;border-color:#dfe5ff;}
.chip.resp{background:#ecfdf5;color:#047857;border-color:#c9f2de;}
.chip.type{background:#f1f5f9;color:#334155;border-color:#e2e8f0;}
.chip.due{background:#f1f5f9;color:#334155;border-color:#e2e8f0;}
.chip.due.soon{background:#fffbeb;color:#92400e;border-color:#fde68a;}
.chip.due.late{background:#fef2f2;color:#b91c1c;border-color:#fecaca;}
.chip.due.done{background:#f0fdf4;color:#166534;border-color:#bbf7d0;}
.chip.file{background:#f5f3ff;color:#5b21b6;border-color:#e4dcff;max-width:240px;}
.chip.file span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.chip.st-pending{background:#e2e8f0;color:#334155;}
.chip.st-progress{background:#fef3c7;color:#92400e;}
.chip.st-partial{background:#dbeafe;color:#1e40af;}
.chip.st-submitted{background:#dcfce7;color:#166534;}
.chip.st-overdue{background:#fee2e2;color:#b91c1c;}
.chip.st-na{background:#f1f5f9;color:#64748b;}
.pbc-notes{margin-top:8px;padding:8px 12px;background:#fffbeb;border:1px solid #fde68a;border-radius:10px;color:#78350f;font-size:13px;}
.pbc-notes b.h{display:block;font-size:10.5px;letter-spacing:.08em;text-transform:uppercase;color:#b45309;margin-bottom:2px;}
.pbc-notes p{margin:0 0 .2rem;}
.pbc-delivery{margin-top:6px;font-size:12.5px;color:#64748b;display:flex;gap:8px;align-items:flex-start;}
.pbc-delivery>div{flex:1;min-width:0;}
.pbc-delivery p{margin:0;}
.pbc-delivery{line-height:1.5;}
.pbc-delivery a{color:#2563eb;}

/* card buttons / select */
div[class*="st-key-pbc-card-"] button{border-radius:10px;font-weight:600;font-size:13px;padding:.2rem .5rem;min-height:2.1rem;}
div[class*="st-key-pbc_docsfull_"] button{background:#eef2ff;border-color:#c7d2fe;color:#3730a3;}
div[class*="st-key-pbc-card-"] [data-baseweb="select"]>div{border-radius:10px;min-height:2.1rem;font-size:13px;}

/* brand accent */
.stTabs [data-baseweb="tab-highlight"],.stTabs .react-aria-SelectionIndicator,[data-testid="stTabs"] .react-aria-SelectionIndicator{background-color:#2a5298!important;height:3px;}
.stTabs [role="tab"][aria-selected="true"],.stTabs [role="tab"][aria-selected="true"] *,[data-testid="stTabs"] [role="tab"][aria-selected="true"] *{color:#1e3c72!important;font-weight:700;}
.stTabs [role="tab"],[data-testid="stTabs"] [role="tab"]{font-weight:600;}
button[kind="primary"],[data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#1e3c72,#2f5fb3)!important;border:0!important;
  color:#fff!important;font-weight:700;box-shadow:0 4px 12px rgba(30,60,114,.28);transition:all .2s ease;}
button[kind="primary"]:hover,[data-testid="stBaseButton-primary"]:hover{filter:brightness(1.12);transform:translateY(-1px);box-shadow:0 8px 18px rgba(30,60,114,.32);}
button[kind="primary"]:disabled,[data-testid="stBaseButton-primary"]:disabled{background:#cbd5e1!important;box-shadow:none;color:#fff!important;}

/* empty state */
.pbc-empty{text-align:center;padding:46px 20px;border:2px dashed #cbd5e1;border-radius:18px;color:#64748b;background:#f8fafc;}
.pbc-empty .big{font-size:42px;}

/* ================= compact + professional overrides ================= */
div[data-testid="stVerticalBlock"]{gap:.45rem;}
.pbc-kpis{gap:8px;margin-bottom:4px;}
.pbc-kpi{border-radius:10px;padding:6px 12px;}
.pbc-kpi .vl{font-size:20px;}
.pbc-kpi .lb{font-size:10px;}
.pbc-kpi .sb{font-size:10.5px;}
.pbc-kpi .ic{font-size:15px;top:6px;}
.pbc-sec{margin:10px 0 4px;padding:5px 12px;border-radius:9px;gap:10px;}
.pbc-sec .num{min-width:26px;height:26px;border-radius:7px;font-size:12px;}
.pbc-sec .tt{font-size:14px;}
.pbc-sec .mt{font-size:11px;}
.pbc-sec .bar{width:100px;height:6px;}
.pbc-sub-h{margin:6px 0 2px 4px;font-size:10.5px;}
div[class*="st-key-pbc-card-"]{border-radius:9px;border-left-width:4px;padding:8px 12px 8px;margin-bottom:0;box-shadow:0 1px 2px rgba(15,23,42,.05);}
div[class*="st-key-pbc-card-"]:hover{transform:none;box-shadow:0 3px 10px rgba(15,23,42,.09);}
.pbc-ref{min-width:36px;padding:3px 7px;border-radius:7px;font-size:11.5px;}
.pbc-desc{font-size:13.5px;line-height:1.42;}
.pbc-desc p{margin:0 0 .15rem;}
.pbc-desc ul,.pbc-desc ol{margin:0 0 .2rem 1.1rem;}
.pbc-desc li{margin:0;}
.pbc-meta{gap:4px;margin-top:5px;}
.chip{padding:1px 8px;font-size:11px;gap:4px;}
.pbc-delivery{font-size:11.5px;margin-top:3px;margin-bottom:2px;}
.pbc-notes{margin-top:5px;padding:5px 9px;font-size:12px;border-radius:7px;}
div[class*="st-key-pbc-card-"] button{min-height:1.8rem;padding:.1rem .4rem;font-size:12px;border-radius:7px;}
div[class*="st-key-pbc-card-"] [data-baseweb="select"]>div{min-height:1.8rem;font-size:12px;border-radius:7px;}
div[class*="st-key-pbc-card-"] div[data-testid="stHorizontalBlock"]{gap:.5rem;}

/* ================= status colours on the dropdown (old + new Streamlit DOM) ================= */
div[class*="st-key-pbc-st-"] input{font-weight:700!important;}
div[class*="st-key-pbc-st-pending-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-pending-"] .react-aria-ComboBox>div[role="group"]{background:#e2e8f0!important;border:1px solid #cbd5e1!important;}
div[class*="st-key-pbc-st-pending-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-pending-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-pending-"] .react-aria-ComboBox svg{color:#334155!important;-webkit-text-fill-color:#334155!important;fill:#334155!important;font-weight:700;}
div[class*="st-key-pbc-st-progress-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-progress-"] .react-aria-ComboBox>div[role="group"]{background:#fef3c7!important;border:1px solid #fcd34d!important;}
div[class*="st-key-pbc-st-progress-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-progress-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-progress-"] .react-aria-ComboBox svg{color:#92400e!important;-webkit-text-fill-color:#92400e!important;fill:#92400e!important;font-weight:700;}
div[class*="st-key-pbc-st-partial-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-partial-"] .react-aria-ComboBox>div[role="group"]{background:#dbeafe!important;border:1px solid #93c5fd!important;}
div[class*="st-key-pbc-st-partial-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-partial-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-partial-"] .react-aria-ComboBox svg{color:#1e40af!important;-webkit-text-fill-color:#1e40af!important;fill:#1e40af!important;font-weight:700;}
div[class*="st-key-pbc-st-submitted-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-submitted-"] .react-aria-ComboBox>div[role="group"]{background:#dcfce7!important;border:1px solid #86efac!important;}
div[class*="st-key-pbc-st-submitted-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-submitted-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-submitted-"] .react-aria-ComboBox svg{color:#166534!important;-webkit-text-fill-color:#166534!important;fill:#166534!important;font-weight:700;}
div[class*="st-key-pbc-st-overdue-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-overdue-"] .react-aria-ComboBox>div[role="group"]{background:#fee2e2!important;border:1px solid #fca5a5!important;}
div[class*="st-key-pbc-st-overdue-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-overdue-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-overdue-"] .react-aria-ComboBox svg{color:#b91c1c!important;-webkit-text-fill-color:#b91c1c!important;fill:#b91c1c!important;font-weight:700;}
div[class*="st-key-pbc-st-na-"] [data-baseweb="select"]>div,div[class*="st-key-pbc-st-na-"] .react-aria-ComboBox>div[role="group"]{background:#f1f5f9!important;border:1px solid #e2e8f0!important;}
div[class*="st-key-pbc-st-na-"] [data-baseweb="select"] *,div[class*="st-key-pbc-st-na-"] .react-aria-ComboBox input,div[class*="st-key-pbc-st-na-"] .react-aria-ComboBox svg{color:#64748b!important;-webkit-text-fill-color:#64748b!important;fill:#64748b!important;font-weight:700;}
div[class*="st-key-pbc-card-"] .react-aria-ComboBox>div[role="group"]{min-height:1.8rem;border-radius:7px;font-size:12px;}
div[class*="st-key-pbc-card-"]{padding-bottom:12px!important;}

/* ================= distinct design: ledger rows, status tiles, progress rules ================= */
.pbc-kpis{gap:0;background:#fff;border:1px solid #dfe6f1;border-radius:10px;overflow:hidden;margin-bottom:6px;}
.pbc-kpi{border:0!important;border-radius:0;box-shadow:none!important;background:transparent;padding:7px 14px;border-right:1px solid #e8edf5!important;transform:none!important;}
.pbc-kpi:last-child{border-right:0!important;}
.pbc-kpi:before{display:none;}
.pbc-kpi .vl{color:var(--c);font-size:21px;font-weight:800;letter-spacing:-.01em;}
.pbc-kpi .lb{color:#6b7a90;}
.pbc-kpi .ic{display:none;}
@media (max-width:1150px){.pbc-kpi{border-bottom:1px solid #e8edf5!important;}}
.pbc-sec{background:linear-gradient(90deg,#16a34a var(--p),#d3dcea var(--p)) left bottom/100% 3px no-repeat!important;
  border:0!important;border-radius:0!important;padding:6px 2px 9px!important;margin:14px 0 2px!important;gap:10px;align-items:baseline;}
.pbc-sec .num{background:none!important;min-width:0;height:auto;color:#1e3c72;font-size:20px;font-weight:800;letter-spacing:-.02em;
  font-variant-numeric:tabular-nums;padding:0;display:inline;}
.pbc-sec .tt{font-size:13px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#12294f;}
.pbc-sec .mt{font-size:11px;color:#6b7a90;font-weight:500;margin-left:auto;}
.pbc-sec .mt b{color:#16a34a;font-weight:800;}
.pbc-sec .bar{display:none;}
.pbc-sub-h{color:#6b7a90;}
.pbc-sub-h:before{content:"";}
div[class*="st-key-pbc-card-"]{border:0!important;border-bottom:1px solid #e6ebf3!important;border-radius:0!important;
  box-shadow:none!important;background:transparent!important;padding:9px 4px 15px!important;margin:0!important;transform:none!important;}
div[class*="st-key-pbc-card-"]:hover{background:#f6f9fe!important;box-shadow:none!important;}
div[class*="st-key-pbc-card-na-"]{opacity:.62;}
.pbc-ref{min-width:38px;padding:4px 6px;border-radius:5px;font-size:11.5px;font-weight:800;color:#fff;background:#64748b;
  font-variant-numeric:tabular-nums;letter-spacing:0;}
.pbc-ref.s-pending{background:#64748b;}
.pbc-ref.s-progress{background:#ca8a04;}
.pbc-ref.s-partial{background:#2563eb;}
.pbc-ref.s-submitted{background:#16a34a;}
.pbc-ref.s-overdue{background:#dc2626;}
.pbc-ref.s-na{background:#a8b3c4;}
.pbc-cmt{margin-top:5px;padding:4px 9px;background:#f1f5fb;border-radius:6px;font-size:11.5px;color:#475569;line-height:1.4;}
.pbc-cmt b{color:#1e3c72;} .pbc-cmt span{color:#94a3b8;margin-right:4px;} .pbc-cmt em{color:#94a3b8;font-style:normal;}
.pbc-cm{display:flex;gap:10px;align-items:flex-start;padding:6px 0;}
.pbc-cm .av{width:30px;height:30px;border-radius:50%;color:#fff;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center;flex:0 0 auto;}
.pbc-cm .bd{background:#f1f5fb;border-radius:4px 12px 12px 12px;padding:7px 12px;flex:1;min-width:0;}
.pbc-cm .hd{font-size:12px;color:#1e3c72;} .pbc-cm .hd span{color:#94a3b8;margin-left:8px;font-size:11px;}
.pbc-cm .tx{font-size:13.5px;color:#1f2937;line-height:1.45;margin-top:2px;word-break:break-word;}
div[class*="st-key-pbc_cmtfull_"] button{background:#fff7ed;border-color:#fed7aa;color:#9a3412;}

/* comment box: right column, under the buttons */
.pbc-cmt{margin-top:2px;padding:6px 10px 7px;background:#fff8e6;border:1px solid #f6e3b0;border-radius:7px;font-size:12px;color:#475569;line-height:1.4;}
.pbc-cmt .lb{font-size:9.5px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#a16207;display:flex;justify-content:space-between;margin-bottom:2px;}
.pbc-cmt .lb i{font-style:normal;font-weight:600;letter-spacing:0;text-transform:none;color:#b8956a;}
.pbc-cmt .who b{color:#1e3c72;font-size:12px;} .pbc-cmt .who span{color:#94a3b8;font-size:10.5px;margin-left:6px;}
.pbc-cmt .tx{color:#334155;font-size:12.5px;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word;}

/* sub-section headings: tinted band so they stand out between rows */
.pbc-sub-h{display:flex;align-items:center;gap:8px;margin:10px 0 4px!important;padding:6px 12px;border-radius:7px;
  background:linear-gradient(90deg,#cfe0fb 0%,#e3edfd 55%,#f1f6fe 100%);color:#16356b;font-size:12px;font-weight:800;
  letter-spacing:.07em;text-transform:uppercase;box-shadow:inset 0 0 0 1px #c3d6f6;}
.pbc-sub-h:before{content:"";width:8px;height:8px;border-radius:2px;background:#2f5fb3;transform:rotate(45deg);flex:0 0 auto;}

/* ================= bold KPI cards ================= */
.pbc-kpis{gap:8px!important;background:none!important;border:0!important;border-radius:0!important;overflow:visible!important;margin-bottom:6px;}
.pbc-kpi{background:var(--c)!important;border:0!important;border-radius:10px!important;padding:9px 14px!important;
  box-shadow:0 3px 8px rgba(15,23,42,.18)!important;position:relative;overflow:hidden;}
.pbc-kpi .vl,.pbc-kpi .lb,.pbc-kpi .sb{color:#fff!important;}
.pbc-kpi .vl{font-size:26px!important;font-weight:800;line-height:1.1;}
.pbc-kpi .lb{opacity:.88;font-size:10.5px!important;}
.pbc-kpi .sb{opacity:.82;}
.pbc-kpi .ic{display:none!important;}
div[class*="st-key-pbc-stbtn-"] button *{font-weight:800!important;}

/* ================= status pill-button + sleek text buttons ================= */
div[class*="st-key-pbc-stbtn-"] button{font-weight:800;font-size:12px;letter-spacing:.01em;border-radius:999px;min-height:1.8rem;
  white-space:nowrap;border:1px solid transparent;transition:filter .15s,transform .15s;}
div[class*="st-key-pbc-stbtn-"] button:hover{filter:brightness(.95);transform:translateY(-1px);}
div[class*="st-key-pbc-stbtn-"] button,div[class*="st-key-pbc-stbtn-"] button *{white-space:nowrap;}
div[class*="st-key-pbc-stbtn-pending-"] button{background:#e2e8f0;border-color:#cbd5e1;}
div[class*="st-key-pbc-stbtn-pending-"] button *{color:#334155!important;}
div[class*="st-key-pbc-stbtn-progress-"] button{background:#fef3c7;border-color:#fcd34d;}
div[class*="st-key-pbc-stbtn-progress-"] button *{color:#92400e!important;}
div[class*="st-key-pbc-stbtn-partial-"] button{background:#dbeafe;border-color:#93c5fd;}
div[class*="st-key-pbc-stbtn-partial-"] button *{color:#1e40af!important;}
div[class*="st-key-pbc-stbtn-submitted-"] button{background:#dcfce7;border-color:#86efac;}
div[class*="st-key-pbc-stbtn-submitted-"] button *{color:#166534!important;}
div[class*="st-key-pbc-stbtn-overdue-"] button{background:#fee2e2;border-color:#fca5a5;}
div[class*="st-key-pbc-stbtn-overdue-"] button *{color:#b91c1c!important;}
div[class*="st-key-pbc-stbtn-na-"] button{background:#f1f5f9;border-color:#e2e8f0;}
div[class*="st-key-pbc-stbtn-na-"] button *{color:#64748b!important;}

div[class*="st-key-pbc_docs"] button,div[class*="st-key-pbc_cmt"] button,div[class*="st-key-pbc_edit_"] button{
  background:#fff;border:1px solid #d5deec;border-radius:7px;font-weight:600;font-size:12px;white-space:nowrap;box-shadow:none;}
div[class*="st-key-pbc_docs"] button *,div[class*="st-key-pbc_cmt"] button *,div[class*="st-key-pbc_edit_"] button *{color:#334155!important;white-space:nowrap;}
div[class*="st-key-pbc_docs"] button:hover,div[class*="st-key-pbc_cmt"] button:hover,div[class*="st-key-pbc_edit_"] button:hover{
  background:#eef3fb;border-color:#9db6e0;}
div[class*="st-key-pbc_docsfull_"] button{background:#eef2ff;border-color:#c7d2fe;}
div[class*="st-key-pbc_docsfull_"] button *{color:#3730a3!important;}
div[class*="st-key-pbc_cmtfull_"] button{background:#fff7ed;border-color:#fed7aa;}
div[class*="st-key-pbc_cmtfull_"] button *{color:#9a3412!important;}

/* ================= glass finish: KPI cards + header ================= */
.pbc-kpi{
  background:linear-gradient(135deg,
      color-mix(in srgb,var(--c) 72%,#fff) 0%,
      var(--c) 48%,
      color-mix(in srgb,var(--c) 80%,#000) 100%)!important;
  box-shadow:
      inset 0 1px 0 rgba(255,255,255,.55),
      inset 0 0 0 1px rgba(255,255,255,.20),
      0 10px 22px -8px color-mix(in srgb,var(--c) 75%,transparent)!important;
  transition:transform .25s ease, box-shadow .25s ease;}
.pbc-kpi:hover{transform:translateY(-2px)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.6),inset 0 0 0 1px rgba(255,255,255,.28),
             0 16px 28px -8px color-mix(in srgb,var(--c) 80%,transparent)!important;}
/* glossy sheen across the top-left */
.pbc-kpi:after{content:"";position:absolute;inset:0;pointer-events:none;border-radius:inherit;
  background:linear-gradient(158deg,rgba(255,255,255,.34) 0%,rgba(255,255,255,.10) 34%,rgba(255,255,255,0) 58%);}
/* (corner circle removed) */
.pbc-kpi:before{display:none!important;}
.pbc-kpi .vl,.pbc-kpi .lb,.pbc-kpi .sb{position:relative;z-index:1;text-shadow:0 1px 2px rgba(0,0,0,.18);}

.pbc-hero{
  background:
    radial-gradient(520px 200px at 88% -30%,rgba(125,211,252,.42),transparent 62%),
    radial-gradient(460px 180px at 8% 130%,rgba(167,139,250,.30),transparent 62%),
    linear-gradient(120deg,#0e2145 0%,#1e3c72 46%,#3367bd 100%)!important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.28),inset 0 0 0 1px rgba(255,255,255,.08),0 12px 26px -10px rgba(30,60,114,.55)!important;}
.pbc-hero:before{content:"";position:absolute;inset:0;pointer-events:none;border-radius:inherit;
  background:linear-gradient(180deg,rgba(255,255,255,.14),rgba(255,255,255,0) 46%);}
.pbc-hero:after{display:none!important;}
.pbc-logo{box-shadow:0 4px 12px rgba(0,0,0,.25),inset 0 0 0 1px rgba(255,255,255,.7)!important;}
.pbc-due{background:linear-gradient(135deg,rgba(255,255,255,.26),rgba(255,255,255,.08))!important;
  border:1px solid rgba(255,255,255,.38)!important;backdrop-filter:blur(10px);
  box-shadow:inset 0 1px 0 rgba(255,255,255,.4),0 4px 12px rgba(0,0,0,.18);}
.pbc-ring{box-shadow:0 0 0 1px rgba(255,255,255,.25),0 4px 12px rgba(0,0,0,.2);}
.pbc-ring>div{background:linear-gradient(145deg,#27508f,#14305f)!important;box-shadow:inset 0 1px 0 rgba(255,255,255,.25);}

/* row layout: ref tile + body in ONE element (fewer DOM nodes = faster) */
.pbc-rowflex{display:flex;gap:12px;align-items:flex-start;}
.pbc-rowflex .pbc-ref{flex:0 0 auto;margin-top:2px;}
.pbc-body{flex:1;min-width:0;}
.pbc-pg{font-size:12px;color:#6b7a90;}
.pbc-pg b{color:#1e3c72;}
/* nested sub-section: second level sits indented below its parent band */
.pbc-sub-h.l1{margin-left:26px!important;margin-top:3px!important;font-size:11px;padding:4px 12px;
  background:linear-gradient(90deg,#e4eefd 0%,#f2f7fe 70%,#fafcff 100%);color:#2b4f8c;box-shadow:inset 0 0 0 1px #d3e1f8;}
.pbc-sub-h.l1:before{width:7px;height:7px;background:transparent;border:2px solid #5b87d1;border-radius:2px;}
.pbc-sub-h.l2{margin-left:52px!important;margin-top:3px!important;font-size:10.5px;padding:3px 10px;background:#f4f8fe;color:#3b5f9a;box-shadow:inset 0 0 0 1px #e0eafa;}

/* ================= KPI cards: original size (stacked label / number / note) ================= */
.pbc-kpis{gap:8px!important;margin-bottom:10px!important;}
.pbc-kpi{display:block!important;padding:9px 14px!important;border-radius:10px!important;}
.pbc-kpi .lb{font-size:10.5px!important;letter-spacing:.06em;line-height:1.3;}
.pbc-kpi .vl{font-size:26px!important;line-height:1.1!important;margin:0!important;}
.pbc-kpi .sb{font-size:11px!important;line-height:1.3;min-height:14px!important;white-space:normal;}

/* ================= soft tinted KPI cards (light background, coloured figures) ================= */
.pbc-kpi{
  background:linear-gradient(135deg,
      color-mix(in srgb,var(--c) 15%,#fff) 0%,
      color-mix(in srgb,var(--c) 8%,#fff) 100%)!important;
  box-shadow:inset 0 0 0 1px color-mix(in srgb,var(--c) 22%,#fff),0 2px 6px -3px color-mix(in srgb,var(--c) 40%,transparent)!important;}
.pbc-kpi:hover{box-shadow:inset 0 0 0 1px color-mix(in srgb,var(--c) 35%,#fff),0 6px 14px -6px color-mix(in srgb,var(--c) 45%,transparent)!important;}
.pbc-kpi:after{background:linear-gradient(160deg,rgba(255,255,255,.55) 0%,rgba(255,255,255,0) 55%)!important;}
.pbc-kpi .lb{color:color-mix(in srgb,var(--c) 80%,#1e293b)!important;opacity:.95!important;text-shadow:none!important;}
.pbc-kpi .vl{color:var(--c)!important;text-shadow:none!important;}
.pbc-kpi .sb{color:#64748b!important;opacity:1!important;text-shadow:none!important;}
</style>
"""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def esc(s):
    return _html.escape(str(s or ""), quote=True)


@lru_cache(maxsize=8192)
def _safe_rich_cached(html):
    return sanitize_html(html or "").replace("$", "&#36;")


def safe_rich(html):
    """Sanitised rich text that is also safe to drop inside st.markdown (no $ math, no blank lines)."""
    return _safe_rich_cached(html or "")


_txt = lru_cache(maxsize=8192)(html_to_text)  # cached plain-text conversion (search, delivery chips)


def fmt_size(n):
    n = float(n or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def fmt_date(iso):
    try:
        return date.fromisoformat(iso).strftime("%d %b %Y").lstrip("0")
    except (ValueError, TypeError):
        return ""


def slug(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def item_state(status, deadline_iso):
    if status == "Submitted":
        return "submitted"
    if status == "Not Applicable":
        return "na"
    try:
        if deadline_iso and date.fromisoformat(deadline_iso) < today():
            return "overdue"
    except ValueError:
        pass
    return {"In Progress": "progress", "Partially Submitted": "partial"}.get(status, "pending")


def days_left(deadline_iso):
    try:
        return (date.fromisoformat(deadline_iso) - today()).days
    except (ValueError, TypeError):
        return None


def deadline_chip(deadline_iso, state):
    if not deadline_iso:
        return '<span class="chip due">⏰ No deadline</span>'
    d = days_left(deadline_iso)
    label = fmt_date(deadline_iso)
    if state in ("submitted", "na"):
        return f'<span class="chip due done">✔ {label}</span>'
    if d is None:
        return f'<span class="chip due">📅 {label}</span>'
    if d < 0:
        return f'<span class="chip due late">⚠ Overdue {-d}d · {label}</span>'
    if d == 0:
        return f'<span class="chip due late">🔥 Due today · {label}</span>'
    if d <= 3:
        return f'<span class="chip due soon">⏳ {d} day{"s" if d != 1 else ""} left · {label}</span>'
    return f'<span class="chip due">📅 {label} · {d} days</span>'


def file_icon(name):
    ext = (name.rsplit(".", 1)[-1] if "." in name else "").lower()
    if ext == "pdf": return "📕"
    if ext in ("xls", "xlsx", "csv", "xlsm"): return "📗"
    if ext in ("doc", "docx", "rtf", "txt"): return "📘"
    if ext in ("ppt", "pptx"): return "📙"
    if ext in ("png", "jpg", "jpeg", "gif", "webp", "bmp", "tif", "tiff"): return "🖼️"
    if ext in ("zip", "rar", "7z"): return "🗜️"
    return "📄"


def dl_button(label, data, file_name, mime, key, **kw):
    """Download button that does not trigger an app rerun (keeps dialogs open) when supported."""
    try:
        return st.download_button(label, data=data, file_name=file_name, mime=mime, key=key,
                                  on_click="ignore", **kw)
    except TypeError:
        return st.download_button(label, data=data, file_name=file_name, mime=mime, key=key, **kw)


@st.cache_data(show_spinner=False)
def _logo_b64():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


def _status_changed(item_id, key, old, ref):
    new = st.session_state.get(key)
    if new and new != old:
        pdb.set_status(item_id, new)
        pdb.log(f"🔄 Status of {ref or 'item #' + str(item_id)} changed: {old} → {new}")
        st.session_state["_pbc_toast"] = f"{ref or 'Item'} → {new}"


# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------
def _select_or_custom(label, options, current, key, custom_label):
    NEW = "＋ Add new…"
    opts = list(options)
    if current and current not in opts:
        opts = [current] + opts
    opts = [""] + opts + [NEW]
    idx = opts.index(current) if current in opts else 0
    sel = st.selectbox(label, opts, index=idx, key=key, format_func=lambda x: x or "— none —")
    if sel == NEW:
        return st.text_input(custom_label, key=key + "_custom").strip()
    return sel


def _edit_body(item_id, sections, dirs, resps):
    item = pdb.get_item(item_id) if item_id else None
    item = item or {k: "" for k in pdb.ITEM_COLS}
    if not item_id:
        item["status"] = "Pending"
        item["deadline"] = ""
    ver = st.session_state.get("_pbc_edit_ver", 0)
    sfx = f"{item_id or 'new'}_{ver}"

    c1, c2, c3 = st.columns([1, 1.4, 1.4])
    ref = c1.text_input("Ref", value=item["ref"], key=f"e_ref_{sfx}", placeholder="e.g. 7.9")
    stat_opts = pdb.STATUSES + ([item["status"]] if item["status"] not in pdb.STATUSES and item["status"] else [])
    status = c2.selectbox("Status", stat_opts, index=stat_opts.index(item["status"]) if item["status"] in stat_opts else 0,
                          key=f"e_status_{sfx}")
    try:
        cur_dl = date.fromisoformat(item["deadline"]) if item["deadline"] else None
    except ValueError:
        cur_dl = None
    deadline = c3.date_input("Deadline (clear for none)", value=cur_dl, key=f"e_dl_{sfx}", format="YYYY-MM-DD")

    c4, c5 = st.columns([1, 1])
    with c4:
        section = _select_or_custom("Section", sections, item["section"], f"e_sec_{sfx}", "New section name (e.g. 15. Leases)")
    with c5:
        subsection = st.text_input("Sub-section (optional)", value=item["subsection"], key=f"e_sub_{sfx}")

    st.caption("Spelling wrong in a list? Fix it once under the ⚙️ Manage Lists tab instead of adding a new option.")
    c6, c7 = st.columns([1, 1])
    with c6:
        directorate = _select_or_custom("Directorate", dirs, item["directorate"], f"e_dir_{sfx}", "New directorate")
    with c7:
        responsible = _select_or_custom("Responsible person", resps, item["responsible"], f"e_resp_{sfx}", "New person")

    st.markdown("##### 📝 Request / schedule required")
    st.caption("Full rich-text editor — bold, italics, colours, lists, headings, links. Formatting from the Word file is kept.")
    desc = st_quill(value=item["description"], placeholder="Describe the document / schedule requested…",
                    html=True, toolbar=ADV_TOOLBAR, key=f"e_desc_{sfx}")
    st.markdown("##### 📨 Type of information / how to deliver")
    info = st_quill(value=item["info_type"], placeholder="e.g. Soft copy, Hard and soft copies…",
                    html=True, toolbar=COMPACT_TOOLBAR, key=f"e_info_{sfx}")
    st.markdown("##### 🗒️ Our notes / response to the auditors")
    notes = st_quill(value=item["notes"], placeholder="Internal notes, explanations, what was sent and when…",
                     html=True, toolbar=ADV_TOOLBAR, key=f"e_notes_{sfx}")

    if st.button("💾 Save changes" if item_id else "➕ Create request item", type="primary",
                 **STRETCH, key=f"e_save_{sfx}"):
        if not (ref or "").strip() and not html_to_text(desc or ""):
            st.error("Please enter a Ref or a description.")
            return
        payload = {
            "ref": (ref or "").strip(),
            "section": section or "General",
            "subsection": (subsection or "").strip(),
            "description": sanitize_html(desc if desc is not None else item["description"]),
            "info_type": sanitize_html(info if info is not None else item["info_type"]),
            "directorate": directorate,
            "responsible": responsible,
            "deadline": deadline.isoformat() if deadline else "",
            "status": status,
            "notes": sanitize_html(notes if notes is not None else item["notes"]),
        }
        new_id = pdb.save_item(payload, item_id)
        pdb.log(f"{'✏️ Updated' if item_id else '➕ Created'} request {payload['ref'] or '#' + str(new_id)}")
        st.session_state["_pbc_edit_ver"] = ver + 1
        st.session_state["_pbc_toast"] = "Saved ✔" if item_id else "New request created ✔"
        st.rerun()

    if item_id:
        with st.expander("🗑️ Delete this request item"):
            n_docs = len(pdb.list_documents(item_id))
            st.warning(f"This permanently removes the item" + (f" and its {n_docs} attached document(s)." if n_docs else "."))
            ok = st.checkbox("Yes, delete it", key=f"e_delok_{sfx}")
            if st.button("Delete permanently", disabled=not ok, key=f"e_del_{sfx}"):
                pdb.delete_item(item_id)
                pdb.log(f"🗑️ Deleted request {item['ref'] or '#' + str(item_id)}")
                st.session_state["_pbc_toast"] = "Item deleted"
                st.rerun()


def edit_item_dialog(item_id, sections, dirs, resps):
    item = pdb.get_item(item_id) if item_id else None
    title = f"✏️ Edit request {item['ref']}" if item and item["ref"] else ("✏️ Edit request" if item_id else "➕ New request item")
    st.dialog(title, width="large")(_edit_body)(item_id, sections, dirs, resps)


def _docs_body(item_id):
    item = pdb.get_item(item_id)
    if not item:
        st.error("Item not found.")
        return
    st.markdown(
        f"<div class='pbc-wrap'><span class='pbc-ref s-{item_state(item['status'], item['deadline'])}'>{esc(item['ref'] or '—')}</span></div>"
        f"<div class='pbc-desc' style='margin:8px 0 4px'>{safe_rich(item['description'])}</div>",
        unsafe_allow_html=True)
    st.divider()

    docs = pdb.list_documents(item_id)
    if docs:
        st.markdown(f"**{len(docs)} document{'s' if len(docs) != 1 else ''} attached**")
        for d in docs:
            with st.container(border=True):
                a, b, c = st.columns([7, 1.5, 1])
                a.markdown(
                    f"{file_icon(d['filename'])} **{d['filename']}**  \n"
                    f"<span style='color:#64748b;font-size:12.5px'>{fmt_size(d['size'])} · added {d['uploaded_at']}"
                    f"{' · ' + esc(d['note']) if d['note'] else ''}</span>", unsafe_allow_html=True)
                with b:
                    dl_button("⬇ Download", pdb.get_document_data(d["id"]), d["filename"],
                              d["mime"] or "application/octet-stream", key=f"dl_{d['id']}",
                              **STRETCH)
                with c:
                    if st.button("🗑️", key=f"deldoc_{d['id']}", help="Remove this document", **STRETCH):
                        pdb.delete_document(d["id"])
                        pdb.log(f"🗑️ Removed document '{d['filename']}' from {item['ref'] or '#' + str(item_id)}")
                        st.session_state["_pbc_reopen_docs"] = item_id
                        st.session_state["_pbc_toast"] = "Document removed"
                        st.rerun()
    else:
        st.info("No documents attached yet. Add the supporting files for this request below.")

    st.markdown("##### ➕ Attach documents")
    up_ver = st.session_state.get("_pbc_up_ver", 0)
    files = st.file_uploader(
        f"Any file type · up to {pdb.MAX_FILE_MB} MB each · you can select several at once",
        accept_multiple_files=True, key=f"pbc_up_{item_id}_{up_ver}")
    note = st.text_input("Optional note for these files (e.g. 'Final signed version')", key=f"pbc_note_{item_id}_{up_ver}")
    n_sel = len(files) if files else 0
    label = f"📎 Attach {n_sel} file{'s' if n_sel != 1 else ''}" if n_sel else "📎 Attach files"
    if st.button(label, type="primary", disabled=not files, **STRETCH, key=f"pbc_attach_{item_id}_{up_ver}"):
        ok_n, msgs = 0, []
        for f in files:
            ok, msg = pdb.add_document(item_id, f.name, f.getvalue(), f.type or "", note.strip())
            ok_n += ok
            if not ok:
                msgs.append(msg)
        if ok_n:
            pdb.log(f"📎 Attached {ok_n} document(s) to {item['ref'] or '#' + str(item_id)}")
        st.session_state["_pbc_up_ver"] = up_ver + 1
        st.session_state["_pbc_reopen_docs"] = item_id
        st.session_state["_pbc_toast"] = f"{ok_n} document(s) attached" + (" · " + " ".join(msgs) if msgs else "")
        st.rerun()


def docs_dialog(item_id):
    item = pdb.get_item(item_id)
    title = f"📎 Documents — {item['ref']}" if item and item["ref"] else "📎 Documents"
    st.dialog(title, width="large")(_docs_body)(item_id)


def _status_body(item_id):
    item = pdb.get_item(item_id)
    if not item:
        st.error("Item not found.")
        return
    cur = item["status"] or "Pending"
    state = item_state(cur, item["deadline"])
    st.markdown(
        f"<div class='pbc-wrap'><span class='pbc-ref s-{state}'>{esc(item['ref'] or '—')}</span></div>"
        f"<div class='pbc-desc' style='margin:6px 0 2px'>{safe_rich(item['description'])}</div>",
        unsafe_allow_html=True)
    st.divider()
    st.markdown(f"Current status: <span class='chip st-{state}' style='font-size:12.5px'>{esc(cur)}</span>",
                unsafe_allow_html=True)
    opts = pdb.STATUSES + ([cur] if cur not in pdb.STATUSES else [])
    ver = st.session_state.get("_pbc_st_ver", 0)
    new = st.radio("Change status to", opts, index=opts.index(cur), key=f"pbc_stnew_{item_id}_{ver}")
    changed = new != cur
    if changed:
        st.warning(f"You are about to change this request from **{cur}** to **{new}**.")
    if st.button("✔ Confirm status change", type="primary", disabled=not changed, **STRETCH,
                 key=f"pbc_stok_{item_id}_{ver}"):
        pdb.set_status(item_id, new)
        pdb.log(f"🔄 Status of {item['ref'] or '#' + str(item_id)} changed: {cur} → {new}")
        st.session_state["_pbc_st_ver"] = ver + 1
        st.session_state["_pbc_toast"] = f"{item['ref'] or 'Item'} → {new}"
        st.rerun()


def status_dialog(item_id):
    item = pdb.get_item(item_id)
    title = f"🔄 Change status — {item['ref']}" if item and item["ref"] else "🔄 Change status"
    st.dialog(title, width="small")(_status_body)(item_id)


AVATAR_COLORS = ["#1e3c72", "#0f766e", "#7c3aed", "#b45309", "#be123c", "#0369a1"]


def _initials(name):
    parts = re.findall(r"[A-Za-z0-9]+", name or "")
    return ("".join(w[0] for w in parts[:2]) or "?").upper()


def _plain_html(text):
    return esc(text).replace("\n", "<br>").replace("$", "&#36;")


def _comments_body(item_id):
    item = pdb.get_item(item_id)
    if not item:
        st.error("Item not found.")
        return
    st.markdown(
        f"<div class='pbc-wrap'><span class='pbc-ref s-{item_state(item['status'], item['deadline'])}'>{esc(item['ref'] or '—')}</span></div>"
        f"<div class='pbc-desc' style='margin:6px 0 2px'>{safe_rich(item['description'])}</div>",
        unsafe_allow_html=True)
    st.divider()

    comments = pdb.list_comments(item_id)
    if comments:
        st.markdown(f"**{len(comments)} comment{'s' if len(comments) != 1 else ''}**")
        for c in comments:
            color = AVATAR_COLORS[sum(map(ord, c["author"] or "?")) % len(AVATAR_COLORS)]
            a, b = st.columns([14, 1], vertical_alignment="top")
            a.markdown(
                f"<div class='pbc-wrap'><div class='pbc-cm'><div class='av' style='background:{color}'>{esc(_initials(c['author']))}</div>"
                f"<div class='bd'><div class='hd'><b>{esc(c['author'])}</b><span>{esc(c['created_at'][:16])}</span></div>"
                f"<div class='tx'>{_plain_html(c['body'])}</div></div></div></div>", unsafe_allow_html=True)
            if b.button("🗑️", key=f"delcm_{c['id']}", help="Delete this comment"):
                pdb.delete_comment(c["id"])
                pdb.log(f"🗑️ Deleted a comment by {c['author']} on {item['ref'] or '#' + str(item_id)}")
                st.session_state["_pbc_reopen_comments"] = item_id
                st.rerun()
    else:
        st.info("No comments yet. Start the discussion below.")

    ver = st.session_state.get("_pbc_cm_ver", 0)
    author = st.text_input("Your name", value=st.session_state.get("_pbc_author", ""),
                           placeholder="e.g. Assoumpta", key=f"pbc_cm_author_{ver}")
    body = st.text_area("Comment", placeholder="Write an update, question or note for the team…",
                        height=100, key=f"pbc_cm_body_{item_id}_{ver}")
    if st.button("💬 Post comment", type="primary", **STRETCH, disabled=not (body or "").strip(),
                 key=f"pbc_cm_post_{item_id}_{ver}"):
        if pdb.add_comment(item_id, author, body):
            st.session_state["_pbc_author"] = (author or "").strip()
            pdb.log(f"💬 {(author or '').strip() or 'Anonymous'} commented on {item['ref'] or '#' + str(item_id)}")
        st.session_state["_pbc_cm_ver"] = ver + 1
        st.session_state["_pbc_reopen_comments"] = item_id
        st.session_state["_pbc_toast"] = "Comment posted"
        st.rerun()


def comments_dialog(item_id):
    item = pdb.get_item(item_id)
    title = f"💬 Comments — {item['ref']}" if item and item["ref"] else "💬 Comments"
    st.dialog(title, width="large")(_comments_body)(item_id)


# ---------------------------------------------------------------------------
# Page sections
# ---------------------------------------------------------------------------
def _hero(df, meta):
    total = len(df)
    na = int((df["status"] == "Not Applicable").sum())
    sub = int((df["status"] == "Submitted").sum())
    pct = round(sub / max(total - na, 1) * 100) if total else 0

    due = meta.get("due_date") or ""
    d = days_left(due)
    if d is None:
        due_html = ""
    else:
        cls = "late" if d < 0 else ("soon" if d <= 3 else "")
        num = f"{abs(d)}" if d != 0 else "Today"
        lab = ("days overdue" if d < 0 else ("days left" if d != 1 else "day left")) if d != 0 else "documents due"
        due_html = f"<div class='pbc-due {cls}'><b>{num}</b><span>{lab}</span><div style='color:#fff;opacity:.8'>{fmt_date(due)}</div></div>"

    logo = _logo_b64()
    logo_html = f"<div class='pbc-logo'><img src='data:image/jpeg;base64,{logo}'/></div>" if logo else "<div class='pbc-logo'>🏥</div>"
    issued = esc(meta.get("date", ""))
    html = (
        "<div class='pbc-wrap'><div class='pbc-hero'>" + logo_html +
        "<div class='pbc-hero-text'><div class='pbc-eyebrow'>Rwanda Medical Supply Ltd · Office of the Auditor General</div>"
        "<div class='pbc-h1'>OAG Preliminary Information Request</div>"
        f"<div class='pbc-sub'>Audit for the year ended 30 June 2026{' · issued ' + issued if issued else ''} · {total} requests · {df['section'].nunique()} sections</div></div>"
        "<div class='pbc-hero-right'>"
        f"<div class='pbc-ring' style='--p:{pct}'><div><b>{pct}%</b><span>submitted</span></div></div>"
        + due_html + "</div></div></div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def _kpis(view):
    total = len(view)
    cnt = view["status"].value_counts()
    sub, na = int(cnt.get("Submitted", 0)), int(cnt.get("Not Applicable", 0))
    prog = int(cnt.get("In Progress", 0) + cnt.get("Partially Submitted", 0))
    pend = int(cnt.get("Pending", 0))
    overdue = sum(1 for _, r in view.iterrows() if item_state(r["status"], r["deadline"]) == "overdue")
    docs, size = int(view["doc_count"].sum()), int(view["doc_bytes"].sum())
    base = max(total - na, 1)

    def card(lb, vl, sb, color, icon):
        return (f"<div class='pbc-kpi' style='--c:{color}'><span class='ic'>{icon}</span>"
                f"<div class='lb'>{lb}</div><div class='vl'>{vl}</div><div class='sb'>{sb}</div></div>")

    html = ("<div class='pbc-wrap'><div class='pbc-kpis'>"
            + card("Total requests", total, f"{na} not applicable" if na else "&nbsp;", "#1e3c72", "📋")
            + card("Submitted", sub, f"{sub / base * 100:.0f}% of applicable", "#15803d", "✅")
            + card("In progress", prog, f"{prog / base * 100:.0f}% of applicable", "#ca8a04", "🔄")
            + card("Pending", pend, f"{pend / base * 100:.0f}% of applicable", "#475569", "⏳")
            + card("Overdue", overdue, "past deadline", "#b91c1c", "🚨")
            + card("Documents", docs, fmt_size(size) + " stored", "#6d28d9", "📎")
            + "</div></div>")
    st.markdown(html, unsafe_allow_html=True)


def _section_header(sec, df_all):
    g = df_all[df_all["section"] == sec]
    total, na = len(g), int((g["status"] == "Not Applicable").sum())
    sub = int((g["status"] == "Submitted").sum())
    pct = sub / max(total - na, 1) * 100
    m = re.match(r"^\s*([\w.]+?)\.?\s+(.*)$", sec)
    num, title = (m.group(1), m.group(2)) if m and len(m.group(1)) <= 4 else ("•", sec)
    st.markdown(
        f"<div class='pbc-wrap'><div class='pbc-sec' style='--p:{pct:.0f}%'><span class='num'>{esc(num.zfill(2) if num.isdigit() else num)}</span>"
        f"<span class='tt'>{esc(title)}</span>"
        f"<span class='mt'><b>{sub}</b> / {total - na} submitted</span></div></div>",
        unsafe_allow_html=True)


@_fragment
def _render_card(r, docs_by_item, comments_by_item, sections, dirs, resps):
    item_id = int(r["id"])
    state = item_state(r["status"], r["deadline"])
    with st.container(key=f"pbc-card-{state}-{item_id}"):
        c_main, c_right = st.columns([5.7, 4.3], vertical_alignment="top")
        c_stat, b1, b2, b3 = c_right.columns([1.65, 1.1, 1.45, 0.95], vertical_alignment="top")

        chips = []
        if r["directorate"]:
            chips.append(f"<span class='chip dir'>🏢 {esc(r['directorate'])}</span>")
        if r["responsible"]:
            chips.append(f"<span class='chip resp'>👤 {esc(r['responsible'])}</span>")
        chips.append(deadline_chip(r["deadline"], state))
        info_plain = _txt(r["info_type"])
        short_info = bool(info_plain) and len(info_plain) <= 45 and "\n" not in info_plain
        if short_info:
            chips.append(f"<span class='chip type'>📄 {esc(info_plain)}</span>")
        files = docs_by_item.get(item_id, [])
        for name, size in files[:3]:
            chips.append(f"<span class='chip file' title='{esc(name)} · {fmt_size(size)}'>{file_icon(name)} <span>{esc(name)}</span></span>")
        if len(files) > 3:
            chips.append(f"<span class='chip file'>+{len(files) - 3} more</span>")

        body = (f"<div class='pbc-wrap pbc-rowflex'><span class='pbc-ref s-{state}'>{esc(r['ref']) or '—'}</span><div class='pbc-body'>"
                f"<div class='pbc-desc'>{safe_rich(r['description']) or '<i>No description</i>'}</div>")
        body += "<div class='pbc-meta'>" + "".join(chips) + "</div>"
        if info_plain and not short_info:
            body += f"<div class='pbc-delivery'><span>📨</span><div>{safe_rich(r['info_type'])}</div></div>"
        if r["notes"]:
            body += f"<div class='pbc-notes'><b class='h'>Our notes</b>{safe_rich(r['notes'])}</div>"
        body += "</div></div>"
        c_main.markdown(body, unsafe_allow_html=True)

        n = int(r["doc_count"])
        nc = len(comments_by_item.get(item_id, []))
        if c_stat.button(f"{r['status'] or 'Pending'}  ▾", key=f"pbc-stbtn-{state}-{item_id}", **STRETCH):
            status_dialog(item_id)
        if b1.button(f"📎 Docs · {n}" if n else "📎 Docs", key=f"{'pbc_docsfull_' if n else 'pbc_docs0_'}{item_id}", **STRETCH):
            docs_dialog(item_id)
        if b2.button(f"💬 Comment · {nc}" if nc else "💬 Comment", key=f"{'pbc_cmtfull_' if nc else 'pbc_cmt0_'}{item_id}", **STRETCH):
            comments_dialog(item_id)
        if b3.button("✏️ Edit", key=f"pbc_edit_{item_id}", **STRETCH):
            edit_item_dialog(item_id, sections, dirs, resps)

        cms = comments_by_item.get(item_id, [])
        if cms:
            la, lb, lts = cms[-1]
            more = f"<i>+{len(cms) - 1} earlier</i>" if len(cms) > 1 else ""
            c_right.markdown(
                f"<div class='pbc-wrap'><div class='pbc-cmt'><div class='lb'>💬 Comment{more}</div>"
                f"<div class='who'><b>{esc(la)}</b><span>{esc(lts[:16])}</span></div>"
                f"<div class='tx'>{_plain_html(lb)}</div></div></div>", unsafe_allow_html=True)


def _tracker_tab(df, docs_by_item, comments_by_item):
    sections = list(dict.fromkeys(df["section"].tolist()))
    dirs = sorted({d for d in df["directorate"] if d})
    resps = sorted({d for d in df["responsible"] if d})

    # ----- filters -----
    f1, f2, f3, f4, f5 = st.columns([2.4, 2, 1.5, 1.5, 1.8])
    q = f1.text_input("Search", placeholder="🔍 Search ref, wording, person, file name…", key="pbc_f_search", label_visibility="collapsed")
    sec_sel = f2.selectbox("Section", ["All sections"] + sections, key="pbc_f_section", label_visibility="collapsed")
    dir_sel = f3.selectbox("Directorate", ["All directorates"] + dirs, key="pbc_f_dir", label_visibility="collapsed")
    resp_sel = f4.selectbox("Responsible", ["All people"] + resps, key="pbc_f_resp", label_visibility="collapsed")
    stat_sel = f5.multiselect("Status", pdb.STATUSES, key="pbc_f_status", placeholder="Any status", label_visibility="collapsed")

    t1, t2, t3, _sp, t5 = st.columns([1.6, 1.6, 1.6, 2.4, 1.6])
    attn = t1.toggle("⚠️ Needs attention", key="pbc_f_attn", help="Overdue or due within 3 days, and not yet submitted")
    nodocs = t2.toggle("📎 No documents yet", key="pbc_f_nodocs")
    hide_sub = t3.toggle("🙈 Hide submitted", key="pbc_f_hide")
    if t5.button("➕ New request", type="primary", **STRETCH, key="pbc_new"):
        edit_item_dialog(None, sections, dirs, resps)

    view = df
    if q.strip():
        needle = q.strip().lower()
        def hit(r):
            files = " ".join(n for n, _ in docs_by_item.get(int(r["id"]), []))
            cm = " ".join(b for _, b, _ in comments_by_item.get(int(r["id"]), []))
            blob = " ".join([r["ref"], _txt(r["description"]), _txt(r["notes"]), r["directorate"],
                             r["responsible"], r["section"], r["subsection"], files, cm]).lower()
            return needle in blob
        view = view[view.apply(hit, axis=1)]
    if sec_sel != "All sections":
        view = view[view["section"] == sec_sel]
    if dir_sel != "All directorates":
        view = view[view["directorate"] == dir_sel]
    if resp_sel != "All people":
        view = view[view["responsible"] == resp_sel]
    if stat_sel:
        view = view[view["status"].isin(stat_sel)]
    if attn:
        def needs(r):
            st_ = item_state(r["status"], r["deadline"])
            d = days_left(r["deadline"])
            return st_ == "overdue" or (st_ not in ("submitted", "na") and d is not None and d <= 3)
        view = view[view.apply(needs, axis=1)] if len(view) else view
    if nodocs:
        view = view[(view["doc_count"] == 0) & (~view["status"].isin(pdb.DONE_STATUSES))]
    if hide_sub:
        view = view[view["status"] != "Submitted"]

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    _kpis(view)

    if view.empty:
        st.markdown("<div class='pbc-wrap'><div class='pbc-empty'><div class='big'>🔎</div>"
                    "<b>No requests match these filters.</b><br>Try clearing the search or switching a toggle off.</div></div>",
                    unsafe_allow_html=True)
        return sections, dirs, resps

    page_view = view  # everything on one page


    last_sec, last_parts = None, []
    for _, r in page_view.iterrows():
        if r["section"] != last_sec:
            last_sec, last_parts = r["section"], []
            _section_header(r["section"], df)
        sub = r["subsection"]
        parts = [x.strip() for x in sub.split("›")] if sub else []
        common = 0
        while common < len(last_parts) and common < len(parts) and last_parts[common] == parts[common]:
            common += 1
        for lvl in range(common, len(parts)):          # a band for every NEW level (4.2, then 4.2.1 below it)
            st.markdown(f"<div class='pbc-wrap'><div class='pbc-sub-h l{min(lvl, 2)}'>{esc(parts[lvl])}</div></div>",
                        unsafe_allow_html=True)
        last_parts = parts
        _render_card(r, docs_by_item, comments_by_item, sections, dirs, resps)
    return sections, dirs, resps


def _import_tab():
    st.subheader("📥 Import another request list")
    t = pdb.totals()
    st.markdown(
        f"Your tracker currently holds **{t['items']} requests**, **{t['documents']} documents** and **{t['comments']} comments**.  \n"
        "Importing is **add-only**: it never deletes, replaces or edits anything that is already here — "
        "statuses, notes, comments and attached documents all stay exactly as they are. New requests are added "
        "to the end of their section (new sections go at the bottom). Bold, italics, lists and links are kept.")
    up = st.file_uploader("OAG request document (.docx)", type=["docx"], key="pbc_import_up")
    if up is None:
        return
    try:
        parsed = parse_request_docx(up)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read that document: {e}")
        return
    items = parsed["items"]
    keep_dups = st.checkbox(
        "Also add requests that look identical to ones already here (as separate copies)",
        key="pbc_import_dups",
        help="Off by default: a request with the same reference and wording as an existing one is skipped. "
             "Turn this on for a new year's list that reuses the same wording.")
    new_n = pdb.import_items(items, dry_run=True, keep_duplicates=keep_dups)
    c1, c2, c3 = st.columns(3)
    c1.metric("Requests found in file", len(items))
    c2.metric("Will be added", new_n)
    c3.metric("Skipped (already here)", len(items) - new_n)
    if not items:
        st.warning("No requests were found in that file. Check that it uses the same table layout as the OAG list.")
    with st.expander("Preview what was read from the file"):
        import pandas as pd
        st.dataframe(pd.DataFrame([{
            "Ref": i["ref"], "Section": i["section"], "Request": html_to_text(i["description"])[:140],
            "Directorate": i["directorate"], "Responsible": i["responsible"], "Deadline": i["deadline"],
        } for i in items]), hide_index=True, **STRETCH)
    if st.button(f"Add {new_n} request(s) — keep everything else", type="primary", disabled=new_n == 0, key="pbc_do_import"):
        before = pdb.totals()
        added = pdb.import_items(items, keep_duplicates=keep_dups)
        if parsed["meta"] and not pdb.get_request_meta():
            import json
            pdb.set_meta("request_meta", json.dumps(parsed["meta"]))
        pdb.log(f"📥 Imported {added} new request(s) from '{up.name}' (kept all {before['items']} existing; "
                f"{len(items) - added} skipped as already present)")
        st.session_state["_pbc_toast"] = f"Added {added} new request(s); all {before['items']} existing kept"
        st.rerun()


def _lists_tab():
    st.subheader("⚙️ Manage dropdown lists")
    st.markdown(
        "Fix a wrong spelling **once** and it is corrected on every request that uses it — no need to create a "
        "second option. If you rename something to a name that already exists, the two are **merged**.")
    labels = list(pdb.LIST_FIELDS.values())
    choice = st.radio("List to edit", labels, horizontal=True, key="pbc_ml_list", label_visibility="collapsed")
    field = [k for k, v in pdb.LIST_FIELDS.items() if v == choice][0]
    ver = st.session_state.get("_pbc_ml_ver", 0)
    values = pdb.list_values(field)
    existing = {v for v, _ in values}

    if not values:
        st.info("Nothing in this list yet.")
        return
    h1, h2, h3, h4 = st.columns([3, 0.9, 4, 1.6])
    h1.markdown("**Current name**")
    h2.markdown("**Requests**")
    h3.markdown("**Correct name**")
    for i, (val, cnt) in enumerate(values):
        c1, c2, c3, c4 = st.columns([3, 0.9, 4, 1.6], vertical_alignment="center")
        c1.markdown(f"<div class='pbc-wrap' style='font-weight:600;color:#1f2937'>{esc(val)}</div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='pbc-wrap'><span class='chip type'>{cnt}</span></div>", unsafe_allow_html=True)
        new = c3.text_input("New name", value=val, key=f"ml_in_{field}_{i}_{ver}", label_visibility="collapsed")
        new_clean = " ".join((new or "").split())
        merging = bool(new_clean) and new_clean != val and new_clean in existing
        label = "Merge" if merging else "Save"
        if c4.button(label, key=f"ml_btn_{field}_{i}_{ver}", disabled=(not new_clean or new_clean == val), **STRETCH,
                     help="Merges into the existing name" if merging else "Rename everywhere"):
            n = pdb.rename_value(field, val, new_clean)
            st.session_state["_pbc_ml_ver"] = ver + 1
            st.session_state["_pbc_toast"] = f"{'Merged' if merging else 'Renamed'} '{val}' → '{new_clean}' ({n} request{'s' if n != 1 else ''})"
            st.rerun()
        if merging:
            c3.caption(f"⚠ '{new_clean}' already exists — these requests will be merged into it.")


def _export_tab(df, meta, comments_by_item=None):
    st.subheader("💾 Export")
    st.markdown("Formatting (bold, italics, lists) is carried into the Excel and Word exports.")
    stamp = datetime.now(KIGALI).strftime("%Y%m%d_%H%M")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**📗 Excel workbook**  \nTracker + summary by section")
        if st.button("Prepare Excel", **STRETCH, key="pbc_prep_xlsx"):
            st.session_state["_x_xlsx"] = (pex.build_excel(df, comments_by_item), f"RMS_OAG_Information_Request_{stamp}.xlsx", datetime.now(KIGALI).strftime("%H:%M"))
        if "_x_xlsx" in st.session_state:
            data, name, ts = st.session_state["_x_xlsx"]
            st.download_button(f"⬇ Download Excel (prepared {ts})", data, name,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", **STRETCH, key="pbc_dl_xlsx")
    with c2:
        st.markdown("**📘 Word report**  \nLandscape table, section headings")
        if st.button("Prepare Word", **STRETCH, key="pbc_prep_docx"):
            st.session_state["_x_docx"] = (pex.build_word(df, meta), f"RMS_OAG_Information_Request_{stamp}.docx", datetime.now(KIGALI).strftime("%H:%M"))
        if "_x_docx" in st.session_state:
            data, name, ts = st.session_state["_x_docx"]
            st.download_button(f"⬇ Download Word (prepared {ts})", data, name,
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               type="primary", **STRETCH, key="pbc_dl_docx")
    with c3:
        st.markdown("**🗜️ All documents (ZIP)**  \nOne folder per request item")
        n_docs = int(df["doc_count"].sum())
        if st.button(f"Prepare ZIP ({n_docs} file{'s' if n_docs != 1 else ''})", **STRETCH,
                     disabled=n_docs == 0, key="pbc_prep_zip"):
            with st.spinner("Packing documents…"):
                data, cnt = pex.build_documents_zip(df)
            st.session_state["_x_zip"] = (data, f"RMS_OAG_Request_Documents_{stamp}.zip", cnt)
        if "_x_zip" in st.session_state:
            data, name, cnt = st.session_state["_x_zip"]
            st.download_button(f"⬇ Download ZIP ({cnt} files, {fmt_size(len(data))})", data, name, "application/zip",
                               type="primary", **STRETCH, key="pbc_dl_zip")


def _log_tab():
    st.subheader("🕘 Activity log")
    st.markdown("Every change made in this tracker.")
    logs = pdb.get_logs()
    if logs.empty:
        st.info("No activity yet.")
        return
    st.dataframe(logs, hide_index=True, **STRETCH, column_config={
        "timestamp": st.column_config.TextColumn("Time", width="small"),
        "message": st.column_config.TextColumn("Activity", width="large")})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def render_page():
    pdb.init()
    pdb.seed_if_needed(SEED_PATH)
    st.markdown(PAGE_CSS, unsafe_allow_html=True)

    toast = st.session_state.pop("_pbc_toast", None)
    if toast:
        st.toast(toast)

    df = pdb.load_items()
    meta = pdb.get_request_meta()
    docs_by_item = pdb.docs_summary()
    comments_by_item = pdb.comments_all()

    if df.empty:
        st.markdown("<div class='pbc-wrap'><div class='pbc-empty'><div class='big'>📭</div><b>No request items yet.</b><br>"
                    "Open the <b>Import</b> tab to load the auditors' Word request list, or add items one by one.</div></div>",
                    unsafe_allow_html=True)
        sections, dirs, resps = [], [], []
    else:
        _hero(df, meta)

    tab_track, tab_import, tab_lists, tab_export, tab_log = st.tabs(
        ["📋 Request Tracker", "📥 Import Request List", "⚙️ Manage Lists", "💾 Export", "🕘 Activity Log"])

    with tab_track:
        if df.empty:
            if st.button("➕ Add first request", type="primary", key="pbc_new_empty"):
                edit_item_dialog(None, [], [], [])
        else:
            sections, dirs, resps = _tracker_tab(df, docs_by_item, comments_by_item)
        rid = st.session_state.pop("_pbc_reopen_docs", None)
        if rid:
            docs_dialog(rid)
        cid = st.session_state.pop("_pbc_reopen_comments", None)
        if cid:
            comments_dialog(cid)
    with tab_import:
        _import_tab()
    with tab_lists:
        _lists_tab()
    with tab_export:
        if df.empty:
            st.info("Nothing to export yet.")
        else:
            _export_tab(df, meta, comments_by_item)
    with tab_log:
        _log_tab()

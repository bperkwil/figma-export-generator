"""
Figma Export Generator
=======================

Upload an AFI Creative Brief (.xlsx) and get a Figma-ready Excel export with
Onsite, Outbound, and Processing Log sheets -- plus a live dashboard showing
stats, detected sheets, market distribution, auto-corrections, duplicate and
ILM-balance checks, and a preview of the generated rows.

The processor uses fuzzy header matching, so briefs with slightly different
column names / layouts still work. Business rules (asset maps, ILM correction,
copy priority) live in clearly-labelled config blocks near the top.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io
import re
import warnings
from collections import Counter
from difflib import SequenceMatcher

import openpyxl
import streamlit as st

warnings.simplefilter("ignore")  # silence openpyxl "unknown extension" noise

# ============================================================
# PAGE CONFIG & THEME
# ============================================================
st.set_page_config(page_title="Figma Export Generator", page_icon="🎨", layout="wide")

CSS = """
<style>
.stApp { background: #0b1020; }
.block-container { padding-top: 1.2rem; max-width: 1500px; }

/* Header band */
.fx-header {
    background: linear-gradient(135deg, #141a33 0%, #1b1130 100%);
    border: 1px solid #2a2f52; border-radius: 14px;
    padding: 18px 24px; margin-bottom: 18px;
}
.fx-header h1 { margin: 0; font-size: 1.5rem; color: #e8e8f0; }
.fx-header .accent { color: #8b7bf0; }
.fx-header .sub { color: #8a8fb0; font-size: 0.85rem; margin-top: 4px; }

/* Stat cards */
.fx-card {
    background: #10152b; border: 1px solid #262c4d; border-radius: 12px;
    padding: 16px 10px; text-align: center; height: 100%;
}
.fx-card .num { font-size: 2.0rem; font-weight: 700; line-height: 1.1; }
.fx-card .lbl { color: #8a8fb0; font-size: 0.72rem; letter-spacing: .08em;
                text-transform: uppercase; margin-top: 4px; }

/* Section titles */
.fx-section { color: #e8e8f0; font-weight: 600; font-size: 1.0rem;
              margin: 18px 0 8px 0; }

/* Pills */
.fx-pill {
    display: inline-block; background: #161b33; border: 1px solid #2d3358;
    border-radius: 20px; padding: 5px 14px; margin: 3px 6px 3px 0;
    color: #cfd3ec; font-size: 0.82rem;
}
.fx-pill .cnt { color: #8b7bf0; font-weight: 700; margin-left: 6px; }

/* Market distribution bar */
.fx-bar { display: flex; width: 100%; border-radius: 8px; overflow: hidden;
          height: 30px; border: 1px solid #262c4d; }
.fx-seg { display: flex; align-items: center; justify-content: center;
          color: #0b1020; font-size: 0.78rem; font-weight: 700; }

/* Alert / info panels */
.fx-panel { border-radius: 12px; padding: 14px 18px; margin: 8px 0;
            font-family: 'SFMono-Regular', Consolas, monospace; font-size: 0.82rem;
            white-space: pre-wrap; line-height: 1.6; }
.fx-warn { background: #241a0e; border: 1px solid #5a4420; color: #f0c674; }
.fx-ok   { background: #0e2418; border: 1px solid #1f5a3a; color: #74d699; }
.fx-title { font-weight: 700; margin-bottom: 6px; display: block; }
.fx-badge { background: #8b7bf0; color: #0b1020; border-radius: 10px;
            padding: 1px 8px; font-size: 0.7rem; font-weight: 700; margin-left: 8px; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# ============================================================
# CONFIGURATION -- ASSET MAPS & RULES  (edit these)
# ============================================================

# Onsite asset map: "<Asset type>|<Device>" -> {code, role}
# role TITLE = emit a row; SUBTITLE / SKIP = do not emit a standalone row.
ONSITE_ASSET_MAP = {
    "Tall Hero|Desktop": {"code": "hp1", "role": "TITLE"},
    "Window Display Header|Mobile": {"code": "hp2", "role": "TITLE"},
    "Window Display Sub-header|Mobile": {"code": "hp2", "role": "SUBTITLE"},
    "Quad Card|Mobile": {"code": "hp10", "role": "TITLE"},
    "Quad Card|Desktop": {"code": "hp8", "role": "TITLE"},
    "ILM|Mobile": {"code": "dp1", "role": "TITLE"},
    "ILM|Desktop": {"code": "dp2", "role": "TITLE"},
    "ILM|Tablet": {"code": "dp3", "role": "TITLE"},
    "Mobile Footer|Mobile": {"code": "hp11", "role": "TITLE"},
    "Mobile Billboard|Mobile": {"code": "hp12", "role": "TITLE"},
    "Shoveler|Desktop": {"code": "hp7", "role": "TITLE"},
    "Single Category Card|Mobile": {"code": "hp9", "role": "TITLE"},
    "Single Category Card|Desktop": {"code": "hp7", "role": "TITLE"},
    "Sparkle|Mobile": {"code": "", "role": "SKIP"},
    "Sparkle|Desktop": {"code": "", "role": "SKIP"},
    "TSTO|Mobile": {"code": "", "role": "SKIP"},
    "TSTO|Desktop": {"code": "", "role": "SKIP"},
    "Ad Topper|Mobile": {"code": "", "role": "SKIP"},
    "Ad Topper|Desktop": {"code": "", "role": "SKIP"},
    "Dashboard Card|Mobile": {"code": "", "role": "SKIP"},
}
# Fallback: if an asset type is unknown, emit a row anyway using this code
# prefix so nothing is silently dropped (shows up as "??" in the preview).
ONSITE_FALLBACK_CODE = "hp?"

BG_COLOURS = {"GP": "/#207936", "GFC": "/#10C800"}

# Outbound placement -> asset code
OUTBOUND_MAP = {
    "Subject": "email1", "Super Subject": "email1", "Hero": "email1",
    "Card image": "email2", "Category Cards": "email3",
    "Banner Image": "email4", "Category Row": "email5",
    "Notification title": "push1", "Notification text": "push1", "CTA": "push1",
}

# Fuzzy header synonyms: canonical field -> list of header variants we accept.
HEADER_SYNONYMS = {
    "country": ["country", "marketplace", "market"],
    "business": ["business", "business type"],
    "merchant": ["merchant", "merchant name", "store"],
    "asset_type": ["asset type", "asset", "placement"],
    "placement_type": ["placement type"],
    "device": ["device type", "device"],
    "targeting": ["targeting", "target", "audience"],
    "email_num": ["email", "email #", "email number"],
    "master_copy": ["master copy", "master copy en", "english copy"],
    "translated_copy": ["translated copy", "translated copy ll", "local language copy",
                         "ll copy", "local language"],
}

# Copy priority: prefer translated (local language) over master (English).
# First non-empty wins, in this order.
COPY_PRIORITY = ["translated_copy", "master_copy"]

# Sheets we never process (templates, validation, reference tabs, prior output).
SKIP_SHEETS = {
    "_old template_ outbound", "data validation", "qc categories",
    "key message", "timeline", "outbound strategy", "paid search tab",
    "asset map", "figma export", "figma export - outbound", "processing log",
}

# Junk placeholder values treated as empty cells.
PLACEHOLDERS = {"", "requestor fill in", "brand to fill in", "do not touch"}


# ============================================================
# FUZZY MATCHING HELPERS
# ============================================================

def normalise_header(h) -> str:
    if h is None:
        return ""
    h = str(h).strip().lower()
    h = re.sub(r"^\[new\]\s*", "", h)       # strip "[New] " prefix
    h = re.sub(r"\s*\(.*?\)\s*", " ", h)    # strip "(EN)" / "(LL)" etc.
    h = re.sub(r"[^a-z0-9 ]+", " ", h)      # punctuation -> space
    h = re.sub(r"\s+", " ", h)
    return h.strip()


def fuzzy_score(a: str, b: str) -> float:
    a, b = normalise_header(a), normalise_header(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    # substring only counts when BOTH strings are substantial (avoids "#"
    # matching "email #" and similar one-character false positives).
    if len(a) >= 3 and len(b) >= 3 and (a in b or b in a):
        return 0.85
    return SequenceMatcher(None, a, b).ratio()


def detect_columns(sheet, header_row: int = 1) -> dict:
    """Return {canonical_field: column_index} using fuzzy header matching."""
    header_fits: dict[str, tuple[int, float]] = {}
    for c in range(1, sheet.max_column + 1):
        cv = sheet.cell(row=header_row, column=c).value
        norm = normalise_header(cv)
        if len(norm) < 2:          # ignore 1-char headers like "#"
            continue
        best_field, best_score = None, 0.6
        for field, synonyms in HEADER_SYNONYMS.items():
            for syn in synonyms:
                s = fuzzy_score(norm, syn)
                if s > best_score:
                    best_score, best_field = s, field
        if best_field:
            # Keep the best-scoring column for each field.
            if best_field not in header_fits or best_score > header_fits[best_field][1]:
                header_fits[best_field] = (c, best_score)
    return {f: c for f, (c, _s) in header_fits.items()}


def cell(sheet, r, c):
    if not c:
        return ""
    v = sheet.cell(row=r, column=c).value
    return "" if v is None else str(v).strip()


def is_placeholder(v: str) -> bool:
    return str(v or "").strip().lower() in PLACEHOLDERS


def clean_copy(text) -> str:
    if text is None:
        return ""
    return str(text).replace("\\n", "\n").replace("\u00a0", " ").strip()


def pick_copy(sheet, r, col_map) -> str:
    """Copy priority: translated (LL) first, then master (EN)."""
    for field in COPY_PRIORITY:
        if field in col_map:
            val = clean_copy(sheet.cell(row=r, column=col_map[field]).value)
            if val:
                return val
    return ""


def find_data_start(sheet, col_map) -> int:
    """First row whose merchant cell holds real (non-placeholder) data."""
    mc = col_map.get("merchant")
    if not mc:
        return 3
    for r in range(2, min(sheet.max_row + 1, 15)):
        v = cell(sheet, r, mc)
        if v and not is_placeholder(v):
            return r
    return 3


# ============================================================
# SHEET CLASSIFICATION
# ============================================================

def classify_sheet(sheet, col_map) -> str:
    # A sheet can carry BOTH an asset_type column (e.g. "Push") and a
    # placement_type column (e.g. "Notification title"); scan both so the
    # push/email keywords are found wherever they live.
    placement_cols = [col_map.get("asset_type"), col_map.get("placement_type")]
    placement_cols = [c for c in placement_cols if c]
    if not placement_cols:
        return "unknown"

    values = set()
    for pc in placement_cols:
        for r in range(2, min(sheet.max_row + 1, 40)):
            v = cell(sheet, r, pc)
            if v and not is_placeholder(v):
                values.add(v)

    email_kw = {"Subject", "Super Subject", "Hero", "Card image",
                "Category Cards", "Banner Image", "Category Row"}
    push_kw = {"Notification title", "Notification text", "CTA"}
    onsite_kw = {"Tall Hero", "Window Display Header", "Window Display Sub-header",
                 "Quad Card", "ILM", "Sparkle", "TSTO", "Shoveler",
                 "Single Category Card", "Mobile Footer", "Mobile Billboard"}

    if values & push_kw and "email_num" not in col_map:
        return "outbound_push"
    if (len(values & email_kw) >= 1) or "email_num" in col_map:
        return "outbound_email"
    if (values & onsite_kw) or "device" in col_map:
        return "onsite"
    return "unknown"


# ============================================================
# ONSITE PROCESSOR
# ============================================================

def process_onsite(sheet, col_map, sheet_name, log):
    rows, corrections = [], []
    start = find_data_start(sheet, col_map)
    ac = col_map.get("asset_type") or col_map.get("placement_type")
    dvc = col_map.get("device")
    mc = col_map.get("merchant")

    raw = []
    for r in range(start, sheet.max_row + 1):
        merch = cell(sheet, r, mc)
        if not merch or is_placeholder(merch):
            continue
        raw.append({
            "rn": r,
            "at": cell(sheet, r, ac),
            "dv": cell(sheet, r, dvc),
            "me": merch,
            "co": cell(sheet, r, col_map.get("country")),
            "bu": cell(sheet, r, col_map.get("business")),
            "tg": cell(sheet, r, col_map.get("targeting")),
        })

    # --- ILM device auto-correction ---
    # A valid ILM set is Desktop / Mobile / Tablet. When a merchant's ILM group
    # has two "Desktop" and no "Tablet", the duplicate Desktop is corrected to
    # Tablet (a common requestor mistake).
    i = 0
    while i < len(raw):
        rr = raw[i]
        if rr["at"] == "ILM":
            group = [rr]
            j = i + 1
            while (j < len(raw) and raw[j]["at"] == "ILM"
                   and raw[j]["me"] == rr["me"] and raw[j]["co"] == rr["co"]
                   and raw[j]["tg"] == rr["tg"]):
                group.append(raw[j])
                j += 1
            devices = [g["dv"] for g in group]
            if len(group) == 3 and devices.count("Desktop") == 2 and "Tablet" not in devices:
                for k, g in enumerate(group):
                    if k > 0 and g["dv"] == "Desktop" and devices[:k + 1].count("Desktop") == 2:
                        corrections.append(
                            f"✓ {g['me']} ({g['co']}) Row {g['rn']} — Desktop → Tablet")
                        g["dv"] = "Tablet"
                        break
            i = j
        else:
            i += 1

    # --- Build export rows ---
    for idx, rr in enumerate(raw):
        key = f"{rr['at']}|{rr['dv']}"
        info = ONSITE_ASSET_MAP.get(key)
        if info:
            role, code = info["role"], info["code"]
            if role in ("SKIP", "SUBTITLE") or not code:
                continue
        else:
            role, code = "TITLE", ONSITE_FALLBACK_CODE  # unknown -> keep, flag

        copy = pick_copy(sheet, rr["rn"], col_map)
        if not copy:
            continue

        m = rr["me"].replace("...", "")
        t = rr["tg"].replace(" + ", "_").replace(" ", "")

        # Pair a Window Display Header with the Sub-header on the next row.
        subtitle = ""
        if rr["at"] == "Window Display Header" and idx + 1 < len(raw):
            nxt = raw[idx + 1]
            if nxt["at"] == "Window Display Sub-header" and nxt["me"] == rr["me"]:
                subtitle = pick_copy(sheet, nxt["rn"], col_map)

        rows.append({
            "/**/": f"{code}_{m}",
            "/*name*/": f"{m}-{rr['co']}-{t}",
            "Marketplace": rr["co"], "Business": rr["bu"], "Merchant": rr["me"],
            "Asset type": rr["at"], "Device Type": rr["dv"], "Asset Code": code,
            "Eyebrow": "--", "Title": copy, "Subtitle": subtitle,
            "Legal": "T&Cs apply", "Image": "", "BGColour": BG_COLOURS.get(rr["bu"], ""),
            "Source": sheet_name,
        })

    log.append({"sheet": sheet_name, "type": "Onsite", "rows": len(rows)})
    return rows, corrections


# ============================================================
# OUTBOUND PROCESSORS (email + push)
# ============================================================

def _finish_group(group, placements, channel_label, asset_code):
    if not group or not any(placements.values()):
        return None
    m = group["merchant"].replace("...", "")
    t = group["targeting"].replace(" + ", "_").replace(" ", "")
    suffix = group.get("email", "").replace(" ", "") or channel_label
    row = {
        "/**/": f"{asset_code}_{m}",
        "/*name*/": f"{m}-{group['country']}-{t}-{suffix}",
        "Marketplace": group["country"], "Business": group["business"],
        "Merchant": group["merchant"],
        "Channel": group.get("email") or channel_label,
        "Targeting": group["targeting"], "Asset Code": asset_code,
    }
    for pn, ct in placements.items():
        row[pn] = ct
        code = OUTBOUND_MAP.get(pn, "")
        if code:
            row[f"{pn} Code"] = code
    return row


def process_outbound(sheet, col_map, sheet_name, kind, log):
    """kind = 'email' or 'push'. Groups consecutive placement rows per asset."""
    rows = []
    start = find_data_start(sheet, col_map)
    pc = col_map.get("placement_type") or col_map.get("asset_type")
    mc = col_map.get("merchant")
    ec = col_map.get("email_num")

    # New group begins on an email number (email) or a "Notification title" (push).
    def starts_group(email_num, placement):
        return bool(email_num) if kind == "email" else (placement == "Notification title")

    channel_label = "Email" if kind == "email" else "Push"
    asset_code = "email1" if kind == "email" else "push1"

    group, placements = None, {}
    for r in range(start, sheet.max_row + 1):
        merch = cell(sheet, r, mc)
        if not merch or is_placeholder(merch):
            continue
        email_num = cell(sheet, r, ec)
        placement = cell(sheet, r, pc)

        if starts_group(email_num, placement):
            done = _finish_group(group, placements, channel_label, asset_code)
            if done:
                done["Source"] = sheet_name
                rows.append(done)
            group = {
                "country": cell(sheet, r, col_map.get("country")),
                "business": cell(sheet, r, col_map.get("business")),
                "merchant": merch,
                "targeting": cell(sheet, r, col_map.get("targeting")),
                "email": email_num if kind == "email" else "",
            }
            placements = {}

        if placement and group is not None:
            copy = pick_copy(sheet, r, col_map)
            if copy:
                placements[placement] = copy

    done = _finish_group(group, placements, channel_label, asset_code)
    if done:
        done["Source"] = sheet_name
        rows.append(done)

    log.append({"sheet": sheet_name, "type": channel_label, "rows": len(rows)})
    return rows


# ============================================================
# MAIN PIPELINE
# ============================================================

def run_pipeline(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    onsite, outbound, corrections, log = [], [], [], []
    detected = []

    for sn in wb.sheetnames:
        if sn.strip().lower() in SKIP_SHEETS:
            continue
        ws = wb[sn]
        if ws.max_row < 2 or ws.max_column < 2:
            continue
        col_map = detect_columns(ws)

        # Guard: a "Pillar" column must not be mistaken for a device column.
        if "device" in col_map:
            hv = ws.cell(row=1, column=col_map["device"]).value
            if hv and "pillar" in str(hv).lower():
                del col_map["device"]

        stype = classify_sheet(ws, col_map)
        if stype == "onsite":
            r, corr = process_onsite(ws, col_map, sn, log)
            onsite.extend(r)
            corrections.extend(corr)
            detected.append((sn, len(r)))
        elif stype == "outbound_email":
            r = process_outbound(ws, col_map, sn, "email", log)
            outbound.extend(r)
            detected.append((sn, len(r)))
        elif stype == "outbound_push":
            r = process_outbound(ws, col_map, sn, "push", log)
            outbound.extend(r)
            detected.append((sn, len(r)))
        else:
            log.append({"sheet": sn, "type": "Skipped", "rows": 0})

    # --- Market distribution ---
    market_counts = Counter(
        r.get("Marketplace", "") for r in (onsite + outbound) if r.get("Marketplace"))

    # --- Duplicate check (onsite): same /**/ + /*name*/ + Device ---
    seen, dup_count = set(), 0
    for r in onsite:
        sig = (r["/**/"], r["/*name*/"], r.get("Device Type", ""))
        if sig in seen:
            dup_count += 1
        seen.add(sig)

    # --- ILM balance: count ILM rows per device ---
    ilm_devices = Counter(
        r.get("Device Type", "") for r in onsite if r.get("Asset type") == "ILM")

    return {
        "onsite": onsite, "outbound": outbound,
        "corrections": corrections, "log": log, "detected": detected,
        "markets": dict(market_counts),
        "dup_count": dup_count, "ilm_devices": dict(ilm_devices),
    }


# ============================================================
# EXCEL OUTPUT
# ============================================================

ONSITE_HEADERS = ["/**/", "/*name*/", "Marketplace", "Business", "Merchant",
                  "Asset type", "Device Type", "Asset Code", "Eyebrow",
                  "Title", "Subtitle", "Legal", "Image", "BGColour", "Source"]


def build_workbook(result: dict) -> bytes:
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Figma Export"
    ws1.append(ONSITE_HEADERS)
    for row in result["onsite"]:
        ws1.append([row.get(h, "") for h in ONSITE_HEADERS])

    ws2 = wb.create_sheet("Figma Export - Outbound")
    if result["outbound"]:
        keys = []
        for row in result["outbound"]:
            for k in row:
                if k not in keys:
                    keys.append(k)
        ws2.append(keys)
        for row in result["outbound"]:
            ws2.append([row.get(h, "") for h in keys])

    ws3 = wb.create_sheet("Processing Log")
    ws3.append(["Sheet Name", "Type", "Rows Generated"])
    for e in result["log"]:
        ws3.append([e["sheet"], e["type"], e["rows"]])
    ws3.append([])
    ws3.append(["--- Auto-Corrections ---"])
    for c in result["corrections"]:
        ws3.append([c])
    ws3.append([])
    ws3.append(["Total Onsite Rows", len(result["onsite"])])
    ws3.append(["Total Outbound Rows", len(result["outbound"])])
    ws3.append(["Total Corrections", len(result["corrections"])])
    ws3.append(["Duplicates after correction", result["dup_count"]])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# DASHBOARD UI
# ============================================================

MARKET_COLORS = ["#2f7bd6", "#e0a020", "#d6453f", "#3bab52", "#8b5cf6",
                 "#e06ab0", "#20bcd6", "#9ad63b"]


def stat_card(col, number, label, color="#8b7bf0"):
    col.markdown(
        f'<div class="fx-card"><div class="num" style="color:{color}">{number}</div>'
        f'<div class="lbl">{label}</div></div>', unsafe_allow_html=True)


def render_dashboard(result: dict, uploaded_name: str, output_bytes: bytes):
    onsite_n = len(result["onsite"])
    outbound_n = len(result["outbound"])
    total_n = onsite_n + outbound_n
    sheets_n = len([d for d in result["detected"]])
    corr_n = len(result["corrections"])

    # Stat cards
    cols = st.columns(5)
    stat_card(cols[0], total_n, "Total Rows", "#8b7bf0")
    stat_card(cols[1], onsite_n, "Onsite", "#cfd3ec")
    stat_card(cols[2], outbound_n, "Outbound", "#cfd3ec")
    stat_card(cols[3], sheets_n, "Sheets Processed", "#74d699")
    stat_card(cols[4], corr_n, "Auto-Corrections", "#f0c674")

    # Sheets detected
    st.markdown('<div class="fx-section">📄 Sheets Detected</div>', unsafe_allow_html=True)
    pills = "".join(
        f'<span class="fx-pill">{sn}<span class="cnt">{n}</span></span>'
        for sn, n in result["detected"])
    st.markdown(pills or "—", unsafe_allow_html=True)

    # Market distribution bar
    st.markdown('<div class="fx-section">🌍 Market Distribution</div>', unsafe_allow_html=True)
    markets = result["markets"]
    if markets:
        total_m = sum(markets.values()) or 1
        segs = ""
        for i, (mk, cnt) in enumerate(sorted(markets.items())):
            pct = cnt / total_m * 100
            color = MARKET_COLORS[i % len(MARKET_COLORS)]
            segs += (f'<div class="fx-seg" style="width:{pct}%;background:{color}">'
                     f'{mk} {cnt}</div>')
        st.markdown(f'<div class="fx-bar">{segs}</div>', unsafe_allow_html=True)
    else:
        st.markdown("—")

    # Alerts & auto-corrections
    st.markdown(
        f'<div class="fx-section">⚠️ Alerts & Auto-Corrections'
        f'<span class="fx-badge">{corr_n} fixes</span></div>', unsafe_allow_html=True)
    if result["corrections"]:
        body = ("<span class='fx-title'>ILM Device Auto-Correct: Duplicate Desktop → Tablet</span>"
                + "\n".join(result["corrections"]))
        st.markdown(f'<div class="fx-panel fx-warn">{body}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="fx-panel fx-ok">No auto-corrections needed.</div>',
                    unsafe_allow_html=True)

    # Duplicate check
    dup = result["dup_count"]
    dup_cls = "fx-ok" if dup == 0 else "fx-warn"
    st.markdown(
        f'<div class="fx-panel {dup_cls}">🔁 Duplicate Check: {dup} remaining '
        f'duplicate(s) after auto-correction</div>', unsafe_allow_html=True)

    # ILM balance
    ilm = result["ilm_devices"]
    if ilm:
        d, m, t = ilm.get("Desktop", 0), ilm.get("Mobile", 0), ilm.get("Tablet", 0)
        balanced = (d == m == t) and d > 0
        cls = "fx-ok" if balanced else "fx-warn"
        verdict = "perfectly matched" if balanced else "not balanced — please review"
        st.markdown(
            f'<div class="fx-panel {cls}">⚖️ ILM Balance: {d} Desktop · {m} Mobile · '
            f'{t} Tablet — {verdict}</div>', unsafe_allow_html=True)

    # Preview tables
    st.markdown('<div class="fx-section">📋 Preview</div>', unsafe_allow_html=True)
    tab1, tab2 = st.tabs([f"Onsite ({onsite_n} rows)", f"Outbound ({outbound_n} rows)"])
    with tab1:
        if result["onsite"]:
            prev = [{"/**/": r["/**/"], "/*name*/": r["/*name*/"], "MKT": r["Marketplace"],
                     "Merchant": r["Merchant"], "Asset": r["Asset type"],
                     "Device": r["Device Type"], "Code": r["Asset Code"],
                     "Title": r["Title"], "Source": r["Source"]}
                    for r in result["onsite"]]
            st.dataframe(prev, use_container_width=True, height=360)
        else:
            st.info("No onsite rows generated.")
    with tab2:
        if result["outbound"]:
            prev = [{"/**/": r["/**/"], "/*name*/": r["/*name*/"],
                     "MKT": r.get("Marketplace", ""), "Merchant": r.get("Merchant", ""),
                     "Channel": r.get("Channel", ""), "Code": r.get("Asset Code", ""),
                     "Source": r.get("Source", "")}
                    for r in result["outbound"]]
            st.dataframe(prev, use_container_width=True, height=360)
        else:
            st.info("No outbound rows generated.")

    # Download
    st.markdown('<div class="fx-section">⬇️ Download</div>', unsafe_allow_html=True)
    out_name = re.sub(r"\.xlsx$", "", uploaded_name, flags=re.IGNORECASE) + "_Figma_Export.xlsx"
    st.download_button(
        "Download Figma Export (.xlsx)", data=output_bytes, file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")


# ============================================================
# APP ENTRY
# ============================================================

def main():
    st.markdown(
        '<div class="fx-header">'
        '<h1>🎨 <span class="accent">Figma Export</span> Generator</h1>'
        '<div class="sub">Creative Brief → Figma-ready Excel · Auto-detect · '
        'Auto-correct · Auto-export</div></div>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Drop your creative brief Excel file here",
                                type=["xlsx"], label_visibility="collapsed")
    if uploaded is None:
        st.info("⬆️ Upload a creative brief (.xlsx) to begin.")
        return

    try:
        file_bytes = uploaded.getvalue()
        with st.spinner("Processing brief..."):
            result = run_pipeline(file_bytes)
            output_bytes = build_workbook(result)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not process the file: {exc}")
        return

    st.success(f"✅ Processed **{uploaded.name}**")
    render_dashboard(result, uploaded.name, output_bytes)


if __name__ == "__main__":
    main()

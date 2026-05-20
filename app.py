"""
OT Cybersecurity Weekly Brief — Streamlit web app
Deployed at share.streamlit.io. Bryan clicks button → scraper runs → branded
brief renders in-page. Powered by Banneker Partners.
"""

import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------- Config ----------

LOOKBACK_DAYS = 7

QUERIES = {
    "EU regulation (NIS2 / CRA / DORA)": [
        "NIS2 directive",
        "Cyber Resilience Act",
        "DORA cybersecurity",
        "ENISA OT",
        "NIS2 enforcement fine",
    ],
    "North American regulation (NERC-CIP / TSA / CISA)": [
        "NERC CIP",
        "NERC CIP compliance",
        "NERC CIP audit penalty",
        "FERC cybersecurity",
        "TSA pipeline cybersecurity directive",
        "CISA OT advisory",
        "CISA critical infrastructure",
        "EPA water cybersecurity",
    ],
    "OT / ICS cybersecurity": [
        "OT cybersecurity",
        "ICS cybersecurity",
        "operational technology security",
        "SCADA security",
        "industrial control system security",
    ],
    "Country-specific (Europe)": [
        "Germany cybersecurity regulation OT",
        "Poland critical infrastructure cybersecurity",
        "UK NIS regulations",
        "Italy NIS2",
        "Austria Cybersicherheit OT",
        "Switzerland critical infrastructure cybersecurity",
        "Turkey energy cybersecurity",
        "Ireland NIS2",
    ],
    "Breaches & incidents": [
        "OT cyberattack",
        "ICS ransomware",
        "critical infrastructure breach",
        "utility cyberattack",
        "water utility cyberattack",
        "power grid cyberattack",
        "oil gas pipeline cyberattack",
        "manufacturing ransomware shutdown",
        "factory ransomware",
        "energy sector breach",
        "SCADA attack",
        "industrial ransomware",
    ],
    "Competitive / market": [
        "Industrial Defender OT",
        "Claroty OT security",
        "Dragos OT cybersecurity",
        "Nozomi Networks",
        "Armis OT",
        "TXOne Networks",
    ],
}

# Banneker palette
NAVY = "#181F64"
DEEP_NAVY = "#0D0537"
DARK_GREY = "#59595B"
PERIWINKLE = "#76A3E3"
ICE_BLUE = "#E4EDF9"
MID_GREY = "#A6A6A6"

# ---------- Scrape ----------

def fetch_rss(query: str) -> list[dict]:
    url = (
        "https://news.google.com/rss/search?"
        f"q={urllib.parse.quote(query)}&hl=en-US&gl=US&ceid=US:en"
    )
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            data = r.read()
    except Exception:
        return []

    items = []
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []

    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub_raw = (item.findtext("pubDate") or "").strip()
        source_el = item.find("source")
        source = source_el.text if source_el is not None and source_el.text else ""
        try:
            pub = parsedate_to_datetime(pub_raw)
            if pub.tzinfo is None:
                pub = pub.replace(tzinfo=timezone.utc)
        except Exception:
            continue
        items.append({
            "title": title,
            "link": link,
            "published": pub,
            "source": source,
        })
    return items


@st.cache_data(ttl=1800, show_spinner=False)
def scrape_all() -> tuple[dict, int]:
    """Run all queries, dedupe, return (grouped, total). Cached 30 min."""
    cutoff = datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)
    grouped: dict[str, list[dict]] = {}
    seen_links: set[str] = set()
    seen_titles: set[str] = set()

    for theme, queries in QUERIES.items():
        items = []
        for q in queries:
            for it in fetch_rss(q):
                if it["published"] < cutoff:
                    continue
                if it["link"] in seen_links:
                    continue
                title_key = it["title"].lower()[:120]
                if title_key in seen_titles:
                    continue
                seen_links.add(it["link"])
                seen_titles.add(title_key)
                items.append(it)
        items.sort(key=lambda x: x["published"], reverse=True)
        grouped[theme] = items

    total = sum(len(v) for v in grouped.values())
    return grouped, total


# ---------- DOCX builder (for download) ----------

def _set_font(rPr, font_name="Inter", size_pt=11, color=None, bold=False, italic=False, underline=False):
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    rPr.append(sz)
    if color is not None:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)
    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)


def _style_run(run, size_pt=11, color_hex="59595B", bold=False, italic=False):
    run.font.name = "Inter"
    run.font.size = Pt(size_pt)
    r, g, b = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
    run.font.color.rgb = RGBColor(r, g, b)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), "Inter")
    rFonts.set(qn("w:hAnsi"), "Inter")


def _add_hyperlink(paragraph, url, text, size_pt=11):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    _set_font(rPr, size_pt=size_pt, color="181F64", underline=True)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def build_docx_bytes(grouped: dict, total: int) -> bytes:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Inter"
    normal.font.size = Pt(11)

    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
        section.footer_distance = Inches(0.4)

    today_pretty = datetime.now().strftime("%B %d, %Y")

    p = doc.add_paragraph()
    _style_run(p.add_run("OT Cybersecurity Weekly Brief"), size_pt=20, color_hex="181F64", bold=True)
    p = doc.add_paragraph()
    _style_run(p.add_run(today_pretty), size_pt=12, color_hex="76A3E3", bold=True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    _style_run(
        p.add_run(f"Past {LOOKBACK_DAYS} days  |  {total} stories  |  {len(grouped)} themes  |  Source: Google News RSS"),
        size_pt=9, color_hex="59595B", italic=True,
    )

    for theme, items in grouped.items():
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
        _style_run(h.add_run(theme), size_pt=13, color_hex="181F64", bold=True)

        if not items:
            empty = doc.add_paragraph()
            _style_run(empty.add_run("No coverage in the lookback window"),
                       size_pt=10, color_hex="A6A6A6", italic=True)
            continue

        for it in items:
            bp = doc.add_paragraph()
            bp.paragraph_format.left_indent = Inches(0.25)
            bp.paragraph_format.first_line_indent = Inches(-0.18)
            bp.paragraph_format.space_after = Pt(2)
            _style_run(bp.add_run("•  "), size_pt=11, color_hex="181F64", bold=True)
            _style_run(bp.add_run(it["published"].strftime("%b %d") + "  "),
                       size_pt=11, color_hex="181F64", bold=True)
            _add_hyperlink(bp, it["link"], it["title"])
            if it["source"]:
                _style_run(bp.add_run(f"  — {it['source']}"),
                           size_pt=10, color_hex="A6A6A6", italic=True)

    footer = doc.sections[0].footer
    fp = footer.paragraphs[0]
    fp.text = ""
    year = datetime.now().year
    _style_run(fp.add_run(f"(C) {year} Banneker Partners, LLC. All Rights Reserved. Confidential."),
               size_pt=8, color_hex="59595B")

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------- UI ----------

st.set_page_config(
    page_title="OT Cyber Weekly Brief — Banneker Partners",
    page_icon="🛡️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(f"""
<style>
  .stApp {{ background: #FFFFFF; }}
  html, body, [class*="css"] {{
    font-family: Inter, Calibri, "Segoe UI", Arial, sans-serif !important;
  }}
  h1 {{ color: {NAVY} !important; font-weight: 700 !important; }}
  h2, h3 {{ color: {NAVY} !important; font-weight: 700 !important; }}
  p, li, div, span, label {{ color: {DARK_GREY}; }}
  .block-container {{ padding-top: 2.5rem; max-width: 760px; }}
  .stButton > button {{
    background-color: {NAVY} !important;
    color: white !important;
    font-weight: 600 !important;
    border: none !important;
    padding: 14px 32px !important;
    border-radius: 4px !important;
    font-size: 16px !important;
    width: 100%;
  }}
  .stButton > button:hover {{
    background-color: {DEEP_NAVY} !important;
    color: white !important;
  }}
  .stDownloadButton > button {{
    background-color: {ICE_BLUE} !important;
    color: {NAVY} !important;
    border: 1px solid {NAVY} !important;
    font-weight: 600 !important;
    padding: 10px 22px !important;
  }}
  a {{ color: {NAVY} !important; text-decoration: underline; }}
  a:hover {{ color: {PERIWINKLE} !important; }}
  .meta {{ color: {MID_GREY}; font-style: italic; font-size: 13px; }}
  .date {{ color: {PERIWINKLE}; font-weight: 600; }}
  .theme-card {{
    border-left: 3px solid {NAVY};
    padding: 4px 0 4px 14px;
    margin: 26px 0 10px;
  }}
  .story-row {{ margin-bottom: 8px; line-height: 1.45; }}
  .story-date {{ color: {NAVY}; font-weight: 600; font-size: 13px; }}
  .story-source {{ color: {MID_GREY}; font-style: italic; font-size: 12px; }}
  .summary-box {{
    background: {ICE_BLUE};
    border-left: 4px solid {NAVY};
    padding: 14px 18px;
    margin: 14px 0 28px;
    border-radius: 2px;
  }}
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("# OT Cybersecurity Weekly Brief")
st.markdown(
    f"<div class='date' style='font-size:15px; margin-top:-8px; margin-bottom:6px;'>"
    f"Banneker Partners · Industrial Defender market intel</div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='meta' style='margin-bottom:20px;'>"
    "Click the button to pull the last 7 days of OT cyber news across EU/NA regulation, "
    "ICS security, breaches, and the competitive landscape. Fresh every click."
    "</div>",
    unsafe_allow_html=True,
)

# Button
go = st.button("Generate this week's brief", type="primary")

if go or st.session_state.get("has_results"):
    if go or "results" not in st.session_state:
        with st.spinner("Pulling stories from Google News..."):
            grouped, total = scrape_all()
        st.session_state["results"] = (grouped, total)
        st.session_state["has_results"] = True
        st.session_state["generated_at"] = datetime.now()
    else:
        grouped, total = st.session_state["results"]

    generated = st.session_state["generated_at"]

    st.markdown(
        f"<div style='margin-top:24px; margin-bottom:4px;'>"
        f"<span class='date' style='font-size:18px;'>{generated.strftime('%B %d, %Y')}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div class='meta' style='margin-bottom:18px;'>"
        f"Past {LOOKBACK_DAYS} days &middot; {total} stories &middot; "
        f"{len(grouped)} themes &middot; Source: Google News RSS"
        f"</div>",
        unsafe_allow_html=True,
    )

    # Download button
    docx_bytes = build_docx_bytes(grouped, total)
    st.download_button(
        label="Download as Word doc",
        data=docx_bytes,
        file_name=f"ot_cyber_brief_{generated.strftime('%Y-%m-%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    # Themed sections
    for theme, items in grouped.items():
        st.markdown(f"<div class='theme-card'><h3 style='margin:0;'>{theme}</h3></div>",
                    unsafe_allow_html=True)
        if not items:
            st.markdown(f"<div class='meta'>No coverage in the lookback window.</div>",
                        unsafe_allow_html=True)
            continue
        for it in items:
            date = it["published"].strftime("%b %d")
            src = f" <span class='story-source'>— {it['source']}</span>" if it["source"] else ""
            st.markdown(
                f"<div class='story-row'>"
                f"<span class='story-date'>{date}</span> &nbsp; "
                f"<a href='{it['link']}' target='_blank'>{it['title']}</a>"
                f"{src}</div>",
                unsafe_allow_html=True,
            )

# Footer
st.markdown("<br><hr style='border:none; border-top:1px solid #E4EDF9; margin-top:40px;'>",
            unsafe_allow_html=True)
st.markdown(
    f"<div style='color:{MID_GREY}; font-size:11px;'>"
    f"(C) {datetime.now().year} Banneker Partners, LLC. All Rights Reserved. Confidential."
    f"</div>",
    unsafe_allow_html=True,
)

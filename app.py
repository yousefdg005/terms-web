"""
AI Legal Workspace — Alignment Studio + Term Extractor
Enhanced & Debugged Version — (c) Youssef Nour Eldin
"""
import base64
import hashlib
import io
import json
import os
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher
from html import escape as html_escape
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from openai import OpenAI

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# =========================================================
# 1) إعداد الصفحة
# =========================================================
st.set_page_config(
    page_title="AI Legal Workspace", page_icon="⚖️",
    layout="wide", initial_sidebar_state="expanded",
)


def load_css(file_name="style.css"):
    try:
        with open(file_name, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        pass


load_css("style.css")


# =========================================================
# 2) الهيدر
# =========================================================
st.markdown("""
<div class="header-container" style="direction: ltr; text-align: left;">
  <div class="header-text">
    <h1 style="color:#EAB308;margin:0;font-size:26px;font-weight:800;">⚖️ AI Legal Workspace</h1>
    <p style="color:#F3F4F6;font-size:15px;margin-top:8px;">Your all-in-one AI toolkit for legal text alignment & term extraction</p>
    <span style="color:#E5E7EB;font-size:13px;background:rgba(234,179,8,.15);padding:6px 18px;border-radius:20px;border:1px solid rgba(234,179,8,.4);">
      Smart AI Engine | Developed by Youssef Nour Eldin 💻
    </span>
  </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# 3) الشريط الجانبي
# =========================================================
st.sidebar.header("⚙️ إعدادات النظام")
st.sidebar.markdown("""
<div style="background:#1E293B;padding:12px;border-radius:8px;border:1px solid #374151;margin-bottom:15px;text-align:center;">
  <span style="color:#E5E7EB;font-size:13px;">ليس لديك مفتاح API؟</span><br>
  <a href="https://openrouter.ai/keys" target="_blank" style="color:#EAB308;font-weight:bold;text-decoration:none;font-size:14px;">
    🔗 الحصول على المفتاح من OpenRouter
  </a>
</div>
""", unsafe_allow_html=True)

api_key = st.sidebar.text_input(
    "🔑 مفتاح OpenRouter API Key:", value="", type="password",
    placeholder="sk-or-v1-...",
)

PAID_MODELS_MAPPING = {
    "🌙 Qwen 3.5 Flash (الأقوى في معالجة اللغة العربية)": "qwen/qwen3.5-flash-02-23",
    "📜 Claude Haiku 5.5 (التزام صارم بالتعليمات)": "anthropic/claude-haiku-5.5",
    "⚡ Gemini 3.8 Flash (السياقات الطويلة)": "google/gemini-3.8-flash",
    "⚖️ GPT-4o Mini (استخراج هيكلي دقيق)": "openai/gpt-4o-mini",
    "🦙 Llama 4 Maverick (نموذج مفتوح)": "meta-llama/llama-4-maverick",
}
st.sidebar.subheader("🤖 اختيار النموذج")
selected_model_label = st.sidebar.selectbox(
    "اختر النموذج:", options=list(PAID_MODELS_MAPPING.keys()), index=0,
)
selected_model_id = PAID_MODELS_MAPPING[selected_model_label]


# =========================================================
# 4) Session State — باستخدام .setdefault لتفادي KeyError
# =========================================================
st.session_state.setdefault("aligned_pairs", [])
st.session_state.setdefault("results_df", None)
st.session_state.setdefault("tmx_bytes", None)
st.session_state.setdefault("xlsx_bytes", None)
st.session_state.setdefault("csv_bytes", None)


# =========================================================
# 5) أدوات التطبيع والتقسيم
# =========================================================
_INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]")
_XML_BAD_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# اختصارات قانونية لا تكسر الجملة
_ABBREV = {
    "mr", "mrs", "ms", "dr", "prof", "inc", "ltd", "co", "corp", "vs", "etc",
    "no", "art", "sec", "vol", "ch", "pp", "ed", "al", "cf", "eg", "ie",
    "paragraph", "para", "u.s", "u.k", "e.u", "u.n", "v", "p",
}


def normalize_text(text: str) -> str:
    """توحيد Unicode + إزالة الأحرف الخفية."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = _INVISIBLE_RE.sub("", text)
    return text.replace("\u00a0", " ")


def segment_sentences(text: str):
    """
    تقسيم الجمل القانونية مع حماية:
      - الأرقام العشرية (1.5)
      - الاختصارات (Art. No. Dr.)
      - الفواصل العربية/الإنجليزية
    """
    if not text:
        return []
    # 1) حماية الأرقام العشرية
    text = re.sub(r"(\d)\.(\d)", r"\1\uE000\2", text)
    # 2) حماية الاختصارات
    def _protect(m):
        w = m.group(1).lower().rstrip(".")
        return m.group(0).replace(".", "\uE000") if w in _ABBREV else m.group(0)
    text = re.sub(r"\b([A-Za-z][A-Za-z\.]{0,9})\.", _protect, text)
    # 3) التقسيم
    parts = re.split(r"(?<=[.!?؟])\s+(?=[A-Za-z\u0600-\u06FF0-9\"'\(\[])|[؛;\n]+", text)
    # 4) الاستعادة + التنظيف
    out = []
    for p in parts:
        p = p.replace("\uE000", ".").strip()
        if p:
            out.append(re.sub(r"\s+", " ", p))
    return out


def extract_text_from_docx_bytes(data: bytes):
    notes = []
    try:
        doc = Document(io.BytesIO(data))
        paras = [normalize_text(p.text.strip()) for p in doc.paragraphs if p.text.strip()]
        for t_idx, table in enumerate(doc.tables, 1):
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    paras.append(normalize_text(" | ".join(cells)))
            notes.append(f"تم استخراج محتوى الجدول #{t_idx}")
        return "\n".join(paras), notes
    except Exception as e:
        return "", [f"خطأ أثناء قراءة Word: {e}"]


# =========================================================
# 6) خوارزمية المحاذاة (DP — Gale-Church مبسّطة)
# =========================================================
def align_by_length(src_segs, tgt_segs, max_merge: int = 2):
    """
    محاذاة قائمة على نِسب الأطوال (Gale-Church مبسطة).
    تسمح بالحركات: 1-1, 1-0, 0-1, 1-2, 2-1, 2-2.
    الإرجاع: list[(en_text, ar_text)]
    """
    n, m = len(src_segs), len(tgt_segs)
    if n == 0 and m == 0:
        return []
    # مسافات log لتفادي الانحياز للجمل الطويلة
    src_len = np.array([max(len(s), 2) for s in src_segs] + [1], dtype=np.float64)
    tgt_len = np.array([max(len(t), 2) for t in tgt_segs] + [1], dtype=np.float64)

    INF = 1e18
    cost = np.full((n + 1, m + 1), INF)
    back = np.full((n + 1, m + 1, 2), -1, dtype=np.int32)
    cost[0, 0] = 0.0

    DEL_PEN, MERGE_PEN = 0.85, 0.20  # عقوبات الإدراج/الحذف/الدمج

    for i in range(n + 1):
        for j in range(m + 1):
            c = cost[i, j]
            if c >= INF:
                continue
            # 1-1
            if i < n and j < m:
                d = abs(np.log(src_len[i] / tgt_len[j]))
                if c + d < cost[i + 1, j + 1]:
                    cost[i + 1, j + 1] = c + d
                    back[i + 1, j + 1] = (i, j)
            # 1-0
            if i < n and c + DEL_PEN < cost[i + 1, j]:
                cost[i + 1, j] = c + DEL_PEN
                back[i + 1, j] = (i, j)
            # 0-1
            if j < m and c + DEL_PEN < cost[i, j + 1]:
                cost[i, j + 1] = c + DEL_PEN
                back[i, j + 1] = (i, j)
            # 1-2 و 2-1
            if max_merge >= 2:
                if i < n and j + 1 < m:
                    d = abs(np.log(src_len[i] / (tgt_len[j] + tgt_len[j + 1]))) + MERGE_PEN
                    if c + d < cost[i + 1, j + 2]:
                        cost[i + 1, j + 2] = c + d
                        back[i + 1, j + 2] = (i, j)
                if i + 1 < n and j < m:
                    d = abs(np.log((src_len[i] + src_len[i + 1]) / tgt_len[j])) + MERGE_PEN
                    if c + d < cost[i + 2, j + 1]:
                        cost[i + 2, j + 1] = c + d
                        back[i + 2, j + 1] = (i, j)
            # 2-2
            if max_merge >= 2 and i + 1 < n and j + 1 < m:
                d = abs(np.log((src_len[i] + src_len[i + 1]) /
                               (tgt_len[j] + tgt_len[j + 1]))) + 2 * MERGE_PEN
                if c + d < cost[i + 2, j + 2]:
                    cost[i + 2, j + 2] = c + d
                    back[i + 2, j + 2] = (i, j)

    # Backtrack
    pairs, i, j = [], n, m
    while i > 0 or j > 0:
        pi, pj = back[i, j]
        if pi < 0 or pj < 0:
            break
        en = " ".join(src_segs[pi:i]).strip()
        ar = " ".join(tgt_segs[pj:j]).strip()
        pairs.append((en, ar))
        i, j = int(pi), int(pj)
    pairs.reverse()
    return pairs


def fuzzy_score(a: str, b: str) -> float:
    """درجة تشابه حروف (لعرضها فقط / Debug)."""
    return SequenceMatcher(None, a or "", b or "").ratio()


# =========================================================
# 7) TMX / Excel / CSV
# =========================================================
def _xml_safe(t):
    return xml_escape(_XML_BAD_RE.sub("", str(t)))


def generate_tmx_file(pairs) -> bytes:
    """TMX 1.4 مع UTF-8 و معالجة XML آمنة."""
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<tmx version="1.4">',
        '<header creationtool="AI Legal Workspace" creationtoolversion="2.1" '
        'segtype="sentence" adminlang="en-us" srclang="en" datatype="plaintext"/>',
        "<body>",
    ]
    for en, ar in pairs:
        parts.append(
            f'  <tu>\n'
            f'    <tuv xml:lang="en"><seg>{_xml_safe(en)}</seg></tuv>\n'
            f'    <tuv xml:lang="ar"><seg>{_xml_safe(ar)}</seg></tuv>\n'
            f'  </tu>'
        )
    parts.append("</body>\n</tmx>")
    return "\n".join(parts).encode("utf-8")


def parse_tmx_bytes(content: bytes):
    """قراءة TMX بشكل آمن حتى لو احتوى BOM أو وسوم مختلفة."""
    pairs = []
    try:
        root = ET.fromstring(content)
    except ET.ParseError as e:
        st.error(f"❌ فشل قراءة TMX: {e}")
        return pairs
    for tu in root.iter():
        if not tu.tag.endswith("tu"):
            continue
        en, ar = "", ""
        for tuv in tu:
            if not tuv.tag.endswith("tuv"):
                continue
            lang = next((v.lower() for k, v in tuv.attrib.items() if "lang" in k.lower()), "")
            seg = next(("".join(s.itertext()) for s in tuv if s.tag.endswith("seg")), "").strip()
            if lang.startswith("en"):
                en = seg
            elif lang.startswith("ar"):
                ar = seg
        if en and ar:
            pairs.append((normalize_text(en), normalize_text(ar)))
    return pairs


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """UTF-8 BOM لضمان فتحه صحيحاً في Excel العربي."""
    return df.to_csv(index=False).encode("utf-8-sig")


# =========================================================
# 8) توليد بطاقات PDF
# =========================================================
def prepare_arabic_text(text):
    if not text or not isinstance(text, str):
        return ""
    return get_display(arabic_reshaper.reshape(text))


_FONT_CANDIDATES = [
    ("fonts/Amiri-Regular.ttf", "fonts/Amiri-Bold.ttf"),
    ("fonts/NotoNaskhArabic-Regular.ttf", "fonts/NotoNaskhArabic-Bold.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
]
_PDF_FONTS = None


def get_pdf_fonts():
    global _PDF_FONTS
    if _PDF_FONTS:
        return _PDF_FONTS
    for reg, bold in _FONT_CANDIDATES:
        if os.path.exists(reg):
            try:
                pdfmetrics.registerFont(TTFont("AppFont", reg))
                pdfmetrics.registerFont(TTFont("AppFont-Bold", bold if os.path.exists(bold) else reg))
                _PDF_FONTS = ("AppFont", "AppFont-Bold")
                return _PDF_FONTS
            except Exception:
                continue
    _PDF_FONTS = ("Helvetica", "Helvetica-Bold")
    return _PDF_FONTS


def _wrap_logical(text, font, size, max_width):
    lines, cur = [], ""
    for word in str(text).split():
        cand = f"{cur} {word}".strip()
        if cur and pdfmetrics.stringWidth(prepare_arabic_text(cand), font, size) > max_width:
            lines.append(cur); cur = word
        else:
            cur = cand
    if cur: lines.append(cur)
    return lines or [""]


def ar_par(text, style, max_width):
    lines = _wrap_logical(text, style.fontName, style.fontSize, max_width)
    html = "<br/>".join(html_escape(prepare_arabic_text(l)) for l in lines)
    return Paragraph(html, style)


def en_par(text, style):
    return Paragraph(html_escape(str(text or "")), style)


def generate_pdf_cards_bytes(df: pd.DataFrame) -> bytes:
    font, font_bold = get_pdf_fonts()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    ss = getSampleStyleSheet()
    title_st = ParagraphStyle("T", parent=ss["Heading1"], fontName=font_bold, fontSize=18,
                              alignment=1, textColor=colors.HexColor("#0F172A"), leading=26)
    sub_st = ParagraphStyle("S", parent=ss["Normal"], fontName=font, fontSize=10,
                            alignment=1, textColor=colors.HexColor("#64748B"), leading=14)
    ct_st = ParagraphStyle("CT", fontName=font_bold, fontSize=10, alignment=2,
                           leading=15, textColor=colors.HexColor("#1E293B"))
    cv_st = ParagraphStyle("CV", fontName=font, fontSize=10, alignment=2,
                           leading=15, textColor=colors.HexColor("#0F172A"))
    cv_en = ParagraphStyle("CVE", parent=cv_st, alignment=0)

    story = [
        ar_par("بطاقات المصطلحات القانونية الموثقة", title_st, 520),
        ar_par("منصة التحليل والاستخراج القانوني الذكي | إعداد: يوسف نورالدين", sub_st, 520),
        Spacer(1, 15),
    ]
    VAL_W, LBL_W, PAD = 360, 160, 6
    vw, lw = VAL_W - 2 * PAD - 4, LBL_W - 2 * PAD - 4
    total = len(df)
    for n, (_, row) in enumerate(df.iterrows(), 1):
        story.append(ar_par(f"بطاقة مصطلح ({n:02d} / {total:02d})", ct_st, 520))
        story.append(Spacer(1, 5))
        L = lambda t: ar_par(t, ct_st, lw)
        A = lambda k, d="": ar_par(str(row.get(k) or d), cv_st, vw)
        data = [
            [en_par(row.get("English Term", ""), cv_en), L("المصطلح بالإنجليزي")],
            [A("Arabic Term"), L("المقابل بالعربي")],
            [A("Sub Domain", "قانون عام"), L("المجال الفرعي")],
            [A("Document Type", "نص قانوني"), L("نوع المستند")],
            [en_par(row.get("Source EN", ""), cv_en), L("السياق الإنجليزي")],
            [A("Source AR"), L("السياق العربي")],
            [A("Arabic Description"), L("الشرح والتعريف")],
        ]
        t = Table(data, colWidths=[VAL_W, LBL_W])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("PADDING", (0, 0), (-1, -1), PAD),
        ]))
        story += [t, Spacer(1, 15),
                  HRFlowable(width="100%", thickness=1, color=colors.HexColor("#E2E8F0"), spaceAfter=15)]
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# 9) التبويبات
# =========================================================
tab1, tab2 = st.tabs(["🧩 Alignment Studio", "🔍 AI Term Extractor"])


# ---------------------------------------------------------
# TAB 1 : المحاذاة
# ---------------------------------------------------------
with tab1:
    st.subheader("📄 الخطوة 1: رفع ملفات Word (.docx)")
    c1, c2 = st.columns(2)
    with c1:
        doc_en = st.file_uploader("ملف المصدر (إنجليزي):", type=["docx"], key="doc_en")
    with c2:
        doc_ar = st.file_uploader("ملف الهدف (عربي):", type=["docx"], key="doc_ar")

    if doc_en and doc_ar:
        text_en, notes_en = extract_text_from_docx_bytes(doc_en.getvalue())
        text_ar, notes_ar = extract_text_from_docx_bytes(doc_ar.getvalue())

        if not (text_en.strip() and text_ar.strip()):
            st.error("⚠️ تعذّر استخراج نص صالح من أحد الملفين.")
        else:
            st.success("✅ تم استخراج النصوص بنجاح.")
            segs_en = segment_sentences(text_en)
            segs_ar = segment_sentences(text_ar)

            with st.expander("🔍 معاينة أول 10 جمل من كل ملف"):
                st.markdown("**EN:**")
                for s in segs_en[:10]: st.write("-", s)
                st.markdown("**AR:**")
                for s in segs_ar[:10]: st.write("-", s)

            st.markdown("---")
            st.subheader("🧠 الخطوة 2: المحاذاة التلقائية (Length-Based DP)")
            st.caption("خوارزمية Gale-Church مبسّطة — تسمح بـ 1-1, 1-2, 2-1, 1-0, 0-1.")

            mode = st.radio(
                "طريقة المحاذاة الأولية:",
                ["⚡ محاذاة ذكية (DP + نِسب الأطوال)", "🔢 محاذاة بسيطة (فهرسة مباشرة)"],
                horizontal=True,
            )

            # نحتفظ بالمحاذاة في session_state لتفادي إعادة الحساب عند كل rerun
            cache_key = hashlib.md5(
                (text_en + "\x00" + text_ar + mode).encode("utf-8")
            ).hexdigest()

            if st.session_state.get("align_cache_key") != cache_key:
                with st.spinner("⏳ يتم إجراء المحاذاة..."):
                    prog = st.progress(0.0)
                    prog.progress(0.4)
                    if mode.startswith("⚡"):
                        auto_pairs = align_by_length(segs_en, segs_ar, max_merge=2)
                    else:
                        maxn = max(len(segs_en), len(segs_ar))
                        e = segs_en + [""] * (maxn - len(segs_en))
                        a = segs_ar + [""] * (maxn - len(segs_ar))
                        auto_pairs = list(zip(e, a))
                    prog.progress(1.0); prog.empty()
                st.session_state["align_cache_key"] = cache_key
                st.session_state["auto_pairs"] = auto_pairs
                st.session_state["edited_pairs_df"] = pd.DataFrame(
                    [{"EN": e, "AR": a} for e, a in auto_pairs]
                )

            df_edit = st.session_state.get("edited_pairs_df")
            if df_edit is None or df_edit.empty:
                df_edit = pd.DataFrame([{"EN": e, "AR": a} for e, a in auto_pairs])

            st.markdown("### ✏️ الخطوة 3: المراجعة اليدوية والتحرير المباشر")
            st.caption(
                "• استخدم ✏️ للتحرير المباشر داخل الخلايا.\n"
                "• ➕ أسفل الجدول لإضافة وحدة جديدة، و 🗑️ لحذف صف.\n"
                "• للدمج/التقسيم: عدّل النص يدوياً ثم أضف/احذف صف."
            )

            edited = st.data_editor(
                df_edit,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "EN": st.column_config.TextColumn("🇬🇧 English (LTR)", width="large"),
                    "AR": st.column_config.TextColumn("🇸🇦 Arabic (RTL)", width="large"),
                },
                key="alignment_editor",
                height=480,
            )

            # ---- زر توليد المخرجات ----
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                if st.button("🚀 اعتماد المحاذاة وتوليد الملفات", use_container_width=True):
                    final_pairs = [
                        (normalize_text(str(r["EN"]).strip()), normalize_text(str(r["AR"]).strip()))
                        for _, r in edited.iterrows()
                        if str(r["EN"]).strip() and str(r["AR"]).strip()
                    ]
                    if not final_pairs:
                        st.error("⚠️ لا يوجد أزواج صالحة.")
                    else:
                        st.session_state["aligned_pairs"] = final_pairs
                        st.session_state["tmx_bytes"]  = generate_tmx_file(final_pairs)
                        df_out = pd.DataFrame(final_pairs, columns=["English", "Arabic"])
                        st.session_state["xlsx_bytes"] = df_to_excel_bytes(df_out)
                        st.session_state["csv_bytes"]  = df_to_csv_bytes(df_out)
                        st.success(f"🎉 تم اعتماد {len(final_pairs)} وحدة محاذاة.")
            with col_b2:
                st.info("💡 بعد الاعتماد، انتقل إلى تبويب **AI Term Extractor** لاستخراج المصطلحات.")

            # ---- أزرار التنزيل ----
            if st.session_state.get("tmx_bytes"):
                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button(
                        "💾 TMX",
                        data=st.session_state["tmx_bytes"],
                        file_name="Aligned_Legal_Memory.tmx",
                        mime="application/x-tmx+xml",
                        use_container_width=True,
                    )
                with d2:
                    st.download_button(
                        "📊 Excel",
                        data=st.session_state["xlsx_bytes"],
                        file_name="Aligned_Pairs.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with d3:
                    st.download_button(
                        "📄 CSV (UTF-8)",
                        data=st.session_state["csv_bytes"],
                        file_name="Aligned_Pairs.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            # === جدول مقارنة نهائي ===
            if st.session_state.get("aligned_pairs"):
                st.markdown("---")
                st.markdown("### 📋 الجدول النهائي للمحاذاة")
                preview = pd.DataFrame(st.session_state["aligned_pairs"], columns=["English (LTR)", "Arabic (RTL)"])
                st.dataframe(preview, use_container_width=True, height=350)


# ---------------------------------------------------------
# TAB 2 : الاستخراج الذكي
# ---------------------------------------------------------
with tab2:
    st.subheader("🔍 استخراج المصطلحات عبر الذكاء الاصطناعي")

    up_tmx = st.file_uploader("ارفع ملف .TMX (اختياري):", type=["tmx"])
    aligned_units = (
        parse_tmx_bytes(up_tmx.read()) if up_tmx is not None
        else st.session_state.get("aligned_pairs", [])
    )

    if aligned_units:
        st.info(f"تم تحميل **{len(aligned_units)}** وحدة ترجمية.")

        if st.button("⚡ بدء الاستخراج الذكي"):
            if not api_key.strip():
                st.error("⚠️ أدخل مفتاح API أولاً.")
            else:
                client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key.strip())

                system_prompt = """You are an expert legal terminologist.
STRICT RULES:
1. Detect exact sub_domain (e.g., "International Public Law") and document_type.
2. Extract ONLY substantive legal concepts (obligations, rights, liabilities, bodies).
3. DO NOT extract preambles, procedural phrases ("Pursuant to", "بعد الاطلاع على"), dates, or article numbers.
4. Output ONLY valid JSON: {"terms":[{"english_term":"","arabic_term":"","sub_domain":"","document_type":"","pos":"","arabic_description":""}]}
"""

                results = []
                prog = st.progress(0.0)
                status = st.empty()

                for i, (en_text, ar_text) in enumerate(aligned_units):
                    status.text(f"تحليل {i+1}/{len(aligned_units)} — {selected_model_id}")
                    user_prompt = f"[EN]: {en_text}\n[AR]: {ar_text}"
                    try:
                        time.sleep(0.6)
                        resp = client.chat.completions.create(
                            model=selected_model_id,
                            messages=[
                                {"role": "system", "content": system_prompt},
                                {"role": "user", "content": user_prompt},
                            ],
                            temperature=0.0,
                            response_format={"type": "json_object"},
                        )
                        raw = json.loads(resp.choices[0].message.content)
                        terms = raw.get("terms", []) if isinstance(raw, dict) else raw
                        if terms and isinstance(terms, list):
                            for item in terms:
                                if not isinstance(item, dict) or not item.get("english_term"):
                                    continue
                                g = lambda k, d="": str(item.get(k) or d).strip()
                                results.append({
                                    "English Term": g("english_term"),
                                    "Arabic Term": g("arabic_term"),
                                    "Sub Domain": g("sub_domain", "قانون عام"),
                                    "Document Type": g("document_type", "نص قانوني"),
                                    "POS": g("pos", "تركيب اسمي"),
                                    "Arabic Description": g("arabic_description"),
                                    "Source EN": en_text,
                                    "Source AR": ar_text,
                                })
                    except Exception as e:
                        st.warning(f"⚠️ وحدة {i+1}: {e}")
                    prog.progress((i + 1) / len(aligned_units))

                status.text("✅ اكتمل الاستخراج.")
                if results:
                    df = pd.DataFrame(results).drop_duplicates(
                        subset=["English Term", "Arabic Term"], keep="first"
                    ).reset_index(drop=True)
                    st.session_state["results_df"] = df
                else:
                    st.error("لم يتم العثور على مصطلحات.")

    # --- عرض النتائج ---
    df_res = st.session_state.get("results_df")
    if df_res is not None and not df_res.empty:
        st.subheader("📊 جدول المصطلحات:")
        st.dataframe(
            df_res[["English Term", "Arabic Term", "Sub Domain", "Document Type", "Arabic Description"]],
            use_container_width=True,
        )

        # تصدير
        xlsx_terms = df_to_excel_bytes(df_res.drop(columns=["POS"], errors="ignore"))
        csv_terms = df_to_csv_bytes(df_res)
        try:
            pdf_bytes = generate_pdf_cards_bytes(df_res)
        except Exception as e:
            pdf_bytes = None
            st.warning(f"تعذّر توليد PDF: {e}")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("📥 Excel", data=xlsx_terms,
                               file_name="Extracted_Legal_Glossary.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with c2:
            st.download_button("📄 CSV (UTF-8)", data=csv_terms,
                               file_name="Extracted_Legal_Glossary.csv",
                               mime="text/csv", use_container_width=True)
        with c3:
            if pdf_bytes:
                st.download_button("📑 PDF Cards", data=pdf_bytes,
                                   file_name="Legal_Terminology_Cards.pdf",
                                   mime="application/pdf", use_container_width=True)


# =========================================================
# 10) Footer
# =========================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center;background:#131B2E;padding:20px;border-radius:12px;border:1px solid #1E293B;">
  <h3 style="color:#EAB308;margin-bottom:10px;">📫 تواصل معي</h3>
  <p style="font-size:15px;color:#E5E7EB;"><b>إعداد وتطوير:</b> يوسف نورالدين</p>
  <p style="font-size:15px;color:#E5E7EB;">
    📱 <a href="tel:0928756918" style="color:#EAB308;text-decoration:none;">0928756918</a> |
    ✉️ <a href="mailto:ywsfalmqmz@gmail.com" style="color:#EAB308;text-decoration:none;">ywsfalmqmz@gmail.com</a>
  </p>
  <hr style="margin:15px auto;width:40%;border-top:1px solid #1E293B;">
  <p style="font-size:12px;color:#9CA3AF;">جميع الحقوق محفوظة © AI Legal Workspace</p>
</div>
""", unsafe_allow_html=True)

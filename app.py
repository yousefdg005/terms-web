import base64
import io
import json
import os
import re
import time
import xml.etree.ElementTree as ET

import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document
import fitz  # PyMuPDF
from openai import OpenAI
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st
import streamlit.components.v1 as components

# =========================================================
# 1. إعدادات صفحة Streamlit
# =========================================================
st.set_page_config(
    page_title="General Legal Terminology Platform",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# 2. تحميل ملف CSS
# =========================================================
def load_css(file_name="style.css"):
  try:
    with open(file_name, "r", encoding="utf-8") as f:
      st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
  except FileNotFoundError:
    pass


load_css("style.css")


# =========================================================
# 3. دالة معالجة النص العربي للـ PDF (ReportLab)
# =========================================================
def prepare_arabic_text(text):
  if not text or not isinstance(text, str):
    return ""
  reshaped_text = arabic_reshaper.reshape(text)
  bidi_text = get_display(reshaped_text)
  return bidi_text


# =========================================================
# 4. الهيدر العام للمنصة (Neutral Branding)
# =========================================================
st.markdown(
    """
<div class="header-container">
    <div class="header-text">
        <h1 style="color: #EAB308; margin: 0; font-size: 26px; font-weight: 800;">⚖️ المنصة العامة للتحليل والاستخراج القانوني وتوليد البطاقات</h1>
        <p style="color: #F3F4F6; font-size: 15px; margin-top: 8px; margin-bottom: 12px;">نظام ذكي متكامل لمحاذاة النصوص القانونية، استخراج المصطلحات الموضوعية، وتوليد بطاقات PDF ومسارد الموارد المعتمدة</p>
        <span style="color: #E5E7EB; font-size: 13px; background: rgba(234, 179, 8, 0.15); padding: 6px 18px; border-radius: 20px; border: 1px solid rgba(234, 179, 8, 0.4);">
            💻 نظام المعالجة والتحليل الذكي | إعداد وتطوير: <b>يوسف نورالدين</b>
        </span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# 5. الشريط الجانبي للإعدادات
# =========================================================
st.sidebar.header("⚙️ إعدادات النظام")

st.sidebar.markdown(
    """
<div style="background-color: #1E293B; padding: 12px; border-radius: 8px; border: 1px solid #374151; margin-bottom: 15px; text-align: center;">
    <span style="color: #E5E7EB; font-size: 13px;">ليس لديك مفتاح API؟</span><br>
    <a href="https://openrouter.ai/keys" target="_blank" style="color: #EAB308; font-weight: bold; text-decoration: none; font-size: 14px; display: inline-block; margin-top: 5px;">
        🔗 الحصول على المفتاح من OpenRouter
    </a>
</div>
""",
    unsafe_allow_html=True,
)

api_key = st.sidebar.text_input(
    "🔑 مفتاح OpenRouter API Key الخاص بك:",
    value="",
    type="password",
    placeholder="sk-or-v1-...",
)

PAID_MODELS_MAPPING = {
    "🌙 Qwen 3.5 Flash (الأقوى في معالجة اللغة العربية)": (
        "qwen/qwen3.5-flash-02-23"
    ),
    "📜 Claude Haiku 5.5 (التزام صارم بتعليمات التصفية)": (
        "anthropic/claude-haiku-5.5"
    ),
    "⚡ Gemini 3.8 Flash (معالجة السياقات الطويلة)": "google/gemini-3.8-flash",
    "⚖️ GPT-4o Mini (استخراج هيكلي دقيق)": "openai/gpt-4o-mini",
    "🦙 Llama 4 Maverick (نموذج مفتوح متقدم)": (
        "meta-llama/llama-4-maverick"
    ),
}

st.sidebar.subheader("🤖 اختيار النموذج المراد استعماله")
selected_model_label = st.sidebar.selectbox(
    "اختر النموذج لتنفيذ عملية التصفية والاستخراج:",
    options=list(PAID_MODELS_MAPPING.keys()),
    index=0,
)
selected_model_id = PAID_MODELS_MAPPING[selected_model_label]


# =========================================================
# 6. إدارة حالة الجلسة (Session State)
# =========================================================
if "aligned_pairs" not in st.session_state:
  st.session_state["aligned_pairs"] = []
if "results_df" not in st.session_state:
  st.session_state["results_df"] = None
if "tmx_bytes" not in st.session_state:
  st.session_state["tmx_bytes"] = None


# =========================================================
# 7. دوال تحويل الـ PDF وقراءة الـ TMX
# =========================================================
MIN_PAGE_CHARS = 25
OCR_DPI = 300

_INVISIBLE_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]")
_XML_BAD_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def normalize_text(text):
  text = unicodedata.normalize("NFKC", text)
  text = _INVISIBLE_RE.sub("", text)
  return text.replace("\u00a0", " ")


def _ocr_page(page):
  pix = page.get_pixmap(dpi=OCR_DPI, colorspace=fitz.csGRAY)
  img = Image.frombytes("L", [pix.width, pix.height], pix.samples)
  return pytesseract.image_to_string(img, lang="ara+eng", config="--oem 1 --psm 6")


@st.cache_data(show_spinner="⏳ جاري استخراج النص من الملف...")
def extract_text_from_pdf_bytes(data):
  pages_text, notes = [], []
  doc = fitz.open(stream=data, filetype="pdf")
  plumber = None
  try:
    plumber = pdfplumber.open(io.BytesIO(data))
  except Exception as e:
    notes.append(f"pdfplumber: {e}")

  for i, page in enumerate(doc):
    txt = ""
    try:
      txt = page.get_text("text", sort=True) or ""
    except Exception as e:
      notes.append(f"صفحة {i+1} (PyMuPDF): {e}")

    if len(txt.strip()) < MIN_PAGE_CHARS and plumber is not None:
      try:
        txt = plumber.pages[i].extract_text() or ""
      except Exception as e:
        notes.append(f"صفحة {i+1} (pdfplumber): {e}")

    if len(txt.strip()) < MIN_PAGE_CHARS:
      try:
        txt = _ocr_page(page)
        notes.append(f"صفحة {i+1}: قُرئت بالـ OCR")
      except Exception as e:
        notes.append(f"صفحة {i+1} (OCR): {e}")

    if txt.strip():
      pages_text.append(normalize_text(txt))

  if plumber is not None:
    plumber.close()
  doc.close()
  return "\n".join(pages_text), notes


def extract_text_from_pdf(pdf_file):
  pdf_file.seek(0)
  return extract_text_from_pdf_bytes(pdf_file.read())


_NEW_UNIT_RE = re.compile(
    r"^\s*(?:"
    r"(?:المادة|مادة|الباب|الفصل|الفرع|القسم)\b"
    r"|(?:article|section|chapter|part|clause)\b"
    r"|\(?\d{1,3}[\).\-–]\s"
    r"|\(?[a-zA-Z]\)\s"
    r"|[أ-ي][\).\-–]\s"
    r")",
    re.IGNORECASE,
)
_PAGE_NOISE_RE = re.compile(
    r"^\s*(?:[-–—]?\s*\d{1,4}\s*[-–—]?|page\s+\d+(?:\s+of\s+\d+)?|صفحة\s+\d+(?:\s+من\s+\d+)?)\s*$",
    re.IGNORECASE,
)
_TERMINAL = (".", "؛", ";", ":", "؟", "?", "!", "。")


def process_pdf_text_to_segments(text):
  segments, buf = [], ""

  def flush():
    nonlocal buf
    if buf.strip():
      segments.append(re.sub(r"\s+", " ", buf).strip())
    buf = ""

  for raw in text.split("\n"):
    line = raw.strip()
    if not line or _PAGE_NOISE_RE.match(line):
      continue
    if buf and (_NEW_UNIT_RE.match(line) or buf.rstrip().endswith(_TERMINAL)):
      flush()
    if buf.endswith("-") and line[:1].islower():
      buf = buf[:-1] + line
    else:
      buf = f"{buf} {line}" if buf else line
  flush()
  return segments


def create_docx_from_text(text):
  doc = Document()
  for paragraph in text.split("\n"):
    if paragraph.strip():
      doc.add_paragraph(paragraph.strip())
  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


def parse_tmx_bytes(tmx_content):
  tree = ET.ElementTree(ET.fromstring(tmx_content))
  root = tree.getroot()
  pairs = []
  for tu in root.iter():
    if tu.tag.endswith("tu"):
      en_text, ar_text = "", ""
      for tuv in tu.iter():
        if tuv.tag.endswith("tuv"):
          lang = ""
          for attr_k, attr_v in tuv.attrib.items():
            if "lang" in attr_k.lower():
              lang = attr_v.lower()
              break
          seg_text = ""
          for elem in tuv.iter():
            if elem.tag.endswith("seg"):
              seg_text = "".join(elem.itertext()).strip()
              break
          if lang.startswith("en"):
            en_text = seg_text
          elif lang.startswith("ar"):
            ar_text = seg_text
      if en_text and ar_text:
        pairs.append((en_text, ar_text))
  return pairs


def _xml_safe(text):
  return xml_escape(_XML_BAD_RE.sub("", str(text)))


def generate_tmx_file(pairs):
  tmx_xml = """<?xml version="1.0" encoding="UTF-8"?>
<tmx version="1.4">
  <header creationtool="General Legal Terminology Platform" creationtoolversion="2.0" segtype="sentence" adminlang="en-us" srclang="en" datatype="plaintext"/>
  <body>
"""
  for en, ar in pairs:
    tmx_xml += f"""    <tu>
      <tuv xml:lang="en"><seg>{_xml_safe(en)}</seg></tuv>
      <tuv xml:lang="ar"><seg>{_xml_safe(ar)}</seg></tuv>
    </tu>\n"""
  tmx_xml += "  </body>\n</tmx>"
  return tmx_xml.encode("utf-8")


# =========================================================
# 8. توليد بطاقات الـ PDF
# =========================================================
_FONT_CANDIDATES = [
    ("fonts/Amiri-Regular.ttf", "fonts/Amiri-Bold.ttf"),
    ("fonts/NotoNaskhArabic-Regular.ttf", "fonts/NotoNaskhArabic-Bold.ttf"),
    (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ),
    (
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf",
    ),
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/tahomabd.ttf"),
    ("/Library/Fonts/Arial Unicode.ttf", "/Library/Fonts/Arial Unicode.ttf"),
    (
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ),
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
        pdfmetrics.registerFont(
            TTFont("AppFont-Bold", bold if os.path.exists(bold) else reg)
        )
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
    if (
        cur
        and pdfmetrics.stringWidth(prepare_arabic_text(cand), font, size)
        > max_width
    ):
      lines.append(cur)
      cur = word
    else:
      cur = cand
  if cur:
    lines.append(cur)
  return lines or [""]


def ar_par(text, style, max_width):
  lines = _wrap_logical(text, style.fontName, style.fontSize, max_width)
  html = "<br/>".join(html_escape(prepare_arabic_text(l)) for l in lines)
  return Paragraph(html, style)


def en_par(text, style):
  return Paragraph(html_escape(str(text)), style)


def generate_pdf_cards_bytes(df):
  font, font_bold = get_pdf_fonts()
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      rightMargin=30,
      leftMargin=30,
      topMargin=30,
      bottomMargin=30,
  )
  story = []
  styles = getSampleStyleSheet()

  title_style = ParagraphStyle(
      "TitleStyle",
      parent=styles["Heading1"],
      fontName=font_bold,
      fontSize=18,
      alignment=1,
      textColor=colors.HexColor("#0F172A"),
      leading=26,
  )
  sub_style = ParagraphStyle(
      "SubStyle",
      parent=styles["Normal"],
      fontName=font,
      fontSize=10,
      alignment=1,
      textColor=colors.HexColor("#64748B"),
      leading=14,
  )
  cell_title_style = ParagraphStyle(
      "CellTitle",
      fontName=font_bold,
      fontSize=10,
      alignment=2,
      leading=15,
      textColor=colors.HexColor("#1E293B"),
  )
  cell_val_style = ParagraphStyle(
      "CellVal",
      fontName=font,
      fontSize=10,
      alignment=2,
      leading=15,
      textColor=colors.HexColor("#0F172A"),
  )
  cell_val_en_style = ParagraphStyle("CellValEn", parent=cell_val_style, alignment=0)

  VAL_W, LBL_W, PAD = 360, 160, 6
  val_text_w = VAL_W - 2 * PAD - 4
  lbl_text_w = LBL_W - 2 * PAD - 4

  story.append(
      ar_par(
          "بطاقات المصطلحات القانونية الموثقة (Legal Terminology Cards)",
          title_style,
          520,
      )
  )
  story.append(
      ar_par(
          "منصة التحليل والاستخراج القانوني الذكي | إعداد: يوسف نورالدين",
          sub_style,
          520,
      )
  )
  story.append(Spacer(1, 15))

  total = len(df)
  for n, (_, row) in enumerate(df.iterrows(), start=1):
    story.append(
        ar_par(f"بطاقة مصطلح رقم: ({n:02d} / {total:02d})", cell_title_style, 520)
    )
    story.append(Spacer(1, 5))

    def lbl(t):
      return ar_par(t, cell_title_style, lbl_text_w)

    def ar_val(key, default=""):
      return ar_par(
          str(row.get(key, default) or default), cell_val_style, val_text_w
      )

    data = [
        [
            en_par(row.get("English Term", ""), cell_val_en_style),
            lbl("المصطلح بالإنجليزي (English Term)"),
        ],
        [ar_val("Arabic Term"), lbl("المقابل بالعربي (Arabic Term)")],
        [ar_val("Sub Domain", "قانون عام"), lbl("المجال الفرعي (Sub-Domain)")],
        [ar_val("Document Type", "نص قانوني عام"), lbl("نوع المستند (Document Type)")],
        [
            en_par(row.get("Source EN", ""), cell_val_en_style),
            lbl("السياق الإنجليزي (Context EN)"),
        ],
        [ar_val("Source AR"), lbl("السياق العربي (Context AR)")],
        [ar_val("Arabic Description"), lbl("الشرح والتعريف (Legal Description)")],
    ]

    t = Table(data, colWidths=[VAL_W, LBL_W])
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("PADDING", (0, 0), (-1, -1), PAD),
        ])
    )
    story.append(t)
    story.append(Spacer(1, 15))
    story.append(
        HRFlowable(
            width="100%",
            thickness=1,
            color=colors.HexColor("#E2E8F0"),
            spaceAfter=15,
        )
    )

  doc.build(story)
  buffer.seek(0)
  return buffer


# =========================================================
# 9. علامات التبويب الرئيسية (Tabs)
# =========================================================
tab1, tab2 = st.tabs(
    ["🟢 استوديو التحويل والمحاذاة (Alignment Studio)", "🔵 الاستخراج الذكي (Terms Extractor)"]
)


# =========================================================
# مكوّن استوديو المحاذاة (مع شريط تمرير داخلي مخصص للوحدات)
# =========================================================
ALIGNMENT_STUDIO_HTML = r"""<!DOCTYPE html>
<html lang="ar"><head><meta charset="utf-8">
<style>
  * { box-sizing: border-box; }
  html, body { height: 100%; margin: 0; padding: 0; background: #0B0F19; color: #E5E7EB; font-family: 'Cairo','Segoe UI',Tahoma,Arial,sans-serif; overflow: hidden; }
  
  .studio-wrapper { display: flex; flex-direction: column; height: 100vh; padding: 10px; }
  
  #bar { display: flex; gap: 10px; align-items: center; margin-bottom: 10px; direction: rtl; flex-shrink: 0; }
  #toast { color: #FDE047; font-size: 13px; }
  
  /* صندوق الحاوي مع شريط تمرير خاص داخلي ولن يطيل الصفحة العامة */
  .alignment-container { 
    display: flex; flex-direction: column; gap: 12px; direction: rtl; 
    flex-grow: 1; overflow-y: auto; padding: 5px 10px 20px 10px;
    border: 1px solid #1F2937; border-radius: 12px; background: #07090E;
  }
  
  /* تخصيص شريط التمرير الداخلي ليكون أنيقاً */
  .alignment-container::-webkit-scrollbar { width: 8px; }
  .alignment-container::-webkit-scrollbar-track { background: #0B0F19; border-radius: 4px; }
  .alignment-container::-webkit-scrollbar-thumb { background: #374151; border-radius: 4px; }
  .alignment-container::-webkit-scrollbar-thumb:hover { background: #EAB308; }

  .segment-row { display: flex; gap: 12px; background: #111827; border: 1px solid #374151; border-radius: 10px; padding: 12px; align-items: center; flex-shrink: 0; }
  .segment-row:hover { border-color: #4B5563; }
  .segment-num { color: #EAB308; font-weight: bold; width: 45px; text-align: center; font-size: 14px; flex-shrink: 0; }
  .cell { flex: 1; display: flex; gap: 6px; align-items: stretch; min-width: 0; border: 2px solid transparent; border-radius: 8px; }
  .cell.drag-over { border: 2px dashed #EAB308; background: rgba(234,179,8,.15); }
  .handle { width: 18px; flex-shrink: 0; display: flex; align-items: center; justify-content: center; background: #1F2937; border-radius: 6px; color: #9CA3AF; cursor: grab; user-select: none; font-size: 14px; }
  .handle:active { cursor: grabbing; color: #EAB308; }
  .card-box { flex: 1; min-width: 0; background: #1E293B; border: 1px solid #475569; border-radius: 8px; padding: 10px 14px; min-height: 55px; font-size: 14px; line-height: 1.7; white-space: pre-wrap; overflow-wrap: anywhere; cursor: text; }
  .card-box:focus { outline: none; border-color: #EAB308; background: #0F172A; }
  .en-box { direction: ltr; text-align: left; color: #93C5FD; }
  .ar-box { direction: rtl; text-align: right; color: #FDE047; }
  .action-btn-group { display: flex; flex-direction: column; gap: 5px; flex-shrink: 0; }
  .btn-ctrl { background: #1F2937; color: #E5E7EB; border: 1px solid #374151; border-radius: 6px; padding: 4px 8px; font-size: 11px; font-weight: bold; cursor: pointer; font-family: inherit; }
  .btn-ctrl:hover { background: #374151; color: #EAB308; }
  .btn-del:hover { background: #991B1B; color: #FFF; }
  .add-row-btn { background: linear-gradient(135deg,#EAB308 0%,#CA8A04 100%); color: #0F172A; font-weight: bold; border: none; border-radius: 8px; padding: 8px 16px; cursor: pointer; font-family: inherit; }
</style></head>
<body>
<div class="studio-wrapper">
  <div id="bar">
    <button class="add-row-btn" id="addBtn">➕ إضافة وحدة محاذاة جديدة</button>
    <span id="toast"></span>
  </div>
  <div class="alignment-container" id="alignmentContainer"></div>
</div>
<script>
  const container = document.getElementById('alignmentContainer');
  const toastEl = document.getElementById('toast');
  let data = [], token = null, drag = null, sendTimer = null;

  function post(type, extra) { window.parent.postMessage(Object.assign({isStreamlitMessage:true, type}, extra||{}), '*'); }
  function pushValue() { post('streamlit:setComponentValue', {value: data.map(d => ({en:d.en, ar:d.ar})), dataType:'json'}); }
  function pushSoon() { clearTimeout(sendTimer); sendTimer = setTimeout(pushValue, 500); }
  window.addEventListener('message', (ev) => {
    if (!ev.data || ev.data.type !== 'streamlit:render') return;
    const args = ev.data.args || {};
    if (args.token !== token) {
      token = args.token;
      data = (args.pairs || []).map(p => ({en: p.en || '', ar: p.ar || ''}));
      renderRows();
    }
  });
  post('streamlit:componentReady', {apiVersion: 1});
  post('streamlit:setFrameHeight', {height: 650});

  function toast(msg) { toastEl.textContent = msg; setTimeout(() => { if (toastEl.textContent === msg) toastEl.textContent = ''; }, 2500); }

  function makeCell(type, index) {
    const cell = document.createElement('div');
    cell.className = 'cell'; cell.dataset.index = index; cell.dataset.type = type;
    const handle = document.createElement('div');
    handle.className = 'handle'; handle.textContent = '⠿'; handle.draggable = true; handle.title = 'اسحب للتبديل';
    const box = document.createElement('div');
    box.className = 'card-box ' + (type === 'en' ? 'en-box' : 'ar-box');
    box.dataset.index = index; box.dataset.type = type;
    box.contentEditable = 'plaintext-only';
    if (box.contentEditable !== 'plaintext-only') box.contentEditable = 'true';
    box.textContent = data[index][type];
    box.addEventListener('input', () => { data[index][type] = box.innerText.replace(/\n$/, ''); pushSoon(); });
    box.addEventListener('paste', (e) => {
      if (box.contentEditable === 'plaintext-only') return;
      e.preventDefault();
      document.execCommand('insertText', false, (e.clipboardData || window.clipboardData).getData('text/plain'));
    });
    handle.addEventListener('dragstart', (e) => { drag = {index, type}; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(index)); });
    cell.addEventListener('dragover', (e) => {
      if (!drag || drag.type !== type) return;
      e.preventDefault(); cell.classList.add('drag-over');
      const th = 50;
      if (e.clientY < th) container.scrollTop -= 15; else if (window.innerHeight - e.clientY < th) container.scrollTop += 15;
    });
    cell.addEventListener('dragleave', () => cell.classList.remove('drag-over'));
    cell.addEventListener('drop', (e) => {
      e.preventDefault(); cell.classList.remove('drag-over');
      if (!drag || drag.type !== type || drag.index === index) { drag = null; return; }
      const t = data[drag.index][type]; data[drag.index][type] = data[index][type]; data[index][type] = t;
      drag = null; renderRows(); pushValue();
    });
    cell.append(handle, box);
    return cell;
  }

  function renderRows() {
    container.textContent = '';
    data.forEach((_, index) => {
      const row = document.createElement('div'); row.className = 'segment-row';
      const num = document.createElement('div'); num.className = 'segment-num'; num.textContent = '[' + (index + 1) + ']';
      const grp = document.createElement('div'); grp.className = 'action-btn-group';
      const sp = document.createElement('button'); sp.className = 'btn-ctrl'; sp.textContent = '✂️ تقسيم'; sp.title = 'تقسيم عند مكان المؤشر';
      const del = document.createElement('button'); del.className = 'btn-ctrl btn-del'; del.textContent = '🗑️ حذف'; del.title = 'حذف هذه الوحدة';
      [sp, del].forEach(b => b.addEventListener('mousedown', e => e.preventDefault()));
      sp.addEventListener('click', () => splitUnit(index));
      del.addEventListener('click', () => deleteRow(index));
      grp.append(sp, del);
      row.append(num, makeCell('en', index), makeCell('ar', index), grp);
      container.appendChild(row);
    });
  }

  function splitUnit(index) {
    const sel = window.getSelection();
    if (!sel.rangeCount || !sel.anchorNode) { toast('ضع المؤشر داخل النص أولاً'); return; }
    const el = sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement;
    const box = el && el.closest('.card-box');
    if (!box || Number(box.dataset.index) !== index) { toast('ضع المؤشر داخل نص هذه الوحدة ثم اضغط تقسيم'); return; }
    const type = box.dataset.type;
    const pre = document.createRange();
    pre.selectNodeContents(box); pre.setEnd(sel.anchorNode, sel.anchorOffset);
    const full = data[index][type];
    const offset = Math.min(pre.toString().length, full.length);
    const part1 = full.slice(0, offset).trim(), part2 = full.slice(offset).trim();
    if (!part2) { toast('المؤشر في نهاية النص، لا يوجد ما يُقسَّم'); return; }
    data[index][type] = part1;
    const nr = {en: '', ar: ''}; nr[type] = part2;
    data.splice(index + 1, 0, nr);
    renderRows(); pushValue();
  }

  function deleteRow(index) { data.splice(index, 1); renderRows(); pushValue(); }
  document.getElementById('addBtn').addEventListener('click', () => {
    data.push({en: '', ar: ''}); renderRows(); pushValue();
    container.scrollTop = container.scrollHeight;
  });
</script></body></html>"""

_COMPONENT_DIR = Path(__file__).parent / "alignment_studio_component"


def _declare_alignment_studio():
  _COMPONENT_DIR.mkdir(exist_ok=True)
  index = _COMPONENT_DIR / "index.html"
  if not index.exists() or index.read_text(encoding="utf-8") != ALIGNMENT_STUDIO_HTML:
    index.write_text(ALIGNMENT_STUDIO_HTML, encoding="utf-8")
  return components.declare_component("alignment_studio", path=str(_COMPONENT_DIR))


alignment_studio = _declare_alignment_studio()


# ---------------------------------------------------------
# Tab 1: استوديو المحاذاة الشامل الاحترافي (MateCat Studio Full)
# ---------------------------------------------------------
with tab1:
  st.subheader("📄 الخطوة 1: رفع ملفات PDF واستخراج النصوص")

  col_pdf1, col_pdf2 = st.columns(2)
  with col_pdf1:
    pdf_en = st.file_uploader(
        "ارفع ملف PDF المصدر (إنجليزي):", type=["pdf"], key="pdf_en"
    )
  with col_pdf2:
    pdf_ar = st.file_uploader(
        "ارفع ملف PDF الهدف (عربي):", type=["pdf"], key="pdf_ar"
    )

  if pdf_en and pdf_ar:
    text_en, notes_en = extract_text_from_pdf_bytes(pdf_en.getvalue())
    text_ar, notes_ar = extract_text_from_pdf_bytes(pdf_ar.getvalue())

    if text_en.strip() and text_ar.strip():
      st.success("✅ تم استخراج النصوص بنجاح!")
      all_notes = [f"EN — {n}" for n in notes_en] + [f"AR — {n}" for n in notes_ar]
      if all_notes:
        with st.expander("ℹ️ ملاحظات الاستخراج (صفحات قُرئت بالـ OCR أو حدثت فيها أخطاء)"):
          for n in all_notes:
            st.write("•", n)

      segs_en = process_pdf_text_to_segments(text_en)
      segs_ar = process_pdf_text_to_segments(text_ar)

      max_len = max(len(segs_en), len(segs_ar))
      segs_en += [""] * (max_len - len(segs_en))
      segs_ar += [""] * (max_len - len(segs_ar))

      st.markdown("---")
      st.subheader("🧩 الخطوة 2: استوديو المحاذاة التفاعلي (MateCat Alignment Studio)")
      st.caption(
          "💡 **الميزات المتاحة:** شريط تمرير خاص بالوحدات داخل النافذة دون إطالة"
          " الصفحة العامة | اسحب من المقبض ⠿ | زر التقسيم ✂️ | 🗑️ حذف | ➕ إضافة وحدة."
      )

      initial_pairs = [{"en": e, "ar": a} for e, a in zip(segs_en, segs_ar)]
      token = hashlib.md5((text_en + "\x00" + text_ar).encode("utf-8")).hexdigest()

      studio_value = alignment_studio(
          pairs=initial_pairs, token=token, key="alignment_studio", default=None
      )
      current_pairs = studio_value if studio_value is not None else initial_pairs

      st.markdown("<br>", unsafe_allow_html=True)

      if st.button("🚀 اعتماد المحاذاة النهائية وتوليد TMX", use_container_width=True):
        final_pairs = [
            (p["en"].strip(), p["ar"].strip())
            for p in current_pairs
            if p["en"].strip() and p["ar"].strip()
        ]
        if final_pairs:
          st.session_state["aligned_pairs"] = final_pairs
          st.session_state["tmx_bytes"] = generate_tmx_file(final_pairs)
        else:
          st.session_state["tmx_bytes"] = None
          st.error("⚠️ يرجى التأكد من وجود جمل صالحة للتقابل.")

      if st.session_state.get("tmx_bytes"):
        st.success(f"🎉 تم اعتماد **{len(st.session_state['aligned_pairs'])}** وحدة محاذاة بنجاح!")
        st.download_button(
            "💾 تنزيل ملف الـ TMX الموحد",
            data=st.session_state["tmx_bytes"],
            file_name="Aligned_Legal_Memory.tmx",
            mime="application/x-tmx+xml",
        )
        st.info("👉 **انتقل الآن إلى 'التبويب الثاني (الاستخراج الذكي)'**!")
    else:
      st.error("⚠️ تعذر استخراج النصوص من ملفات الـ PDF.")


# ---------------------------------------------------------
# Tab 2: الاستخراج الذكي وتوليد البطاقات
# ---------------------------------------------------------
with tab2:
  st.subheader("🔍 استخراج المصطلحات عبر الذكاء الاصطناعي")

  uploaded_tmx = st.file_uploader(
      "ارفع ملف الذاكرة (.tmx) المحاذى جاهزاً (أو استخدم المحاذاة من التبويب"
      " الأول):",
      type=["tmx"],
  )

  if uploaded_tmx is not None:
    aligned_units = parse_tmx_bytes(uploaded_tmx.read())
  else:
    aligned_units = st.session_state["aligned_pairs"]

  if aligned_units:
    st.info(f"تم تحميل **{len(aligned_units)}** وحدة ترجمية جاهزة للاستخراج.")

    if st.button("⚡ بدء الاستخراج الذكي والتصفية الصارمة"):
      if not api_key.strip():
        st.error("⚠️ يرجى إدخال مفتاح OpenRouter API Key أولاً.")
      else:
        client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=api_key.strip(),
        )

        system_prompt = """You are an expert legal terminologist specializing in general legal systems, international law, contracts, and legislations.

STRICT FILTERING & AUTOMATIC CONTEXT DETECTION RULES:

1. AUTOMATIC CONTEXT & SUB-DOMAIN DETECTION:
   - Identify the exact "sub_domain" (e.g., "International Public Law", "Contractual Obligations", "Commercial Law", "Maritime Law").
   - Identify the "document_type" (e.g., "UN Convention", "Legislative Act", "Commercial Agreement").

2. POSITIVE SCOPE:
   - Extract ONLY substantive legal concepts, legal obligations, rights, liabilities, statutory bodies, and judicial principles.

3. STRICT NEGATIVE DIRECTIVES (قواعد الاستبعاد الصارمة):
   - DO NOT extract preamble, procedural, or introductory phrases (e.g., "Having reviewed", "Pursuant to", "In accordance with", "بعد الاطلاع على", "بناء على").
   - DO NOT extract dates, article numbers, or general administrative filler words (e.g., "Decision", "Date", "Law", "Article").

4. OUTPUT FORMAT:
   - Output MUST be valid JSON as an object with key "terms", containing an array of objects:
     {"english_term": "", "arabic_term": "", "sub_domain": "", "document_type": "", "pos": "", "arabic_description": ""}
"""

        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, (en_text, ar_text) in enumerate(aligned_units):
          status_text.text(
              f"جاري تحليل الوحدة ({i+1}/{len(aligned_units)})... عبر نموذج"
              f" `{selected_model_id}`"
          )
          user_prompt = f"[EN]: {en_text}\n[AR]: {ar_text}"

          try:
            time.sleep(1.0)
            response = client.chat.completions.create(
                model=selected_model_id,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.0,
                response_format={"type": "json_object"},
            )

            raw = json.loads(response.choices[0].message.content)
            terms = raw.get("terms", []) if isinstance(raw, dict) else raw

            def _s(item, key, default=""):
              return str(item.get(key) or default).strip()

            if terms and isinstance(terms, list):
              for item in terms:
                if isinstance(item, dict) and item.get("english_term"):
                  results.append({
                      "English Term": _s(item, "english_term"),
                      "Arabic Term": _s(item, "arabic_term"),
                      "Sub Domain": _s(item, "sub_domain", "قانون عام"),
                      "Document Type": _s(item, "document_type", "نص قانوني"),
                      "POS": _s(item, "pos", "تركيب اسمي"),
                      "Arabic Description": _s(item, "arabic_description"),
                      "Source EN": en_text,
                      "Source AR": ar_text,
                  })
          except Exception as e:
            st.warning(f"تنبيه في الوحدة {i+1}: {e}")

          progress_bar.progress((i + 1) / len(aligned_units))

        status_text.text("✅ اكتمل الاستخراج والتصفية بنجاح!")

        if results:
          df = pd.DataFrame(results)
          df.drop_duplicates(
              subset=["English Term", "Arabic Term"], inplace=True
          )
          st.session_state["results_df"] = df
        else:
          st.error("لم يتم العثور على مصطلحات موضوعية بعد التصفية.")

  if st.session_state["results_df"] is not None:
    df_res = st.session_state["results_df"]

    st.subheader("📊 جدول المصطلحات المستخرجة والمصنفة:")
    st.dataframe(
        df_res[[
            "English Term",
            "Arabic Term",
            "Sub Domain",
            "Document Type",
            "Arabic Description",
        ]],
        use_container_width=True,
    )

    excel_buffer = io.BytesIO()
    excel_df = df_res.drop(columns=["POS"], errors="ignore")
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
      excel_df.to_excel(writer, index=False)
    excel_buffer.seek(0)

    try:
      pdf_buffer = generate_pdf_cards_bytes(df_res)
    except Exception as e:
      pdf_buffer = None
      st.warning(f"تعذر توليد بطاقات PDF: {e}")

    col_d1, col_d2 = st.columns(2)
    with col_d1:
      st.download_button(
          label="📥 تنزيل مسرد المصطلحات (Excel)",
          data=excel_buffer,
          file_name="Extracted_Legal_Glossary.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    with col_d2:
      if pdf_buffer is not None:
        st.download_button(
            label="📑 تنزيل البطاقات الترجمية (PDF Cards)",
            data=pdf_buffer,
            file_name="Legal_Terminology_Cards.pdf",
            mime="application/pdf",
        )

# =========================================================
# 10. التذييل بأسفل الصفحة
# =========================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    """
<div style="text-align: center; background-color: #131B2E; padding: 20px; border-radius: 12px; border: 1px solid #1E293B;">
    <h3 style="color: #EAB308; margin-bottom: 10px;">📫 تواصل معي (Contact & Support)</h3>
    <p style="font-size: 15px; color: #E5E7EB; margin: 5px;"><b>إعداد وتطوير:</b> يوسف نورالدين</p>
    <p style="font-size: 15px; color: #E5E7EB; margin: 5px;">📱 <b>الهاتف:</b> <a href="tel:0928756918" style="color: #EAB308; text-decoration: none;">0928756918</a> | ✉️ <b>البريد:</b> <a href="mailto:ywsfalmqmz@gmail.com" style="color: #EAB308; text-decoration: none;">ywsfalmqmz@gmail.com</a></p>
    <hr style="margin: 15px auto; width: 40%; border-top: 1px solid #1E293B;">
    <p style="font-size: 12px; color: #9CA3AF;">جميع الحقوق محفوظة © المنصة العامة للتحليل والاستخراج القانوني وتوليد البطاقات</p>
</div>
""",
    unsafe_allow_html=True,
)

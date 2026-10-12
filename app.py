import base64
import io
import json
import os
import re
from string import Template
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
# 3. دالة معالجة النص العربي للـ PDF
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
# 6. إدارة حالة الجلسة
# =========================================================
if "aligned_pairs" not in st.session_state:
  st.session_state["aligned_pairs"] = []
if "results_df" not in st.session_state:
  st.session_state["results_df"] = None


# =========================================================
# 7. دوال تحويل الـ PDF وقراءة الـ TMX وتجزئة الفقرات
# =========================================================
def extract_text_from_pdf(pdf_file):
  text = ""
  try:
    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
      for page in pdf.pages:
        extracted = page.extract_text()
        if extracted and extracted.strip():
          text += extracted + "\n"

    if not text.strip():
      pdf_file.seek(0)
      doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
      for page in doc:
        pix = page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_text = pytesseract.image_to_string(img, lang="ara+eng")
        if ocr_text and ocr_text.strip():
          text += ocr_text + "\n"

  except Exception as e:
    print(f"Error in hybrid PDF extraction: {e}")

  return text


def process_pdf_text_to_segments(raw_text):
  if not raw_text:
    return []
  cleaned = re.sub(r"(?<!\n)\n(?!\n)", " ", raw_text)
  raw_segments = re.split(
      r"\n+|\n?•\s*|\n?(?=مادة\s+\d+|Article\s+\d+|Resolution\s+No\.|قرار\s+رقم)",
      cleaned,
  )
  return [seg.strip() for seg in raw_segments if seg.strip()]


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


def generate_tmx_file(pairs):
  tmx_xml = """<?xml version="1.0" encoding="UTF-8"?>
<tmx version="1.4">
  <header creationtool="General Legal Terminology Platform" creationtoolversion="2.0" segtype="sentence" adminlang="en-us" srclang="en" datatype="plaintext"/>
  <body>
"""
  for en, ar in pairs:
    tmx_xml += f"""    <tu>
      <tuv xml:lang="en"><seg>{en}</seg></tuv>
      <tuv xml:lang="ar"><seg>{ar}</seg></tuv>
    </tu>\n"""
  tmx_xml += "  </body>\n</tmx>"
  return tmx_xml.encode("utf-8")


# =========================================================
# 8. دالة توليد بطاقات الـ PDF الفاخرة
# =========================================================
def generate_pdf_cards_bytes(df):
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
      fontName="Helvetica-Bold",
      fontSize=18,
      alignment=1,
      textColor=colors.HexColor("#0F172A"),
  )
  sub_style = ParagraphStyle(
      "SubStyle",
      parent=styles["Normal"],
      fontName="Helvetica",
      fontSize=10,
      alignment=1,
      textColor=colors.HexColor("#64748B"),
  )
  cell_title_style = ParagraphStyle(
      "CellTitle",
      fontName="Helvetica-Bold",
      fontSize=10,
      alignment=2,
      textColor=colors.HexColor("#1E293B"),
  )
  cell_val_style = ParagraphStyle(
      "CellVal",
      fontName="Helvetica",
      fontSize=10,
      alignment=2,
      textColor=colors.HexColor("#0F172A"),
  )

  story.append(
      Paragraph(
          prepare_arabic_text(
              "بطاقات المصطلحات القانونية الموثقة (Legal Terminology Cards)"
          ),
          title_style,
      )
  )
  story.append(
      Paragraph(
          prepare_arabic_text(
              "منصة التحليل والاستخراج القانوني الذكي | إعداد: يوسف نورالدين"
          ),
          sub_style,
      )
  )
  story.append(Spacer(1, 15))

  for idx, row in df.iterrows():
    card_title = Paragraph(
        prepare_arabic_text(f"بطاقة مصطلح رقم: ({idx+1:02d} / {len(df)})"),
        cell_title_style,
    )
    story.append(card_title)
    story.append(Spacer(1, 5))

    data = [
        [
            Paragraph(str(row.get("English Term", "")), cell_val_style),
            Paragraph(
                prepare_arabic_text("المصطلح بالإنجليزي (English Term)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(
                prepare_arabic_text(str(row.get("Arabic Term", ""))),
                cell_val_style,
            ),
            Paragraph(
                prepare_arabic_text("المقابل بالعربي (Arabic Term)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(
                prepare_arabic_text(str(row.get("Sub Domain", "قانون عام"))),
                cell_val_style,
            ),
            Paragraph(
                prepare_arabic_text("المجال الفرعي (Sub-Domain)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(
                prepare_arabic_text(
                    str(row.get("Document Type", "نص قانوني عام"))
                ),
                cell_val_style,
            ),
            Paragraph(
                prepare_arabic_text("نوع المستند (Document Type)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(str(row.get("Source EN", "")), cell_val_style),
            Paragraph(
                prepare_arabic_text("السياق الإنجليزي (Context EN)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(
                prepare_arabic_text(str(row.get("Source AR", ""))),
                cell_val_style,
            ),
            Paragraph(
                prepare_arabic_text("السياق العربي (Context AR)"),
                cell_title_style,
            ),
        ],
        [
            Paragraph(
                prepare_arabic_text(str(row.get("Arabic Description", ""))),
                cell_val_style,
            ),
            Paragraph(
                prepare_arabic_text("الشرح والتعريف (Legal Description)"),
                cell_title_style,
            ),
        ],
    ]

    t = Table(data, colWidths=[360, 160])
    t.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("PADDING", (0, 0), (-1, -1), 6),
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
    text_en = extract_text_from_pdf(pdf_en)
    text_ar = extract_text_from_pdf(pdf_ar)

    if text_en.strip() and text_ar.strip():
      st.success("✅ تم استخراج النصوص بنجاح!")

      lines_en = process_pdf_text_to_segments(text_en)
      lines_ar = process_pdf_text_to_segments(text_ar)

      max_len = max(len(lines_en), len(lines_ar))
      while len(lines_en) < max_len:
        lines_en.append("")
      while len(lines_ar) < max_len:
        lines_ar.append("")

      st.markdown("---")
      st.subheader(
          "🧩 الخطوة 2: استوديو المحاذاة التفاعلي (MateCat Alignment Studio)"
      )
      st.caption(
          "💡 **الميزات المتاحة:** السحب والإفلات الحر مع التمرير التلقائي"
          " للشاشة | زر التقسيم ✂️ عند مؤشر الكتابة | زر الحذف 🗑️ | زر إضافة"
          " وحدة جديدة."
      )

      pairs_json = json.dumps(
          [{"en": en, "ar": ar} for en, ar in zip(lines_en, lines_ar)]
      )

      # أسلوب محمي عبر Template لتجنب SyntaxError الخاص بالـ f-string
      html_template = Template("""
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@600;700&display=swap');
                * { box-sizing: border-box; }
                body { font-family: 'Cairo', sans-serif; background-color: #0B0F19; color: #E5E7EB; margin: 0; padding: 10px; }
                
                .alignment-container { display: flex; flex-direction: column; gap: 12px; direction: rtl; padding-bottom: 50px; }
                
                .segment-row { 
                    display: flex; gap: 12px; background: #111827; border: 1px solid #374151; 
                    border-radius: 10px; padding: 12px; align-items: center; position: relative;
                }
                .segment-row:hover { border-color: #4B5563; }
                
                .segment-num { color: #EAB308; font-weight: bold; width: 45px; text-align: center; font-size: 14px; flex-shrink: 0; }
                
                .card-box { 
                    flex: 1; background: #1E293B; border: 1px solid #475569; border-radius: 8px; 
                    padding: 10px 14px; min-height: 55px; cursor: grab; transition: all 0.2s ease;
                }
                .card-box:active { cursor: grabbing; border-color: #EAB308; box-shadow: 0 0 10px rgba(234, 179, 8, 0.3); }
                .card-box[contenteditable="true"]:focus { outline: none; border-color: #EAB308; background: #0F172A; }
                
                .en-box { direction: ltr; text-align: left; color: #93C5FD; font-size: 14px; }
                .ar-box { direction: rtl; text-align: right; color: #FDE047; font-size: 14px; }
                
                .drag-over { border: 2px dashed #EAB308 !important; background: rgba(234, 179, 8, 0.15) !important; }
                
                .action-btn-group { display: flex; flex-direction: column; gap: 5px; flex-shrink: 0; }
                .btn-ctrl { 
                    background: #1F2937; color: #E5E7EB; border: 1px solid #374151; border-radius: 6px; 
                    padding: 4px 8px; font-size: 11px; font-weight: bold; cursor: pointer; transition: 0.2s; 
                }
                .btn-ctrl:hover { background: #374151; color: #EAB308; }
                .btn-del:hover { background: #991B1B; color: #FFF; }
                
                .add-row-btn {
                    background: linear-gradient(135deg, #EAB308 0%, #CA8A04 100%);
                    color: #0F172A; font-weight: bold; border: none; border-radius: 8px;
                    padding: 10px 20px; cursor: pointer; font-family: 'Cairo', sans-serif;
                    margin-bottom: 15px; display: inline-block;
                }
            </style>

            <button class="add-row-btn" onclick="addNewRow()">➕ إضافة وحدة محاذاة جديدة</button>
            <div class="alignment-container" id="alignmentContainer"></div>

            <script>
                let data = $PAIRS_JSON;
                const container = document.getElementById('alignmentContainer');

                function renderRows() {
                    container.innerHTML = '';
                    data.forEach((item, index) => {
                        const row = document.createElement('div');
                        row.className = 'segment-row';

                        row.innerHTML = `
                            <div class="segment-num">[${index + 1}]</div>
                            <div class="card-box en-box" draggable="true" data-index="${index}" data-type="en" contenteditable="true">${item.en}</div>
                            <div class="card-box ar-box" draggable="true" data-index="${index}" data-type="ar" contenteditable="true">${item.ar}</div>
                            <div class="action-btn-group">
                                <button class="btn-ctrl" onclick="splitUnit(${index})" title="تقسيم عند مكان المؤشر">✂️ تقسيم</button>
                                <button class="btn-ctrl btn-del" onclick="deleteRow(${index})" title="حذف هذه الوحدة">🗑️ حذف</button>
                            </div>
                        `;
                        container.appendChild(row);
                    });

                    addDragAndEditListeners();
                }

                let draggedIndex = null;
                let draggedType = null;

                function addDragAndEditListeners() {
                    document.querySelectorAll('.card-box').forEach(box => {
                        box.addEventListener('dragstart', (e) => {
                            draggedIndex = e.target.getAttribute('data-index');
                            draggedType = e.target.getAttribute('data-type');
                            e.dataTransfer.setData('text/plain', e.target.innerText);
                        });

                        box.addEventListener('dragover', (e) => {
                            e.preventDefault();
                            if (e.target.getAttribute('data-type') === draggedType) {
                                e.target.classList.add('drag-over');
                            }
                            
                            const scrollThreshold = 80;
                            if (e.clientY < scrollThreshold) {
                                window.scrollBy(0, -15);
                            } else if (window.innerHeight - e.clientY < scrollThreshold) {
                                window.scrollBy(0, 15);
                            }
                        });

                        box.addEventListener('dragleave', (e) => {
                            e.target.classList.remove('drag-over');
                        });

                        box.addEventListener('drop', (e) => {
                            e.preventDefault();
                            e.target.classList.remove('drag-over');
                            const targetIndex = e.target.getAttribute('data-index');
                            const targetType = e.target.getAttribute('data-type');

                            if (draggedIndex !== null && targetType === draggedType) {
                                const temp = data[draggedIndex][draggedType];
                                data[draggedIndex][draggedType] = data[targetIndex][draggedType];
                                data[targetIndex][draggedType] = temp;
                                renderRows();
                            }
                        });

                        box.addEventListener('input', (e) => {
                            const idx = e.target.getAttribute('data-index');
                            const type = e.target.getAttribute('data-type');
                            data[idx][type] = e.target.innerText;
                        });
                    });
                }

                function splitUnit(index) {
                    const selection = window.getSelection();
                    if (!selection.rangeCount) return;

                    const range = selection.getRangeAt(0);
                    const focusNode = selection.focusNode;
                    
                    if (!focusNode) return;

                    let targetBox = focusNode.parentElement;
                    while (targetBox && !targetBox.classList.contains('card-box')) {
                        targetBox = targetBox.parentElement;
                    }

                    if (targetBox) {
                        const type = targetBox.getAttribute('data-type');
                        const text = targetBox.innerText;
                        const offset = selection.focusOffset;

                        const part1 = text.substring(0, offset).trim();
                        const part2 = text.substring(offset).trim();

                        if (part2) {
                            data[index][type] = part1;
                            const newRow = { en: "", ar: "" };
                            newRow[type] = part2;
                            data.splice(index + 1, 0, newRow);
                            renderRows();
                        }
                    }
                }

                function deleteRow(index) {
                    data.splice(index, 1);
                    renderRows();
                }

                function addNewRow() {
                    data.push({ en: "", ar: "" });
                    renderRows();
                    window.scrollTo(0, document.body.scrollHeight);
                }

                renderRows();
            </script>
            """)

      html_code = html_template.substitute(PAIRS_JSON=pairs_json)
      components.html(html_code, height=800, scrolling=True)

      st.markdown("<br>", unsafe_allow_html=True)

      if st.button(
          "🚀 اعتماد المحاذاة النهائية وتوليد TMX", use_container_width=True
      ):
        final_pairs = [
            (en.strip(), ar.strip())
            for en, ar in zip(lines_en, lines_ar)
            if en.strip() and ar.strip()
        ]
        if final_pairs:
          st.session_state["aligned_pairs"] = final_pairs
          tmx_data = generate_tmx_file(final_pairs)

          st.success(f"🎉 تم اعتماد **{len(final_pairs)}** وحدة محاذاة بنجاح!")
          st.download_button(
              "💾 تنزيل ملف الـ TMX الموحد",
              data=tmx_data,
              file_name="Aligned_Legal_Memory.tmx",
              mime="application/x-tmx+xml",
          )
          st.info("👉 **انتقل الآن إلى 'التبويب الثاني (الاستخراج الذكي)'**!")
        else:
          st.error("⚠️ يرجى التأكد من وجود جمل صالحة للتقابل.")
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
   - Identify the exact "sub_domain" (e.g., "International Public Law", "Contractual Obligations", "Commercial Law", "Administrative Law").
   - Identify the "document_type" (e.g., "UN Convention", "Legislative Act", "Commercial Agreement", "Executive Decision").

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
            data = raw.get("terms", []) if isinstance(raw, dict) else raw

            if data and isinstance(data, list):
              for item in data:
                if isinstance(item, dict) and "english_term" in item:
                  results.append({
                      "English Term": item.get("english_term", "").strip(),
                      "Arabic Term": item.get("arabic_term", "").strip(),
                      "Sub Domain": item.get("sub_domain", "قانون عام").strip(),
                      "Document Type": item.get(
                          "document_type", "نص قانوني"
                      ).strip(),
                      "POS": item.get("pos", "تركيب اسمي").strip(),
                      "Arabic Description": item.get(
                          "arabic_description", ""
                      ).strip(),
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

  # عرض النتائج والتصدير
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

    pdf_buffer = generate_pdf_cards_bytes(df_res)

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

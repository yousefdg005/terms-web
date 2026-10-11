import base64
import io
import json
import os
import re
import time
import xml.etree.ElementTree as ET
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
from openai import OpenAI
import pandas as pd
import streamlit as st

# =========================================================
# 1. إعدادات صفحة Streamlit (يجب أن تكون أول دالة تنفيذياً)
# =========================================================
st.set_page_config(
    page_title="مستخرج المصطلحات القانونية | المجمع القانوني الليبي",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",  # جعل النافذة الجانبية مغلقة تلقائياً
)


# =========================================================
# 2. تحميل ملف الـ CSS الخارجي
# =========================================================
def load_css(file_name="style.css"):
  try:
    with open(file_name, "r", encoding="utf-8") as f:
      st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
  except FileNotFoundError:
    pass


load_css("style.css")


# =========================================================
# 3. تحويل الصور المحلية إلى Base64 للهيدر
# =========================================================
def get_image_base64(file_path):
  try:
    with open(file_path, "rb") as f:
      data = f.read()
    return base64.b64encode(data).decode()
  except Exception:
    return ""


depart_b64 = get_image_base64("depart.jpg")
moma_b64 = get_image_base64("moma.jpg")

depart_img_tag = (
    f'<img src="data:image/jpeg;base64,{depart_b64}" class="header-logo"'
    ' alt="قسم الترجمة">'
    if depart_b64
    else ""
)
moma_img_tag = (
    f'<img src="data:image/jpeg;base64,{moma_b64}" class="header-logo"'
    ' alt="المجمع القانوني">'
    if moma_b64
    else ""
)

# =========================================================
# 4. الهيدر الرئيسي المصمم بـ HTML العصري
# =========================================================
st.markdown(
    f"""
<div class="header-container">
    <div class="header-logo-wrapper">
        {depart_img_tag}
    </div>
    <div class="header-text">
        <h1 style="color: #EAB308; margin: 0; font-size: 25px; font-weight: 700;">⚖️ منصة استخراج مسارد للمصطلحات القانونية وتوليد البطاقات الترجمية</h1>
        <p style="color: #9CA3AF; font-size: 15px; margin-top: 8px; margin-bottom: 12px;">منصة ذكية متخصصة في أتمتة استخراج المصطلحات القانونية المصفاة المعتمدة لـ <b>المجمع القانوني الليبي</b> من ذاكرات الترجمة (<code>.tmx</code>)</p>
        <span style="color: #E5E7EB; font-size: 14px; background: rgba(234, 179, 8, 0.1); padding: 5px 15px; border-radius: 20px; border: 1px solid rgba(234, 179, 8, 0.3);">
            🎓 إعداد وتطوير: <b>يوسف نورالدين</b>
        </span>
    </div>
    <div class="header-logo-wrapper">
        {moma_img_tag}
    </div>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# 5. الشريط الجانبي للإعدادات والنماذج المأجورة حصراً عبر OpenRouter
# =========================================================
st.sidebar.header("⚙️ إعدادات النظام")

# بطاقة زر الانتقال المباشر لـ OpenRouter
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

# حقل المفتاح
api_key = st.sidebar.text_input(
    "🔑 مفتاح OpenRouter API Key الخاص بك:",
    value="",
    type="password",
    placeholder="أدخل مفتاحك هنا (sk-or-v1-...)",
    help="أدخل مفتاح OpenRouter المدفوع الخاص بك لتشغيل النموذج المختار.",
)

st.sidebar.markdown("---")

# قائمة النماذج المعتمدة من OpenRouter (دعم قوي للعربية وتصفية دقيقة)
PAID_MODELS_MAPPING = {
    "🌙 Qwen 3.5 Flash (الأقوى في معالجة اللغة العربية)": (
        "qwen/qwen3.5-flash-02-23"
    ),
    "📜 Claude Haiku 5.5 (التزام صارم بتعليمات الاستبعاد والتصفية)": (
        "anthropic/claude-haiku-5.5"
    ),
    "⚡ Gemini 3.8 Flash (معالجة الملفات الطويلة وسياق 1M)": (
        "google/gemini-3.8-flash"
    ),
    "⚖️ GPT-4o Mini (استخراج هيكلي دقيق وموثوق)": "openai/gpt-4o-mini",
    "🦙 Llama 4 Maverick (نموذج مفتوح متقدم متعدد اللغات)": (
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
st.sidebar.caption(f"النموذج النشط حالياً: `{selected_model_id}`")

# =========================================================
# 6. الدوال الأساسية المعالجة
# =========================================================
def sanitize_arabic_text(text):
  if not isinstance(text, str):
    return text
  cleaned = re.sub(r'[\u4e00-\u9fff\u0400-\u04FF]', '', text).strip()
  return cleaned if cleaned else 'لا يوجد وصف'


def parse_tmx_bytes(tmx_content):
  tree = ET.ElementTree(ET.fromstring(tmx_content))
  root = tree.getroot()
  pairs = []

  for tu in root.iter():
    if tu.tag.endswith('tu'):
      en_text, ar_text = '', ''
      for tuv in tu.iter():
        if tuv.tag.endswith('tuv'):
          lang = ''
          for attr_k, attr_v in tuv.attrib.items():
            if 'lang' in attr_k.lower():
              lang = attr_v.lower()
              break

          seg_text = ''
          for elem in tuv.iter():
            if elem.tag.endswith('seg'):
              seg_text = ''.join(elem.itertext()).strip()
              break

          if lang.startswith('en'):
            en_text = seg_text
          elif lang.startswith('ar'):
            ar_text = seg_text

      if en_text and ar_text:
        pairs.append((en_text, ar_text))
  return pairs


def generate_docx_bytes(df):
  doc = Document()
  title = doc.add_paragraph(
      'بطاقات المصطلحات القانونية الموثقة (Terminological Cards)'
  )
  title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  title.runs[0].font.size = Pt(18)
  title.runs[0].font.bold = True
  title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

  subtitle = doc.add_paragraph(
      'المرجعية والمصدر: المجمع القانوني الليبي (شبكة التشريعات'
      ' الليبية - lawsociety.ly) | إعداد: يوسف نورالدين'
  )
  subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
  subtitle.runs[0].font.size = Pt(11)
  subtitle.runs[0].font.italic = True

  doc.add_paragraph()

  for idx, row in df.iterrows():
    card_hdr = doc.add_paragraph(f'بطاقة مصطلح رقم: ({idx+1:02d} / {len(df)})')
    card_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    card_hdr.runs[0].font.bold = True
    card_hdr.runs[0].font.color.rgb = RGBColor(180, 0, 0)

    table = doc.add_table(rows=8, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    pos_text = (
        str(row['POS'])
        if 'POS' in row and pd.notna(row['POS'])
        else 'تركيب اسمي مركب'
    )

    fields = [
        ('المصطلح بالإنجليزي (English Term)', str(row['English Term'])),
        ('المقابل بالعربي (Arabic Term)', str(row['Arabic Term'])),
        ('النوع النحوي (Part of Speech)', pos_text),
        (
            'المجال التخصصي (Sub-Domain)',
            'قانون إداري وتشريع مؤسسي / التشريعات النافذة',
        ),
        ('السياق الأصلي الإنجليزي (Context EN)', str(row['Source EN'])),
        ('السياق الأصلي العربي (Context AR)', str(row['Source AR'])),
        (
            'الشرح والتعريف (Legal Description)',
            str(row['Arabic Description']),
        ),
        (
            'المرجع والتوثيق (Reference)',
            'المجمع القانوني الليبي (شبكة التشريعات الليبية)',
        ),
    ]

    for r_idx, (label, val) in enumerate(fields):
      row_cells = table.rows[r_idx].cells
      p_label = row_cells[0].paragraphs[0]
      p_label.alignment = WD_ALIGN_PARAGRAPH.RIGHT
      r_lbl = p_label.add_run(label)
      r_lbl.font.bold = True
      r_lbl.font.color.rgb = RGBColor(0, 51, 102)

      shd = parse_xml(r'<w:shd {} w:fill="F0F4F8"/>'.format(nsdecls('w')))
      row_cells[0]._tc.get_or_add_tcPr().append(shd)

      p_val = row_cells[1].paragraphs[0]
      p_val.alignment = WD_ALIGN_PARAGRAPH.RIGHT
      p_val.add_run(val)

      row_cells[0].width = Inches(2.3)
      row_cells[1].width = Inches(4.2)

    doc.add_paragraph()

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# =========================================================
# 7. إدارة حالة الجلسة (st.session_state) للحفاظ على المخرجات
# =========================================================
if 'results_df' not in st.session_state:
  st.session_state['results_df'] = None

# =========================================================
# 8. واجهة التفاعل ورفع الملفات ومعالجتها
# =========================================================
uploaded_file = st.file_uploader(
    '📂 اختر ملف الذاكرة الترجمية (.tmx) للبدء:', type=['tmx']
)

if uploaded_file is not None:
  if not api_key.strip():
    st.error(
        '⚠️ يرجى إدخال مفتاح OpenRouter API Key الخاص بك في الشريط الجانبي لتفعيل'
        ' المنصة.'
    )
  else:
    st.success(f'تم تحميل الملف بنجاح: **{uploaded_file.name}**')

    if st.button('🚀 بدء استخراج المصطلحات وتوليد البطاقات'):
      client = OpenAI(
          base_url='https://openrouter.ai/api/v1',
          api_key=api_key.strip(),
          default_headers={
              'HTTP-Referer': 'https://github.com',
              'X-Title': 'Libyan Legal Assembly Terminology Platform',
          },
      )

      tmx_bytes = uploaded_file.read()
      aligned_units = parse_tmx_bytes(tmx_bytes)
      st.info(
          f'تم قراءة **{len(aligned_units)}** وحدة ترجمة محاذاة من الملف.'
      )

      # System Prompt مصفى بصرامة مخصص للمجمع القانوني الليبي
      system_prompt = """You are an elite legal terminologist specialized in English-Arabic legal texts for the "Libyan Legal Assembly" (المجمع القانوني الليبي).

STRICT FILTERING & EXTRACTION RULES (قواعد التصفية الصارمة):

1. POSITIVE SCOPE (ما يجب استخراجه):
   - Extract ONLY substantive legal concepts, legal obligations, rights, liabilities, torts, and judicial principles.
   - Extract official legal institutions and legislative bodies (e.g., "House of Representatives / مجلس النواب").
   - Extract complete compound legal phrases with legal force (e.g., "Granting confidence / منح الثقة").

2. STRICT NEGATIVE DIRECTIVES (قواعد الاستبعاد الصارمة - يمنع استخراج ما يلي كلياً):
   - DO NOT extract preamble, procedural, or introductory phrases (العبارات الديباجية والإجرائية).
     * EXPLICITLY REJECT AND IGNORE phrases such as:
       - "Having reviewed / بعد الاطلاع"
       - "Pursuant to / استناداً إلى"
       - "In accordance with / بناءً على"
       - "Whereas / حيث أن"
       - "Having taken note of / بعد الإطلاع على"
   - DO NOT extract common administrative words, dates, or non-substantive general vocabulary (e.g., "Decision", "Date", "Article", "Law", "Meeting", "Document").
   - DO NOT extract connecting words, transitional phrases, or generic action verbs unless part of a defined legal concept.

3. STANDARDIZATION & REFERENCE:
   - Standardize Arabic terms: Strip contextual pronouns, prepositions, and inflections.
   - Reference Attribution: All extracted terms belong to "المجمع القانوني الليبي (شبكة التشريعات الليبية)".

4. CATEGORIZATION (pos):
   - Categorize 'pos' accurately as one of:
     - "اسم (Noun)"
     - "تركيب اسمي مركب (Compound Noun Phrase)"
     - "فعل / عبارة فعلية (Verb / Verb Phrase)"
     - "حرف / شبه جملة (Preposition / Phrase)"

5. OUTPUT FORMAT:
   - Output MUST be valid JSON as an object with a key "terms", containing an array of objects with keys:
     "english_term", "arabic_term", "pos", "arabic_description".
"""

      results = []
      seen_terms = set()
      progress_bar = st.progress(0)
      status_text = st.empty()

      for i, (en_text, ar_text) in enumerate(aligned_units):
        status_text.text(
            f'جاري معالجة الوحدة ({i+1}/{len(aligned_units)})... عبر نموذج'
            f' `{selected_model_id}`'
        )
         user_prompt = f'[EN]: {en_text}\n[AR]: {ar_text}'

        max_retries = 3
        for attempt in range(max_retries):
          try:
            time.sleep(2.0)  # زمن التهدئة بين الطلبات

            response = client.chat.completions.create(
                model=selected_model_id,
                messages=[
                    {'role': 'system', 'content': system_prompt},
                    {'role': 'user', 'content': user_prompt},
                ],
                temperature=0.0,
                response_format={'type': 'json_object'},
            )

            content = response.choices[0].message.content
            raw = json.loads(content)
            data = raw.get('terms', []) if isinstance(raw, dict) else raw

            if data:
              for item in data:
                if isinstance(item, dict) and 'english_term' in item:
                  en_term = str(item.get('english_term', '')).strip()
                  ar_term = str(item.get('arabic_term', '')).strip()
                  pos_val = str(
                      item.get('pos', 'تركيب اسمي مركب')
                  ).strip()
                  ar_desc = sanitize_arabic_text(
                      item.get('arabic_description', '')
                  )

                  term_key = (en_term.lower(), ar_term)
                  if term_key not in seen_terms and en_term and ar_term:
                    seen_terms.add(term_key)
                    results.append({
                        'English Term': en_term,
                        'Arabic Term': ar_term,
                        'POS': pos_val,
                        'Arabic Description': ar_desc,
                        'Source EN': en_text,
                        'Source AR': ar_text,
                    })
            break

          except Exception as e:
            if '429' in str(e):
              if attempt < max_retries - 1:
                status_text.text(
                    f'⏳ تجاوز الحد المؤقت.. انتظار 10 ثوانٍ لإعادة محاولة'
                    f' الوحدة {i+1}...'
                )
                time.sleep(10.0)
              else:
                st.warning(
                    f'تجاوز الحد المسموح للوحدة {i+1} بعد عدة محاولات.'
                )
            else:
              st.warning(f'تنبيه أثناء معالجة الوحدة {i+1}: {e}')
              break

        progress_bar.progress((i + 1) / len(aligned_units))

      status_text.text('✅ اكتملت المعالجة والتصفية بنجاح!')

      if results:
        df = pd.DataFrame(results)
        df.drop_duplicates(
            subset=['English Term', 'Arabic Term'], inplace=True
        )
        st.session_state['results_df'] = df
      else:
        st.session_state['results_df'] = None
        st.error(
            'لم يتم العثور على أي مصطلحات قانونية موضوعية بعد تطبيق قواعد'
            ' التصفية الصارمة.'
        )

# =========================================================
# 9. عرض النتائج وتنزيل الملفات (محفوظة بـ session_state)
# =========================================================
if st.session_state['results_df'] is not None:
  df = st.session_state['results_df']

  st.subheader('📊 عرض تفاعلي لمخرجات المصطلحات المستخرجة والمصفاة:')
  st.dataframe(
      df[['English Term', 'Arabic Term', 'POS', 'Arabic Description']],
      use_container_width=True,
  )

  # تجهيز التنزيل لـ Excel
  # تجهيز التنزيل لـ Excel (مع استبعاد عمود POS)
excel_buffer = io.BytesIO()
excel_df = df.drop(columns=['POS'], errors='ignore')
with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
    excel_df.to_excel(writer, index=False)
excel_buffer.seek(0)

  # تجهيز التنزيل لـ Word
  docx_buffer = generate_docx_bytes(df)

  col1, col2 = st.columns(2)
  with col1:
    st.download_button(
        label='📥 تنزيل قاعدة البيانات (Excel)',
        data=excel_buffer,
        file_name='Libyan_Legal_Terms_Extracted.xlsx',
        mime=(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ),
    )
  with col2:
    st.download_button(
        label='📄 تنزيل البطاقات الترجمية (Word Docx)',
        data=docx_buffer,
        file_name='Terminological_Cards_Libyan_Law.docx',
        mime=(
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ),
    )

# =========================================================
# 10. التذييل بأسفل الصفحة (Dark Mode Compatible)
# =========================================================
st.markdown('<br><br>', unsafe_allow_html=True)
st.markdown(
    """
<div style="text-align: center; background-color: #131B2E; padding: 20px; border-radius: 12px; border: 1px solid #1E293B;">
    <h3 style="color: #EAB308; margin-bottom: 10px;">📫 تواصل معي (Contact Me)</h3>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;"><b>إعداد وتطوير:</b> يوسف نورالدين</p>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;">📱 <b>رقم الهاتف:</b> <a href="tel:0928756918" style="color: #EAB308; text-decoration: none;">0928756918</a></p>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;">✉️ <b>البريد الإلكتروني:</b> <a href="mailto:ywsfalmqmz@gmail.com" style="color: #EAB308; text-decoration: none;">ywsfalmqmz@gmail.com</a></p>
    <hr style="margin: 15px auto; width: 50%; border-top: 1px solid #1E293B;">
    <p style="font-size: 12px; color: #9CA3AF;">جميع الحقوق محفوظة © منصة أتمتة البطاقات الترجمية - المجمع القانوني الليبي</p>
</div>
""",
    unsafe_allow_html=True,
)

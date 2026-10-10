import os
import io
import xml.etree.ElementTree as ET
import json
import re
import time
import base64
import pandas as pd
import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# =========================================================
# 1. إعدادات صفحة Streamlit (يجب أن تكون أول دالة تنفيذياً)
# =========================================================
st.set_page_config(
    page_title="مستخرج المصطلحات القانونية | يوسف نورالدين",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed"  # جعل النافذة الجانبية مغلقة تلقائياً
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

depart_img_tag = f'<img src="data:image/jpeg;base64,{depart_b64}" class="header-logo" alt="قسم الترجمة">' if depart_b64 else ''
moma_img_tag = f'<img src="data:image/jpeg;base64,{moma_b64}" class="header-logo" alt="المجمع القانوني">' if moma_b64 else ''

# =========================================================
# 4. الهيدر الرئيسي المصمم بـ HTML العصري
# =========================================================
st.markdown(f"""
<div class="header-container">
    <div class="header-logo-wrapper">
        {depart_img_tag}
    </div>
    <div class="header-text">
        <h1 style="color: #EAB308; margin: 0; font-size: 25px; font-weight: 700;">⚖️ منصة استخراج مسارد للمصطلحات القانونية وتوليد البطاقات الترجمية</h1>
        <p style="color: #9CA3AF; font-size: 15px; margin-top: 8px; margin-bottom: 12px;">منصة ذكية متخصصة في أتمتة استخراج المصطلحات القانونية وبناء البطاقات الترجمية الموثقة من ذاكرات الترجمة (<code>.tmx</code>)</p>
        <span style="color: #E5E7EB; font-size: 14px; background: rgba(234, 179, 8, 0.1); padding: 5px 15px; border-radius: 20px; border: 1px solid rgba(234, 179, 8, 0.3);">
            🎓 إعداد وتطوير: <b>يوسف نورالدين</b>
        </span>
    </div>
    <div class="header-logo-wrapper">
        {moma_img_tag}
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 5. الشريط الجانبي للإعدادات والنماذج (منزلق ومغلق تلقائياً)
# =========================================================
st.sidebar.header("⚙️ إعدادات النظام")

# بطاقة زر الانتقال المباشر لـ OpenRouter
st.sidebar.markdown("""
<div style="background-color: #1E293B; padding: 12px; border-radius: 8px; border: 1px solid #374151; margin-bottom: 15px; text-align: center;">
    <span style="color: #E5E7EB; font-size: 13px;">ليس لديك مفتاح API؟</span><br>
    <a href="https://openrouter.ai/keys" target="_blank" style="color: #EAB308; font-weight: bold; text-decoration: none; font-size: 14px; display: inline-block; margin-top: 5px;">
        🔗 الحصول على المفتاح من OpenRouter
    </a>
</div>
""", unsafe_allow_html=True)

# حقل المفتاح
api_key = st.sidebar.text_input(
    "🔑 مفتاح OpenRouter API Key الخاص بك:",
    value="",
    type="password",
    placeholder="أدخل مفتاحك هنا (sk-or-v1-...)",
    help="أدخل مفتاح OpenRouter الخاص بك لتشغيل النموذج اللغوي."
)

# دليل إرشادي تفاعلي
with st.sidebar.expander("❓ كيف تحصل على مفتاح API مجاني؟"):
    st.markdown("""
    **الخطوات بسيطة جداً:**
    1. اضغط على الرابط أعلاه للانتقال إلى [OpenRouter.ai/keys](https://openrouter.ai/keys).
    2. قم بتسجيل الدخول واضغط على **Create Key**.
    3. انسخ المفتاح المكون من رموز وألصقه في الحقل أعلاه.
    
    ---
    * 🟢 **النماذج المجانية:** تعمل فوراً ودون شحن رصيد.
    * 💳 **النماذج المدفوعة:** تتطلب شحن رصيد بسيط في حسابك.
    """)

st.sidebar.markdown("---")

# قائمة النماذج المتاحة عبر OpenRouter
AVAILABLE_MODELS = {
    "🟢 Google Gemma 4 31B (مجاني بالكامل - موصى به)": "google/gemma-4-31b-it:free",
    "🟢 Meta Llama 3.3 70B (مجاني بالكامل - قوي وسريع)": "meta-llama/llama-3.3-70b-instruct:free",
    "🟢 Qwen 2.5 Coder 32B (مجاني بالكامل - دقيق هيكلياً)": "qwen/qwen-2.5-coder-32b-instruct:free",
    "🟢 DeepSeek R1 (مجاني بالكامل - استدلالي متقدم)": "deepseek/deepseek-r1:free",
    "🟡 OpenRouter Auto Free Router (موجه مجاني تلقائي)": "openrouter/free",
    "💳 OpenAI GPT-4o Mini (مدفوع - اقتصادي)": "openai/gpt-4o-mini",
    "💳 Anthropic Claude 3.5 Sonnet (مدفوع - أعلى دقة)": "anthropic/claude-3.5-sonnet"
}

st.sidebar.subheader("🤖 اختيار نموذج الذكاء الاصطناعي")

selected_model_label = st.sidebar.selectbox(
    "اختر النموذج المطلوب للاستخراج:",
    options=list(AVAILABLE_MODELS.keys()),
    index=0
)

model_name = AVAILABLE_MODELS[selected_model_label]
st.sidebar.caption(f"النموذج النشط حالياً: `{model_name}`")

# =========================================================
# 6. الدوال الأساسية المعالجة
# =========================================================
def sanitize_arabic_text(text):
    if not isinstance(text, str):
        return text
    cleaned = re.sub(r'[\u4e00-\u9fff\u0400-\u04FF]', '', text).strip()
    return cleaned if cleaned else "لا يوجد وصف"

def parse_tmx_bytes(tmx_content):
    tree = ET.ElementTree(ET.fromstring(tmx_content))
    root = tree.getroot()
    pairs = []
    
    for tu in root.iter():
        if tu.tag.endswith('tu'):
            en_text, ar_text = "", ""
            for tuv in tu.iter():
                if tuv.tag.endswith('tuv'):
                    lang = ""
                    for attr_k, attr_v in tuv.attrib.items():
                        if 'lang' in attr_k.lower():
                            lang = attr_v.lower()
                            break
                    
                    seg_text = ""
                    for elem in tuv.iter():
                        if elem.tag.endswith('seg'):
                            seg_text = "".join(elem.itertext()).strip()
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
    title = doc.add_paragraph("بطاقات المصطلحات القانونية الموثقة (Terminological Cards)")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(18)
    title.runs[0].font.bold = True
    title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

    subtitle = doc.add_paragraph("المرجعية والمصدر: المجمع القانوني الليبي (www.lawsociety.ly) | إعداد: يوسف نورالدين")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(11)
    subtitle.runs[0].font.italic = True

    doc.add_paragraph()

    for idx, row in df.iterrows():
        card_hdr = doc.add_paragraph(f"بطاقة مصطلح رقم: ({idx+1:02d} / {len(df)})")
        card_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        card_hdr.runs[0].font.bold = True
        card_hdr.runs[0].font.color.rgb = RGBColor(180, 0, 0)
        
        table = doc.add_table(rows=8, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        pos_text = str(row['POS']) if 'POS' in row and pd.notna(row['POS']) else "تركيب اسمي مركب"

        fields = [
            ("المصطلح بالإنجليزي (English Term)", str(row['English Term'])),
            ("المقابل بالعربي (Arabic Term)", str(row['Arabic Term'])),
            ("النوع النحوي (Part of Speech)", pos_text),
            ("المجال التخصصي (Sub-Domain)", "قانون إداري وتنفيذي / تنظيم المؤسسات والتعليم"),
            ("السياق الأصلي الإنجليزي (Context EN)", str(row['Source EN'])),
            ("السياق الأصلي العربي (Context AR)", str(row['Source AR'])),
            ("الشرح والتعريف (Legal Description)", str(row['Arabic Description'])),
            ("المرجع والتوثيق (Reference)", "المجمع القانوني الليبي (lawsociety.ly)")
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
# 7. واجهة التفاعل ورفع الملفات
# =========================================================
uploaded_file = st.file_uploader("📂 اختر ملف tmx", type=["tmx"])

if uploaded_file is not None:
    if not api_key.strip():
        st.error("⚠️ يرجى إدخال مفتاح OpenRouter API Key الخاص بك في الشريط الجانبي لتفعيل المنصة.")
    else:
        st.success(f"تم تحميل الملف بنجاح: **{uploaded_file.name}**")
        
        if st.button("🚀 بدء استخراج المصطلحات وتوليد البطاقات"):
            client = OpenAI(
                base_url="https://openrouter.ai/api/v1",
                api_key=api_key.strip(),
                default_headers={
                    "HTTP-Referer": "https://github.com",
                    "X-Title": "Legal Terminology Extractor GUI"
                }
            )

            tmx_bytes = uploaded_file.read()
            aligned_units = parse_tmx_bytes(tmx_bytes)
            st.info(f"تم قراءة **{len(aligned_units)}** وحدة ترجمة محاذاة من الملف.")

            system_prompt = """You are an expert legal terminologist specializing in English-Arabic legal texts.

Strict Rules:
1. Extract ONLY domain-specific substantive legal terms and complete compound legal phrases.
2. Ignore metadata headers and administrative labels.
3. Standardize Arabic terms: Strip contextual pronouns and inflections.
4. Categorize 'pos' accurately as one of:
   - "اسم (Noun)"
   - "تركيب اسمي مركب (Compound Noun Phrase)"
   - "فعل / عبارة فعلية (Verb / Verb Phrase)"
   - "حرف / شبه جملة (Preposition / Phrase)"
5. Provide a concise, accurate legal explanation under "arabic_description" STRICTLY in Modern Standard Arabic.
6. Output MUST be valid JSON as an object with a key "terms", containing an array of objects with keys: "english_term", "arabic_term", "pos", "arabic_description"."""

            results = []
            seen_terms = set()
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, (en_text, ar_text) in enumerate(aligned_units):
                status_text.text(f"جاري معالجة الوحدة ({i+1}/{len(aligned_units)})...")
                user_prompt = f"[EN]: {en_text}\n[AR]: {ar_text}"
                
                try:
                    time.sleep(1.0)
                    response = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_prompt}
                        ],
                        temperature=0.0,
                        response_format={"type": "json_object"}
                    )
                    
                    content = response.choices[0].message.content
                    raw = json.loads(content)
                    data = raw.get("terms", []) if isinstance(raw, dict) else raw

                    if data:
                        for item in data:
                            if isinstance(item, dict) and "english_term" in item:
                                en_term = str(item.get("english_term", "")).strip()
                                ar_term = str(item.get("arabic_term", "")).strip()
                                pos_val = str(item.get("pos", "تركيب اسمي مركب")).strip()
                                ar_desc = sanitize_arabic_text(item.get("arabic_description", ""))
                                
                                term_key = (en_term.lower(), ar_term)
                                if term_key not in seen_terms and en_term and ar_term:
                                    seen_terms.add(term_key)
                                    results.append({
                                        "English Term": en_term,
                                        "Arabic Term": ar_term,
                                        "POS": pos_val,
                                        "Arabic Description": ar_desc,
                                        "Source EN": en_text,
                                        "Source AR": ar_text
                                    })
                except Exception as e:
                    st.warning(f"تنبيه أثناء معالجة الوحدة {i+1}: {e}")

                progress_bar.progress((i + 1) / len(aligned_units))

            status_text.text("✅ اكتملت المعالجة بنجاح!")
            
            if results:
                df = pd.DataFrame(results)
                df.drop_duplicates(subset=["English Term", "Arabic Term"], inplace=True)
                
                st.subheader("📊 عرض تفاعلي لمخرجات المصطلحات المستخرجة:")
                st.dataframe(df[["English Term", "Arabic Term", "POS", "Arabic Description"]], use_container_width=True)

                # تجهيز التنزيل لـ Excel
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                excel_buffer.seek(0)

                # تجهيز التنزيل لـ Word
                docx_buffer = generate_docx_bytes(df)

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 تنزيل قاعدة البيانات (Excel)",
                        data=excel_buffer,
                        file_name="Legal_Terms_Extracted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.download_button(
                        label="📄 تنزيل البطاقات الترجمية (Word Docx)",
                        data=docx_buffer,
                        file_name="Terminological_Cards_Libyan_Law.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error("لم يتم العثور على أي مصطلحات قانونية مستخرجة من الملف.")

# =========================================================
# 8. التذييل بأسفل الصفحة (Dark Mode Compatible)
# =========================================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; background-color: #131B2E; padding: 20px; border-radius: 12px; border: 1px solid #1E293B;">
    <h3 style="color: #EAB308; margin-bottom: 10px;">📫 تواصل معي (Contact Me)</h3>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;"><b>إعداد وتطوير:</b> يوسف نورالدين</p>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;">📱 <b>رقم الهاتف:</b> <a href="tel:0928756918" style="color: #EAB308; text-decoration: none;">0928756918</a></p>
    <p style="font-size: 16px; color: #E5E7EB; margin: 5px;">✉️ <b>البريد الإلكتروني:</b> <a href="mailto:ywsfalmqmz@gmail.com" style="color: #EAB308; text-decoration: none;">ywsfalmqmz@gmail.com</a></p>
    <hr style="margin: 15px auto; width: 50%; border-top: 1px solid #1E293B;">
    <p style="font-size: 12px; color: #9CA3AF;">جميع الحقوق محفوظة © منصة أتمتة البطاقات الترجمية القانونية</p>
</div>
""", unsafe_allow_html=True)
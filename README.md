
# ⚖️ Legal Terminology Extractor & Flashcard Generator

An intelligent AI-powered web application designed for the Libyan Legal Assembly (Libyan Legislation Network) to extract domain-specific legal terminology from manually aligned Translation Memory files (.tmx) in English and Arabic. It automatically applies advanced preprocessing, strict negative directives, and custom prompt engineering via OpenRouter API to generate downloadable Excel glossaries and formatted Word terminological cards.

---

## 📋 Prerequisites & Usage Requirements

To process files successfully, users must provide:

* **A Manually Aligned Parallel TMX File (.tmx):** Upload an aligned bilingual translation memory file containing manually paired English-Arabic translation units.
* **OpenRouter API Key:** Input a valid OpenRouter API key (`sk-or-v1-...`) in the sidebar to authorize model requests.

---

## 🤖 Models & Account Credit

Supported Suite: Users can select from a curated list of high-precision LLMs via OpenRouter, optimized for Arabic and structured output parsing:

* `qwen/qwen3.5-flash-02-23` (Qwen 3.5 Flash)
* `anthropic/claude-haiku-5.5` (Claude Haiku 5.5)
* `google/gemini-3.8-flash` (Gemini 3.8 Flash)
* `openai/gpt-4o-mini` (GPT-4o Mini)
* `meta-llama/llama-4-maverick` (Llama 4 Maverick)

> 💡 **Important Note:** Maintaining a small credit balance (e.g., as low as $2.00) on your OpenRouter account completely bypasses free-tier rate limits (Error 429) and prevents service interruption during peak server traffic.

---

## ⚙️ Terminology Extraction Pipeline (Pre-processing & Process)

To achieve ultimate precision, the extraction workflow follows rigorous processing stages:

* **Pre-processing Stage (المعالجة الأولية):** Reading the translation memory files where users have manually aligned the texts, cleaning raw units, removing XML/HTML tags and markup symbols, and filtering out invalid or empty segments to prepare pure parallel text blocks.
* **Process Stage (المعالجة التحليلية العميقة):** Enforcing Strict Negative Directives and advanced Prompt Engineering to automatically discard procedural preambles, administrative filler words, and dates, while guiding the AI models to isolate and standardize substantive legal concepts and institutional titles.

---

## 🚀 Key Features

* 📂 **TMX Parsing & Manual Alignment Support:** Processing user-aligned bilingual translation units from uploaded .tmx files.
* 🛡️ **Strict Negative Directives & Prompt Engineering:** Advanced prompt design combined with strict filtering rules to eliminate preambles/procedural phrases and dates, isolating true legal concepts.
* 🤖 **Multi-LLM Integration:** Flexible model selection tailored for legal translation tasks via OpenRouter API.
* 💾 **State Persistence (`st.session_state`):** Protects extracted datasets and interactive preview tables from wiping during file downloads or UI interactions.
* 🎨 **Modern Branded UI/UX:** Dark-mode interface with custom CSS styling, tab configuration, background image overlay, and institutional logos (`depart.jpg` & `moma.jpg`).
* 📊 **Multi-Format Export:**
  * **Excel Spreadsheet (.xlsx):** Filtered legal term glossary (automatically excludes the POS column).
  * **Word Document (.docx):** Formatted legal terminological cards attributed to the Libyan Legal Assembly.

---

## 💻 Core Code Snippets & System Implementation

### 1. Pre-processing: XML Cleaning & TMX Parsing
Extracting raw parallel units and cleaning XML/HTML tags and markup noise:

```python
import xml.etree.ElementTree as ET
import re

def parse_and_clean_tmx(tmx_file):
    """
    Pre-processing: Extract raw parallel units and clean XML/HTML tags.
    """
    tree = ET.parse(tmx_file)
    root = tree.getroot()
    parallel_units = []

    for tu in root.findall('.//tu'):
        tuv_en = tu.find(".//tuv[@xml:lang='en']/seg", namespaces={'xml': '[http://www.w3.org/XML/1998/namespace](http://www.w3.org/XML/1998/namespace)'})
        tuv_ar = tu.find(".//tuv[@xml:lang='ar']/seg", namespaces={'xml': '[http://www.w3.org/XML/1998/namespace](http://www.w3.org/XML/1998/namespace)'})
        
        if tuv_en is not None and tuv_ar is not None:
            clean_en = re.sub(r'<[^>]+>', '', tuv_en.text or '').strip()
            clean_ar = re.sub(r'<[^>]+>', '', tuv_ar.text or '').strip()
            
            if clean_en and clean_ar:
                parallel_units.append({'en': clean_en, 'ar': clean_ar})
                
    return parallel_units

2. Process: Prompt Engineering & Strict Negative Directives
Guided prompt formulation enforcing strict negative directives to filter out noise and isolate substantive legal terms:
SYSTEM_PROMPT = """
You are an expert legal terminologist specializing in Libyan and Arab legal systems.
Your task is to extract substantive legal terms, institutions, and domain-specific concepts from the provided parallel text.

### STRICT NEGATIVE DIRECTIVES:
1. DO NOT extract procedural preambles or administrative boilerplate phrases (e.g., "Having reviewed", "Pursuant to", "بعد الاطلاع على", "بناء على ما عرضه").
2. DO NOT extract generic dates, numbers, or article references (e.g., "Law No. 5", "المادة 12").
3. DO NOT extract common non-legal verbs or conjunctions.

### EXTRACTION RULES:
- Focus exclusively on substantive legal concepts, statutory bodies, rights, and liabilities.
- Provide the exact source English term and its standardized Arabic legal equivalent.
- Output MUST strictly adhere to JSON format.
"""

3. OpenRouter API Integration
Executing LLM inference via OpenRouter API with dynamic model choice:
import openai

def extract_legal_terms(cleaned_units, api_key, model_name="qwen/qwen3.5-flash-02-23"):
    client = openai.OpenAI(
        base_url="[https://openrouter.ai/api/v1](https://openrouter.ai/api/v1)",
        api_key=api_key,
    )
    
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Extract terms from these parallel units: {cleaned_units}"}
        ],
        temperature=0.1
    )
    
    return response.choices[0].message.content

🛠️ Tech Stack
 * Language: Python
 * Framework: Streamlit
 * LLM Provider: OpenRouter / OpenAI SDK
 * Data & Document Processing: Pandas, OpenPyXL, Python-docx, ElementTree
👤 Developer & Academic Context
 * Developer: Yousef Nuralddin Bashir Elmgamez
 * Department: Translation Studies, Faculty of Languages, University of Tripoli
 * Reference Body: Libyan Legal Assembly (Libyan Legislation Network - lawsociety.ly)
 * Contact: +218 92 8756918 | ywsfalmqmz@gmail.com
⚖️ منصة استخراج المصطلحات القانونية وتوليد البطاقات الترجمية
منصة ذكية مخصصة لصالح المجمع القانوني الليبي (شبكة التشريعات الليبية) تعتمد على معالجة اللغات الطبيعية ونماذج الذكاء الاصطناعي لاستخراج المصطلحات القانونية الموضوعية المصفاة من ملفات الذاكرة الترجمية (.tmx) التي قام المستخدم بإنشائها ومحاذاتها يدوياً، مع تطبيق هندسة الأوامر وقواعد التصفية الدقيقة، وتصدير النتائج إلى مسارد Excel وبطاقات معجمية موثقة بصيغة Word.
📋 متطلبات التشغيل والمستلزمات
لتشغيل المنصة واستخراج المصطلحات بنجاح، يحتاج المستخدم إلى:
 * ملف ذاكرة ترجمية محاذى يدوياً (.tmx): رفع ملف TMX ثنائي اللغة يحتوي على وحدات ترجمة تم إعدادها ومحاذاتها يدوياً من قبل المستخدم.
 * مفتاح API من OpenRouter: إدخال مفتاح API الخاص بالمستخدم (sk-or-v1-...) في الشريط الجانبي لتفعيل عملية الاستخراج.
🤖 النماذج ورصيد الحساب
النماذج المتاحة: تتيح المنصة التبديل بين مجموعة نماذج مدفوعة ومتقدمة عبر OpenRouter تتميز بالدقة الفائقة في معالجة اللغة العربية والالتزام بالمخرجات الهيكلية:
 * qwen/qwen3.5-flash-02-23 (Qwen 3.5 Flash)
 * anthropic/claude-haiku-5.5 (Claude Haiku 5.5)
 * google/gemini-3.8-flash (Gemini 3.8 Flash)
 * openai/gpt-4o-mini (GPT-4o Mini)
 * meta-llama/llama-4-maverick (Llama 4 Maverick)
> 💡 ملاحظة هامة: وجود رصيد رمزي بسيط (حتى لو 2 دولار فقط) في حساب OpenRouter الخاص بك يلغي قيود معدل الطلبات (Rate Limits / Error 429) كلياً ويمنع حدوث الأخطاء أو التوقف أثناء أوقات الضغط العالي.
> 
⚙️ آلية العمل ومراحل المعالجة (Pre-processing & Process)
تعتمد المنصة على خط إنتاج برمجي دقيق يمر بمرحلتين أساسيتين:
 * مرحلة المعالجة الأولية (Pre-processing): قراءة ملفات الذاكرة الترجمية المحاذاة يدوياً، تنظيف النصوص من وسوم XML والرموز التشويقية، وتصفية الوحدات غير الصالحة للتحليل لضمان نقاء النصوص الموازية.
 * مرحلة المعالجة والتحليل الذكي (Process): توظيف هندسة الأوامر (Prompt Engineering) المتقدمة مع تطبيق قواعد التصفية الصارمة (Strict Negative Directives) لاستبعاد عبارات الديباجة، المقدمات الإجرائية، والتواريخ، وتوجيه النماذج للتركيز الحصري على استخراج وتوحيد صياغة المفاهيم والمؤسسات القانونية الموضوعية بدقة متناهية.
🚀 المميزات الرئيسية
 * 📂 دعم ملفات TMX المحاذاة يدوياً: معالجة وحدات الترجمة الثنائية التي أعدها المستخدم بعناية.
 * 🛡️ هندسة الأوامر والتصفية الصارمة (Strict Negative Directives & Prompt Engineering): استخدام تعليمات برمجية واضحة لاستبعاد عبارات الديباجة والإجراءات والتواريخ والتركيز على المفاهيم الموضوعية.
 * 🤖 دعم متعدد للنماذج: إمكانية اختيار النموذج المناسب لمعالجة النصوص القانونية عبر OpenRouter.
 * 💾 حفظ النتائج وحمايتها (st.session_state): ضمان استمرار عرض النتائج في الجدول التفاعلي ومنع اختفائها أثناء التفاعل أو التحميل.
 * 🎨 واجهة عربية تفاعلية عصرية: تصميم Dark Mode متجاوب مع تنسيقات CSS، عنوان النافذة والأيقونة، خلفية شفافة عصرية، وشعارات موثقة للقسم والمجمع.
 * 📄 تصدير متعدد الصيغ:
   * ملف Excel (.xlsx): مسرد مصطلحي مصفى (مستبعد منه عمود POS تلقائياً).
   * وثيقة Word (.docx): بطاقات ترجمية موثقة ومعتمدة باسم المجمع القانوني الليبي.
💻 نماذج الكود الأساسية وآلية التنفيذ
1. مرحلة المعالجة الأولية: تنظيف وسوم XML وقراءة TMX
استخراج الوحدات الموازية وتنظيفها من الوسوم التشويقية:
import xml.etree.ElementTree as ET
import re

def parse_and_clean_tmx(tmx_file):
    """
    Pre-processing: Extract raw parallel units and clean XML/HTML tags.
    """
    tree = ET.parse(tmx_file)
    root = tree.getroot()
    parallel_units = []

    for tu in root.findall('.//tu'):
        tuv_en = tu.find(".//tuv[@xml:lang='en']/seg", namespaces={'xml': '[http://www.w3.org/XML/1998/namespace](http://www.w3.org/XML/1998/namespace)'})
        tuv_ar = tu.find(".//tuv[@xml:lang='ar']/seg", namespaces={'xml': '[http://www.w3.org/XML/1998/namespace](http://www.w3.org/XML/1998/namespace)'})
        
        if tuv_en is not None and tuv_ar is not None:
            clean_en = re.sub(r'<[^>]+>', '', tuv_en.text or '').strip()
            clean_ar = re.sub(r'<[^>]+>', '', tuv_ar.text or '').strip()
            
            if clean_en and clean_ar:
                parallel_units.append({'en': clean_en, 'ar': clean_ar})
                
    return parallel_units

2. مرحلة المعالجة والتحليل: هندسة الأوامر والتصفية الصارمة
البرومبت الموجه للنموذج عبر OpenRouter لتطبيق الاستبعاد الصارم واستخراج المفاهيم الموضوعية فقط:
SYSTEM_PROMPT = """
You are an expert legal terminologist specializing in Libyan and Arab legal systems.
Your task is to extract substantive legal terms, institutions, and domain-specific concepts from the provided parallel text.

### STRICT NEGATIVE DIRECTIVES:
1. DO NOT extract procedural preambles or administrative boilerplate phrases (e.g., "Having reviewed", "Pursuant to", "بعد الاطلاع على", "بناء على ما عرضه").
2. DO NOT extract generic dates, numbers, or article references (e.g., "Law No. 5", "المادة 12").
3. DO NOT extract common non-legal verbs or conjunctions.

### EXTRACTION RULES:
- Focus exclusively on substantive legal concepts, statutory bodies, rights, and liabilities.
- Provide the exact source English term and its standardized Arabic legal equivalent.
- Output MUST strictly adhere to JSON format.
"""

3. الربط البرمجي عبر OpenRouter API
كود استدعاء نماذج الذكاء الاصطناعي ومعالجة البيانات:
import openai

def extract_legal_terms(cleaned_units, api_key, model_name="qwen/qwen3.5-flash-02-23"):
    client = openai.OpenAI(
        base_url="[https://openrouter.ai/api/v1](https://openrouter.ai/api/v1)",
        api_key=api_key,
    )
    
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Extract terms from these parallel units: {cleaned_units}"}
        ],
        temperature=0.1
    )
    
    return response.choices[0].message.content

🛠️ التقنيات المستخدمة
 * Language: Python
 * Framework: Streamlit
 * LLM Provider: OpenRouter / OpenAI SDK
 * Data & Document Processing: Pandas, OpenPyXL, Python-docx, ElementTree
👤 إعداد وتطوير
 * المطور: يوسف نورالدين بشير المقعمز
 * التخصص: شعبة الترجمة - كلية اللغات - جامعة طرابلس
 * المرجعية التشريعية: المجمع القانوني الليبي (شبكة التشريعات الليبية - lawsociety.ly)
 * التواصل: +218 92 8756918 | ywsfalmqmz@gmail.com


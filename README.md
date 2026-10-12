إليك ملف `README.md` الشامل والمُحدَّث بناءً على التعديلات الأخيرة، مع التنويه بأن "استوديو المحاذاة" لا يزال تحت التطوير (Beta). قمت بتجهيزه باللغتين العربية والإنجليزية ليكون جاهزاً للنسخ مباشرة:

```markdown
# ⚖️ AI Legal Workspace
**منصة التحليل والاستخراج القانوني وتوليد البطاقات عبر الذكاء الاصطناعي**

منصة ذكية متكاملة مصممة للمترجمين القانونيين والمحامين، توفر أدوات متطورة لمحاذاة النصوص القانونية (عربي/إنجليزي)، استخراج المصطلحات المتخصصة بدقة عالية باستخدام نماذج الذكاء الاصطناعي، وتوليد بطاقات ترجمية (PDF) ومسارد (Excel) احترافية.

---

## ✨ المميزات الرئيسية (Features)

*   🟢 **استوديو التحويل والمحاذاة - Alignment Studio (Beta ⚠️):**
    *   *ملاحظة: هذه الأداة لا تزال تحت التطوير (نسخة تجريبية).*
    *   استخراج هجين للنصوص من ملفات PDF مع دعم متقدم لتقنية التعرف البصري على الحروف (OCR) للغة العربية.
    *   تجزئة ذكية للنصوص القانونية (حسب المواد والبنود).
    *   واجهة تفاعلية مرنة لسحب وإفلات، تقسيم، وحذف وحدات المحاذاة.
    *   توليد وتصدير ملفات ذاكرة الترجمة (TMX).
*   🔵 **الاستخراج الذكي للمصطلحات - AI Term Extractor:**
    *   استخراج آلي للمصطلحات القانونية الموضوعية وتصنيفها (المجال الفرعي، نوع المستند).
    *   تصفية صارمة لاستبعاد العبارات التمهيدية والإجرائية.
    *   دعم نماذج ذكاء اصطناعي متعددة عبر (OpenRouter) مثل Qwen, Claude, Gemini, Llama, و GPT-4o.
*   📊 **التصدير والتقارير:**
    *   تصدير النتائج كمسرد مصطلحات Excel.
    *   توليد بطاقات ترجمية أنيقة ومنسقة بصيغة PDF لدعم المراجعة والدراسة.

## 🚀 متطلبات التشغيل (Prerequisites)

تأكد من تثبيت مكتبات بايثون التالية لتشغيل المنصة:
```bash
pip install streamlit pandas openai pdfplumber PyMuPDF pytesseract reportlab arabic-reshaper python-bidi python-docx openpyxl

```

*(ملاحظة: يتطلب استخراج النصوص عبر الصور تثبيت محرك `Tesseract OCR` على نظام التشغيل الخاص بك).*

## 🛠️ كيفية الاستخدام

1. قم بتشغيل المنصة محلياً عبر الأمر:
```bash
streamlit run app.py

```


2. أدخل مفتاح `OpenRouter API Key` الخاص بك في الشريط الجانبي.
3. **استوديو المحاذاة (التبويب الأول):** ارفع ملفات PDF، قم بتعديل المحاذاة يدوياً، وحمّل ملف `TMX`.
4. **الاستخراج الذكي (التبويب الثاني):** ارفع ملف `TMX` (أو اعتمد على المحاذاة الحالية)، اختر نموذج الذكاء الاصطناعي، وابدأ الاستخراج.
5. حمل مسرد (Excel) أو بطاقات (PDF) من قسم النتائج.

## 👨‍💻 المطور

**إعداد وتطوير:** يوسف نورالدين (Yousef Nour Eldin)

---

---

# ⚖️ AI Legal Workspace

**Your all-in-one AI toolkit for legal text alignment and term extraction**

An integrated, AI-powered platform designed for legal translators and professionals. It provides advanced tools for aligning legal texts (English/Arabic), strictly extracting specialized terminology using Large Language Models, and generating professional PDF flashcards and Excel glossaries.

---

## ✨ Key Features

* 🟢 **Alignment Studio (Beta ⚠️):**
* *Note: This tool is currently under active development (Beta).*
* Hybrid text extraction from PDFs with robust Arabic OCR fallback support.
* Smart segmentation of legal texts (by articles, clauses, etc.).
* Interactive drag-and-drop UI to easily align, split, or delete translation units.
* Generate and export Translation Memory (.TMX) files.


* 🔵 **AI Term Extractor:**
* Automated extraction of substantive legal concepts and categorization (Sub-domain, Document type).
* Strict filtering to exclude procedural and preamble boilerplate phrases.
* Powered by multiple LLMs via OpenRouter (Qwen, Claude, Gemini, Llama, GPT-4o).


* 📊 **Professional Export:**
* Export extracted terminology to a clean Excel glossary.
* Generate beautifully styled PDF terminology flashcards for study and reference.



## 🚀 Prerequisites

Ensure you have the following Python packages installed:

```bash
pip install streamlit pandas openai pdfplumber PyMuPDF pytesseract reportlab arabic-reshaper python-bidi python-docx openpyxl

```

*(Note: To use the OCR fallback features, you must have the `Tesseract OCR` engine installed on your OS).*

## 🛠️ How to Use

1. Run the application locally:
```bash
streamlit run app.py

```


2. Enter your `OpenRouter API Key` in the sidebar.
3. **Alignment Studio (Tab 1):** Upload your PDFs, adjust the alignment using the interactive UI, and download the `TMX` file.
4. **Term Extractor (Tab 2):** Upload your `TMX` file (or use the aligned output from Tab 1), select an AI model, and start the extraction.
5. Download your final Excel glossary or PDF cards from the results section.

## 👨‍💻 Developer

**Developed by:** Youssef Nour Eldin

```

```

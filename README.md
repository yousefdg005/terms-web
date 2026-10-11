⚖️ Legal Terminology Extractor & Flashcard Generator
An intelligent AI-powered web application designed for the Libyan Legal Assembly (Libyan Legislation Network) to extract domain-specific legal terminology from Translation Memory files (.tmx) in English and Arabic. It automatically applies strict filtering rules, standardizes terms, and generates downloadable Excel glossaries and formatted Word terminological cards via OpenRouter API.

📋 Prerequisites & Usage Requirements
To process files successfully, users must provide:

A Parallel TMX File (.tmx): Upload an aligned bilingual translation memory file containing English-Arabic translation units.

OpenRouter API Key: Input a valid OpenRouter API key (sk-or-v1-...) in the sidebar to authorize model requests.

🤖 Models & Account Credit
Supported Suite: Users can select from a curated list of high-precision LLMs via OpenRouter, optimized for Arabic and structured output parsing:

qwen/qwen3.5-flash-02-23 (Qwen 3.5 Flash)

anthropic/claude-haiku-5.5 (Claude Haiku 5.5)

google/gemini-3.8-flash (Gemini 3.8 Flash)

openai/gpt-4o-mini (GPT-4o Mini)

meta-llama/llama-4-maverick (Llama 4 Maverick)

💡 Important Note: Maintaining a small credit balance (e.g., as low as $2.00) on your OpenRouter account completely bypasses free-tier rate limits (Error 429) and prevents service interruption during peak server traffic.

🚀 Key Features
📂 TMX Parsing: Automated extraction and alignment of English-Arabic translation units from uploaded .tmx files.

🛡️ Strict Negative Directives: Automatic filtering of preambles/procedural phrases (e.g., Having reviewed, Pursuant to), dates, and generic administrative words, isolating substantive legal concepts and institutions.

🤖 Multi-LLM Integration: Flexible model selection tailored for legal translation tasks via OpenRouter API.

💾 State Persistence (st.session_state): Protects extracted datasets and interactive preview tables from wiping during file downloads or UI interactions.

🎨 Modern Branded UI/UX: Dark-mode interface with custom CSS styling, tab configuration (terms extractor & icon.png), background image overlay (back.png/back.jpg), and institutional logos (depart.jpg & moma.jpg).

📊 Multi-Format Export:

Excel Spreadsheet (.xlsx): Filtered legal term glossary (automatically excludes the POS column).

Word Document (.docx): Formatted legal terminological cards attributed to the Libyan Legal Assembly.

🛠️ Tech Stack
Language: Python

Framework: Streamlit

LLM Provider: OpenRouter / OpenAI SDK

Data & Document Processing: Pandas, OpenPyXL, Python-docx, ElementTree

👤 Developer & Academic Context
Developer: Yousef Nuralddin Bashir Elmgamez

Department: Translation Studies, Faculty of Languages, University of Tripoli

Reference Body: Libyan Legal Assembly (Libyan Legislation Network - lawsociety.ly)

Contact: +218 92 8756918 | ywsfalmqmz@gmail.com

⚖️ منصة استخراج المصطلحات القانونية وتوليد البطاقات الترجمية
منصة ذكية مخصصة لصالح المجمع القانوني الليبي (شبكة التشريعات الليبية) تعتمد على معالجة اللغات الطبيعية ونماذج الذكاء الاصطناعي لاستخراج المصطلحات القانونية الموضوعية المصفاة من ملفات الذاكرة الترجمية (.tmx)، مع إمكانية تصدير النتائج إلى مسارد Excel وبطاقات معجمية موثقة بصيغة Word.

📋 متطلبات التشغيل والمستلزمات
لتشغيل المنصة واستخراج المصطلحات بنجاح، يحتاج المستخدم إلى:

ملف الذاكرة الترجمية (.tmx): رفع ملف TMX محاذى يحتوي على نصوص موازية باللغتين (إنجليزي - عربي).

مفتاح API من OpenRouter: إدخال مفتاح API الخاص بالمستخدم (sk-or-v1-...) في الشريط الجانبي لتفعيل عملية الاستخراج.

🤖 النماذج ورصيد الحساب
النماذج المتاحة: تتيح المنصة التبديل بين مجموعة نماذج مدفوعة ومتقدمة عبر OpenRouter تتميز بالدقة الفائقة في معالجة اللغة العربية والالتزام بالمخرجات الهيكلية:

qwen/qwen3.5-flash-02-23 (Qwen 3.5 Flash)

anthropic/claude-haiku-5.5 (Claude Haiku 5.5)

google/gemini-3.8-flash (Gemini 3.8 Flash)

openai/gpt-4o-mini (GPT-4o Mini)

meta-llama/llama-4-maverick (Llama 4 Maverick)

💡 ملاحظة هامة: وجود رصيد رمزي بسيط (حتى لو 2 دولار فقط) في حساب OpenRouter الخاص بك يلغي قيود معدل الطلبات (Rate Limits / Error 429) كلياً ويمنع حدوث الأخطاء أو التوقف أثناء أوقات الضغط العالي.

🚀 المميزات الرئيسية
📂 دعم ملفات TMX: قراءة وتحليل وحدات الترجمة المحاذاة (عربي - إنجليزي) فور رفع الملف.

🛡️ تصفية صارمة (Strict Negative Directives): استبعاد تلقائي لعبارات الديباجة والمدمدات الإجرائية (مثل بعد الاطلاع، استناداً إلى) والكلمات الإدارية العامة والتواريخ، والتركيز الحصري على المفاهيم والمؤسسات القانونية الموضوعية.

🤖 دعم متعدد للنماذج: إمكانية اختيار النموذج المناسب لمعالجة النصوص القانونية عبر OpenRouter.

💾 حفظ النتائج وحمايتها (st.session_state): ضمان استمرار عرض النتائج في الجدول التفاعلي ومنع اختفائها أثناء التفاعل أو التحميل.

🎨 واجهة عربية تفاعلية عصرية: تصميم Dark Mode متجاوب مع تنسيقات CSS، عنوان النافذة والأيقونة (terms extractor / icon.png)، خلفية شفافة عصرية (back.png/back.jpg)، وشعارات موثقة للقسم والمجمع (depart.jpg / moma.jpg).

📄 تصدير متعدد الصيغ:

ملف Excel (.xlsx): مسرد مصطلحي مصفى (مستبعد منه عمود POS تلقائياً).

وثيقة Word (.docx): بطاقات ترجمية موثقة ومعتمدة باسم المجمع القانوني الليبي.

🛠️ التقنيات المستخدمة
Python

Streamlit

OpenAI / OpenRouter API

Pandas & OpenPyXL

Python-Docx & ElementTree

👤 إعداد وتطوير
المطور: يوسف نورالدين بشير المقعمز

التخصص: شعبة الترجمة - كلية اللغات - جامعة طرابلس

المرجعية التشريعية: المجمع القانوني الليبي (شبكة التشريعات الليبية - lawsociety.ly)

التواصل: 0928756918 | ywsfalmqmz@gmail.com

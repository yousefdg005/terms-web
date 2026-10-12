⚖️ Legal Terminology Extractor & Flashcard Generator
An intelligent AI-powered web application designed for the Libyan Legal Assembly (Libyan Legislation Network) to extract domain-specific legal terminology from manually aligned Translation Memory files (.tmx) in English and Arabic. It automatically applies advanced preprocessing, strict negative directives, and custom prompt engineering via OpenRouter API to generate downloadable Excel glossaries and formatted Word terminological cards.
📋 Prerequisites & Usage Requirements
To process files successfully, users must provide:
A Manually Aligned Parallel TMX File (.tmx): Upload an aligned bilingual translation memory file containing manually paired English-Arabic translation units.
OpenRouter API Key: Input a valid OpenRouter API key (sk-or-v1-...) in the sidebar to authorize model requests.
🤖 Models & Account Credit
Supported Suite: Users can select from a curated list of high-precision LLMs via OpenRouter, optimized for Arabic and structured output parsing:
qwen/qwen3.5-flash-02-23 (Qwen 3.5 Flash)
anthropic/claude-haiku-5.5 (Claude Haiku 5.5)
google/gemini-3.8-flash (Gemini 3.8 Flash)
openai/gpt-4o-mini (GPT-4o Mini)
meta-llama/llama-4-maverick (Llama 4 Maverick)
💡 Important Note: Maintaining a small credit balance (e.g., as low as $2.00) on your OpenRouter account completely bypasses free-tier rate limits (Error 429) and prevents service interruption during peak server traffic.
⚙️ Terminology Extraction Pipeline (Pre-processing & Process)
To achieve ultimate precision, the extraction workflow follows rigorous processing stages:
Pre-processing Stage (المعالجة الأولية): Reading the translation memory files where users have manually aligned the texts, cleaning raw units, removing XML/XML tags and markup symbols, and filtering out invalid or empty segments to prepare pure parallel text blocks.
Process Stage (المعالجة التحليلية العميقة): Enforcing Strict Negative Directives and advanced Prompt Engineering to automatically discard procedural preambles, administrative filler words, and dates, while guiding the AI models to isolate and standardize substantive legal concepts and institutional titles.
🚀 Key Features
📂 TMX Parsing & Manual Alignment Support: Processing user-aligned bilingual translation units from uploaded .tmx files.
🛡️ Strict Negative Directives & Prompt Engineering: Advanced prompt design combined with strict filtering rules to eliminate preambles/procedural phrases and dates, isolating true legal concepts.
🤖 Multi-LLM Integration: Flexible model selection tailored for legal translation tasks via OpenRouter API.
💾 State Persistence (st.session_state): Protects extracted datasets and interactive preview tables from wiping during file downloads or UI interactions.
🎨 Modern Branded UI/UX: Dark-mode interface with custom CSS styling, tab configuration, background image overlay, and institutional logos (depart.jpg & moma.jpg).
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
منصة ذكية مخصصة لصالح المجمع القانوني الليبي (شبكة التشريعات الليبية) تعتمد على معالجة اللغات الطبيعية ونماذج الذكاء الاصطناعي لاستخراج المصطلحات القانونية الموضوعية المصفاة من ملفات الذاكرة الترجمية (.tmx) التي قام المستخدم بإنشائها ومحاذاتها يدوياً، مع تطبيق هندسة الأوامر وقواعد التصفية الدقيقة، وتصدير النتائج إلى مسارد Excel وبطاقات معجمية موثقة بصيغة Word.
📋 متطلبات التشغيل والمستلزمات
لتشغيل المنصة واستخراج المصطلحات بنجاح، يحتاج المستخدم إلى:
ملف ذاكرة ترجمية محاذى يدوياً (.tmx): رفع ملف TMX ثنائي اللغة يحتوي على وحدات ترجمة تم إعدادها ومحاذاتها يدوياً من قبل المستخدم.
مفتاح API من OpenRouter: إدخال مفتاح API الخاص بالمستخدم (sk-or-v1-...) في الشريط الجانبي لتفعيل عملية الاستخراج.
🤖 النماذج ورصيد الحساب
النماذج المتاحة: تتيح المنصة التبديل بين مجموعة نماذج مدفوعة ومتقدمة عبر OpenRouter تتميز بالدقة الفائقة في معالجة اللغة العربية والالتزام بالمخرجات الهيكلية:
qwen/qwen3.5-flash-02-23 (Qwen 3.5 Flash)
anthropic/claude-haiku-5.5 (Claude Haiku 5.5)
google/gemini-3.8-flash (Gemini 3.8 Flash)
openai/gpt-4o-mini (GPT-4o Mini)
meta-llama/llama-4-maverick (Llama 4 Maverick)
💡 ملاحظة هامة: وجود رصيد رمزي بسيط (حتى لو 2 دولار فقط) في حساب OpenRouter الخاص بك يلغي قيود معدل الطلبات (Rate Limits / Error 429) كلياً ويمنع حدوث الأخطاء أو التوقف أثناء أوقات الضغط العالي.
⚙️ آلية العمل ومراحل المعالجة (Pre-processing & Process)
تعتمد المنصة على خط إنتاج برمجي دقيق يمر بمرحلتين أساسيتين:
مرحلة المعالجة الأولية (Pre-processing): قراءة ملفات الذاكرة الترجمية المحاذاة يدوياً، تنظيف النصوص من وسوم XML والرموز التشويقية، وتصفية الوحدات غير الصالحة للتحليل لضمان نقاء النصوص الموازية.
مرحلة المعالجة والتحليل الذكي (Process): توظيف هندسة الأوامر (Prompt Engineering) المتقدمة مع تطبيق قواعد التصفية الصارمة (Strict Negative Directives) لاستبعاد عبارات الديباجة، المقدمات الإجرائية، والتواريخ، وتوجيه النماذج للتركيز الحصري على استخراج وتوحيد صياغة المفاهيم والمؤسسات القانونية الموضوعية بدقة متناهية.
🚀 المميزات الرئيسية
📂 دعم ملفات TMX المحاذاة يدوياً: معالجة وحدات الترجمة الثنائية التي أعدها المستخدم بعناية.
🛡️ هندسة الأوامر والتصفية الصارمة (Strict Negative Directives & Prompt Engineering): استخدام تعليمات برمجية واضحة لاستبعاد عبارات الديباجة والإجراءات والتواريخ والتركيز على المفاهيم الموضوعية.
🤖 دعم متعدد للنماذج: إمكانية اختيار النموذج المناسب لمعالجة النصوص القانونية عبر OpenRouter.
💾 حفظ النتائج وحمايتها (st.session_state): ضمان استمرار عرض النتائج في الجدول التفاعلي ومنع اختفائها أثناء التفاعل أو التحميل.
🎨 واجهة عربية تفاعلية عصرية: تصميم Dark Mode متجاوب مع تنسيقات CSS، عنوان النافذة والأيقونة، خلفية شفافة عصرية، وشعارات موثقة للقسم والمجمع.
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
# ⚖️ Legal Terminology Extractor & Flashcard Generator

An intelligent AI-powered web application designed to extract domain-specific legal terminology from Translation Memory files (`.tmx`) in English and Arabic. It automatically generates standardized terminological glossaries and downloadable Word flashcards using Large Language Models via OpenRouter API.

## 📋 Prerequisites & Usage Requirements
To process files successfully, users must provide:
1. **A Parallel TMX File (`.tmx`):** Must contain aligned bilingual translation units (English - Arabic).
2. **OpenRouter API Key:** Users need to input their own OpenRouter API key (`sk-or-v1-...`) in the sidebar to authorize requests.

## 🤖 Models & Account Credit
- **Model Support:** The application supports both fully free models and paid models via OpenRouter.
- 💡 **Important Note:** Maintaining a small credit balance (e.g., as low as **$2.00**) on your OpenRouter account completely bypasses free-tier rate limits (Error 429) and prevents errors during peak server traffic.

## 🚀 Key Features
- 📂 **TMX Parsing:** Automated extraction and alignment of English-Arabic translation units.
- 🤖 **Multi-LLM Integration:** Seamless integration with OpenRouter API (supporting Gemma, Llama 3.3, Qwen 2.5, and Auto-Free Routers).
- 🎨 **Modern Arabic UI/UX:** Dark-mode interface with native RTL support, responsive mobile navigation, and customizable settings.
- 📊 **Multi-Format Export:** Export extracted legal terms into structured Excel spreadsheets (`.xlsx`) and formatted Word documents (`.docx`).

## 🛠️ Tech Stack
- **Language:** Python
- **Framework:** Streamlit
- **LLM Provider:** OpenRouter / OpenAI SDK
- **Data Processing:** Pandas, OpenPyXL, Python-docx

## 👤 Developer & Academic Context
- **Developer:** Yousef Nuralddin Elmgamez
- **Department:** Translation Studies, Faculty of Languages, University of Tripoli
- # ⚖️ منصة استخراج المصطلحات القانونية وتوليد البطاقات الترجمية

منصة ذكية تعتمد على معالجة اللغات الطبيعية ونماذج الذكاء الاصطناعي لاستخراج المصطلحات القانونية المتخصصة من ملفات الذاكرة الترجمية (`.tmx`)، مع إمكانية تصدير النتائج إلى ملفات Excel وبطاقات معجمية بصيغة Word.

## 📋 متطلبات التشغيل والمستلزمات
لتشغيل المنصة واستخراج المصطلحات بنجاح، يحتاج المستخدم إلى:
1. **ملف الذاكرة الترجمية (`.tmx`):** يحتوي على نصوص موازية ومحاذاة باللغتين (إنجليزي - عربي).
2. **مفتاح API خاص من OpenRouter:** يتوجب إدخال مفتاح API الخاص بالمستخدم (`sk-or-v1-...`) في الشريط الجانبي لتفعيل المعالجة.

## 🤖 النماذج ورصيد الحساب
- **النماذج المتاحة:** تتيح المنصة الاختيار بين نماذج مجانية بالكامل وأخرى مدفوعة فائقة السرعة والدقة عبر OpenRouter.
- 💡 **ملاحظة هامة:** وجود رصيد رمزي بسيط (حتى لو **2 دولار فقط**) في حساب OpenRouter الخاص بك يلغي قيود معدل الطلبات (Rate Limits / Error 429) كلياً ويمنع حدوث الأخطاء أو التوقف أثناء أوقات الذروة والضغط العالي.

## 🚀 المميزات الرئيسية
- 📂 **دعم ملفات TMX:** تحليل وحساب وحدات الترجمة المحاذاة (عربي - إنجليزي).
- 🤖 **دعم متعدد للنماذج:** التكامل مع خوادم OpenRouter المتقدمة (مثل Gemma, Llama 3.3, Qwen 2.5, و Auto Free Router).
- 🎨 **واجهة عربية تفاعلية:** تصميم عصري يدعم اتجاه (RTL) والقائمة الجانبية المخصصة.
- 📄 **تصدير متعدد الصيغ:** استخراج البيانات كملفات Excel ووثائق Word مقسمة ببطاقات ترجمية.

## 🛠️ التقنيات المستخدمة
- **Python**
- **Streamlit**
- **OpenAI / OpenRouter API**
- **Pandas & OpenPyXL**
- **Python-Docx**

## 👤 إعداد وتطوير
- **المطور:** يوسف نورالدين المقعمز
- **التخصص:** شعبة الترجمة - كلية اللغات - جامعة طرابلس
- 

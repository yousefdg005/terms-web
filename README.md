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

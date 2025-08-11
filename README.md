# 📊 PPTX Inconsistency Finder

An **AI-enabled Python tool** that processes any multi-slide PowerPoint presentation (`.pptx`) and flags **factual or logical inconsistencies** across slides.

It can detect:

- 🔢 **Conflicting numerical data**  
  e.g., revenue figures that don’t match, percentages that don’t add up.
- 📄 **Contradictory textual claims**  
  e.g., "Market is highly competitive" vs "Few competitors".
- 📅 **Timeline mismatches**  
  e.g., conflicting dates or forecasts.

The tool uses **Google Gemini 2.5 Flash** for AI-powered cross-slide analysis and **OCR** for text extraction from images inside slides.

---

## 📂 Features
- Works with `.pptx` PowerPoint files.
- Extracts text from both native text boxes and images (via OCR).
- Detects factual/logical inconsistencies across slides.
- AI-assisted reasoning via Gemini 2.5 Flash API.
- Option to run **without AI** for deterministic checks only.
- Outputs a **structured JSON report** with slide references.

---

## 🚀 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/YOUR_USERNAME/ppt_inconsistency_finder.git
   cd ppt_inconsistency_finder

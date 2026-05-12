# 📄 Industrial OCR Pipeline: Azure AI & GPT-4

[![Azure](https://img.shields.io/badge/Azure-Document%20Intelligence-0089D6?style=for-the-badge&logo=microsoft-azure)](https://azure.microsoft.com/en-us/services/form-recognizer/)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4-412991?style=for-the-badge&logo=openai)](https://openai.com/)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python)](https://www.python.org/)

An enterprise-grade document extraction pipeline that leverages **Azure Document Intelligence** (formerly Form Recognizer) for layout-aware OCR and **Azure OpenAI (GPT-4)** for intelligent data structuring. This tool is designed to process complex industrial certificates, invoices, and forms into highly formatted, multi-sheet Excel reports.

---

## 🚀 Key Features

- **🔍 Layout-Aware Extraction**: Uses Azure's `prebuilt-layout` model to accurately capture text, nested tables, and selection marks (checkboxes).
- **🧠 Semantic Intelligence**: Employs GPT-4 to interpret raw OCR text, mapping it to structured business fields like Part Numbers, Serial Numbers, and Customer Details.
- **📊 Consolidated Reporting**: Automatically generates a professional Excel workbook with:
  - **Certificate Details**: A high-level overview of all processed documents.
  - **Equipment Line Items**: Granular table data extracted from each form.
  - **Consolidated Register**: A unified database of all parts across the entire batch.
- **✨ Automated Formatting**: Produces publication-ready Excel files with frozen panes, auto-adjusted columns, and professional styling.

---

## 📁 Project Architecture

```text
ocr3/
├── data/
│   ├── input/          # Place raw PDFs/Images here
│   └── output/         # Structured Excel reports are generated here
├── utils/
│   ├── azure_ocr.py     # Document Intelligence API wrapper
│   ├── azure_openai.py  # GPT-4 structuring logic
│   └── excel_exporter.py # Professional Excel formatting engine
├── main.py              # Orchestration entry point
├── .env                 # API Credentials (ignored by git)
└── requirements.txt     # Dependency manifest
```

---

## 🛠️ Setup & Installation

### 1. Prerequisites
- Python 3.8 or higher
- Azure Subscription with:
  - Document Intelligence Resource
  - Azure OpenAI Resource (or standard OpenAI API)

### 2. Clone & Install
```bash
git clone <your-repo-url>
cd ocr3
pip install -r requirements.txt
```

### 3. Configuration
Create a `.env` file in the root directory and add your credentials:
```env
AZURE_FORM_RECOGNIZER_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_FORM_RECOGNIZER_KEY=your_key_here

AZURE_OPENAI_API_KEY=your_key_here
AZURE_OPENAI_ENDPOINT=https://your-openai.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4
```

---

## 📖 Usage

1. **Prepare Data**: Drop your PDF or Image files (JPG, PNG) into the `data/input/` directory.
2. **Execute Pipeline**:
   ```bash
   python main.py
   ```
3. **Review Results**: Open `data/output/consolidated_ocr_report.xlsx` to view your structured data.

---

## 🛡️ Security & Privacy
This pipeline is designed for enterprise use. By using Azure AI services, your data remains within your Azure boundary, ensuring compliance with data privacy regulations. **Never commit your `.env` file to version control.**

---

## ⚖️ License
Distributed under the MIT License. See `LICENSE` for more information.

---

<p align="center">
  Built with ❤️ for High-Performance Document Processing
</p>

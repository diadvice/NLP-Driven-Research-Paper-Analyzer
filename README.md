# NLP-Driven Research Paper Analyzer

## Overview
The **NLP-Driven Research Paper Analyzer** is a Python-based tool designed to extract key information from research papers in PDF and Word formats. Leveraging natural language processing (NLP) techniques, this project automates the extraction of essential data such as authors, publication dates, conclusions, keywords, and GitHub links, streamlining the research workflow.

## Features
- **Text Extraction**: Efficiently extracts text from PDF and Word documents.
- **NLP Integration**: Utilizes SpaCy for identifying authors and publication dates.
- **Conclusion Parsing**: Smartly extracts conclusions, stopping at specified keywords like "Acknowledgments" or "References."
- **Google Sheets Automation**: Integrates with Google Sheets to automatically update a research database.
- **Keyword Extraction**: Identifies and compiles the most relevant keywords from the text.

## Technology Stack
- **Python**: Main programming language used for development.
- **Libraries**:
  - `PyPDF2`: For PDF text extraction.
  - `python-docx`: For Word document text extraction.
  - `SpaCy`: For natural language processing.
  - `gspread`: For interacting with Google Sheets.
  - `oauth2client`: For Google Sheets API authentication.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/NLP-Driven-Research-Paper-Analyzer.git
   cd NLP-Driven-Research-Paper-Analyzer

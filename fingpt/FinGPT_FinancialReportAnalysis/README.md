# Financial Report Analysis Project

## Overview

This project provides tools for analyzing financial reports, specifically annual reports (10-K), using advanced language models such as GPT-4 or other locally deployed Large Language Models (LLM). It's designed to help users generate comprehensive analysis reports in PDF format, offering insights into a company's financial health and performance over the fiscal year.

## Features

- **PDF Report Generation**: Automatically generate detailed analysis reports in PDF format for annual financial statements.
- **GPT-4 and LLM Support**: Utilize the power of GPT-4 or any locally deployed LLM for deep and insightful analysis.
- **RAG Support**: The ability to utilize the power of RAG for question-answering and summarization tasks.
- **Automated Financial Tables**: Extract key metrics from uploaded PDFs and populate an Excel workbook using
  `utils.financial_report_parser`.
- **Customizable Analysis**: Users can modify the analysis scope by choosing different company symbols and models.
- **Easy to Use**: Designed with simplicity in mind, simply run all cells in the provided notebook to get your report.

## Requirements

Before starting, ensure you have the following installed:
- Python 3.11 or later
- Jupyter Notebook
- Necessary Python packages (pandas, matplotlib, openai, etc.)

Obtain the sec-api (which is used to grab the 10-K report) from https://sec-api.io/profile for free.

(Optional) Obtain the fmp api for target price (paid) from https://site.financialmodelingprep.com/developer/docs/dashboard.

## Getting Started

To begin analyzing financial reports:

0. **(optional) Prepare the local LLM**:
   If you want to run the analysis with the locally deployed models, please download Ollama and have it running: https://ollama.com/download.
   Also, download the model you want to use in the list of available models: https://ollama.com/library with command:
   ```bash
    ollama run <model_name>
    ```

1. **Open the Notebook**:
   Launch Jupyter Notebook and open the `reportanalysis.ipynb` file:
   ```
   jupyter notebook reportanalysis.ipynb
   ```
   All the necessary libraries and dependencies are already imported in the notebook.

2. **Configure the Notebook**:
   Modify the `company symbol` and `models` variables within the notebook to suit the analysis you wish to perform.

3. **Run the Analysis**:
   Execute all cells in the notebook to generate your financial report analysis in PDF format.

## Extracting Structured Tables from PDFs

The module `fingpt.FinGPT_FinancialReportAnalysis.utils.financial_report_parser` automates the process of extracting
numbers from annual reports and populating them into an Excel template.

### Super-simple "do it on my Mac" guide

If computers feel confusing, follow these baby steps exactly. Each line that looks like `this` is something you copy
and paste into the **Terminal** app on your Mac.

1. **Open Terminal**  
   Click the smiling face in your dock → type “Terminal” in the search box → press **Return**.

2. **Make sure Python is ready**  
   In the Terminal window type:
   ```bash
   python3 --version
   ```
   If you see a version number (for example `Python 3.11.5`), you are good. If it says “command not found,” download
   and install Python from https://www.python.org/downloads/mac-osx/, then open Terminal again.

3. **Go to the project folder**  
   If you downloaded this project as a zip, double-click the zip in Finder once. Then in Terminal type `cd `
   (with a space at the end) and drag the project folder into the Terminal window so the path appears, then press
   **Return**.

4. **Create a safe sandbox for Python packages**  
   ```bash
   python3 -m venv venv
   ```

5. **Turn the sandbox on**  
   ```bash
   source venv/bin/activate
   ```
   You will know it worked because you will see `(venv)` at the start of the Terminal line.

6. **Teach Python all the tricks it needs**  
   ```bash
   pip install -r requirements.txt
   ```

7. **Put your PDF reports in a folder**  
   Make a new folder (for example `my_pdfs`) and move every report you want to read into that folder.

8. **Run the magic sorter**  
   ```bash
   python3 -m fingpt.FinGPT_FinancialReportAnalysis.utils.financial_report_parser \
       /path/to/my_pdfs \
       /path/to/your_template.xlsx \
       /path/to/save_results.xlsx
   ```
   *Tip:* the easiest way to fill in each `/path/to/...` is to type the command, then drag the folder or file from
   Finder into the Terminal window.

9. **Open the finished Excel file**  
   Double-click the file you picked in the last line (for example `save_results.xlsx`) to see all the numbers filled in
   for you. Check the `ExtractionNotes` sheet to review any numbers the script could not find.

Take a breath—you just ran the whole project from your Mac Terminal!

1. Create an Excel workbook that contains a sheet called `Mapping`. Each row of this sheet should define a financial
   category with the following columns:

   | Column | Description |
   | --- | --- |
   | `Category` | Display name of the metric (e.g. `Total Revenue`). |
   | `Pattern` | Semicolon-separated keywords to locate the number in the PDF. Wrap custom regular expressions in `/` (for example, `/Revenue\s+\$?([\d,.]+)/`). |
   | `Sheet` *(optional)* | Worksheet where the value should be written. Defaults to `AutoFilled`. |
   | `Cell` *(optional)* | Cell address (e.g. `B4`). When omitted the writer looks for the row that contains the category name and fills the next column. |
   | `Scale` *(optional)* | Multiplier applied to the parsed value (for example `1e6` when numbers are reported in millions). |
   | `Aggregation` *(optional)* | How to aggregate multiple matches (`first`, `sum`, or `average`). |

2. Place one or more PDF files that contain the financial statements inside a directory.
3. Run the command below to populate a copy of your template:

   ```bash
   python -m fingpt.FinGPT_FinancialReportAnalysis.utils.financial_report_parser \
       /path/to/pdf_directory \
       /path/to/template.xlsx \
       /path/to/output.xlsx
   ```

The script creates a worksheet called `ExtractionNotes` that summarises every match and highlights the categories for
which no numbers were detected, making the review process straightforward.

## Contributing

We welcome contributions and suggestions! Please open an issue or submit a pull request with your improvements.

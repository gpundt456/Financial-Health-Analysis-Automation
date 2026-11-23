# Financial-Health-Analysis-Automation

This is an n8n workflow that extracts financial data from documents such as balance sheets, income statements, and cash flow statements and turns the extracted data into a readable financial health analysis Excel workbook.

## Why I Built This

I wanted to build an automation that could take financial statements and turn them into a report that someone without a financial background could understand.

Financial statements contain a lot of useful information, but understanding financial ratios, comparing different reporting periods, and identifying potential strengths or weaknesses can be time-consuming.

I built this workflow to automate that process using n8n, locally hosted LLMs, JavaScript, and Python. The goal was to create something that could extract financial information, perform calculations, explain the results in plain English, and automatically generate and deliver an Excel report.

## Features

- **Financial document uploads:** Uses an n8n form where users can submit financial statements and enter an email address.
- **Multiple file formats:** Supports PDF, PNG, JPG, XLSX, and CSV financial documents.
- **Multiple statement types:** Processes balance sheets, income statements, and cash flow statements.
- **Local AI document extraction:** Uses LLMs running through Ollama to identify financial statements and extract financial data.
- **Financial data validation:** Checks extracted financial information for missing values, inconsistencies, and accounting discrepancies.
- **Data consolidation:** Organizes financial accounts and reporting periods into a structured dataset.
- **Financial ratio calculations:** Calculates available financial metrics related to liquidity, leverage, profitability, and cash flow.
- **Financial health scoring:** Generates a financial health score based on the available financial information.
- **Data coverage reporting:** Identifies which financial categories can and cannot be assessed when statements are missing.
- **AI financial explanations:** Uses a local LLM to explain financial results in language intended for readers without a financial background.
- **Automated Excel reporting:** Uses Python to generate a formatted financial analysis workbook with charts, financial statements, ratios, and explanations.
- **Automatic email delivery:** Sends the completed Excel report to the email address entered in the submission form.

## Technologies Used

- **n8n:** Workflow automation and document processing
- **Ollama:** Running local large language models
- **Qwen 3.5 9B:** Text-based financial data extraction and financial analysis
- **Qwen3-VL 8B Instruct:** Vision-based extraction from financial statement images
- **JavaScript:** Data parsing, validation, consolidation, and financial calculations within n8n
- **Python:** Automated Excel report generation
- **FastAPI:** Local HTTP service connecting n8n to the Python report generator
- **openpyxl:** Creating and formatting Excel workbooks
- **SMTP:** Automated report delivery by email

## How It Works

1. **Document submission:** A user opens the n8n form and uploads a balance sheet, with the option to include an income statement and cash flow statement. The user also provides an email address.
2. **Document classification:** The workflow identifies the uploaded files and routes them through the appropriate extraction process.
3. **AI extraction:** Ollama processes the extracted text or images and converts the financial information into structured JSON.
4. **Validation:** JavaScript Code nodes check the extracted statements, reporting periods, numerical values, and accounting relationships.
5. **Consolidation:** Financial information from the uploaded documents is combined into a standardized dataset.
6. **Financial calculations:** The workflow calculates financial ratios, evaluates available financial metrics, and produces a financial health score.
7. **AI interpretation:** A local LLM analyzes the calculated financial results and generates explanations, observations, and limitations.
8. **Excel generation:** The structured financial information is sent to a local Python FastAPI server, which generates a formatted Excel workbook.
9. **Email delivery:** The completed workbook is returned to n8n and emailed to the user as an attachment.

## Example Output

The generated Excel workbook includes six main worksheets:

| Worksheet | Purpose |
|---|---|
| Overview | Summary of financial condition, health score, and major findings |
| Health Checks | Financial ratios, performance indicators, and scoring details |
| Statements | Extracted financial statement data and comparisons |
| Insights | AI-generated analysis, observations, and recommendations |
| Money Words | Plain-English explanations of financial terminology |
| About This Report | Methodology, data coverage, validation results, and limitations |

The report is intended to make financial information easier to understand without requiring the reader to interpret every accounting figure manually.

## Requirements

This project is designed to run locally rather than relying on a paid cloud LLM API.

To replicate the workflow, you will need:

- Python 3.10 or newer
- Node.js
- A local n8n installation
- Ollama
- The required Qwen models downloaded through Ollama
- The Python Excel report generator included in this repository
- An SMTP email account configured in n8n
- A public HTTPS endpoint if the upload form needs to be accessible outside your local network

A computer capable of running the local language models is also required. Performance will depend on the available system RAM, GPU, and model configuration.

## Installing the Local LLMs

This workflow uses two Ollama models:

**Qwen 3.5 9B** for processing financial text and generating financial explanations.

```bash
ollama pull qwen3.5:9b
```

**Qwen3-VL 8B Instruct** for extracting financial information from images.

```bash
ollama pull qwen3-vl:8b-instruct
```

The workflow is configured to communicate with Ollama at:

```text
http://127.0.0.1:11434/api/generate
```

Both models must be installed and available through Ollama for the corresponding workflow branches to function.

If using different model versions, the model names in the n8n HTTP Request nodes must be updated.

## Python Excel Report Generator

The workflow uses a separate Python program to generate the final Excel workbook.

This is necessary because n8n performs the extraction, calculations, and analysis, while Python handles the more advanced spreadsheet formatting and report creation.

Install the Python dependencies:

```bash
pip install fastapi uvicorn openpyxl
```

Start the Python report service from the directory containing `excel_server.py`:

```bash
python -m uvicorn excel_server:app --host 127.0.0.1 --port 8000
```

The n8n workflow sends the report data to:

```text
http://127.0.0.1:8000/generate-report
```

The Python service returns an Excel workbook, which n8n attaches to the outgoing email.

The server must be running whenever the workflow generates a report.

## Setting Up the n8n Workflow

1. Install and start n8n locally.
2. Download the workflow JSON file from this repository.
3. Open n8n and import the workflow using **Import from File**.
4. Install and start Ollama.
5. Download the required Qwen models.
6. Start the Python Excel report generator.
7. Configure your SMTP email credentials in the **Send an Email** node.
8. Update the sender email address to your own address.
9. Open the form trigger and obtain your own form URL.
10. Activate the workflow and test it using publicly available or fictional financial statements.

### Important: Local Configuration

The exported workflow is configured around a locally hosted environment.

Its Ollama and Python HTTP Request nodes use `127.0.0.1`, which refers to the computer running the services.

If n8n is running on a different machine or inside a container, these addresses may need to be changed.

The workflow will not automatically work through someone else's computer just by importing the JSON. The required local services, models, and email credentials must also be configured.

## Form Trigger and Public Access

The workflow starts with an n8n Form Trigger.

The form allows users to upload financial documents and specify where the completed report should be delivered.

n8n generates form URLs for the workflow's configured environment. A person importing this project should obtain their own test or production form URL rather than using the URL from my local installation.

For access over the internet, a public HTTPS endpoint is required.

I use a locally hosted n8n environment with zrok tunneling for public webhook access. The separate launcher project can be found here:

[n8n-zrok-launcher](https://github.com/gpundt456/n8n-zrok-launcher)

The launcher project is separate from the financial analysis workflow and isn't required when another suitable hosting or HTTPS tunneling solution is available.

## Limitations

- AI extraction is not guaranteed to be accurate and should be checked against the original financial statements.
- Results depend on the quality, formatting, and completeness of the uploaded documents.
- Missing financial statements can limit which ratios and categories can be calculated.
- Financial health scores are generated using the workflow's own scoring methodology and should not be interpreted as official credit ratings.
- Some documents may require manual review when extracted values are inconsistent or incomplete.
- Local LLM inference can be resource-intensive, especially when processing large documents or images.
- The analysis is intended for informational and educational purposes and is not a substitute for professional financial, accounting, or investment advice.

## Security and Privacy

The workflow is intended to process publicly available, fictional, or otherwise non-confidential financial documents.

The public repository should not contain personal financial documents, real user submissions, email credentials, passwords, or private API keys.

Although the LLM inference runs locally, the form and email delivery involve network services. Users should configure appropriate authentication, access controls, and data retention settings before using the workflow with real financial information.

The exported workflow should also be reviewed for account-specific configuration before being imported into another n8n installation.

## Development Challenges

Some of the problems I worked through while developing this project included:

- Getting local LLMs to reliably extract structured financial information from different document formats.
- Handling inconsistent account names across financial statements.
- Validating accounting relationships without incorrectly rejecting small rounding differences.
- Combining financial data from multiple uploaded documents and reporting periods.
- Distinguishing financial ratios expressed as percentages from ratios expressed as multiples.
- Creating a financial health score when some categories cannot be evaluated.
- Preventing missing data from being presented as zero or as a fully assessed financial category.
- Generating professional Excel reports that remain understandable to readers without a financial background.
- Connecting n8n to a separate Python reporting service and returning the generated workbook for email delivery.

## AI Assistance

Although I conceptualized this project idea and data extraction structure entirely on my own, this project was developed with AI assistance, including help writing Python and JavaScript code, developing system prompts, troubleshooting errors, and improving the Excel report generator.

I designed the automation around my own goals, connected the different systems, tested workflows, identified problems in the outputs, and iteratively improved the project with AI assistance.

AI tools were used throughout development rather than only for individual code suggestions.

The project represents my experience building and troubleshooting a complete local AI automation system, including document extraction, structured data processing, financial analysis, and automated report generation. 

InvoiceLens

## AI-Powered Invoice Tracking and Analysis

InvoiceLens is a simple tool that helps small-business owners track invoices and understand what needs their attention.

Instead of manually reviewing a spreadsheet, users can upload invoice data, view a dashboard, filter invoices, and ask questions in natural language.

## What It Does

* Tracks invoices and their payment status
* Shows total invoice value and outstanding amount
* Identifies overdue invoices
* Allows users to filter invoices
* Answers invoice-related questions using natural language
* Helps business owners quickly identify what needs attention

## Example Questions

Users can ask:

* "Where is my money stuck?"
* "Which invoices need my attention?"
* "Which suppliers do we owe money to?"
* "Are there any invoices that look problematic?"
* "What should I follow up on?"

## How It Works

1. User uploads invoice data.
2. Python reads and cleans the data.
3. InvoiceLens calculates key invoice metrics.
4. The user asks a question in normal language.
5. The AI interprets the question using the invoice data.
6. InvoiceLens returns a simple business-friendly answer.

The AI is instructed to use the supplied invoice data and avoid inventing invoice facts.

## Technology

* Python
* Streamlit
* Pandas
* Ollama
* Llama 3.2
* OpenPyXL
* PyPDF
* python-docx

## How to Run Locally

### 1. Activate the virtual environment

On Windows:

```powershell
.venv\Scripts\activate
```

### 2. Install the required packages

```powershell
pip install -r requirements.txt
```

### 3. Start Ollama

Make sure the Llama 3.2 model is installed:

```powershell
ollama pull llama3.2
```

Start the model:

```powershell
ollama run llama3.2
```

Keep this running.

### 4. Start InvoiceLens

Open another terminal in the project folder:

```powershell
streamlit run app.py
```

The application will open in your browser.

## Sample Data

The included `invoice.xlsx` file contains fictional invoice data created for demonstration purposes.

No real customer or confidential financial information is used.

## Project Limitations

This is a prototype created for demonstration purposes.

The current version uses a local AI model through Ollama.

A production version could include:

* Database storage
* User authentication
* Cloud deployment
* Email payment reminders
* Automatic invoice ingestion
* Accounting software integrations
* More advanced duplicate detection
* Audit logs

## Future Improvements

The next version could make the system more reliable by separating AI question interpretation from the actual financial calculations.

For example:

**User question → AI understands intent → Python calculates the facts → AI explains the result**

This would reduce the possibility of AI generating incorrect financial information.

## Project Goal

InvoiceLens is not intended to replace accounting software.

Its purpose is to make invoice information easier to understand and help business owners quickly identify what requires attention.

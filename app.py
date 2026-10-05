import re
import streamlit as st
import pandas as pd
import ollama

from pypdf import PdfReader
from docx import Document


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="InvoiceLens",
    page_icon="🐬",
    layout="wide"
)


# =========================================================
# HEADER
# =========================================================

st.title("InvoiceLens")
st.caption("Your invoices, understood.")

st.divider()


# =========================================================
# READ PDF
# =========================================================

def read_pdf(file):

    reader = PdfReader(file)

    pages = []

    for page_number, page in enumerate(
        reader.pages,
        start=1
    ):

        text = page.extract_text() or ""

        pages.append({
            "page": page_number,
            "text": text
        })

    return pages


# =========================================================
# READ WORD
# =========================================================

def read_word(file):

    document = Document(file)

    parts = []

    # Read paragraphs
    for paragraph in document.paragraphs:

        text = paragraph.text.strip()

        if text:
            parts.append(text)

    # Read tables
    for table in document.tables:

        for row in table.rows:

            values = []

            for cell in row.cells:

                value = cell.text.strip()

                if value:
                    values.append(value)

            if values:

                parts.append(
                    " | ".join(values)
                )

    return "\n".join(parts)


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(file):

    return pd.read_excel(file)


# =========================================================
# CLEAN AMOUNT
# =========================================================

def clean_amount(value):

    if value is None:
        return None

    try:

        if pd.isna(value):
            return None

    except:
        pass

    value = str(value)

    value = (
        value
        .replace("₹", "")
        .replace("$", "")
        .replace("€", "")
        .replace("£", "")
        .replace("Rs.", "")
        .replace("Rs", "")
        .replace(",", "")
        .replace("■", "")
        .strip()
    )

    match = re.search(
        r"\d+(?:\.\d+)?",
        value
    )

    if not match:
        return None

    try:
        return float(match.group())

    except:
        return None


# =========================================================
# CLEAN DATE
# =========================================================

def clean_date(value):

    if value is None:
        return pd.NaT

    try:

        if pd.isna(value):
            return pd.NaT

    except:
        pass

    return pd.to_datetime(
        str(value),
        errors="coerce",
        dayfirst=True
    )


# =========================================================
# FIND INVOICE NUMBER
# =========================================================

def find_invoice_number(text):

    # -----------------------------------------------------
    # FIRST: LOOK FOR "INVOICE NO"
    # -----------------------------------------------------

    patterns = [

        r"Invoice\s+No\.?\s*[:#\-]?\s*([A-Za-z0-9][A-Za-z0-9./_-]*)",

        r"Invoice\s+Number\s*[:#\-]?\s*([A-Za-z0-9][A-Za-z0-9./_-]*)",

        r"Invoice\s+#\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9./_-]*)",

        r"Invoice\s+ID\s*[:#\-]?\s*([A-Za-z0-9][A-Za-z0-9./_-]*)",

        r"Tax\s+Invoice\s*[:#\-]?\s*([A-Za-z0-9][A-Za-z0-9./_-]*)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            return match.group(1).strip()


    # -----------------------------------------------------
    # FALLBACK
    # -----------------------------------------------------

    fallback_patterns = [

        # OTS/2026-020
        r"\b[A-Z]{2,10}/\d{4}[-/]\d+\b",

        # INV-1001
        r"\b[A-Z]{2,10}-\d{3,}\b",

        # SI/2026/001
        r"\b[A-Z]{1,10}/\d{2,4}/\d{2,}\b",

        # ABC_2026_001
        r"\b[A-Z]{2,10}_\d{2,4}_\d{2,}\b"
    ]

    for pattern in fallback_patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            return match.group(0).strip()


    return None


# =========================================================
# FIND DATE
# =========================================================

def find_date(text, labels):

    for label in labels:

        pattern = (
            rf"{re.escape(label)}"
            r"\s*[:\-]?\s*"
            r"(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})"
        )

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            return clean_date(
                match.group(1)
            )

    return pd.NaT


# =========================================================
# FIND VALUE AFTER LABEL
# =========================================================

def find_value(
    text,
    labels,
    stop_labels=None
):

    if stop_labels is None:
        stop_labels = []


    for label in labels:

        # -------------------------------------------------
        # If we know what comes after the value
        # -------------------------------------------------

        if stop_labels:

            stops = "|".join(
                re.escape(x)
                for x in stop_labels
            )

            pattern = (
                rf"{re.escape(label)}"
                rf"\s*[:\-]?\s*"
                rf"(.*?)"
                rf"(?=\s+(?:{stops})\b|$)"
            )

        else:

            pattern = (
                rf"{re.escape(label)}"
                rf"\s*[:\-]?\s*"
                rf"([^\n]+)"
            )


        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = re.sub(
                r"\s+",
                " ",
                value
            )

            return value


    return None


# =========================================================
# FIND VENDOR / CUSTOMER
# =========================================================

def find_vendor(text):

    # Customer
    value = find_value(
        text,
        [
            "Customer",
            "Bill To",
            "Client"
        ],
        [
            "Customer GSTIN",
            "GSTIN",
            "Place",
            "Invoice Date",
            "Due Date",
            "Payment Terms"
        ]
    )

    if value:
        return value


    # Vendor
    value = find_value(
        text,
        [
            "Vendor",
            "Supplier",
            "Seller",
            "From"
        ],
        [
            "GSTIN",
            "Invoice Date",
            "Due Date",
            "Payment Terms"
        ]
    )

    if value:
        return value


    return "Unknown"


# =========================================================
# FIND AMOUNT
# =========================================================

def find_amount(text):

    patterns = [

        r"Grand\s+Total\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"Total\s+Amount\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"Invoice\s+Total\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"Net\s+Amount\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"Total\s*[:\-]?\s*[₹$€€£]?\s*([\d,]+(?:\.\d+)?)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            return clean_amount(
                match.group(1)
            )


    return None


# =========================================================
# FIND GST / TAX
# =========================================================

def find_gst(text):

    patterns = [

        r"GST\s*(?:@\s*\d+(?:\.\d+)?\s*%)?\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"GST\s+Amount\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)",

        r"Tax\s+Amount\s*[:\-]?\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            return clean_amount(
                match.group(1)
            )


    return None


# =========================================================
# FIND PAYMENT TERMS
# =========================================================

def find_payment_terms(text):

    return find_value(
        text,
        [
            "Payment Terms",
            "Payment Term",
            "Terms"
        ],
        [
            "S.No.",
            "S.No",
            "Item",
            "Description",
            "Amount",
            "Total"
        ]
    )


# =========================================================
# FIND PAYMENT STATUS
# =========================================================

def find_status(
    text,
    due_date
):

    lower_text = text.lower()


    # -----------------------------------------------------
    # EXPLICIT STATUS
    # -----------------------------------------------------

    status_patterns = [

        r"Payment\s+Status\s*[:\-]?\s*(Paid|Pending|Unpaid|Overdue|Settled|Due)",

        r"Status\s*[:\-]?\s*(Paid|Pending|Unpaid|Overdue|Settled|Due)"
    ]


    for pattern in status_patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:

            status = match.group(1).lower()

            if status in [
                "paid",
                "settled"
            ]:

                return "Paid"

            if status == "overdue":

                return "Overdue"

            return "Pending"


    # -----------------------------------------------------
    # STATUS WORDS
    # -----------------------------------------------------

    if re.search(
        r"\bfully\s+paid\b",
        lower_text
    ):

        return "Paid"


    if re.search(
        r"\bpaid\b",
        lower_text
    ):

        return "Paid"


    if re.search(
        r"\boverdue\b",
        lower_text
    ):

        return "Overdue"


    if re.search(
        r"\bunpaid\b",
        lower_text
    ):

        return "Pending"


    if re.search(
        r"\bpending\b",
        lower_text
    ):

        return "Pending"


    # -----------------------------------------------------
    # USE DUE DATE
    # -----------------------------------------------------

    if due_date is not None:

        try:

            if not pd.isna(due_date):

                today = pd.Timestamp.today().normalize()

                if due_date < today:

                    return "Overdue"

                return "Pending"

        except:
            pass


    return "Unknown"


# =========================================================
# PARSE ONE INVOICE
# =========================================================

def parse_invoice(
    text,
    page=None
):

    invoice_no = find_invoice_number(
        text
    )

    if not invoice_no:

        return None


    invoice_date = find_date(
        text,
        [
            "Invoice Date",
            "Invoice Dt",
            "Bill Date"
        ]
    )


    # Only use generic "Date" if needed
    if pd.isna(invoice_date):

        invoice_date = find_date(
            text,
            ["Date"]
        )


    due_date = find_date(
        text,
        [
            "Due Date",
            "Payment Due",
            "Due"
        ]
    )


    vendor = find_vendor(
        text
    )


    amount = find_amount(
        text
    )


    gst = find_gst(
        text
    )


    payment_terms = find_payment_terms(
        text
    )


    status = find_status(
        text,
        due_date
    )


    return {

        "Invoice No": invoice_no,

        "Vendor / Customer": vendor,

        "Invoice Date": invoice_date,

        "Due Date": due_date,

        "Amount": amount,

        "GST / Tax": gst,

        "Payment Terms": (
            payment_terms
            if payment_terms
            else "Unknown"
        ),

        "Status": status,

        "Page": page
    }


# =========================================================
# PDF PROCESSOR
# =========================================================

def process_pdf(file):

    pages = read_pdf(
        file
    )

    invoices = []

    extracted_text = []


    for page in pages:

        text = page["text"]


        extracted_text.append(
            f"--- PAGE {page['page']} ---\n{text}"
        )


        invoice = parse_invoice(
            text,
            page["page"]
        )


        if invoice:

            invoices.append(
                invoice
            )


    return (
        pd.DataFrame(invoices),
        "\n\n".join(extracted_text)
    )


# =========================================================
# WORD PROCESSOR
# =========================================================

def process_word(file):

    text = read_word(
        file
    )


    invoices = []


    # -----------------------------------------------------
    # Find invoice beginnings
    # -----------------------------------------------------

    patterns = [

        r"Invoice\s+No\.?\s*[:#\-]?\s*[A-Za-z0-9][A-Za-z0-9./_-]*",

        r"Invoice\s+Number\s*[:#\-]?\s*[A-Za-z0-9][A-Za-z0-9./_-]*",

        r"Invoice\s+#\s*[A-Za-z0-9][A-Za-z0-9./_-]*",

        r"\b[A-Z]{2,10}/\d{4}[-/]\d+\b",

        r"\b[A-Z]{2,10}-\d{3,}\b"
    ]


    matches = []


    for pattern in patterns:

        found = re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        for match in found:

            matches.append(match)


    matches.sort(
        key=lambda x: x.start()
    )


    # Remove overlapping / duplicate matches
    unique_matches = []

    last_end = -1

    for match in matches:

        if match.start() >= last_end:

            unique_matches.append(
                match
            )

            last_end = match.end()


    # -----------------------------------------------------
    # If only one invoice
    # -----------------------------------------------------

    if not unique_matches:

        invoice = parse_invoice(
            text
        )

        if invoice:

            invoices.append(
                invoice
            )


    # -----------------------------------------------------
    # Multiple invoices
    # -----------------------------------------------------

    else:

        for i, match in enumerate(
            unique_matches
        ):

            start = match.start()


            if i + 1 < len(
                unique_matches
            ):

                end = unique_matches[
                    i + 1
                ].start()

            else:

                end = len(text)


            invoice_text = text[
                start:end
            ]


            invoice = parse_invoice(
                invoice_text,
                i + 1
            )


            if invoice:

                invoices.append(
                    invoice
                )


    return (
        pd.DataFrame(invoices),
        text
    )


# =========================================================
# EXCEL PROCESSOR
# =========================================================

def process_excel(file):

    df = read_excel(
        file
    )


    df = df.copy()


    # -----------------------------------------------------
    # NORMALIZE COLUMN NAMES
    # -----------------------------------------------------

    df.columns = [

        str(column)
        .strip()
        .lower()
        .replace(" ", "_")
        .replace("-", "_")
        .replace("/", "_")

        for column in df.columns
    ]


    # -----------------------------------------------------
    # FIND COLUMNS
    # -----------------------------------------------------

    def find_column(possible_names):

        for name in possible_names:

            if name in df.columns:

                return name

        return None


    invoice_column = find_column(
        [
            "invoice_no",
            "invoice_number",
            "invoice",
            "invoice_id",
            "bill_no",
            "bill_number"
        ]
    )


    vendor_column = find_column(
        [
            "vendor",
            "vendor_name",
            "supplier",
            "supplier_name",
            "customer",
            "customer_name",
            "client",
            "client_name"
        ]
    )


    invoice_date_column = find_column(
        [
            "invoice_date",
            "bill_date",
            "date"
        ]
    )


    due_date_column = find_column(
        [
            "due_date",
            "payment_due",
            "due"
        ]
    )


    amount_column = find_column(
        [
            "amount",
            "total",
            "total_amount",
            "invoice_amount",
            "grand_total"
        ]
    )


    gst_column = find_column(
        [
            "gst",
            "gst_amount",
            "tax",
            "tax_amount"
        ]
    )


    status_column = find_column(
        [
            "status",
            "payment_status",
            "invoice_status"
        ]
    )


    # -----------------------------------------------------
    # NO INVOICE COLUMN
    # -----------------------------------------------------

    if invoice_column is None:

        return pd.DataFrame()


    # -----------------------------------------------------
    # BUILD RESULT
    # -----------------------------------------------------

    result = pd.DataFrame()


    result["Invoice No"] = (
        df[invoice_column]
        .astype(str)
        .str.strip()
    )


    # Vendor
    if vendor_column:

        result["Vendor / Customer"] = (
            df[vendor_column]
            .fillna("Unknown")
            .astype(str)
            .str.strip()
        )

    else:

        result["Vendor / Customer"] = "Unknown"


    # Invoice date
    if invoice_date_column:

        result["Invoice Date"] = (
            df[invoice_date_column]
            .apply(clean_date)
        )

    else:

        result["Invoice Date"] = pd.NaT


    # Due date
    if due_date_column:

        result["Due Date"] = (
            df[due_date_column]
            .apply(clean_date)
        )

    else:

        result["Due Date"] = pd.NaT


    # Amount
    if amount_column:

        result["Amount"] = (
            df[amount_column]
            .apply(clean_amount)
            .fillna(0)
        )

    else:

        result["Amount"] = 0


    # GST
    if gst_column:

        result["GST / Tax"] = (
            df[gst_column]
            .apply(clean_amount)
        )

    else:

        result["GST / Tax"] = None


    # Status
    if status_column:

        result["Status"] = (
            df[status_column]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.title()
        )

    else:

        result["Status"] = "Unknown"


    # Payment terms
    result["Payment Terms"] = "Unknown"

    # Page
    result["Page"] = None


    # -----------------------------------------------------
    # NORMALIZE STATUS
    # -----------------------------------------------------

    for index, row in result.iterrows():

        status = str(
            row["Status"]
        ).lower().strip()


        if status in [
            "paid",
            "settled"
        ]:

            result.at[
                index,
                "Status"
            ] = "Paid"


        elif status == "overdue":

            result.at[
                index,
                "Status"
            ] = "Overdue"


        elif status in [
            "pending",
            "unpaid",
            "due"
        ]:

            result.at[
                index,
                "Status"
            ] = "Pending"


        else:

            result.at[
                index,
                "Status"
            ] = find_status(
                "",
                row["Due Date"]
            )


    return result


# =========================================================
# AI
# =========================================================

def ask_invoice_ai(
    question,
    df
):

    ai_df = df.copy()


    for column in [
        "Invoice Date",
        "Due Date"
    ]:

        if column in ai_df.columns:

            ai_df[column] = (
                ai_df[column]
                .astype(str)
            )


    invoice_data = ai_df.to_dict(
        orient="records"
    )


    system_prompt = """
You are InvoiceLens.

You are an AI invoice analyst
helping a small business owner.

The user can ask questions naturally.

Examples:

Where is my money stuck?

Who hasn't paid?

What is overdue?

Which invoices need attention?

Which customer owes the most?

What are my biggest invoices?

How much is outstanding?

Give me a summary.

Use ONLY the supplied invoice data.

Never invent:

- invoice numbers
- companies
- amounts
- dates
- statuses

If status is Unknown, clearly say
that the document did not contain
enough information to determine payment status.

Keep answers simple and business-friendly.
"""


    user_prompt = f"""
Invoice data:

{invoice_data}


User question:

{question}


Answer using only the invoice data.
"""


    try:

        response = ollama.chat(

            model="llama3.2",

            messages=[

                {
                    "role": "system",
                    "content": system_prompt
                },

                {
                    "role": "user",
                    "content": user_prompt
                }

            ]
        )


        return response[
            "message"
        ]["content"]


    except Exception as error:

        return (
            "InvoiceLens could not connect "
            "to the local AI.\n\n"
            f"Error: {error}"
        )


# =========================================================
# UPLOAD
# =========================================================

uploaded_file = st.file_uploader(
    "Upload your invoice file",
    type=[
        "pdf",
        "docx",
        "xlsx",
        "xls"
    ]
)


if not uploaded_file:

    st.info(
        "Upload a PDF, Word, or Excel invoice file."
    )

    st.stop()


# =========================================================
# FILE TYPE
# =========================================================

file_type = (
    uploaded_file.name
    .split(".")[-1]
    .lower()
)


df = pd.DataFrame()

source = ""


# =========================================================
# PDF
# =========================================================

if file_type == "pdf":

    source = "PDF"

    try:

        df, extracted_text = process_pdf(
            uploaded_file
        )


        with st.expander(
            "🔎 Show extracted PDF content"
        ):

            st.text(
                extracted_text
            )


    except Exception as error:

        st.error(
            f"Could not read PDF file: {error}"
        )

        st.stop()


# =========================================================
# WORD
# =========================================================

elif file_type == "docx":

    source = "Word"

    try:

        df, extracted_text = process_word(
            uploaded_file
        )


        with st.expander(
            "🔎 Show extracted Word content"
        ):

            st.text(
                extracted_text
            )


    except Exception as error:

        st.error(
            f"Could not read Word file: {error}"
        )

        st.stop()


# =========================================================
# EXCEL
# =========================================================

elif file_type in [
    "xlsx",
    "xls"
]:

    source = "Excel"

    try:

        df = process_excel(
            uploaded_file
        )


    except Exception as error:

        st.error(
            f"Could not read Excel file: {error}"
        )

        st.stop()


# =========================================================
# EMPTY CHECK
# =========================================================

if df.empty:

    st.error(
        f"InvoiceLens read the {source} file, "
        "but could not identify any invoices."
    )

    st.info(
        "Try a document containing labels such as "
        "'Invoice No', 'Invoice Number', "
        "'Invoice #', or an invoice-style number."
    )

    st.stop()


# =========================================================
# CLEAN
# =========================================================

df = df[
    df["Invoice No"]
    .astype(str)
    .str.strip()
    .ne("")
].copy()


df["Amount"] = (
    pd.to_numeric(
        df["Amount"],
        errors="coerce"
    )
    .fillna(0)
)


# =========================================================
# SUCCESS
# =========================================================

st.success(
    f"✓ {source} processed successfully — "
    f"{len(df)} invoice(s) detected."
)


# =========================================================
# DASHBOARD
# =========================================================

st.subheader(
    "Overview"
)


total_invoices = len(df)


total_value = df[
    "Amount"
].sum()


paid_df = df[
    df["Status"] == "Paid"
]


pending_df = df[
    df["Status"] == "Pending"
]


overdue_df = df[
    df["Status"] == "Overdue"
]


unknown_df = df[
    df["Status"] == "Unknown"
]


paid_amount = paid_df[
    "Amount"
].sum()


pending_amount = pending_df[
    "Amount"
].sum()


overdue_amount = overdue_df[
    "Amount"
].sum()


outstanding_amount = (
    pending_amount +
    overdue_amount
)


col1, col2, col3, col4 = st.columns(4)


with col1:

    st.metric(
        "Total Invoices",
        total_invoices
    )


with col2:

    st.metric(
        "Total Value",
        f"₹{total_value:,.0f}"
    )


with col3:

    st.metric(
        "Outstanding",
        f"₹{outstanding_amount:,.0f}"
    )


with col4:

    st.metric(
        "Overdue",
        len(overdue_df)
    )


# =========================================================
# STATUS
# =========================================================

st.divider()

st.subheader(
    "Payment Status"
)


c1, c2, c3, c4 = st.columns(4)


with c1:

    st.metric(
        "Paid",
        len(paid_df),
        f"₹{paid_amount:,.0f}"
    )


with c2:

    st.metric(
        "Pending",
        len(pending_df),
        f"₹{pending_amount:,.0f}"
    )


with c3:

    st.metric(
        "Overdue",
        len(overdue_df),
        f"₹{overdue_amount:,.0f}"
    )


with c4:

    st.metric(
        "Unknown",
        len(unknown_df)
    )


# =========================================================
# WARNING
# =========================================================

if len(overdue_df) > 0:

    st.warning(
        f"⚠️ {len(overdue_df)} invoice(s) "
        "are overdue and require attention."
    )


if len(unknown_df) > 0:

    st.info(
        f"ℹ️ {len(unknown_df)} invoice(s) "
        "do not contain enough payment information "
        "to determine their status."
    )


# =========================================================
# INVOICE TABLE
# =========================================================

st.divider()

st.subheader(
    "Invoices"
)


filter_status = st.selectbox(
    "Filter by status",
    [
        "All",
        "Paid",
        "Pending",
        "Overdue",
        "Unknown"
    ]
)


if filter_status == "All":

    filtered_df = df

else:

    filtered_df = df[
        df["Status"] == filter_status
    ]


st.dataframe(
    filtered_df,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# AI QUESTIONS
# =========================================================

st.divider()

st.subheader(
    "💬 Ask InvoiceLens"
)


st.caption(
    "Ask naturally. You don't need special commands."
)


question = st.text_input(
    "What would you like to know?",
    placeholder=(
        "Example: Where is my money stuck?"
    )
)


if question:

    with st.spinner(
        "InvoiceLens is analyzing your invoices..."
    ):

        answer = ask_invoice_ai(
            question,
            df
        )


    st.markdown(
        answer
    )
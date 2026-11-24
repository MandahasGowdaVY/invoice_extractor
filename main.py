import io
import json
import os
import time
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from google import genai
from google.genai import errors, types
from openpyxl import Workbook
from pydantic import BaseModel, field_validator
from sqlmodel import Field, Session, SQLModel, create_engine, select

load_dotenv()

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
app = FastAPI(title="AI Invoice Extractor")

MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.6-flash",
]

# --- 1. Database Setup ---
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./invoices.db")
connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)


class InvoiceRecord(SQLModel, table=True):
    id: int | None = Field(default=None, primary_key=True)
    vendor_name: str
    invoice_number: str | None = None
    invoice_date: str | None = None
    currency: str = "INR"
    subtotal: float | None = None
    tax_amount: float | None = None
    total_amount: float
    items_json: str = "[]"


SQLModel.metadata.create_all(engine)


def get_session():
    with Session(engine) as session:
        yield session


# --- 2. AI Extraction Schemas ---
class LineItem(BaseModel):
    description: str
    quantity: float | None = None
    unit_price: float | None = None
    amount: float | None = None


CURRENCY_MAP = {
    "$": "USD",
    "US$": "USD",
    "₹": "INR",
    "RS": "INR",
    "RS.": "INR",
    "€": "EUR",
    "£": "GBP",
}


class InvoiceData(BaseModel):
    vendor_name: str
    invoice_number: str | None = None
    invoice_date: str | None = None
    currency: str = "INR"
    subtotal: float | None = None
    tax_amount: float | None = None
    total_amount: float
    items: list[LineItem] = []

    @field_validator("currency", mode="before")
    @classmethod
    def normalize_currency(cls, v: str) -> str:
        if not v:
            return "INR"
        cleaned = v.strip().upper()
        return CURRENCY_MAP.get(cleaned, cleaned)


# --- 3. Frontend Web UI ---
HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>AI Invoice Extractor</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-100 min-h-screen">
  <div class="max-w-5xl mx-auto px-4 py-10">
    <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-8">
      <div>
        <h1 class="text-2xl font-bold tracking-tight">AI Invoice &amp; Receipt Extractor</h1>
        <p class="text-slate-400 text-sm mt-1">Upload any PDF or receipt photo to extract structured data and export to Excel.</p>
      </div>
      <div class="flex gap-3">
        <a href="/docs" class="px-4 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-sm font-medium">API Docs</a>
        <a href="/export" class="px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-medium">Download Excel (.xlsx)</a>
      </div>
    </div>

    <!-- Upload Card -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl p-6 mb-8">
      <form id="uploadForm" class="flex flex-col sm:flex-row items-center gap-4">
        <input
          id="fileInput"
          type="file"
          accept=".pdf,image/jpeg,image/png,image/webp"
          required
          class="block w-full text-sm text-slate-300 file:mr-4 file:py-2.5 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-indigo-600 file:text-white hover:file:bg-indigo-500 cursor-pointer"
        />
        <button
          id="submitBtn"
          type="submit"
          class="w-full sm:w-auto px-6 py-2.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 font-medium text-sm whitespace-nowrap"
        >
          Extract Invoice
        </button>
      </form>
      <p id="statusText" class="text-sm text-slate-400 mt-3 hidden"></p>
    </div>

    <!-- Extracted Invoices Table -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
      <div class="px-6 py-4 border-b border-slate-800 flex items-center justify-between">
        <h2 class="font-semibold text-base">Extracted Records</h2>
        <span id="countBadge" class="text-xs text-slate-400">0 records</span>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-left text-sm">
          <thead class="bg-slate-800/60 text-slate-300 text-xs uppercase">
            <tr>
              <th class="px-6 py-3">Vendor</th>
              <th class="px-6 py-3">Invoice #</th>
              <th class="px-6 py-3">Date</th>
              <th class="px-6 py-3">Tax</th>
              <th class="px-6 py-3">Total</th>
              <th class="px-6 py-3">Items</th>
            </tr>
          </thead>
          <tbody id="invoiceTableBody" class="divide-y divide-slate-800">
            <tr><td colspan="6" class="px-6 py-8 text-center text-slate-500">Loading invoices...</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <script>
    const form = document.getElementById("uploadForm");
    const fileInput = document.getElementById("fileInput");
    const submitBtn = document.getElementById("submitBtn");
    const statusText = document.getElementById("statusText");
    const tableBody = document.getElementById("invoiceTableBody");
    const countBadge = document.getElementById("countBadge");

    async function loadInvoices() {
      const res = await fetch("/invoices");
      const data = await res.json();
      countBadge.textContent = `${data.length} record${data.length === 1 ? "" : "s"}`;

      if (!data.length) {
        tableBody.innerHTML = '<tr><td colspan="6" class="px-6 py-8 text-center text-slate-500">No invoices extracted yet. Upload one above!</td></tr>';
        return;
      }

      tableBody.innerHTML = data.map(inv => {
        const items = JSON.parse(inv.items_json || "[]");
        return `
          <tr class="hover:bg-slate-800/40">
            <td class="px-6 py-4 font-medium text-white">${inv.vendor_name}</td>
            <td class="px-6 py-4 text-slate-400">${inv.invoice_number || "—"}</td>
            <td class="px-6 py-4 text-slate-400">${inv.invoice_date || "—"}</td>
            <td class="px-6 py-4 text-slate-300">${inv.tax_amount != null ? inv.currency + " " + inv.tax_amount.toFixed(2) : "—"}</td>
            <td class="px-6 py-4 font-semibold text-emerald-400">${inv.currency} ${inv.total_amount.toFixed(2)}</td>
            <td class="px-6 py-4 text-slate-400">${items.length} item${items.length === 1 ? "" : "s"}</td>
          </tr>
        `;
      }).join("");
    }

    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      if (!fileInput.files.length) return;

      const formData = new FormData();
      formData.append("file", fileInput.files[0]);

      submitBtn.disabled = true;
      submitBtn.textContent = "Extracting...";
      statusText.classList.remove("hidden", "text-red-400", "text-emerald-400");
      statusText.classList.add("text-slate-400");
      statusText.textContent = "Reading document with Gemini AI...";

      try {
        const res = await fetch("/extract", { method: "POST", body: formData });
        const result = await res.json();
        if (!res.ok) throw new Error(result.detail || "Extraction failed");

        statusText.classList.replace("text-slate-400", "text-emerald-400");
        statusText.textContent = `Extracted ${result.vendor_name} (${result.currency} ${result.total_amount})`;
        form.reset();
        await loadInvoices();
      } catch (err) {
        statusText.classList.replace("text-slate-400", "text-red-400");
        statusText.textContent = err.message;
      } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = "Extract Invoice";
      }
    });

    loadInvoices();
  </script>
</body>
</html>"""


# --- 4. API Endpoints ---
@app.get("/", response_class=HTMLResponse)
def home():
    return HTML_PAGE


@app.post("/extract", response_model=InvoiceData)
async def extract_invoice(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    allowed_types = {"application/pdf", "image/jpeg", "image/png", "image/webp"}
    if file.content_type not in allowed_types:
        raise HTTPException(
            status_code=400,
            detail="Unsupported file type. Upload a PDF, JPG, PNG, or WEBP file.",
        )

    file_bytes = await file.read()
    last_error = None

    for model_name in MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=file_bytes, mime_type=file.content_type),
                    "Extract all key invoice details and line items accurately from this document.",
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=InvoiceData,
                ),
            )
            parsed = InvoiceData.model_validate_json(response.text)

            record = InvoiceRecord(
                vendor_name=parsed.vendor_name,
                invoice_number=parsed.invoice_number,
                invoice_date=parsed.invoice_date,
                currency=parsed.currency,
                subtotal=parsed.subtotal,
                tax_amount=parsed.tax_amount,
                total_amount=parsed.total_amount,
                items_json=json.dumps([item.model_dump() for item in parsed.items]),
            )
            session.add(record)
            session.commit()

            return parsed
        except (errors.ServerError, errors.ClientError) as e:
            last_error = e
            time.sleep(1)
            continue

    raise HTTPException(
        status_code=503,
        detail=f"All models are currently busy. Please try again in a few seconds. ({last_error})",
    )


@app.get("/invoices")
def list_invoices(session: Session = Depends(get_session)):
    return session.exec(select(InvoiceRecord).order_by(InvoiceRecord.id.desc())).all()


@app.get("/export")
def export_excel(session: Session = Depends(get_session)):
    records = session.exec(select(InvoiceRecord).order_by(InvoiceRecord.id.desc())).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Invoices"
    ws.append([
        "ID",
        "Vendor Name",
        "Invoice Number",
        "Date",
        "Currency",
        "Subtotal",
        "Tax Amount",
        "Total Amount",
        "Item Count",
    ])

    for r in records:
        items = json.loads(r.items_json)
        ws.append([
            r.id,
            r.vendor_name,
            r.invoice_number or "",
            r.invoice_date or "",
            r.currency,
            r.subtotal or 0,
            r.tax_amount or 0,
            r.total_amount,
            len(items),
        ])

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=invoices.xlsx"},
    )
"""
invoice_schema.py -- Pydantic V2 accounting schema for Nexus invoice parsing.
Standardizes currency, dates, and accounting document classifications.
"""
import re
from typing import Optional, Literal
from pydantic import BaseModel, Field, field_validator

class NexusExtractedInvoice(BaseModel):
    vendor_name: str = Field(default="", description="Name of the billing vendor")
    invoice_number: str = Field(default="", description="Invoice or account reference number")
    invoice_date: Optional[str] = Field(default=None, description="Date on the invoice (YYYY-MM-DD)")
    due_date: Optional[str] = Field(default=None, description="Payment due date if stated")
    amount: str = Field(default="0.00", description="Total balance or amount due")
    property_name: Optional[str] = Field(default=None, description="Billed property or building name")
    service_address: Optional[str] = Field(default=None, description="Job site or property service address")
    vendor_address: Optional[str] = Field(default=None, description="Remit-to or mailing address of vendor")
    account_number: Optional[str] = Field(default=None, description="Utility or customer account number")
    doc_type: Literal[
        "Standard Vendor Invoice",
        "Utility (Electric)",
        "Utility (Gas)",
        "Utility (Water/Sewer)",
        "Utility (Trash/Waste)",
        "Telecom / Internet",
        "Maintenance & Repairs",
        "Supplies & Materials",
        "Legal & Eviction",
        "Software & IT",
        "Overdue / Late Notice",
        "Disconnection Notice",
        "Statement of Account",
        "Statement",
        "Notice",
        "Credit Memo"
    ] = "Standard Vendor Invoice"

    @field_validator("amount", mode="before")
    @classmethod
    def clean_amount(cls, v):
        if not v:
            return "0.00"
        s = str(v).strip()
        # Clean currency characters, commas, and trailing symbols
        cleaned = re.sub(r"[^\d.-]", "", s)
        try:
            val = float(cleaned)
            return f"{val:.2f}"
        except Exception:
            return "0.00"

    @field_validator("invoice_number", mode="before")
    @classmethod
    def clean_inv_num(cls, v):
        if v is None:
            return ""
        # Remove whitespace to maintain continuity (e.g. '110 181 023 430' -> '110181023430')
        return "".join(str(v).split())

    @field_validator("account_number", mode="before")
    @classmethod
    def clean_acct_num(cls, v):
        if v is None:
            return None
        # Remove whitespace to maintain continuity
        s = "".join(str(v).split())
        return s if s else None

    @field_validator("vendor_name", mode="before")
    @classmethod
    def clean_vendor(cls, v):
        if v is None:
            return ""
        return str(v).strip()

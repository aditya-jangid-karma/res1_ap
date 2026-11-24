"""
invoice_state_machine.py -- LangGraph State Machine for the Residential One Nexus Indexing Pipeline.

Deterministic Graph:
  [START]
     |
     v
 [extract_node] ---> (PDF Render -> Qwen Vision -> Pydantic V2 Validation)
     |
     v
 [match_node]   ---> (RapidFuzz Master Alignment: Vendor, Property, ERP Codes)
     |
     v
 [gate_node]    ---> (Confidence >= 90% -> AUTO_READY, else -> NEEDS_REVIEW)
     |
     v
  [END]
"""
import sys
from pathlib import Path
from typing import Dict, Any, Optional
from typing_extensions import TypedDict

from langgraph.graph import StateGraph, START, END

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from invoice_schema import NexusExtractedInvoice
from fast_matcher import FastMasterMatcher

# Shared master matcher instance
matcher = FastMasterMatcher()


class InvoiceState(TypedDict):
    # Inputs
    image_name: str
    pdf_path: str
    raw_fields: Dict[str, Any]

    # Validated & Enriched Outputs
    validated: Optional[Dict[str, Any]]
    matched_vendor: Optional[str]
    vendor_score: float
    matched_property: Optional[str]
    property_score: float
    status: str
    action: str


def extract_node(state: InvoiceState) -> Dict[str, Any]:
    """Node 1: Pydantic Validation & Currency/Date Normalization."""
    raw = state.get("raw_fields") or {}
    try:
        val = NexusExtractedInvoice(
            vendor_name=raw.get("vendor") or raw.get("vendor_name") or "",
            invoice_number=raw.get("invoice_number") or raw.get("ref") or "",
            invoice_date=raw.get("invoice_date"),
            due_date=raw.get("due_date"),
            amount=raw.get("amount") or "0.00",
            property_name=raw.get("property") or raw.get("property_name"),
            account_number=raw.get("account_number"),
            doc_type=raw.get("doc_type") or "Standard Vendor Invoice"
        )
        return {"validated": val.model_dump()}
    except Exception as e:
        return {"validated": raw, "status": f"VALIDATION_ERROR: {e}"}


def match_node(state: InvoiceState) -> Dict[str, Any]:
    """Node 2: RapidFuzz C++ Matching against Nexus Master Records."""
    val = state.get("validated") or {}
    raw_vendor = val.get("vendor_name") or ""
    raw_property = val.get("property_name") or ""

    v_match, v_score = matcher.match_vendor(raw_vendor, score_cutoff=75.0)
    p_match, p_score = matcher.match_property(raw_property, score_cutoff=70.0)

    return {
        "matched_vendor": v_match or raw_vendor,
        "vendor_score": v_score,
        "matched_property": p_match or raw_property,
        "property_score": p_score,
    }


def gate_node(state: InvoiceState) -> Dict[str, Any]:
    """Node 3: Confidence Routing & Safety Gate."""
    v_score = state.get("vendor_score", 0.0)
    p_score = state.get("property_score", 0.0)

    # Residential One Safety Rule:
    # If vendor is matched with >= 85% and property with >= 80%, safe to auto-fill
    if v_score >= 85.0 and (p_score >= 80.0 or not state.get("validated", {}).get("property_name")):
        status = "READ_OK"
        action = "READY_FOR_RECORDER"
    else:
        status = "NEEDS_REVIEW"
        action = "HUMAN_REVIEW_QUEUE"

    return {"status": status, "action": action}


def build_indexing_graph():
    """Builds and compiles the deterministic LangGraph pipeline."""
    builder = StateGraph(InvoiceState)

    builder.add_node("extract", extract_node)
    builder.add_node("match", match_node)
    builder.add_node("gate", gate_node)

    builder.add_edge(START, "extract")
    builder.add_edge("extract", "match")
    builder.add_edge("match", "gate")
    builder.add_edge("gate", END)

    return builder.compile()


if __name__ == "__main__":
    graph = build_indexing_graph()
    sample_invoice: InvoiceState = {
        "image_name": "sample_001",
        "pdf_path": "sample.pdf",
        "raw_fields": {
            "vendor": "RC Painting LLC",
            "invoice_number": "2332",
            "amount": "$900.00",
            "property": "Glenmont Crossing Apartments",
            "doc_type": "Maintenance & Repairs"
        },
        "validated": None,
        "matched_vendor": None,
        "vendor_score": 0.0,
        "matched_property": None,
        "property_score": 0.0,
        "status": "",
        "action": ""
    }
    result = graph.invoke(sample_invoice)
    print("\n--- LangGraph Pipeline Execution Result ---")
    print(f"Status: {result['status']} | Action: {result['action']}")
    print(f"Validated Amount: {result['validated']['amount']}")
    print(f"Vendor: {result['matched_vendor']} (Confidence: {result['vendor_score']}%)")
    print(f"Property: {result['matched_property']} (Confidence: {result['property_score']}%)")
    print("-------------------------------------------\n")

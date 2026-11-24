"""
fast_matcher.py -- RapidFuzz C++ accelerated matcher for Nexus Vendors and Properties.
Replaces slow string searches with high-speed token_sort_ratio and partial_ratio.
"""
from pathlib import Path
from typing import Optional, Tuple
import pandas as pd
from rapidfuzz import process, fuzz

HERE = Path(__file__).resolve().parent
PROJECT = HERE.parent

class FastMasterMatcher:
    def __init__(self, vendors_csv=None, properties_csv=None):
        v_path = vendors_csv or (PROJECT / "masters" / "vendors.csv")
        p_path = properties_csv or (PROJECT / "masters" / "properties.csv")

        self.vendors_df = pd.DataFrame()
        self.properties_df = pd.DataFrame()
        self.vendor_names = []
        self.property_names = []

        if Path(v_path).exists():
            try:
                self.vendors_df = pd.read_csv(v_path)
                for col in ["vendor_name", "Vendor_Name", "name", "VendorName"]:
                    if col in self.vendors_df.columns:
                        self.vendor_names = [str(x).strip() for x in self.vendors_df[col].dropna().unique() if str(x).strip()]
                        break
            except Exception:
                pass

        if Path(p_path).exists():
            try:
                self.properties_df = pd.read_csv(p_path)
                for col in ["property_name", "Property_Name", "name", "PropertyName"]:
                    if col in self.properties_df.columns:
                        self.property_names = [str(x).strip() for x in self.properties_df[col].dropna().unique() if str(x).strip()]
                        break
            except Exception:
                pass

    def match_vendor(self, raw_vendor: str, score_cutoff: float = 85.0) -> Tuple[Optional[str], float]:
        """Rapid fuzzy token match for vendor names against Nexus masters."""
        if not raw_vendor or not self.vendor_names:
            return None, 0.0
        
        match = process.extractOne(
            str(raw_vendor).strip(),
            self.vendor_names,
            scorer=fuzz.token_sort_ratio,
            score_cutoff=score_cutoff
        )
        if match:
            best_name, score, _ = match
            return best_name, round(score, 1)
        return None, 0.0

    def match_property(self, raw_property: str, score_cutoff: float = 80.0) -> Tuple[Optional[str], float]:
        """Rapid fuzzy token match for property names against Nexus masters."""
        if not raw_property or not self.property_names:
            return None, 0.0

        match = process.extractOne(
            str(raw_property).strip(),
            self.property_names,
            scorer=fuzz.WRatio,
            score_cutoff=score_cutoff
        )
        if match:
            best_name, score, _ = match
            return best_name, round(score, 1)
        return None, 0.0

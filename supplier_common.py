"""Shared constants and data model for the supplier scorecard and risk analysis tool.

Used by generate_data.py, validate_data.py, and analyze_suppliers.py so all
three scripts agree on what a valid supplier record looks like.
"""

from pydantic import BaseModel, ConfigDict, Field

CATEGORIES = [
    "Aluminium Extrusions",
    "Electronic Components",
    "Fasteners",
    "Industrial Adhesives",
    "Packaging",
    "Plastic Housings",
    "Steel Sheet",
    "Wiring Harnesses",
]

COUNTRIES = ["Germany", "Poland", "Turkey", "China", "Vietnam", "Italy", "Portugal", "Mexico", "India", "Greece"]

RAW_DATA_PATH = "data/suppliers_raw.csv"
VALIDATED_DATA_PATH = "data/suppliers_validated.csv"
VALIDATION_REPORT_PATH = "data/validation_report.txt"
OUTPUT_WORKBOOK_PATH = "output/supplier_analysis.xlsx"

CSV_FIELDS = [
    "supplier_id",
    "supplier_name",
    "country",
    "category",
    "is_incumbent",
    "unit_price",
    "lead_time_mean_days",
    "lead_time_std_days",
    "defect_rate",
    "disruption_prob",
    "annual_volume_units",
]


class Supplier(BaseModel):
    """A qualified supplier for one purchased category.

    annual_volume_units is the category's yearly requirement this supplier has quoted for
    (the same figure for every supplier in a category); is_incumbent marks the one you use today.
    """

    model_config = ConfigDict(extra="ignore")

    supplier_id: str = Field(min_length=1)
    supplier_name: str = Field(min_length=1)
    country: str = Field(min_length=1)
    category: str = Field(min_length=1)
    is_incumbent: bool
    unit_price: float = Field(gt=0)
    lead_time_mean_days: float = Field(gt=0)
    lead_time_std_days: float = Field(ge=0)
    defect_rate: float = Field(ge=0, lt=1)  # share of delivered units that are defective
    disruption_prob: float = Field(ge=0, le=1)  # yearly probability of a major outage
    annual_volume_units: int = Field(gt=0)

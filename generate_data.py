"""Generate a fake supplier dataset with Faker.

Produces a CSV of supplier records (price, lead-time mean and variability,
defect rate, disruption probability) grouped by purchased category, with one
incumbent per category. The rest of the pipeline (validate_data.py,
analyze_suppliers.py) consumes it. Cheaper suppliers are made a little
less reliable on average, so price and reliability genuinely trade off.
A small percentage of rows are deliberately corrupted (negative prices,
out-of-range rates, blanks, duplicate IDs) so validate_data.py has real work
to do - pass --error-rate 0 for a perfectly clean dataset.
"""

import argparse
import csv
import os
import random

from faker import Faker

from supplier_common import CATEGORIES, COUNTRIES, CSV_FIELDS, RAW_DATA_PATH


def generate_category(fake: Faker, category: str, start_index: int, suppliers_per_category: int) -> list[dict]:
    base_price = random.uniform(2, 120)
    annual_volume = random.randint(2_000, 400_000)
    rows = []

    for k in range(suppliers_per_category):
        # price_factor < 1 means a cheaper supplier; cheaper tends to mean less reliable.
        price_factor = random.uniform(0.85, 1.2)
        reliability = (price_factor - 0.85) / 0.35  # 0 = cheapest, 1 = priciest
        lead_mean = random.uniform(7, 55)
        rows.append(
            {
                "supplier_id": f"SUP-{start_index + k:04d}",
                "supplier_name": fake.company(),
                "country": random.choice(COUNTRIES),
                "category": category,
                "is_incumbent": k == 0,
                "unit_price": round(base_price * price_factor, 2),
                "lead_time_mean_days": round(lead_mean, 1),
                "lead_time_std_days": round(lead_mean * random.uniform(0.05, 0.45) * (1.3 - 0.6 * reliability), 1),
                "defect_rate": round(random.uniform(0.002, 0.06) * (1.4 - 0.8 * reliability), 4),
                "disruption_prob": round(random.uniform(0.01, 0.15) * (1.3 - 0.5 * reliability), 3),
                "annual_volume_units": annual_volume,
            }
        )
    return rows


def corrupt_row(row: dict) -> dict:
    """Introduce one realistic data-quality issue into a row, in place."""
    issue = random.choice(["negative_price", "bad_defect_rate", "blank_lead_time", "bad_disruption_prob", "blank_country"])
    if issue == "negative_price":
        row["unit_price"] = -abs(row["unit_price"])
    elif issue == "bad_defect_rate":
        row["defect_rate"] = 1.8
    elif issue == "blank_lead_time":
        row["lead_time_mean_days"] = ""
    elif issue == "bad_disruption_prob":
        row["disruption_prob"] = -0.1
    elif issue == "blank_country":
        row["country"] = ""
    return row


def generate_dataset(categories: int, suppliers_per_category: int, error_rate: float, seed: int | None) -> list[dict]:
    if seed is not None:
        random.seed(seed)
        Faker.seed(seed)
    fake = Faker()

    rows: list[dict] = []
    for i, category in enumerate(CATEGORIES[:categories]):
        rows.extend(generate_category(fake, category, len(rows) + 1, suppliers_per_category))

    n_errors = round(len(rows) * error_rate)
    for row in random.sample(rows, k=min(n_errors, len(rows))):
        corrupt_row(row)

    # Inject a duplicate supplier ID to exercise the validator's duplicate check.
    if error_rate > 0 and len(rows) > 1:
        dup = dict(rows[0])
        dup["supplier_id"] = rows[-1]["supplier_id"]
        rows.append(dup)

    return rows


def write_csv(rows: list[dict], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate fake supplier data for the scorecard pipeline.")
    parser.add_argument("--categories", type=int, default=6, help=f"Number of purchased categories, 1-{len(CATEGORIES)} (default: 6)")
    parser.add_argument("--suppliers-per-category", type=int, default=4, help="Qualified suppliers per category (default: 4)")
    parser.add_argument(
        "--error-rate",
        type=float,
        default=0.05,
        help="Fraction of rows (0-1) to deliberately corrupt, to exercise validate_data.py (default: 0.05)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible output")
    parser.add_argument("--output", type=str, default=RAW_DATA_PATH, help="Path to write the raw CSV to")
    args = parser.parse_args()

    if not 1 <= args.categories <= len(CATEGORIES):
        parser.error(f"--categories must be between 1 and {len(CATEGORIES)}")

    rows = generate_dataset(args.categories, args.suppliers_per_category, args.error_rate, args.seed)
    write_csv(rows, args.output)
    print(f"Generated {len(rows)} rows ({round(args.error_rate * 100)}% target error rate) -> {args.output}")


if __name__ == "__main__":
    main()

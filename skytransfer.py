"""Populate vendor info on operational workbooks using SKU substring matches."""

from __future__ import annotations

from datetime import datetime
import numbers
from pathlib import Path
from typing import Any, cast

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SKU_WORKBOOK = BASE_DIR / "SKU Transfer Total List.xlsx"
OUTPUT_DIR = BASE_DIR / "newStuffs"
SOURCE_MONTH = datetime.now().strftime("%m.%Y")
SOURCE_MONTH = (datetime.now().replace(day=1) - pd.DateOffset(months=1)).strftime("%m.%Y")

# Configuration per source sheet: file name, lookup column, and vendor columns.
CATEGORY_CONFIG = (
	{
		"name": "Calls",
		"source": BASE_DIR / f"Calls {SOURCE_MONTH}.xlsx",
		"lookup": "OD SKU",
		"vendor_id": "Vendor ID",
		"vendor_name": "Vendor Name",
	},
	{
		"name": "Emails",
		"source": BASE_DIR / f"Emails {SOURCE_MONTH}.xlsx",
		"lookup": "SKU",
		"vendor_id": "Vendor ID",
		"vendor_name": "Vendor Name",
	},
	{
		"name": "Gift Cards",
		"source": BASE_DIR / f"Gift Cards {SOURCE_MONTH}.xlsx",
		"lookup": "Item",
		"vendor_id": "Vendor ID",
		"vendor_name": "Vendor Name",
	},
	{
		"name": "Shipments",
		"source": BASE_DIR / f"Shipments {SOURCE_MONTH}.xlsx",
		"lookup": "Item",
		"vendor_id": "Vendor ID",
		"vendor_name": "Vendor Name",
	},
	{
		"name": "Storage",
		"source": BASE_DIR / f"Storage {SOURCE_MONTH}.xlsx",
		"lookup": "Name",
		"vendor_id": "Vendor ID",
		"vendor_name": "Vendor Name",
	},
)


def _normalize_column_name(value: str) -> str:
	return str(value).strip().lower()


def _find_column(df: pd.DataFrame, desired: str) -> str:
	desired_norm = _normalize_column_name(desired)
	for column in df.columns:
		if _normalize_column_name(column) == desired_norm:
			return column
	raise KeyError(f"Column '{desired}' not found. Available columns: {list(df.columns)}")


def _sanitize_sku(raw_value) -> str:
	if pd.isna(raw_value):
		return ""
	if isinstance(raw_value, numbers.Integral):
		return str(int(raw_value))
	if isinstance(raw_value, numbers.Real):
		as_float = float(raw_value)
		if as_float.is_integer():
			return str(int(as_float))
	return str(raw_value).strip()


def load_sku_catalog() -> list[dict[str, object]]:
	df = pd.read_excel(SKU_WORKBOOK)
	sku_col = _find_column(df, "SKU contains")
	vendor_id_col = _find_column(df, "New Vendor ID")
	vendor_name_col = _find_column(df, "New Vendor Name")

	catalog: list[dict[str, object]] = []
	for _, row in df.iterrows():
		sku = _sanitize_sku(row[sku_col])
		if not sku:
			continue
		catalog.append(
			{
				"sku": sku,
				"vendor_id": row[vendor_id_col],
				"vendor_name": row[vendor_name_col],
			}
		)

	# Match longer SKU strings first to avoid premature partial hits.
	catalog.sort(key=lambda entry: len(entry["sku"]), reverse=True)
	return catalog


def match_vendor(text_value, catalog: list[dict[str, object]]) -> tuple[object, object] | None:
	if pd.isna(text_value):
		return None
	candidate = str(text_value).strip()
	if not candidate or candidate.lower() == "nan":
		return None
	for entry in catalog:
		sku = entry["sku"]
		if sku and sku in candidate:
			return entry["vendor_id"], entry["vendor_name"]
	return None


def process_category(config: dict[str, object], catalog: list[dict[str, object]]) -> tuple[int, int]:
	df = pd.read_excel(config["source"])
	lookup_col = _find_column(df, config["lookup"])
	vendor_id_col = _find_column(df, config["vendor_id"])
	vendor_name_col = _find_column(df, config["vendor_name"])

	updated = 0
	for idx, value in df[lookup_col].items():
		match = match_vendor(value, catalog)
		if match is None:
			continue
		vendor_id, vendor_name = match
		df.at[idx, vendor_id_col] = cast(Any, vendor_id)
		df.at[idx, vendor_name_col] = cast(Any, vendor_name)
		updated += 1

	OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
	destination = OUTPUT_DIR / f"New {config['name']}.xlsx"
	df.to_excel(destination, index=False)
	return updated, len(df)


def main() -> None:
	catalog = load_sku_catalog()
	if not catalog:
		raise RuntimeError("SKU catalog is empty. Check 'SKU Transfer Total List.xlsx'.")

	for config in CATEGORY_CONFIG:
		updated, total = process_category(config, catalog)
		print(f"{config['name']}: updated {updated} of {total} rows -> New {config['name']}.xlsx")


if __name__ == "__main__":
	main()
"""Data-backed per-employee weekly / monthly / lifetime productivity.

Each metric is grouped by category and field, never conflating heterogeneous work
units (measurements, drawings, changes). Archived reports excluded. For duplicate
saved versions of the same date/company, use only the latest update.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any


def period_bounds(period: str, today: date) -> tuple[str | None, str | None]:
    if period == "week":
        return (today - timedelta(days=today.weekday())).isoformat(), today.isoformat()
    if period == "month":
        return today.replace(day=1).isoformat(), today.isoformat()
    if period == "all":
        return None, None
    raise ValueError(f"Unknown reporting period: {period}")


def latest_daily_snapshots(database: Any, start: str | None, end: str | None) -> list[dict[str, Any]]:
    reports = database.list_reports(start, end)
    by_date_company: dict[tuple[str, str], dict[str, Any]] = {}
    for meta in reports:
        report_date = str(meta.get("report_date", ""))
        key = report_date, str(meta.get("company_id", ""))
        existing = by_date_company.get(key)
        if existing is None or str(meta.get("updated_at", "")) > str(existing.get("updated_at", "")):
            by_date_company[key] = meta
    result = []
    for meta in sorted(by_date_company.values(), key=lambda x: str(x.get("report_date", ""))):
        value = database.get_report(meta["id"])
        if value is None or not isinstance(value.get("snapshot"), dict):
            continue
        result.append(value["snapshot"])
    return result


def employee_productivity(database: Any, config: dict[str, Any], period: str, today: date) -> dict[str, Any]:
    start, end = period_bounds(period, today)
    snapshots = latest_daily_snapshots(database, start, end)
    employees: dict[tuple[str, str], dict[str, Any]] = {}
    detail: dict[tuple[str, str, str, str], Decimal] = defaultdict(Decimal)

    def employee(category_id: str, name: str, label: str) -> dict[str, Any]:
        key = (category_id, name)
        if key not in employees:
            employees[key] = {"category_id": category_id, "category": label,
                              "name": name, "records": 0, "dates": set(), "total": Decimal(0)}
        return employees[key]

    for category_id, category in config.get("categories", {}).items():
        if not isinstance(category, dict):
            continue
        for person in category.get("people", []):
            if str(person).strip():
                employee(str(category_id), str(person).strip(), str(category.get("label", category_id)))

    for snapshot in snapshots:
        date_key = str(snapshot.get("date", ""))
        cfg = snapshot.get("config", {})
        categories = cfg.get("categories", {}) if isinstance(cfg, dict) else {}
        for row in snapshot.get("rows", []):
            if not isinstance(row, dict):
                continue
            category_id = str(row.get("category", ""))
            name = str(row.get("name", "")).strip()
            if not category_id or not name:
                continue
            category = categories.get(category_id, {})
            label = str(row.get("category_label") or category.get("label") or category_id)
            record = employee(category_id, name, label)
            record["records"] += 1
            record["dates"].add(date_key)
            for field in row.get("fields", []):
                if not isinstance(field, dict) or str(field.get("kind")) not in {"number", "currency"}:
                    continue
                field_id = str(field.get("id", "")).strip()
                if not field_id:
                    continue
                raw = field.get("value", "")
                try:
                    value = Decimal(str(raw or 0).replace(",", ""))
                except (InvalidOperation, ValueError):
                    continue
                if not value.is_finite():
                    continue
                detail[(category_id, name, field_id, str(field.get("label", field_id)))] += value
                if field.get("kind") == "number":
                    record["total"] += value

    people_out = []
    for item in employees.values():
        people_out.append({
            "category": item["category"], "category_id": item["category_id"],
            "name": item["name"], "records": item["records"], "days": len(item["dates"]),
            "total": item["total"]
        })
    people_out.sort(key=lambda x: (x["category"], -x["total"], x["name"]))
    field_out = [
        {"category_id": k[0], "person": k[1], "field_id": k[2],
         "field": k[3], "amount": value}
        for k, value in detail.items()
    ]
    field_out.sort(key=lambda x: (x["category_id"], x["person"], x["field"]))
    return {"period": period, "start": start, "end": end,
            "days": len({str(x.get("date", "")) for x in snapshots}), "reports": len(snapshots),
            "employees": people_out, "fields": field_out}


def format_total(value: Decimal) -> str:
    return format(value.quantize(Decimal("0.01")), "f").rstrip("0").rstrip(".") or "0"

"""Validation functions for bill data."""


def evaluate_price_value(price_value: float | int | str) -> float | None:
    """Convert a price value to float. Return None on failure."""
    if isinstance(price_value, (int, float)):
        return float(price_value)

    if not isinstance(price_value, str):
        return None

    price_value = price_value.strip()
    try:
        return float(price_value.replace(",", "."))
    except ValueError:
        return None


def validate_bill_total(
    bill_data: dict[str, float | int | str | list],
) -> dict[str, bool | float | str] | None:
    """Validate that item prices sum to the declared total. Return None on error."""
    items: list = bill_data.get("items", [])  # type: ignore[assignment]
    declared_total_raw = bill_data.get("total")
    declared_total: float | int | str | None = (
        declared_total_raw if not isinstance(declared_total_raw, list) else None
    )

    if declared_total is None:
        return None

    if not items:
        return None

    calculated_sum: float = 0.0
    for item in items:
        price: float | int | str | None = item.get("item_price")
        if price is None:
            return None

        evaluated_price: float | None = evaluate_price_value(price)
        if evaluated_price is None:
            return None

        calculated_sum += evaluated_price

    declared_total_value: float | None = evaluate_price_value(declared_total)
    if declared_total_value is None:
        return None

    difference: float = abs(calculated_sum - declared_total_value)
    is_valid: bool = difference < 0.01

    result: dict[str, bool | float | str] = {
        "valid": is_valid,
        "calculated_sum": round(calculated_sum, 2),
        "declared_total": round(declared_total_value, 2),
        "difference": round(difference, 2),
        "message": "✓ Valid total" if is_valid else "⚠ Price mismatch",
    }

    return result

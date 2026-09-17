import math


def balanced_percentages(
    values: list[float],
    decimals: int = 1,
) -> list[float]:
    """Round a part-to-whole composition while keeping the display total at 100%."""
    clean_values = [max(0.0, float(value or 0)) for value in values]
    total = sum(clean_values)
    if total <= 0:
        return [0.0 for _ in clean_values]
    scale = 10**decimals
    target_units = 100 * scale
    raw_units = [value / total * target_units for value in clean_values]
    rounded_units = [math.floor(value) for value in raw_units]
    remaining = target_units - sum(rounded_units)
    order = sorted(
        range(len(raw_units)),
        key=lambda index: raw_units[index] - rounded_units[index],
        reverse=True,
    )
    for index in order[:remaining]:
        rounded_units[index] += 1
    return [value / scale for value in rounded_units]

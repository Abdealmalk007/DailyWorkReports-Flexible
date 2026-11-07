from __future__ import annotations

import math
import random
from statistics import median
from typing import Any, Iterable


def _numbers(values: Iterable[Any]) -> list[float]:
    result: list[float] = []
    for value in values:
        try:
            number = float(value)
        except (TypeError, ValueError):
            continue
        if math.isfinite(number):
            result.append(number)
    return result


def robust_anomalies(
    rows: list[dict[str, Any]], value_key: str = "total", label_key: str = "label"
) -> list[dict[str, Any]]:
    values = _numbers(row.get(value_key) for row in rows)
    if len(values) < 5:
        return []
    center = median(values)
    mad = median(abs(value - center) for value in values)
    if mad <= 1e-9:
        return []
    result = []
    for row in rows:
        value = float(row.get(value_key) or 0)
        score = 0.6745 * (value - center) / mad
        if abs(score) >= 3.5:
            result.append(
                {
                    "label": str(row.get(label_key, "")),
                    "value": round(value, 2),
                    "score": round(score, 2),
                    "direction": "مرتفع" if score > 0 else "منخفض",
                }
            )
    return sorted(result, key=lambda item: abs(item["score"]), reverse=True)


def neural_forecast(values: Iterable[Any], lookback: int = 3) -> dict[str, Any]:
    series = _numbers(values)
    lookback = max(2, min(7, int(lookback)))
    if len(series) < lookback + 4:
        return {
            "available": False,
            "prediction": None,
            "lower_bound": None,
            "upper_bound": None,
            "confidence": 0,
            "samples": max(0, len(series) - lookback),
            "reason": "يلزم سبعة أيام على الأقل لبناء توقع داخلي.",
            "model": "daily-local-neural-v2",
        }
    lower, upper = min(series), max(series)
    span = upper - lower
    if span <= 1e-9:
        return {
            "available": True,
            "prediction": round(lower, 2),
            "lower_bound": round(lower, 2),
            "upper_bound": round(lower, 2),
            "confidence": 95,
            "samples": len(series) - lookback,
            "reason": "الإنتاجية التاريخية مستقرة.",
            "model": "daily-local-neural-v2",
        }
    normalized = [(value - lower) / span for value in series]
    samples = [
        (normalized[index - lookback : index], normalized[index])
        for index in range(lookback, len(normalized))
    ]
    hidden_size = min(6, lookback + 2)
    generator = random.Random(1100 + len(series) * 13)
    input_weights = [
        [generator.uniform(-0.35, 0.35) for _ in range(lookback)]
        for _ in range(hidden_size)
    ]
    hidden_bias = [0.0] * hidden_size
    output_weights = [generator.uniform(-0.35, 0.35) for _ in range(hidden_size)]
    output_bias = 0.0

    def forward(inputs: list[float]) -> tuple[list[float], float]:
        hidden = [
            math.tanh(
                hidden_bias[node]
                + sum(input_weights[node][offset] * inputs[offset] for offset in range(lookback))
            )
            for node in range(hidden_size)
        ]
        return hidden, output_bias + sum(
            output_weights[node] * hidden[node] for node in range(hidden_size)
        )

    rate = 0.045
    for _epoch in range(420):
        for inputs, expected in samples:
            hidden, output = forward(inputs)
            error = max(-2.0, min(2.0, output - expected))
            old_output = list(output_weights)
            for node in range(hidden_size):
                output_weights[node] -= rate * error * hidden[node]
            output_bias -= rate * error
            for node in range(hidden_size):
                gradient = error * old_output[node] * (1.0 - hidden[node] ** 2)
                for offset in range(lookback):
                    input_weights[node][offset] -= rate * gradient * inputs[offset]
                hidden_bias[node] -= rate * gradient
        rate = max(0.012, rate * 0.995)
    errors = [abs(forward(inputs)[1] - expected) for inputs, expected in samples]
    normalized_prediction = max(0.0, min(1.5, forward(normalized[-lookback:])[1]))
    prediction = max(0.0, lower + normalized_prediction * span)
    mean_error = sum(errors) / max(1, len(errors))
    confidence = int(max(25, min(92, (1 - min(1.0, mean_error)) * 76 + min(16, len(samples)))))
    return {
        "available": True,
        "prediction": round(prediction, 2),
        "lower_bound": round(max(0.0, prediction - mean_error * span * 1.65), 2),
        "upper_bound": round(max(0.0, prediction + mean_error * span * 1.65), 2),
        "confidence": confidence,
        "samples": len(samples),
        "reason": "توقع محلي مع نطاق احتمالي محسوب من خطأ التدريب.",
        "model": "daily-local-neural-v2",
    }


def scenario_forecast(
    values: Iterable[Any], adjustment_percent: float = 0.0
) -> dict[str, Any]:
    """Apply a bounded what-if adjustment to the local forecast and its interval."""
    forecast = neural_forecast(values)
    adjustment = max(-80.0, min(300.0, float(adjustment_percent or 0.0)))
    factor = 1.0 + adjustment / 100.0
    result = dict(forecast)
    result["adjustment_percent"] = round(adjustment, 2)
    result["model"] = "daily-scenario-v2"
    if result.get("available"):
        for key in ("prediction", "lower_bound", "upper_bound"):
            result[key] = round(max(0.0, float(result.get(key) or 0.0) * factor), 2)
        result["reason"] = (
            f"سيناريو محلي يطبق تغيراً مقداره {adjustment:+.1f}% على التوقع الأساسي."
        )
    return result


def report_quality_score(snapshot: dict[str, Any]) -> dict[str, Any]:
    """Score report completeness and consistency without modifying the report."""
    if not isinstance(snapshot, dict):
        return {"score": 0, "level": "ضعيف", "issues": ["بنية التقرير غير صالحة"]}
    rows = [row for row in snapshot.get("rows", []) if isinstance(row, dict)]
    issues: list[str] = []
    score = 0.0
    if rows:
        score += 20
    else:
        issues.append("لا توجد سجلات عمل")

    identity_keys = ("date", "company_name", "user_name", "code", "page")
    identity_filled = sum(bool(str(snapshot.get(key, "")).strip()) for key in identity_keys)
    score += 20 * identity_filled / len(identity_keys)
    if identity_filled < len(identity_keys):
        issues.append("بيانات تعريف التقرير غير مكتملة")

    if rows:
        named = sum(bool(str(row.get("name", "")).strip()) for row in rows)
        score += 20 * named / len(rows)
        if named < len(rows):
            issues.append("بعض السجلات دون اسم أو جهة")

        field_values = 0
        field_filled = 0
        row_keys: list[str] = []
        for row in rows:
            row_keys.append(
                f"{row.get('category', '')}|{str(row.get('name', '')).strip().casefold()}|{row.get('total', 0)}"
            )
            for field in row.get("fields", []):
                if not isinstance(field, dict):
                    continue
                field_values += 1
                if str(field.get("value", "")).strip():
                    field_filled += 1
        score += 30 * field_filled / max(1, field_values)
        if field_values and field_filled < field_values:
            issues.append("توجد حقول عمل فارغة")
        unique_ratio = len(set(row_keys)) / len(row_keys)
        score += 10 * unique_ratio
        if unique_ratio < 1:
            issues.append("توجد سجلات متشابهة قد تكون مكررة")

    score = int(round(max(0.0, min(100.0, score))))
    if score >= 85:
        level = "ممتاز"
    elif score >= 70:
        level = "جيد"
    elif score >= 50:
        level = "متوسط"
    else:
        level = "ضعيف"
    if not issues:
        issues.append("لا توجد ملاحظات جودة جوهرية")
    return {"score": score, "level": level, "issues": issues, "model": "report-quality-v2"}


def productivity_insights(analytics: dict[str, Any]) -> list[dict[str, str]]:
    insights: list[dict[str, str]] = []
    forecast = analytics.get("forecast") or {}
    if forecast.get("available"):
        insights.append(
            {
                "level": "info",
                "title": "توقع اليوم التالي",
                "message": (
                    f"يتوقع المحرك نحو {float(forecast['prediction']):.1f} وحدة عمل "
                    f"بثقة تقديرية {int(forecast['confidence'])}%."
                ),
            }
        )
        if forecast.get("lower_bound") is not None:
            insights.append(
                {
                    "level": "info",
                    "title": "نطاق التوقع",
                    "message": (
                        f"النطاق المحتمل من {float(forecast['lower_bound']):.1f} إلى "
                        f"{float(forecast['upper_bound']):.1f} وحدة."
                    ),
                }
            )
    else:
        insights.append({"level": "info", "title": "التعلم المحلي", "message": str(forecast.get("reason", ""))})
    anomalies = analytics.get("anomalies") or []
    if anomalies:
        labels = "، ".join(str(item.get("label", "")) for item in anomalies[:3])
        insights.append({"level": "warning", "title": "تغيرات غير اعتيادية", "message": f"رُصد تغير يحتاج مراجعة في: {labels}."})
    categories = analytics.get("categories") or []
    if categories:
        leader = categories[0]
        insights.append(
            {
                "level": "success",
                "title": "الفئة الأعلى",
                "message": f"{leader.get('label', '')} تتصدر بإجمالي {leader.get('total', 0)}.",
            }
        )
    quality = analytics.get("quality") or {}
    if quality.get("reports"):
        average = float(quality.get("average") or 0)
        attention = int(quality.get("needs_attention") or 0)
        insights.append(
            {
                "level": "success" if average >= 85 and not attention else "warning",
                "title": "جودة التقارير",
                "message": (
                    f"متوسط الجودة {average:.1f}/100، والتقارير التي تحتاج مراجعة: {attention}."
                ),
            }
        )
    if not analytics.get("reports"):
        insights.append({"level": "warning", "title": "الفترة المختارة", "message": "لا توجد تقارير كافية للتحليل ضمن هذه الفترة."})
    return insights

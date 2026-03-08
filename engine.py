
from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from math import ceil, isfinite
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


class CalculationError(Exception):
    pass


PRODUCT_COLUMN_MAP = {
    "sku": "SKU",
    "артикул": "SKU",
    "product_name": "Product Name",
    "наименование": "Product Name",
    "product name": "Product Name",
    "model": "Model",
    "модель": "Model",
    "current_price_rub": "Current Price RUB",
    "цена": "Current Price RUB",
    "cost_rub": "Cost RUB",
    "себестоимость": "Cost RUB",
    "weight_kg": "Weight kg",
    "вес": "Weight kg",
    "length_cm": "Length cm",
    "длина": "Length cm",
    "width_cm": "Width cm",
    "ширина": "Width cm",
    "height_cm": "Height cm",
    "высота": "Height cm",
    "manual_category": "Manual Category",
    "категория": "Manual Category",
    "target_margin_pct_override": "Target Margin % Override",
    "storage_cost_rub_override": "Storage Cost RUB Override",
    "inbound_cost_rub_override": "Inbound Cost RUB Override",
    "is_kgt": "Is KGT",
    "notes": "Notes",
}

REQUIRED_PRODUCT_COLUMNS = [
    "SKU",
    "Product Name",
    "Cost RUB",
    "Weight kg",
    "Length cm",
    "Width cm",
    "Height cm",
]

DEFAULT_SETTINGS = {
    "default_model": "FBS",
    "target_margin_pct": 20.0,
    "buyout_rate_pct": 95.0,
    "tax_rate_pct": 6.0,
    "ad_rate_pct": 8.0,
    "defect_rate_pct": 1.0,
    "transfer_schedule": "Еженедельно, с отсрочкой в 2 недели",
    "reverse_same_locality_share_pct": 0.0,
    "fbs_avg_orders_per_day": 20,
    "avg_items_per_order": 1.2,
    "fby_storage_cost_rub_per_unit": 0.0,
    "fby_inbound_self_delivery_rub_per_unit": 0.0,
    "default_commission_rate_pct": 8.0,
    "kgt_weight_threshold_kg": 25.0,
    "kgt_sum_sides_threshold_cm": 150.0,
}


def _normalize_name(name: str) -> str:
    return str(name).strip().lower().replace("\n", " ").replace("  ", " ")




def _blankish(x: Any) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    return str(x).strip() == ""


def _clean_string(x: Any) -> str:
    if _blankish(x):
        return ""
    return str(x).strip()


def read_products_excel(file_like) -> pd.DataFrame:
    df = pd.read_excel(file_like, sheet_name=0)
    if df.empty:
        raise CalculationError("В файле товаров нет строк.")
    rename_map = {}
    for col in df.columns:
        key = _normalize_name(col)
        rename_map[col] = PRODUCT_COLUMN_MAP.get(key, col)
    df = df.rename(columns=rename_map)

    missing = [c for c in REQUIRED_PRODUCT_COLUMNS if c not in df.columns]
    if missing:
        raise CalculationError(f"В файле товаров не хватает обязательных столбцов: {', '.join(missing)}")

    for optional in [
        "Model",
        "Current Price RUB",
        "Manual Category",
        "Target Margin % Override",
        "Storage Cost RUB Override",
        "Inbound Cost RUB Override",
        "Is KGT",
        "Notes",
    ]:
        if optional not in df.columns:
            df[optional] = None

    return df


def load_settings(file_or_path) -> Dict[str, Any]:
    params = DEFAULT_SETTINGS.copy()
    wb = load_workbook(file_or_path, data_only=True)
    ws = wb["Settings"] if "Settings" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    for row in rows:
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        value = row[1]
        if key in params and value is not None:
            params[key] = value
    # types
    params["default_model"] = str(params["default_model"]).strip().upper()
    for key in params:
        if key == "default_model" or key == "transfer_schedule":
            continue
        if isinstance(params[key], str):
            try:
                params[key] = float(str(params[key]).replace(",", "."))
            except Exception:
                pass
    return params


def load_commissions(path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Commissions")
    df.columns = [str(c).strip() for c in df.columns]
    needed = {"Category", "Commission Rate %"}
    if not needed.issubset(df.columns):
        raise CalculationError("В файле commissions.xlsx нужен лист Commissions со столбцами Category и Commission Rate %.")
    df["Category_norm"] = df["Category"].astype(str).str.strip().str.lower()
    return df


def load_keyword_rules(path: str | Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Keyword Rules")
    df.columns = [str(c).strip() for c in df.columns]
    needed = {"Keyword", "Category", "Priority"}
    if not needed.issubset(df.columns):
        raise CalculationError("В файле keyword_rules.xlsx нужен лист Keyword Rules со столбцами Keyword, Category, Priority.")
    df["keyword_norm"] = df["Keyword"].astype(str).str.strip().str.lower()
    df["Category_norm"] = df["Category"].astype(str).str.strip().str.lower()
    df = df.sort_values(["Priority", "Keyword"])
    return df


def load_tariffs(path: str | Path) -> Dict[str, Any]:
    general = pd.read_excel(path, sheet_name="General")
    payment_transfer = pd.read_excel(path, sheet_name="Payment Transfer")
    fbs_pickup = pd.read_excel(path, sheet_name="FBS Pickup")

    general_map = {str(k).strip(): v for k, v in zip(general["parameter"], general["value"])}
    payment_map = {
        str(row["frequency_label"]).strip(): float(row["rate_pct"])
        for _, row in payment_transfer.iterrows()
        if pd.notna(row["frequency_label"])
    }
    pickup_rows = fbs_pickup.to_dict("records")

    return {
        "general": general_map,
        "payment_transfer": payment_map,
        "fbs_pickup": pickup_rows,
    }


def detect_category(product_name: str, keyword_rules_df: pd.DataFrame) -> Tuple[str, str]:
    name_norm = str(product_name).strip().lower()
    for _, row in keyword_rules_df.iterrows():
        kw = row["keyword_norm"]
        if kw and kw in name_norm:
            return row["Category"], f"keyword:{row['Keyword']}"
    return "UNMAPPED", "unmapped"


def lookup_commission_rate(category: str, commissions_df: pd.DataFrame, default_rate_pct: float) -> Tuple[float, bool]:
    norm = str(category).strip().lower()
    match = commissions_df.loc[commissions_df["Category_norm"] == norm]
    if match.empty:
        return float(default_rate_pct), False
    val = match.iloc[0]["Commission Rate %"]
    if pd.isna(val):
        return float(default_rate_pct), False
    return float(val), True


def calc_volume_liters(length_cm: float, width_cm: float, height_cm: float) -> int:
    return max(1, ceil((float(length_cm) * float(width_cm) * float(height_cm)) / 1000.0))


def calc_volumetric_weight_kg(length_cm: float, width_cm: float, height_cm: float) -> float:
    return (float(length_cm) * float(width_cm) * float(height_cm)) / 5000.0


def calc_middle_mile_rub(volume_liters: int, tariffs: Dict[str, Any]) -> float:
    g = tariffs["general"]
    base = float(g["middle_mile_base_rub"])
    extra_1_160 = float(g["middle_mile_up_to_160_rub_per_l"])
    extra_160_plus = float(g["middle_mile_over_160_rub_per_l"])
    liters = int(volume_liters)
    if liters <= 1:
        return base
    if liters <= 160:
        return base + (liters - 1) * extra_1_160
    return base + 159 * extra_1_160 + (liters - 160) * extra_160_plus


def calc_fbs_pickup_fee_per_unit(avg_orders_per_day: float, avg_items_per_order: float, tariffs: Dict[str, Any]) -> float:
    day_orders = max(1.0, float(avg_orders_per_day))
    items_per_order = max(1.0, float(avg_items_per_order))
    chosen_fee = 0.0
    for row in tariffs["fbs_pickup"]:
        low = float(row["min_orders"])
        high = float(row["max_orders"]) if pd.notna(row["max_orders"]) else float("inf")
        if low <= day_orders <= high:
            chosen_fee = float(row["fee_rub"])
            break
    return chosen_fee / day_orders / items_per_order


def calc_fbs_processing_fee_per_unit(is_kgt: bool, volumetric_weight_kg: float, avg_items_per_order: float, tariffs: Dict[str, Any]) -> float:
    g = tariffs["general"]
    if is_kgt:
        return max(0.0, volumetric_weight_kg) * float(g["fbs_kgt_processing_rub_per_kg"])
    return float(g["fbs_regular_processing_rub_per_order"]) / max(1.0, float(avg_items_per_order))


def price_based_delivery_fee(price_rub: float, tariffs: Dict[str, Any]) -> float:
    g = tariffs["general"]
    rate = float(g["delivery_to_buyer_rate_pct"]) / 100.0
    cap = float(g["delivery_to_buyer_cap_rub"])
    return min(float(price_rub) * rate, cap)


def _to_float(x: Any, default: float = 0.0) -> float:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return default
    if isinstance(x, bool):
        return float(x)
    try:
        return float(str(x).replace(",", ".").strip())
    except Exception:
        return default


def _is_trueish(x: Any) -> bool:
    if isinstance(x, bool):
        return x
    s = str(x).strip().lower()
    return s in {"1", "true", "yes", "y", "да", "x"}


def compute_profit_for_price(
    price_rub: float,
    product: Dict[str, Any],
    settings: Dict[str, Any],
    tariffs: Dict[str, Any],
    commission_rate_pct: float,
) -> Dict[str, float]:
    buyout = max(0.01, min(1.0, _to_float(settings["buyout_rate_pct"]) / 100.0))
    ad_rate = _to_float(settings["ad_rate_pct"]) / 100.0
    tax_rate = _to_float(settings["tax_rate_pct"]) / 100.0
    defect_rate = _to_float(settings["defect_rate_pct"]) / 100.0
    transfer_rate = float(tariffs["payment_transfer"][settings["transfer_schedule"]]) / 100.0
    reverse_same_locality_share = _to_float(settings["reverse_same_locality_share_pct"]) / 100.0
    payment_acceptance = float(tariffs["general"]["payment_acceptance_rub_per_sku"])
    returns_handling = float(tariffs["general"]["returns_handling_rub_per_shipment"])

    price_rub = float(price_rub)
    cost_rub = _to_float(product["Cost RUB"])
    volume_liters = int(product["Volume liters"])
    middle_mile = calc_middle_mile_rub(volume_liters, tariffs)
    delivery_to_buyer = price_based_delivery_fee(price_rub, tariffs)

    model = product["Resolved Model"]
    fbs_pickup = 0.0
    fbs_processing = 0.0
    if model == "FBS":
        fbs_pickup = calc_fbs_pickup_fee_per_unit(settings["fbs_avg_orders_per_day"], settings["avg_items_per_order"], tariffs)
        fbs_processing = calc_fbs_processing_fee_per_unit(
            bool(product["Is KGT Resolved"]),
            _to_float(product["Volumetric weight kg"]),
            settings["avg_items_per_order"],
            tariffs,
        )

    if product.get("Storage Cost RUB Override") is not None and str(product.get("Storage Cost RUB Override")).strip() != "":
        storage_cost = _to_float(product.get("Storage Cost RUB Override"), 0.0)
    else:
        storage_cost = _to_float(settings["fby_storage_cost_rub_per_unit"]) if model == "FBY" else 0.0

    if product.get("Inbound Cost RUB Override") is not None and str(product.get("Inbound Cost RUB Override")).strip() != "":
        inbound_cost = _to_float(product.get("Inbound Cost RUB Override"), 0.0)
    else:
        inbound_cost = _to_float(settings["fby_inbound_self_delivery_rub_per_unit"]) if model == "FBY" else 0.0

    expected_revenue = price_rub * buyout
    commission = price_rub * (commission_rate_pct / 100.0) * buyout
    payment_transfer = price_rub * transfer_rate * buyout
    tax = price_rub * tax_rate * buyout
    ad_cost = price_rub * ad_rate * buyout
    defect_cost = cost_rub * defect_rate

    reverse_middle_mile_factor = (1.0 - buyout) * (1.0 - reverse_same_locality_share)
    reverse_middle_mile = middle_mile * reverse_middle_mile_factor
    reverse_handling_cost = returns_handling * (1.0 - buyout)

    expected_cogs = cost_rub * buyout

    total_cost = (
        expected_cogs
        + commission
        + payment_transfer
        + payment_acceptance
        + delivery_to_buyer
        + middle_mile
        + reverse_middle_mile
        + reverse_handling_cost
        + tax
        + ad_cost
        + defect_cost
        + fbs_pickup
        + fbs_processing
        + storage_cost
        + inbound_cost
    )
    profit = expected_revenue - total_cost
    margin_pct = (profit / expected_revenue * 100.0) if expected_revenue > 0 else None

    return {
        "Expected Revenue RUB": expected_revenue,
        "Commission RUB": commission,
        "Payment Transfer RUB": payment_transfer,
        "Payment Acceptance RUB": payment_acceptance,
        "Delivery to Buyer RUB": delivery_to_buyer,
        "Middle Mile RUB": middle_mile,
        "Reverse Handling RUB": reverse_handling_cost,
        "Reverse Middle Mile RUB": reverse_middle_mile,
        "Tax RUB": tax,
        "Ads RUB": ad_cost,
        "Defect Cost RUB": defect_cost,
        "FBS Pickup RUB": fbs_pickup,
        "FBS Processing RUB": fbs_processing,
        "Storage RUB": storage_cost,
        "Inbound to Yandex WH RUB": inbound_cost,
        "Expected COGS RUB": expected_cogs,
        "Total Cost RUB": total_cost,
        "Profit RUB": profit,
        "Margin %": margin_pct,
    }


def solve_recommended_price(
    product: Dict[str, Any],
    settings: Dict[str, Any],
    tariffs: Dict[str, Any],
    commission_rate_pct: float,
    target_margin_pct: float,
    seed_price: float,
) -> float | None:
    buyout = max(0.01, min(1.0, _to_float(settings["buyout_rate_pct"]) / 100.0))
    if buyout <= 0:
        return None

    target_margin = max(0.0, min(0.95, float(target_margin_pct) / 100.0))

    def margin_for(price: float) -> float:
        metrics = compute_profit_for_price(price, product, settings, tariffs, commission_rate_pct)
        revenue = metrics["Expected Revenue RUB"]
        if revenue <= 0:
            return -999.0
        return metrics["Profit RUB"] / revenue

    low = 1.0
    high = max(seed_price * 2.0 if seed_price and seed_price > 0 else 1000.0, 1000.0)
    tries = 0
    while margin_for(high) < target_margin and tries < 40:
        high *= 1.5
        tries += 1
        if high > 1_000_000:
            return None

    for _ in range(80):
        mid = (low + high) / 2.0
        if margin_for(mid) >= target_margin:
            high = mid
        else:
            low = mid
    return round(high, 2)


def calculate_unit_economics(
    products_df: pd.DataFrame,
    settings: Dict[str, Any],
    tariffs: Dict[str, Any],
    commissions_df: pd.DataFrame,
    keyword_rules_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    results = []
    warnings = []

    for _, row in products_df.iterrows():
        product = row.to_dict()
        resolved_model = (_clean_string(product.get("Model")) or _clean_string(settings["default_model"])).upper()
        if resolved_model not in {"FBS", "FBY"}:
            resolved_model = settings["default_model"]

        manual_category = _clean_string(product.get("Manual Category"))
        if manual_category:
            category = manual_category
            category_source = "manual"
        else:
            category, category_source = detect_category(product["Product Name"], keyword_rules_df)

        commission_rate_pct, commission_is_exact = lookup_commission_rate(
            category, commissions_df, float(settings["default_commission_rate_pct"])
        )

        sum_sides = _to_float(product["Length cm"]) + _to_float(product["Width cm"]) + _to_float(product["Height cm"])
        auto_kgt = (
            _to_float(product["Weight kg"]) > _to_float(settings["kgt_weight_threshold_kg"])
            or sum_sides > _to_float(settings["kgt_sum_sides_threshold_cm"])
        )
        resolved_kgt = _is_trueish(product.get("Is KGT")) if not _blankish(product.get("Is KGT")) else auto_kgt

        volume_liters = calc_volume_liters(product["Length cm"], product["Width cm"], product["Height cm"])
        volumetric_weight_kg = calc_volumetric_weight_kg(product["Length cm"], product["Width cm"], product["Height cm"])

        product["Resolved Model"] = resolved_model
        product["Resolved Category"] = category
        product["Category Source"] = category_source
        product["Commission Rate %"] = commission_rate_pct
        product["Commission Exact Match"] = commission_is_exact
        product["Volume liters"] = volume_liters
        product["Volumetric weight kg"] = round(volumetric_weight_kg, 2)
        product["Sum sides cm"] = round(sum_sides, 2)
        product["Is KGT Resolved"] = resolved_kgt

        current_price = _to_float(product.get("Current Price RUB"), 0.0)
        target_margin_pct = _to_float(product.get("Target Margin % Override"), _to_float(settings["target_margin_pct"]))

        metrics = compute_profit_for_price(
            price_rub=current_price if current_price > 0 else 0.0,
            product=product,
            settings=settings,
            tariffs=tariffs,
            commission_rate_pct=commission_rate_pct,
        ) if current_price > 0 else {}

        rec_price = solve_recommended_price(
            product=product,
            settings=settings,
            tariffs=tariffs,
            commission_rate_pct=commission_rate_pct,
            target_margin_pct=target_margin_pct,
            seed_price=max(current_price, _to_float(product["Cost RUB"]) * 2.0),
        )

        out = {
            "SKU": product["SKU"],
            "Product Name": product["Product Name"],
            "Model": resolved_model,
            "Category": category,
            "Category Source": category_source,
            "Commission Rate %": round(commission_rate_pct, 2),
            "Commission Rate Status": "exact" if commission_is_exact else "fallback",
            "Current Price RUB": current_price if current_price > 0 else None,
            "Target Margin %": target_margin_pct,
            "Cost RUB": _to_float(product["Cost RUB"]),
            "Weight kg": _to_float(product["Weight kg"]),
            "Length cm": _to_float(product["Length cm"]),
            "Width cm": _to_float(product["Width cm"]),
            "Height cm": _to_float(product["Height cm"]),
            "Volume liters": volume_liters,
            "Volumetric weight kg": round(volumetric_weight_kg, 2),
            "Sum sides cm": round(sum_sides, 2),
            "Is KGT": resolved_kgt,
        }
        for key in [
            "Expected Revenue RUB",
            "Commission RUB",
            "Payment Transfer RUB",
            "Payment Acceptance RUB",
            "Delivery to Buyer RUB",
            "Middle Mile RUB",
            "Reverse Handling RUB",
            "Reverse Middle Mile RUB",
            "Tax RUB",
            "Ads RUB",
            "Defect Cost RUB",
            "FBS Pickup RUB",
            "FBS Processing RUB",
            "Storage RUB",
            "Inbound to Yandex WH RUB",
            "Expected COGS RUB",
            "Total Cost RUB",
            "Profit RUB",
            "Margin %",
        ]:
            out[key] = round(metrics[key], 2) if metrics and metrics.get(key) is not None else None

        out["Recommended Price RUB"] = rec_price
        if rec_price is not None and current_price > 0:
            out["Recommended Price Delta RUB"] = round(rec_price - current_price, 2)
            out["Recommended Price Delta %"] = round((rec_price / current_price - 1.0) * 100.0, 2) if current_price else None
        else:
            out["Recommended Price Delta RUB"] = None
            out["Recommended Price Delta %"] = None

        out["Notes"] = product.get("Notes")

        if category == "UNMAPPED":
            warnings.append(
                {
                    "SKU": product["SKU"],
                    "Warning": "Не удалось определить категорию по названию. Использован fallback commission.",
                }
            )
        elif not commission_is_exact:
            warnings.append(
                {
                    "SKU": product["SKU"],
                    "Warning": f"Для категории '{category}' не найдена точная комиссия. Использован default_commission_rate_pct.",
                }
            )
        if current_price <= 0:
            warnings.append(
                {
                    "SKU": product["SKU"],
                    "Warning": "Не указана текущая цена. Текущая прибыль/маржа не посчитаны, только рекомендованная цена.",
                }
            )

        results.append(out)

    result_df = pd.DataFrame(results)
    warnings_df = pd.DataFrame(warnings)

    percent_cols = ["Commission Rate %", "Target Margin %", "Margin %", "Recommended Price Delta %"]
    for col in percent_cols:
        if col in result_df.columns:
            result_df[col] = result_df[col].round(2)

    return result_df, warnings_df


def _apply_sheet_style(ws, currency_cols=None, percent_cols=None, freeze="A2"):
    currency_cols = set(currency_cols or [])
    percent_cols = set(percent_cols or [])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")

    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[1].height = 28

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top")
            if cell.column in currency_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
            if cell.column in percent_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '0.00'
    for idx, col_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for c in col_cells:
            value = "" if c.value is None else str(c.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 28)


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def build_result_workbook(
    result_df: pd.DataFrame,
    warnings_df: pd.DataFrame,
    settings: Dict[str, Any],
    tariffs: Dict[str, Any],
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    ws.append(list(result_df.columns))
    for row in result_df.itertuples(index=False):
        ws.append(list(row))

    settings_ws = wb.create_sheet("Settings Used")
    settings_ws.append(["parameter", "value"])
    for k, v in settings.items():
        settings_ws.append([k, v])

    tariff_ws = wb.create_sheet("Tariffs Snapshot")
    tariff_ws.append(["parameter", "value"])
    for k, v in tariffs["general"].items():
        tariff_ws.append([k, v])
    tariff_ws.append([])
    tariff_ws.append(["payment_transfer_schedule", "rate_pct"])
    for k, v in tariffs["payment_transfer"].items():
        tariff_ws.append([k, v])

    warn_ws = wb.create_sheet("Warnings")
    if warnings_df.empty:
        warn_ws.append(["Info"])
        warn_ws.append(["Без предупреждений"])
    else:
        warn_ws.append(list(warnings_df.columns))
        for row in warnings_df.itertuples(index=False):
            warn_ws.append(list(row))

    source_ws = wb.create_sheet("Sources")
    source_ws.append(["Title", "URL"])
    for title, url in [
        ("Тарифы FBS", "https://yandex.ru/support/marketplace/ru/introduction/rates/models/fbs"),
        ("Тарифы FBY", "https://yandex.ru/support/marketplace/ru/introduction/rates/models/fby"),
        ("Прием и перевод платежей", "https://yandex.ru/support/marketplace/ru/introduction/rates/acquiring"),
        ("Калькулятор дохода", "https://partner.market.yandex.ru/welcome/cost-calculator"),
    ]:
        source_ws.append([title, url])

    currency_headers = {
        "Current Price RUB", "Cost RUB", "Expected Revenue RUB", "Commission RUB", "Payment Transfer RUB",
        "Payment Acceptance RUB", "Delivery to Buyer RUB", "Middle Mile RUB", "Reverse Handling RUB",
        "Reverse Middle Mile RUB", "Tax RUB", "Ads RUB", "Defect Cost RUB", "FBS Pickup RUB",
        "FBS Processing RUB", "Storage RUB", "Inbound to Yandex WH RUB", "Expected COGS RUB",
        "Total Cost RUB", "Profit RUB", "Recommended Price RUB", "Recommended Price Delta RUB",
    }
    percent_headers = {"Commission Rate %", "Target Margin %", "Margin %", "Recommended Price Delta %"}
    header_to_idx = {cell.value: cell.column for cell in ws[1]}
    currency_cols = {header_to_idx[h] for h in currency_headers if h in header_to_idx}
    percent_cols = {header_to_idx[h] for h in percent_headers if h in header_to_idx}

    _apply_sheet_style(ws, currency_cols=currency_cols, percent_cols=percent_cols)
    _apply_sheet_style(settings_ws)
    _apply_sheet_style(tariff_ws)
    _apply_sheet_style(warn_ws)
    _apply_sheet_style(source_ws)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

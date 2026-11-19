"""
Lógica reutilizable para construir reportes pensionales (independiente de Streamlit).

Las reglas de negocio replican las de `pages/1_📄_ Load _PDF_Colpensiones.py`:
- `clean_to_num`: misma normalización de montos en formato colombiano ("$ 1.641.000").
- `build_weighted_monthly_report`: misma ponderación del IBC por mes
  (Σ días × IBC / días del mes; si ningún empleador llega a 30 días se suman con tope 30,
  en caso contrario se toma el máximo).
- `build_gaps_report`: misma detección de lagunas (meses ausentes = 30 días,
  meses parciales = 30 - días cotizados) y fusión de segmentos consecutivos.
"""

import math

import pandas as pd

MAX_DAYS_PER_MONTH = 30
# Separador de empresas cuando un mes tiene varios aportantes (la primera es la principal)
EMPLOYER_SEPARATOR = " / "


def is_really_empty(val) -> bool:
    """Indica si un valor está vacío a efectos prácticos (NaN, cadena vacía o cero)."""
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return True
    s = str(val).strip().lower()
    return s in ["", "-", "$ 0", "0", "0,00", "0.00", "$ 0,00", "$ 0.00"]


def clean_to_num(val) -> float:
    """Convierte montos con separadores de miles/decimales mixtos a float."""
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("$", "").replace(" ", "").strip()
    if not s or s == "-":
        return 0.0
    if "." in s and "," in s:
        if s.find(".") < s.find(","):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in [1, 2]:
            s = parts[0].replace(".", "") + "." + parts[1]
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts[-1]) == 3 and len(parts) > 1:
            s = s.replace(".", "")
        elif len(parts) == 2 and len(parts[1]) in [1, 2]:
            pass
        else:
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def round_half_up(value: float) -> int:
    """Redondeo comercial (como Excel): 0,5 sube. `round()` de Python redondea al par."""
    return int(math.floor(value + 0.5))


def build_weighted_monthly_report(
    df: pd.DataFrame,
    *,
    month_col: str,
    employer_col: str,
    days_col: str,
    ibc_col: str,
    document: str | int | None = None,
) -> pd.DataFrame:
    """
    Consolida las cotizaciones en una fila por mes con el IBC ponderado por días.

    Cada fila de entrada es un aporte real (no hay rangos expandidos), por lo que
    varias filas del MISMO empleador en un mes se SUMAN (días con tope 30 y
    Σ días × IBC) en lugar de tomar el máximo como en Colpensiones.

    Devuelve columnas: DOCUMENTO, EMPRESA, FECHA INICIAL, FECHA FINAL, IBC, DIAS.
    """
    columns = ["DOCUMENTO", "EMPRESA", "FECHA INICIAL", "FECHA FINAL", "IBC", "DIAS"]
    if df.empty:
        return pd.DataFrame(columns=columns)

    work = df[[month_col, employer_col, days_col, ibc_col]].copy()
    work["_days"] = pd.to_numeric(work[days_col], errors="coerce").fillna(0)
    work["_ibc"] = work[ibc_col].apply(clean_to_num)
    work["_weighted"] = work["_days"] * work["_ibc"]
    work["_period"] = pd.to_datetime(work[month_col]).dt.to_period("M")

    final_rows = []
    for period, group in work.groupby("_period", sort=True):
        per_employer = group.groupby(employer_col, sort=False).agg(
            days=("_days", "sum"), weighted=("_weighted", "sum")
        )
        per_employer["days"] = per_employer["days"].clip(upper=MAX_DAYS_PER_MONTH)

        # Misma regla que Colpensiones: si nadie llega a 30 días se suman (tope 30).
        if (per_employer["days"] < MAX_DAYS_PER_MONTH).all():
            month_days = min(per_employer["days"].sum(), MAX_DAYS_PER_MONTH)
        else:
            month_days = per_employer["days"].max()

        if month_days <= 0:
            continue

        ibc_weighted = per_employer["weighted"].sum() / month_days

        start_date = period.start_time
        if month_days < MAX_DAYS_PER_MONTH:
            end_date = start_date + pd.offsets.Day(int(month_days) - 1)
        else:
            end_date = period.end_time

        # Empresas del mes, primero la de mayor aporte ponderado.
        employers = per_employer.sort_values("weighted", ascending=False).index.tolist()

        final_rows.append(
            {
                "DOCUMENTO": document,
                "EMPRESA": EMPLOYER_SEPARATOR.join(employers),
                "FECHA INICIAL": start_date.date(),
                "FECHA FINAL": end_date.date(),
                "IBC": round_half_up(ibc_weighted),
                "DIAS": int(month_days),
            }
        )

    return pd.DataFrame(final_rows, columns=columns)


def build_gaps_report(
    df_final: pd.DataFrame,
    *,
    start_col: str = "FECHA INICIAL",
    end_col: str = "FECHA FINAL",
    days_col: str = "DIAS",
) -> pd.DataFrame:
    """
    Detecta lagunas pensionales (meses sin cotización o parcialmente cotizados)
    entre el primer y el último mes del reporte final, base 30 días por mes.
    """
    columns = ["FECHA INICIAL", "FECHA FINAL", "DÍAS FALTANTES (Base 30)", "SEMANAS"]
    if df_final is None or df_final.empty:
        return pd.DataFrame(columns=columns)

    df_g = df_final.copy()
    df_g["_ini"] = pd.to_datetime(df_g[start_col])
    df_g["_fin"] = pd.to_datetime(df_g[end_col])
    df_g["_period"] = df_g["_ini"].dt.to_period("M")
    df_g = df_g.sort_values("_ini").reset_index(drop=True)

    existing = {row["_period"]: row for _, row in df_g.iterrows()}
    all_months = pd.period_range(df_g["_period"].min(), df_g["_period"].max(), freq="M")

    # 1. Segmentos de laguna mes a mes
    segments = []
    for period in all_months:
        if period not in existing:
            segments.append(
                {
                    "start": period.start_time,
                    "end": period.end_time,
                    "days": MAX_DAYS_PER_MONTH,
                }
            )
            continue
        row = existing[period]
        if row[days_col] < MAX_DAYS_PER_MONTH:
            gap_start = row["_fin"] + pd.offsets.Day(1)
            if gap_start.month == row["_fin"].month:
                segments.append(
                    {
                        "start": gap_start,
                        "end": period.end_time,
                        "days": MAX_DAYS_PER_MONTH - row[days_col],
                    }
                )

    if not segments:
        return pd.DataFrame(columns=columns)

    # 2. Fusionar segmentos consecutivos
    merged = [segments[0].copy()]
    for seg in segments[1:]:
        current = merged[-1]
        if (seg["start"] - current["end"]).days <= 1:
            current["end"] = seg["end"]
            current["days"] += seg["days"]
        else:
            merged.append(seg.copy())

    # 3. Formato de salida
    gap_rows = [
        {
            "FECHA INICIAL": r["start"].date(),
            "FECHA FINAL": r["end"].date(),
            "DÍAS FALTANTES (Base 30)": int(r["days"]),
            "SEMANAS": round(r["days"] / 7, 2),
        }
        for r in merged
    ]

    return pd.DataFrame(gap_rows, columns=columns)

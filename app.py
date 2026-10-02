# ==============================================================
# LA TINKA AI - VERSION ACTUALIZADA
# ==============================================================
# Reemplaza la versión anterior de app.py.
#
# Mejoras principales:
# 1) Recarga real del Excel subido (sin caché obsoleto).
# 2) Detecta automáticamente el último sorteo.
# 3) Muestra alerta si el archivo está desactualizado.
# 4) Detecta los premios "reventados" mediante resaltado verde
#    en la columna Premio y permite usar esa información como
#    variable histórica disponible antes del siguiente sorteo.
# 5) Usa la hoja Modificaciones para determinar el universo
#    histórico de bolillas por fecha.
# 6) Backtesting walk-forward configurable.
# 7) Comparación contra una referencia aleatoria bajo las
#    mismas restricciones.
# 8) No utiliza el monto/reventón del sorteo objetivo para
#    predecir ese mismo sorteo (evita look-ahead).
#
# INSTALACIÓN:
# pip install streamlit pandas numpy openpyxl
#
# EJECUCIÓN:
# streamlit run app.py
# ==============================================================

from __future__ import annotations

import hashlib
import io
import os
from collections import Counter
from datetime import date, datetime, timedelta
from itertools import combinations
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from zoneinfo import ZoneInfo


# ==============================================================
# CONFIGURACIÓN
# ==============================================================

st.set_page_config(
    page_title="La Tinka AI",
    page_icon="🎱",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "3.3 - backtesting robusto con pd.NA/NaN y premios incompletos"

DEFAULT_FILES = [
    "La_Tinka_Todos_Los_Sorteos_1994_2026.xlsx",
    "Resultados la Tinka 1994 a 2026.xlsx",
    "Resultados_la_Tinka_1994_2026_revisado.xlsx",
]

BALL_COLUMNS = [
    "Bolilla 1",
    "Bolilla 2",
    "Bolilla 3",
    "Bolilla 4",
    "Bolilla 5",
    "Bolilla 6",
]

DATE_CANDIDATES = ["Fecha", "fecha", "FECHA"]
DRAW_CANDIDATES = ["Sorteo", "sorteo", "SORTEO"]


# ==============================================================
# PROGRAMACIÓN
# ==============================================================

DIAS_SORTEO = {2: "miércoles", 6: "domingo"}


def proximo_sorteo():
    ahora = datetime.now(ZoneInfo("America/Lima"))
    fecha = ahora.date()

    while fecha.weekday() not in DIAS_SORTEO:
        fecha += timedelta(days=1)

    return fecha, DIAS_SORTEO[fecha.weekday()]


# ==============================================================
# UTILIDADES
# ==============================================================

def normalize_text(text) -> str:
    if text is None:
        return ""
    return (
        str(text)
        .strip()
        .lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
        .replace("_", " ")
    )


def normalize_ball_column(name) -> str:
    s = normalize_text(name).replace(" ", "")
    replacements = {
        "bolilla1": "Bolilla 1",
        "bolilla2": "Bolilla 2",
        "bolilla3": "Bolilla 3",
        "bolilla4": "Bolilla 4",
        "bolilla5": "Bolilla 5",
        "bolilla6": "Bolilla 6",
        "bolilla01": "Bolilla 1",
        "bolilla02": "Bolilla 2",
        "bolilla03": "Bolilla 3",
        "bolilla04": "Bolilla 4",
        "bolilla05": "Bolilla 5",
        "bolilla06": "Bolilla 6",
    }
    return replacements.get(s, str(name).strip())


def stable_seed(*values) -> int:
    raw = "|".join(map(str, values))
    return int(hashlib.sha256(raw.encode("utf-8")).hexdigest()[:8], 16)


def find_file() -> Optional[Path]:
    possible = []
    for filename in DEFAULT_FILES:
        possible += [
            Path(filename),
            Path("datos") / filename,
            Path("/mnt/data") / filename,
        ]
    for path in possible:
        if path.exists():
            return path
    return None


def file_signature(file_obj_or_path) -> str:
    """Firma del contenido: evita que Streamlit reutilice un Excel viejo."""
    if isinstance(file_obj_or_path, (str, Path)):
        p = Path(file_obj_or_path)
        stat = p.stat()
        return f"{p.resolve()}|{stat.st_size}|{stat.st_mtime_ns}"

    data = file_obj_or_path.getvalue()
    return hashlib.sha256(data).hexdigest()


# ==============================================================
# CARGA DEL EXCEL
# ==============================================================

@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes, signature: str) -> Dict[str, pd.DataFrame]:
    """Lee el contenido del Excel. La firma obliga a recalcular cuando cambia."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return {
        sheet: pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
        for sheet in xls.sheet_names
    }


def read_excel_fresh(source) -> Dict[str, pd.DataFrame]:
    """Entrada única para archivos locales o UploadedFile."""
    if isinstance(source, (str, Path)):
        p = Path(source)
        data = p.read_bytes()
        sig = file_signature(p)
    else:
        data = source.getvalue()
        sig = hashlib.sha256(data).hexdigest()

    return load_excel_cached(data, sig)


# ==============================================================
# DETECCIÓN Y PREPARACIÓN
# ==============================================================

def detect_date_column(df: pd.DataFrame):
    normalized = {normalize_text(c): c for c in df.columns}
    for candidate in DATE_CANDIDATES:
        if normalize_text(candidate) in normalized:
            return normalized[normalize_text(candidate)]
    for c in df.columns:
        if "fecha" in normalize_text(c):
            return c
    return None


def detect_draw_column(df: pd.DataFrame):
    normalized = {normalize_text(c): c for c in df.columns}
    for candidate in DRAW_CANDIDATES:
        if normalize_text(candidate) in normalized:
            return normalized[normalize_text(candidate)]
    for c in df.columns:
        if "sorteo" in normalize_text(c) and "extra" not in normalize_text(c):
            return c
    return None


def detect_ball_columns(df: pd.DataFrame) -> List[str]:
    found = {}
    for c in df.columns:
        nc = normalize_ball_column(c)
        if nc in BALL_COLUMNS:
            found[nc] = c

    missing = [c for c in BALL_COLUMNS if c not in found]
    if missing:
        raise ValueError(
            "No se encontraron las seis columnas de bolillas. "
            f"Faltan: {', '.join(missing)}"
        )
    return [found[c] for c in BALL_COLUMNS]


def prepare_results(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    date_col = detect_date_column(df)
    draw_col = detect_draw_column(df)
    ball_cols = detect_ball_columns(df)

    if date_col is None:
        raise ValueError("No se encontró la columna Fecha.")

    rename = {date_col: "Fecha"}
    if draw_col:
        rename[draw_col] = "Sorteo"

    for i, c in enumerate(ball_cols, 1):
        rename[c] = f"Bolilla {i}"

    df = df.rename(columns=rename)

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=True)

    if "Sorteo" not in df.columns:
        df["Sorteo"] = np.arange(1, len(df) + 1)

    df["Sorteo"] = pd.to_numeric(df["Sorteo"], errors="coerce")

    for c in BALL_COLUMNS:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    before = len(df)
    df = df.dropna(subset=["Fecha"] + BALL_COLUMNS).copy()
    dropped = before - len(df)

    for c in BALL_COLUMNS:
        df[c] = df[c].astype(int)

    df = df.sort_values(["Fecha", "Sorteo"]).reset_index(drop=True)

    # Conservamos columnas auxiliares del archivo.
    if "Premio" in df.columns:
        df["Premio"] = pd.to_numeric(df["Premio"], errors="coerce")

    df.attrs["rows_dropped"] = dropped
    return df


# ==============================================================
# DETECCIÓN DE REVENTONES POR COLOR
# ==============================================================

def is_green_fill(cell) -> bool:
    """Detecta el resaltado verde usado en el Excel del usuario."""
    fill = cell.fill
    if not fill or fill.fill_type != "solid":
        return False

    fg = fill.fgColor

    # En el Excel analizado, el verde corresponde al tema 9.
    if fg.type == "theme" and fg.theme == 9:
        return True

    if fg.type == "rgb" and fg.rgb:
        rgb = str(fg.rgb)[-6:].upper()
        try:
            r = int(rgb[0:2], 16)
            g = int(rgb[2:4], 16)
            b = int(rgb[4:6], 16)
            return g > r * 1.15 and g > b * 1.05 and g > 90
        except Exception:
            return False

    return False


@st.cache_data(show_spinner=False)
def detect_green_prizes(file_bytes: bytes, signature: str) -> pd.DataFrame:
    """Lee estilo de la columna Premio sin alterar los datos."""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    if "Resultados" not in wb.sheetnames:
        return pd.DataFrame(columns=["Sorteo", "Fecha", "ReventoVerde"])

    ws = wb["Resultados"]
    headers = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}

    date_col = headers.get("Fecha")
    draw_col = headers.get("Sorteo")
    prize_col = headers.get("Premio")

    if not date_col or not draw_col or not prize_col:
        return pd.DataFrame(columns=["Sorteo", "Fecha", "ReventoVerde"])

    records = []
    for row in range(2, ws.max_row + 1):
        date_value = ws.cell(row, date_col).value
        draw_value = ws.cell(row, draw_col).value
        prize_cell = ws.cell(row, prize_col)

        if draw_value is None:
            continue

        records.append(
            {
                "Sorteo": pd.to_numeric(draw_value, errors="coerce"),
                "Fecha": pd.to_datetime(date_value, errors="coerce"),
                "ReventoVerde": bool(is_green_fill(prize_cell)),
            }
        )

    return pd.DataFrame(records)


def attach_prize_features(
    df: pd.DataFrame,
    file_bytes: bytes,
    signature: str,
) -> pd.DataFrame:
    """
    Integra premios/reventones respetando datos faltantes.

    ReventoVerde:
      True  = confirmado por resaltado verde.
      False = existe premio documentado y no está marcado en verde.
      NA    = no existe información suficiente.
    """
    out = df.copy().sort_values(["Fecha", "Sorteo"]).reset_index(drop=True)

    if "Premio" not in out.columns:
        out["Premio"] = np.nan
    out["Premio"] = pd.to_numeric(out["Premio"], errors="coerce")

    green = detect_green_prizes(file_bytes, signature)

    if not green.empty:
        green = green.dropna(subset=["Sorteo"]).copy()
        green["Sorteo"] = pd.to_numeric(
            green["Sorteo"], errors="coerce"
        ).astype("Int64")
        out["Sorteo"] = pd.to_numeric(
            out["Sorteo"], errors="coerce"
        ).astype("Int64")

        green_map = green[["Sorteo", "ReventoVerde"]].drop_duplicates("Sorteo")
        out = out.merge(green_map, on="Sorteo", how="left", suffixes=("", "_green"))

        if "ReventoVerde_green" in out.columns:
            # Caso en que la tabla base ya traía una columna con ese nombre.
            out["ReventoVerde"] = out["ReventoVerde_green"].astype("boolean")
            out = out.drop(columns=["ReventoVerde_green"], errors="ignore")
        elif "ReventoVerde" in out.columns:
            # Caso normal: la columna proviene directamente de green_map.
            out["ReventoVerde"] = out["ReventoVerde"].astype("boolean")
        else:
            out["ReventoVerde"] = pd.Series(
                pd.NA, index=out.index, dtype="boolean"
            )
    else:
        out["ReventoVerde"] = pd.Series(
            pd.NA, index=out.index, dtype="boolean"
        )

    # Si el monto del premio existe y no hay marca verde, podemos clasificar
    # el reventón como "no marcado" dentro del archivo. Si el monto falta,
    # permanece desconocido.
    out.loc[
        out["ReventoVerde"].isna() & out["Premio"].notna(),
        "ReventoVerde"
    ] = False

    out["ReventoConocido"] = out["ReventoVerde"].notna()

    out["PremioAnterior"] = out["Premio"].shift(1)
    out["PremioCambioPct"] = (
        (out["Premio"] - out["PremioAnterior"])
        / out["PremioAnterior"].replace(0, np.nan)
    )

    # Variables disponibles ANTES del sorteo actual.
    out["ReventoAnterior"] = out["ReventoVerde"].shift(1).astype("boolean")
    out["PremioConocidoAntes"] = out["PremioAnterior"]

    streak = []
    n = 0
    known = True

    for value in out["ReventoVerde"].tolist():
        streak.append(float(n) if known else np.nan)

        if pd.isna(value):
            # No sabemos si hubo o no reventón en este período.
            # No continuamos acumulando una racha inventada.
            known = False
        elif bool(value):
            n = 0
            known = True
        else:
            n += 1
            known = True

    out["SorteosSinReventarAntes"] = streak

    last_revent_date = None
    days_since = []

    for _, row in out.iterrows():
        if last_revent_date is None:
            days_since.append(np.nan)
        else:
            days_since.append(
                (row["Fecha"] - last_revent_date).days
            )

        if (
            pd.notna(row["ReventoVerde"])
            and bool(row["ReventoVerde"])
        ):
            last_revent_date = row["Fecha"]

    out["DiasDesdeUltimoReventonAntes"] = days_since
    out["PremioDisponible"] = out["Premio"].notna()

    return out


# ==============================================================
# UNIVERSO HISTÓRICO
# ==============================================================

def prepare_universe_schedule(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Usa la hoja Modificaciones.
    La regla utilizada es: para una fecha determinada, toma la última
    Fecha Inicio que ya entró en vigencia.
    """
    if "Modificaciones" not in sheets:
        return pd.DataFrame()

    mod = sheets["Modificaciones"].copy()

    date_start = next(
        (c for c in mod.columns if normalize_text(c) == "fecha inicio"),
        None,
    )
    total_col = next(
        (c for c in mod.columns if "total bolillas" in normalize_text(c)),
        None,
    )

    if date_start is None or total_col is None:
        return pd.DataFrame()

    mod["Fecha Inicio"] = pd.to_datetime(mod[date_start], errors="coerce", dayfirst=True)
    mod["Total Bolillas"] = pd.to_numeric(mod[total_col], errors="coerce")
    mod = mod.dropna(subset=["Fecha Inicio", "Total Bolillas"]).copy()
    mod["Total Bolillas"] = mod["Total Bolillas"].astype(int)
    return mod.sort_values("Fecha Inicio").reset_index(drop=True)


def universe_for_date(
    target_date,
    schedule: pd.DataFrame,
    fallback: int = 53,
) -> int:
    if schedule.empty:
        return fallback

    target = pd.Timestamp(target_date)
    valid = schedule[schedule["Fecha Inicio"] <= target]

    if valid.empty:
        return int(schedule.iloc[0]["Total Bolillas"])

    return int(valid.iloc[-1]["Total Bolillas"])


def add_historical_universe(df: pd.DataFrame, schedule: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["UniversoVigente"] = [
        universe_for_date(d, schedule) for d in out["Fecha"]
    ]
    return out


def get_ball_universe(df: pd.DataFrame) -> List[int]:
    values = set()
    for c in BALL_COLUMNS:
        values.update(pd.to_numeric(df[c], errors="coerce").dropna().astype(int).tolist())
    return sorted(values)


def get_historical_universe_until(
    df: pd.DataFrame,
    cutoff_date=None,
) -> List[int]:
    if cutoff_date is not None:
        d = df[df["Fecha"] < pd.Timestamp(cutoff_date)]
    else:
        d = df

    return get_ball_universe(d)


def valid_numbers_for_training(
    train: pd.DataFrame,
    cutoff_date,
    schedule: pd.DataFrame,
) -> List[int]:
    """Universo oficial vigente en el momento del objetivo, restringido al histórico disponible."""
    n = universe_for_date(cutoff_date, schedule)
    # El universo oficial es 1..n. No usamos solo números observados:
    # un número válido puede no haber aparecido todavía.
    return list(range(1, n + 1))


def safe_bool(value):
    """Convierte valores a bool sin fallar con NaN o pd.NA."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return bool(value)


def prize_coverage(df: pd.DataFrame) -> Dict:
    """Calcula la cobertura real de premios y reventones."""
    total = len(df)

    if total == 0:
        return {
            "total": 0,
            "premios": 0,
            "cobertura_premio_pct": 0.0,
            "reventones_confirmados": 0,
            "reventones_conocidos": 0,
            "cobertura_reventon_pct": 0.0,
        }

    premios = int(df["Premio"].notna().sum()) if "Premio" in df else 0
    conocidos = (
        int(df["ReventoVerde"].notna().sum())
        if "ReventoVerde" in df else 0
    )
    confirmados = (
        int(df["ReventoVerde"].fillna(False).astype(bool).sum())
        if "ReventoVerde" in df else 0
    )

    return {
        "total": total,
        "premios": premios,
        "cobertura_premio_pct": premios / total * 100,
        "reventones_confirmados": confirmados,
        "reventones_conocidos": conocidos,
        "cobertura_reventon_pct": conocidos / total * 100,
    }


# ==============================================================
# VALIDACIÓN
# ==============================================================

def validate_data(df: pd.DataFrame) -> Dict:
    duplicate_draws = int(df["Sorteo"].duplicated(keep=False).sum())
    exact_duplicates = int(df.duplicated().sum())

    repeated_balls = 0
    out_of_range = 0

    for _, row in df.iterrows():
        vals = [int(row[c]) for c in BALL_COLUMNS]
        if len(set(vals)) < 6:
            repeated_balls += 1
        if any(v < 1 or v > 100 for v in vals):
            out_of_range += 1

    historical_combos = {
        tuple(sorted(int(row[c]) for c in BALL_COLUMNS))
        for _, row in df.iterrows()
    }

    return {
        "records": len(df),
        "min_date": df["Fecha"].min(),
        "max_date": df["Fecha"].max(),
        "duplicate_draws": duplicate_draws,
        "exact_duplicates": exact_duplicates,
        "repeated_balls": repeated_balls,
        "out_of_range": out_of_range,
        "unique_combinations": len(historical_combos),
        "universe_min": min(get_ball_universe(df)) if len(df) else None,
        "universe_max": max(get_ball_universe(df)) if len(df) else None,
        "reventones": (
            int(df["ReventoVerde"].fillna(False).astype(bool).sum())
            if "ReventoVerde" in df else 0
        ),
        "premio_cobertura_pct": prize_coverage(df)["cobertura_premio_pct"],
        "reventon_cobertura_pct": prize_coverage(df)["cobertura_reventon_pct"],
    }


# ==============================================================
# ESTADÍSTICAS DE NÚMEROS
# ==============================================================

def calculate_frequency(df: pd.DataFrame, decay_lambda: float = 0.0) -> Counter:
    freq = Counter()

    if df.empty:
        return freq

    max_date = df["Fecha"].max()

    for _, row in df.iterrows():
        weight = 1.0

        if decay_lambda > 0:
            days = (max_date - row["Fecha"]).days
            weight = np.exp(-decay_lambda * days)

        for c in BALL_COLUMNS:
            freq[int(row[c])] += weight

    return freq


def calculate_delay(df: pd.DataFrame, universe: List[int]) -> Dict[int, float]:
    if df.empty:
        return {n: np.nan for n in universe}

    last_date = df["Fecha"].max()
    result = {}

    for n in universe:
        dates = df.loc[
            (df[BALL_COLUMNS] == n).any(axis=1), "Fecha"
        ]
        if dates.empty:
            result[n] = np.nan
        else:
            result[n] = float((last_date - dates.max()).days)

    return result


def minmax_dict(values: Dict[int, float], neutral: float = 0.5) -> Dict[int, float]:
    finite = [v for v in values.values() if pd.notna(v) and np.isfinite(v)]

    if not finite:
        return {k: neutral for k in values}

    lo, hi = min(finite), max(finite)

    if hi == lo:
        return {k: neutral for k in values}

    return {
        k: neutral if pd.isna(v) else (v - lo) / (hi - lo)
        for k, v in values.items()
    }


def calculate_contextual_ball_scores(
    df: pd.DataFrame,
    universe: List[int],
    current_reventon_context: bool,
    blend_context: float = 0.20,
) -> Dict[int, float]:
    """
    Compara la frecuencia de cada número en sorteos cuyo estado previo
    fue igual al contexto actual. Es una variable experimental/descriptiva,
    no una afirmación de causalidad.
    """
    base = calculate_ball_scores(df, universe, 0.65, 0.35)

    if "ReventoAnterior" not in df.columns:
        return base

    known_context = df["ReventoAnterior"].notna()
    contextual = df[
        known_context
        & (df["ReventoAnterior"].astype(bool) == bool(current_reventon_context))
    ].copy()

    # Evitar que una muestra contextual demasiado pequeña domine el modelo.
    if len(contextual) < 20:
        return base

    context_freq = calculate_frequency(contextual, 0)
    context_values = {n: context_freq[n] for n in universe}
    context_norm = minmax_dict(context_values)

    return {
        n: (1 - blend_context) * base[n] + blend_context * context_norm[n]
        for n in universe
    }


def calculate_ball_scores(
    df: pd.DataFrame,
    universe: List[int],
    historical_weight: float = 0.65,
    recent_weight: float = 0.35,
) -> Dict[int, float]:
    hist = calculate_frequency(df, 0)
    recent = calculate_frequency(df, 0.003)

    hist_values = {n: hist[n] for n in universe}
    recent_values = {n: recent[n] for n in universe}

    h = minmax_dict(hist_values)
    r = minmax_dict(recent_values)

    return {
        n: historical_weight * h[n] + recent_weight * r[n]
        for n in universe
    }


# ==============================================================
# PARES
# ==============================================================

def calculate_pair_frequency(df: pd.DataFrame) -> Counter:
    pairs = Counter()

    for _, row in df.iterrows():
        nums = sorted(int(row[c]) for c in BALL_COLUMNS)
        for pair in combinations(nums, 2):
            pairs[pair] += 1

    return pairs


def pair_score(combo: Tuple[int, ...], pairs: Counter) -> float:
    if not combo:
        return 0.0
    vals = [np.log1p(pairs[p]) for p in combinations(sorted(combo), 2)]
    return float(np.mean(vals)) if vals else 0.0


# ==============================================================
# ESTRUCTURA
# ==============================================================

def get_low_threshold(universe_size: int) -> int:
    return universe_size // 2


def structure_distribution(combo: Tuple[int, ...], universe_size: int) -> Dict:
    nums = sorted(combo)

    odd = sum(n % 2 for n in nums)
    consecutive = sum(
        1 for a, b in zip(nums, nums[1:]) if b == a + 1
    )
    low = sum(n <= get_low_threshold(universe_size) for n in nums)

    return {
        "odd": odd,
        "consecutive": consecutive,
        "low": low,
        "sum": sum(nums),
    }


def structure_statistics(df: pd.DataFrame, universe_size: int) -> Dict[str, Tuple[float, float]]:
    values = {"odd": [], "consecutive": [], "low": [], "sum": []}

    for _, row in df.iterrows():
        combo = tuple(int(row[c]) for c in BALL_COLUMNS)
        row_universe = int(row.get("UniversoVigente", universe_size))
        s = structure_distribution(combo, row_universe)
        for key in values:
            values[key].append(s[key])

    result = {}
    for key, arr in values.items():
        if not arr:
            result[key] = (0.0, 1.0)
        else:
            result[key] = (float(np.mean(arr)), float(np.std(arr) or 1.0))
    return result


def structure_score(
    combo: Tuple[int, ...],
    stats: Dict[str, Tuple[float, float]],
    universe_size: int,
) -> float:
    s = structure_distribution(combo, universe_size)
    z = []

    for key in ["odd", "consecutive", "low", "sum"]:
        mean, std = stats[key]
        z.append(abs(s[key] - mean) / std if std else 0.0)

    return float(np.exp(-0.5 * np.mean(z)))


# ==============================================================
# COMBINACIONES
# ==============================================================

def weighted_sample_without_replacement(
    universe: List[int],
    scores: Dict[int, float],
    k: int,
    rng: np.random.Generator,
) -> Tuple[int, ...]:
    available = list(universe)
    chosen = []

    for _ in range(k):
        weights = np.array(
            [max(float(scores.get(n, 0.0)), 1e-9) for n in available],
            dtype=float,
        )
        weights /= weights.sum()

        idx = int(rng.choice(len(available), p=weights))
        chosen.append(available.pop(idx))

    return tuple(sorted(chosen))


def validate_combination(
    combo: Tuple[int, ...],
    universe: List[int],
    min_sum: Optional[int] = None,
    max_sum: Optional[int] = None,
    max_consecutive: Optional[int] = None,
    exclude_historical: bool = False,
    historical_combinations: Optional[set] = None,
) -> bool:
    if len(combo) != 6 or len(set(combo)) != 6:
        return False

    if any(n not in universe for n in combo):
        return False

    total = sum(combo)

    if min_sum is not None and total < min_sum:
        return False
    if max_sum is not None and total > max_sum:
        return False

    consecutive = sum(
        1 for a, b in zip(sorted(combo), sorted(combo)[1:]) if b == a + 1
    )

    if max_consecutive is not None and consecutive > max_consecutive:
        return False

    if exclude_historical and historical_combinations is not None:
        if tuple(sorted(combo)) in historical_combinations:
            return False

    return True


def combination_score(
    combo: Tuple[int, ...],
    ball_scores: Dict[int, float],
    pairs: Counter,
    structure_stats_: Dict,
    universe_size: int,
    w_ball: float = 0.60,
    w_pair: float = 0.25,
    w_structure: float = 0.15,
) -> float:
    ball = float(np.mean([ball_scores.get(n, 0.5) for n in combo]))

    raw_pair = pair_score(combo, pairs)
    pair_scale = max(1.0, np.log1p(max(pairs.values(), default=1)))
    pair = min(1.0, raw_pair / pair_scale)

    structure = structure_score(combo, structure_stats_, universe_size)

    return (
        w_ball * ball
        + w_pair * pair
        + w_structure * structure
    )


def generate_candidates(
    train: pd.DataFrame,
    universe: List[int],
    n_candidates: int = 5000,
    seed: int = 12345,
    min_sum: Optional[int] = None,
    max_sum: Optional[int] = None,
    max_consecutive: Optional[int] = None,
    exclude_historical: bool = False,
    w_ball: float = 0.60,
    w_pair: float = 0.25,
    w_structure: float = 0.15,
    use_reventon_context: bool = False,
    current_reventon_context: bool = False,
) -> Tuple[pd.DataFrame, Dict]:
    rng = np.random.default_rng(seed)

    if use_reventon_context:
        ball_scores = calculate_contextual_ball_scores(
            train,
            universe,
            current_reventon_context,
            blend_context=0.20,
        )
    else:
        ball_scores = calculate_ball_scores(train, universe)

    pairs = calculate_pair_frequency(train)

    # El universo puede haber cambiado históricamente.
    # Para la generación actual usamos el tamaño vigente.
    universe_size = max(universe)
    stats = structure_statistics(train, universe_size)

    historical_combinations = {
        tuple(sorted(int(row[c]) for c in BALL_COLUMNS))
        for _, row in train.iterrows()
    }

    candidates = {}
    attempts = 0
    max_attempts = max(1000, n_candidates * 30)

    while len(candidates) < n_candidates and attempts < max_attempts:
        attempts += 1

        combo = weighted_sample_without_replacement(
            universe, ball_scores, 6, rng
        )

        if not validate_combination(
            combo,
            universe,
            min_sum=min_sum,
            max_sum=max_sum,
            max_consecutive=max_consecutive,
            exclude_historical=exclude_historical,
            historical_combinations=historical_combinations,
        ):
            continue

        score = combination_score(
            combo,
            ball_scores,
            pairs,
            stats,
            universe_size,
            w_ball=w_ball,
            w_pair=w_pair,
            w_structure=w_structure,
        )

        candidates[combo] = score

    rows = [
        {"Combinación": combo, "Score": score}
        for combo, score in candidates.items()
    ]

    result = pd.DataFrame(rows)

    if not result.empty:
        result = result.sort_values(
            "Score", ascending=False
        ).reset_index(drop=True)

    info = {
        "attempts": attempts,
        "generated": len(result),
        "acceptance_rate": len(result) / attempts if attempts else 0,
        "requested": n_candidates,
    }

    return result, info


def select_top_combinations(
    candidates: pd.DataFrame,
    top_n: int = 3,
    diversity: bool = True,
    max_overlap: int = 4,
) -> pd.DataFrame:
    if candidates.empty:
        return candidates

    selected = []

    for _, row in candidates.iterrows():
        combo = tuple(row["Combinación"])

        if diversity:
            if any(
                len(set(combo) & set(prev["Combinación"])) > max_overlap
                for prev in selected
            ):
                continue

        selected.append(row)

        if len(selected) >= top_n:
            break

    return pd.DataFrame(selected).reset_index(drop=True)


# ==============================================================
# REFERENCIA ALEATORIA
# ==============================================================

def random_combinations(
    universe: List[int],
    n: int,
    rng: np.random.Generator,
    min_sum: Optional[int] = None,
    max_sum: Optional[int] = None,
    max_consecutive: Optional[int] = None,
    exclude: Optional[set] = None,
) -> List[Tuple[int, ...]]:
    results = []
    seen = set()
    attempts = 0
    max_attempts = max(1000, n * 50)

    while len(results) < n and attempts < max_attempts:
        attempts += 1
        combo = tuple(sorted(rng.choice(universe, size=6, replace=False).tolist()))

        if combo in seen:
            continue

        if not validate_combination(
            combo,
            universe,
            min_sum=min_sum,
            max_sum=max_sum,
            max_consecutive=max_consecutive,
            exclude_historical=False,
        ):
            continue

        if exclude and combo in exclude:
            continue

        seen.add(combo)
        results.append(combo)

    return results


# ==============================================================
# BACKTESTING
# ==============================================================

def evaluate_combination(combo: Tuple[int, ...], actual: Tuple[int, ...]) -> int:
    return len(set(combo) & set(actual))


def evaluate_ticket_set(
    tickets: List[Tuple[int, ...]],
    actual: Tuple[int, ...],
) -> Dict:
    hits = [evaluate_combination(c, actual) for c in tickets]

    return {
        "max_hits": max(hits) if hits else 0,
        "mean_hits": float(np.mean(hits)) if hits else 0.0,
        "hits_2_plus": int(sum(h >= 2 for h in hits)),
        "hits_3_plus": int(sum(h >= 3 for h in hits)),
        "hits_4_plus": int(sum(h >= 4 for h in hits)),
    }


def backtest_strategy(
    df: pd.DataFrame,
    schedule: pd.DataFrame,
    min_train: int = 200,
    step: int = 1,
    top_n: int = 10,
    n_candidates: int = 1500,
    min_sum: Optional[int] = None,
    max_sum: Optional[int] = None,
    max_consecutive: Optional[int] = None,
    strategy: str = "Modelo",
    exclude_historical: bool = False,
    use_reventon_feature: bool = False,
) -> pd.DataFrame:
    data = df.sort_values(["Fecha", "Sorteo"]).reset_index(drop=True)

    records = []

    for test_index in range(min_train, len(data), max(1, step)):
        train = data.iloc[:test_index].copy()
        actual_row = data.iloc[test_index]

        actual = tuple(int(actual_row[c]) for c in BALL_COLUMNS)

        # Universo correcto para el momento del sorteo objetivo.
        universe_size = universe_for_date(
            actual_row["Fecha"], schedule
        )
        universe = list(range(1, universe_size + 1))

        # Reventón: solo variables de filas anteriores.
        # No se utiliza ReventoVerde de actual_row.
        if use_reventon_feature:
            # Se incluye como contexto únicamente si existe información
            # suficiente ANTES del sorteo objetivo.
            prev_revent_raw = train.iloc[-1]["ReventoVerde"]
            streak_raw = train.iloc[-1]["SorteosSinReventarAntes"]

            prev_revent_value = safe_bool(prev_revent_raw)
            if prev_revent_value is None or pd.isna(streak_raw):
                # Información incompleta: no inventamos el contexto.
                prev_revent = "desconocido"
                streak = "desconocido"
            else:
                prev_revent = bool(prev_revent_raw)
                streak = int(float(streak_raw)) + (0 if prev_revent else 1)

            seed_context = (
                f"{strategy}|{actual_row['Fecha'].date()}|"
                f"{prev_revent}|{streak}"
            )
        else:
            seed_context = f"{strategy}|{actual_row['Fecha'].date()}"

        seed = stable_seed("backtest", seed_context, test_index)

        if strategy == "Aleatorio":
            rng = np.random.default_rng(seed)
            tickets = random_combinations(
                universe,
                top_n,
                rng,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
            )

        else:
            candidates, info = generate_candidates(
                train=train,
                universe=universe,
                n_candidates=n_candidates,
                seed=seed,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
                exclude_historical=exclude_historical,
                use_reventon_context=use_reventon_feature,
                current_reventon_context=(
                    safe_bool(train.iloc[-1]["ReventoVerde"])
                    if use_reventon_feature and "ReventoVerde" in train.columns
                    else False
                ),
            )

            if strategy == "Frecuencia":
                # Generación basada solamente en frecuencia histórica.
                # Reordena candidatos con score de frecuencia.
                scores = calculate_ball_scores(train, universe, 1.0, 0.0)
                if not candidates.empty:
                    candidates = candidates.copy()
                    candidates["StrategyScore"] = candidates["Combinación"].apply(
                        lambda c: float(np.mean([scores[n] for n in c]))
                    )
                    candidates = candidates.sort_values(
                        "StrategyScore", ascending=False
                    )
            elif strategy == "Frecuencia + Recencia":
                scores = calculate_ball_scores(train, universe, 0.65, 0.35)
                if not candidates.empty:
                    candidates = candidates.copy()
                    candidates["StrategyScore"] = candidates["Combinación"].apply(
                        lambda c: float(np.mean([scores[n] for n in c]))
                    )
                    candidates = candidates.sort_values(
                        "StrategyScore", ascending=False
                    )

            selected = select_top_combinations(
                candidates, top_n=top_n, diversity=True
            )
            tickets = [
                tuple(x) for x in selected["Combinación"].tolist()
            ] if not selected.empty else []

        metrics = evaluate_ticket_set(tickets, actual)

        records.append(
            {
                "Fecha": actual_row["Fecha"],
                "Sorteo": actual_row["Sorteo"],
                "Estrategia": strategy,
                "Tickets": len(tickets),
                **metrics,
            }
        )

    return pd.DataFrame(records)


def summarize_backtest(results: pd.DataFrame) -> pd.DataFrame:
    if results.empty:
        return pd.DataFrame()

    grouped = []

    for strategy, g in results.groupby("Estrategia"):
        grouped.append(
            {
                "Estrategia": strategy,
                "Pruebas": len(g),
                "Promedio máximo aciertos": g["max_hits"].mean(),
                "Mediana máximo aciertos": g["max_hits"].median(),
                "% pruebas con ≥2": (g["max_hits"] >= 2).mean() * 100,
                "% pruebas con ≥3": (g["max_hits"] >= 3).mean() * 100,
                "% pruebas con ≥4": (g["max_hits"] >= 4).mean() * 100,
                "Máximo observado": g["max_hits"].max(),
            }
        )

    return pd.DataFrame(grouped)


# ==============================================================
# TABLA DE FRECUENCIAS
# ==============================================================

def frequency_table(df: pd.DataFrame, universe: List[int]) -> pd.DataFrame:
    freq = calculate_frequency(df, 0)
    delay = calculate_delay(df, universe)

    total_balls = len(df) * 6

    rows = []
    for n in universe:
        rows.append(
            {
                "Bolilla": n,
                "Apariciones": int(freq[n]),
                "% sobre bolillas": (
                    freq[n] / total_balls * 100 if total_balls else 0
                ),
                "Atraso días": delay[n],
            }
        )

    return pd.DataFrame(rows).sort_values(
        ["Apariciones", "Bolilla"], ascending=[False, True]
    ).reset_index(drop=True)


# ==============================================================
# EXPORTACIÓN
# ==============================================================

def candidates_to_dataframe(df_candidates: pd.DataFrame) -> pd.DataFrame:
    if df_candidates.empty:
        return pd.DataFrame()

    rows = []
    for rank, (_, row) in enumerate(df_candidates.iterrows(), 1):
        combo = tuple(row["Combinación"])
        item = {
            "Ranking": rank,
            "Combinación": "-".join(f"{n:02d}" for n in combo),
            "Score": row.get("Score", row.get("StrategyScore", np.nan)),
        }
        for i, n in enumerate(combo, 1):
            item[f"Bolilla {i}"] = n
        rows.append(item)

    return pd.DataFrame(rows)


# ==============================================================
# INTERFAZ
# ==============================================================

def main():

    st.title("🎱 LA TINKA AI")
    st.caption(
        f"Sistema de análisis histórico y generación de combinaciones · {APP_VERSION}"
    )

    st.info(
        "Este sistema no afirma predecir el resultado de La Tinka. "
        "Genera combinaciones y permite evaluar históricamente si determinadas "
        "reglas producen resultados distintos a referencias aleatorias."
    )

    # ----------------------------------------------------------
    # SIDEBAR: ARCHIVO
    # ----------------------------------------------------------

    st.sidebar.header("📁 Datos")

    if st.sidebar.button("🔄 Limpiar caché y recargar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    uploaded = st.sidebar.file_uploader(
        "Cargar Excel actualizado",
        type=["xlsx", "xlsm"],
        help="Sube el archivo Excel más reciente. El contenido nuevo se detecta automáticamente.",
    )

    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        file_sig = hashlib.sha256(file_bytes).hexdigest()
        source_name = uploaded.name
        try:
            sheets = read_excel_fresh(uploaded)
        except Exception as exc:
            st.error(
                "No se pudo leer el Excel. "
                f"Detalle técnico: {type(exc).__name__}: {exc}"
            )
            st.stop()
    else:
        default_path = find_file()

        if default_path is None:
            st.error(
                "No se encontró el Excel. Sube el archivo actualizado desde la barra lateral."
            )
            st.stop()

        file_bytes = default_path.read_bytes()
        file_sig = file_signature(default_path)
        source_name = default_path.name
        try:
            sheets = read_excel_fresh(default_path)
        except Exception as exc:
            st.error(
                "No se pudo leer el Excel predeterminado. "
                f"Detalle técnico: {type(exc).__name__}: {exc}"
            )
            st.stop()

    if "Resultados" not in sheets:
        st.error(
            "El Excel debe contener una hoja llamada 'Resultados'. "
            f"Hojas encontradas: {', '.join(sheets.keys())}"
        )
        st.stop()

    try:
        df = prepare_results(sheets["Resultados"])
        df = attach_prize_features(df, file_bytes, file_sig)

        schedule = prepare_universe_schedule(sheets)
        df = add_historical_universe(df, schedule)

    except Exception as exc:
        st.error(f"Error al procesar el Excel: {exc}")
        st.stop()

    validation = validate_data(df)

    # ----------------------------------------------------------
    # ESTADO DE ACTUALIZACIÓN
    # ----------------------------------------------------------

    st.sidebar.divider()
    st.sidebar.subheader("🔄 Estado de datos")

    latest_date = df["Fecha"].max()
    latest_rows = df[df["Fecha"] == latest_date]

    latest_draw = (
        latest_rows["Sorteo"].max()
        if not latest_rows.empty else None
    )

    st.sidebar.success(
        f"Datos cargados hasta:\n\n**{latest_date.strftime('%d/%m/%Y')}**"
    )

    if latest_draw is not None:
        st.sidebar.metric("Último sorteo", int(latest_draw))

    st.sidebar.caption(f"Archivo: {source_name}")
    st.sidebar.caption(f"Registros: {len(df):,}")
    cov = prize_coverage(df)
    st.sidebar.caption(
        f"Reventones confirmados: {cov['reventones_confirmados']} · "
        f"Premios disponibles: {cov['premios']:,}/{cov['total']:,} "
        f"({cov['cobertura_premio_pct']:.1f}%)"
    )

    # Control de frescura frente al último miércoles/domingo conocido.
    hoy = datetime.now(ZoneInfo("America/Lima")).date()
    if latest_date.date() < hoy - timedelta(days=4):
        st.sidebar.warning(
            "El archivo podría estar desactualizado. "
            "Verifica que hayas subido el Excel más reciente."
        )

    # ----------------------------------------------------------
    # MÉTRICAS
    # ----------------------------------------------------------

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Sorteos", f"{len(df):,}")
    c2.metric("Desde", df["Fecha"].min().strftime("%d/%m/%Y"))
    c3.metric("Hasta", df["Fecha"].max().strftime("%d/%m/%Y"))
    c4.metric("Último sorteo", int(latest_draw) if latest_draw else "-")
    c5.metric("Premios con dato", f"{cov['cobertura_premio_pct']:.1f}%")

    # ----------------------------------------------------------
    # MENÚ
    # ----------------------------------------------------------

    menu = st.sidebar.radio(
        "Sección",
        [
            "Inicio",
            "Estadísticas",
            "Generador",
            "Backtesting",
            "Reventones",
            "Diagnóstico",
            "Datos",
        ],
    )

    # ==========================================================
    # INICIO
    # ==========================================================

    if menu == "Inicio":
        st.header("🏠 Resumen")

        st.write(
            "La aplicación está utilizando el contenido completo del Excel "
            "subido en esta sesión."
        )

        if df.attrs.get("rows_dropped", 0):
            st.warning(
                f"Se descartaron {df.attrs['rows_dropped']} filas por "
                "fecha/bolillas incompletas."
            )

        st.subheader("Últimos sorteos")

        cols = ["Fecha", "Sorteo"] + BALL_COLUMNS
        if "Boliyapa" in df.columns:
            cols.append("Boliyapa")
        if "Premio" in df.columns:
            cols.append("Premio")
        cols.append("ReventoVerde")

        st.dataframe(
            df[cols].sort_values(["Fecha", "Sorteo"], ascending=False).head(10),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Universo histórico")
        if not schedule.empty:
            st.dataframe(
                schedule,
                use_container_width=True,
                hide_index=True,
            )

    # ==========================================================
    # ESTADÍSTICAS
    # ==========================================================

    elif menu == "Estadísticas":
        st.header("📊 Estadísticas")

        universe_size = universe_for_date(
            df["Fecha"].max(), schedule
        )
        universe = list(range(1, universe_size + 1))

        freq = frequency_table(df, universe)
        st.dataframe(freq, use_container_width=True, hide_index=True)

        st.subheader("Frecuencia")
        st.bar_chart(freq.set_index("Bolilla")["Apariciones"])

        st.subheader("Estructura histórica")
        stats = structure_statistics(df, universe_size)

        cols = st.columns(4)
        labels = {
            "odd": "Impares",
            "consecutive": "Pares consecutivos",
            "low": "Números bajos",
            "sum": "Suma",
        }

        for col, key in zip(cols, stats):
            mean, std = stats[key]
            col.metric(labels[key], f"{mean:.2f}", f"DE {std:.2f}")

    # ==========================================================
    # GENERADOR
    # ==========================================================

    elif menu == "Generador":
        st.header("🎯 Generador de combinaciones")

        current_universe_size = universe_for_date(
            df["Fecha"].max(), schedule
        )
        current_universe = list(range(1, current_universe_size + 1))

        st.write(
            f"Universo vigente según la hoja **Modificaciones**: "
            f"**1–{current_universe_size}**."
        )

        c1, c2, c3 = st.columns(3)

        n_candidates = c1.number_input(
            "Candidatos",
            min_value=100,
            max_value=50000,
            value=5000,
            step=100,
        )

        top_n = c2.number_input(
            "Combinaciones finales",
            min_value=1,
            max_value=20,
            value=3,
        )

        diversity = c3.checkbox(
            "Diversificar",
            value=True,
        )

        st.subheader("Filtros")

        c1, c2, c3 = st.columns(3)

        use_sum = c1.checkbox("Limitar suma", value=False)
        use_consecutive = c2.checkbox(
            "Limitar consecutivos",
            value=False,
        )
        exclude_hist = c3.checkbox(
            "Excluir combinaciones históricas",
            value=False,
            help="Por independencia matemática, una combinación ya sorteada sigue siendo posible.",
        )

        use_reventon_gen = st.checkbox(
            "Usar contexto histórico de reventones",
            value=False,
            help=(
                "Ajusta ligeramente las frecuencias según si el sorteo anterior "
                "reventó. Es experimental y descriptivo; no implica causalidad."
            ),
        )

        min_sum = None
        max_sum = None
        max_consecutive = None

        if use_sum:
            c1, c2 = st.columns(2)
            min_sum = c1.number_input("Suma mínima", 0, 600, 90)
            max_sum = c2.number_input("Suma máxima", 0, 600, 195)

        if use_consecutive:
            max_consecutive = st.number_input(
                "Máximo de pares consecutivos",
                0,
                5,
                1,
            )

        seed_mode = st.radio(
            "Semilla",
            ["Automática", "Manual"],
            horizontal=True,
        )

        if seed_mode == "Manual":
            seed = int(st.number_input("Seed", value=12345, step=1))
        else:
            target_date, target_day = proximo_sorteo()
            seed = stable_seed(
                "generador",
                df["Fecha"].max(),
                latest_draw,
                target_date,
                target_day,
            )

        if st.button("🚀 Generar combinaciones", type="primary"):
            candidates, info = generate_candidates(
                train=df,
                universe=current_universe,
                n_candidates=int(n_candidates),
                seed=seed,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
                exclude_historical=exclude_hist,
                use_reventon_context=use_reventon_gen,
                current_reventon_context=(
                    safe_bool(df.iloc[-1]["ReventoVerde"])
                    if use_reventon_gen
                    else False
                ),
            )

            selected = select_top_combinations(
                candidates,
                top_n=int(top_n),
                diversity=diversity,
            )

            if selected.empty:
                st.error(
                    "No se encontraron combinaciones. "
                    "Reduce las restricciones o aumenta los candidatos."
                )
            else:
                st.success(
                    f"Se generaron {info['generated']:,} candidatos válidos "
                    f"en {info['attempts']:,} intentos "
                    f"(aceptación {info['acceptance_rate']*100:.2f}%)."
                )

                st.subheader("Combinaciones seleccionadas")

                export_df = candidates_to_dataframe(selected)

                for i, row in selected.iterrows():
                    combo = tuple(row["Combinación"])
                    st.markdown(
                        f"### #{i+1}  "
                        f"**{' - '.join(f'{n:02d}' for n in combo)}**"
                    )
                    st.caption(f"Score: {row['Score']:.5f}")

                st.dataframe(
                    export_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "⬇️ Descargar CSV",
                    export_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name="combinaciones_la_tinka.csv",
                    mime="text/csv",
                )

    # ==========================================================
    # BACKTESTING
    # ==========================================================

    elif menu == "Backtesting":
        st.header("🧪 Backtesting")

        st.write(
            "El sistema toma cada sorteo histórico como si fuera futuro: "
            "entrena únicamente con sorteos anteriores y luego compara "
            "las combinaciones generadas con el resultado real."
        )

        c1, c2, c3, c4 = st.columns(4)

        min_train = c1.number_input(
            "Sorteos mínimos de entrenamiento",
            50,
            2000,
            200,
            10,
        )

        step = c2.number_input(
            "Paso entre pruebas",
            1,
            50,
            1,
            1,
            help="1 prueba cada sorteo. Un valor mayor acelera, pero evalúa menos sorteos.",
        )

        top_n_bt = c3.number_input(
            "Tickets por prueba",
            1,
            50,
            10,
            1,
        )

        n_candidates_bt = c4.number_input(
            "Candidatos por prueba",
            200,
            10000,
            1500,
            100,
        )

        st.subheader("Restricciones de la prueba")

        c1, c2, c3 = st.columns(3)

        use_sum_bt = c1.checkbox("Usar rango de suma", value=False)
        use_cons_bt = c2.checkbox("Usar límite consecutivos", value=False)
        exclude_hist_bt = c3.checkbox(
            "Excluir combinaciones históricas",
            value=False,
        )

        min_sum_bt = max_sum_bt = max_cons_bt = None

        if use_sum_bt:
            c1, c2 = st.columns(2)
            min_sum_bt = c1.number_input("Suma mínima BT", 0, 600, 90)
            max_sum_bt = c2.number_input("Suma máxima BT", 0, 600, 195)

        if use_cons_bt:
            max_cons_bt = st.number_input(
                "Máximo consecutivos BT",
                0,
                5,
                1,
            )

        use_reventon = st.checkbox(
            "Registrar contexto de reventón anterior",
            value=True,
            help=(
                "Solo usa información de sorteos anteriores al objetivo. "
                "No utiliza el reventón del sorteo que se está evaluando."
            ),
        )

        if st.button("▶️ Ejecutar backtesting", type="primary"):
            strategies = [
                "Aleatorio",
                "Modelo",
                "Frecuencia",
                "Frecuencia + Recencia",
            ]

            all_results = []

            progress = st.progress(0)

            for idx, strategy in enumerate(strategies):
                result = backtest_strategy(
                    df=df,
                    schedule=schedule,
                    min_train=int(min_train),
                    step=int(step),
                    top_n=int(top_n_bt),
                    n_candidates=int(n_candidates_bt),
                    min_sum=min_sum_bt,
                    max_sum=max_sum_bt,
                    max_consecutive=max_cons_bt,
                    strategy=strategy,
                    exclude_historical=exclude_hist_bt,
                    use_reventon_feature=use_reventon,
                )
                all_results.append(result)
                progress.progress((idx + 1) / len(strategies))

            results = pd.concat(all_results, ignore_index=True)
            summary = summarize_backtest(results)

            st.subheader("Resumen")
            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True,
            )

            st.subheader("Resultados por sorteo")
            st.dataframe(
                results.sort_values(["Fecha", "Estrategia"]),
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "⬇️ Descargar resultados del backtesting",
                results.to_csv(index=False).encode("utf-8-sig"),
                file_name="backtesting_la_tinka.csv",
                mime="text/csv",
            )

            st.info(
                "Las métricas son descriptivas. La comparación no demuestra "
                "que una estrategia tenga mayor probabilidad futura; sirve "
                "para comprobar cómo se comportó históricamente."
            )

    # ==========================================================
    # REVENTONES
    # ==========================================================

    elif menu == "Reventones":
        st.header("💰 Reventones y premios")

        green_df = df[df["ReventoVerde"]].copy()

        st.metric(
            "Sorteos marcados como reventados",
            len(green_df),
        )

        if green_df.empty:
            st.warning(
                "No se detectaron resaltados verdes en la columna Premio."
            )
        else:
            show_cols = [
                "Fecha",
                "Sorteo",
                "Premio",
                "SorteosSinReventarAntes",
                "DiasDesdeUltimoReventonAntes",
            ]

            st.dataframe(
                green_df[show_cols].sort_values(
                    ["Fecha", "Sorteo"], ascending=False
                ),
                use_container_width=True,
                hide_index=True,
            )

            st.subheader("Intervalos entre reventones")

            dates = green_df["Fecha"].sort_values().tolist()
            if len(dates) >= 2:
                intervals = [
                    (dates[i] - dates[i - 1]).days
                    for i in range(1, len(dates))
                ]

                c1, c2, c3 = st.columns(3)
                c1.metric("Promedio días", f"{np.mean(intervals):.1f}")
                c2.metric("Mediana días", f"{np.median(intervals):.1f}")
                c3.metric("Máximo días", f"{max(intervals)}")

                st.bar_chart(pd.Series(intervals, name="Días"))

            st.subheader("Premios")
            prize = green_df[["Fecha", "Sorteo", "Premio"]].dropna()
            if not prize.empty:
                st.line_chart(
                    prize.sort_values("Fecha").set_index("Fecha")["Premio"]
                )

            st.info(
                "El resaltado verde se interpreta como la marca histórica "
                "de 'reventó'. Para el modelo y el backtesting, esta marca "
                "solo puede utilizarse como información disponible antes "
                "del siguiente sorteo."
            )

    # ==========================================================
    # DIAGNÓSTICO
    # ==========================================================

    elif menu == "Diagnóstico":
        st.header("🔎 Diagnóstico")

        st.write(
            "Estos controles permiten detectar problemas del archivo antes "
            "de utilizarlo para generar combinaciones."
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("Filas válidas", validation["records"])
        c2.metric("Duplicados de sorteo", validation["duplicate_draws"])
        c3.metric("Filas repetidas", validation["exact_duplicates"])
        c4.metric("Bolillas repetidas en fila", validation["repeated_balls"])

        st.subheader("Rango observado")
        st.write(
            f"Bolillas observadas en el histórico: "
            f"{validation['universe_min']}–{validation['universe_max']}"
        )

        st.subheader("Universo vigente por período")

        if schedule.empty:
            st.warning("No se pudo leer la hoja Modificaciones.")
        else:
            st.dataframe(
                schedule,
                use_container_width=True,
                hide_index=True,
            )

        st.subheader("Reventones detectados")
        st.write(
            f"Se detectaron **{validation['reventones']}** celdas de Premio "
            "con el resaltado verde."
        )

        st.subheader("Último registro")
        st.dataframe(
            df.tail(1),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================
    # DATOS
    # ==========================================================

    elif menu == "Datos":
        st.header("🗃️ Datos utilizados por el modelo")

        st.caption(
            "Esta tabla muestra los datos ya preparados por la aplicación. "
            "El campo ReventoVerde proviene del resaltado del Excel."
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )


if __name__ == "__main__":
    main()

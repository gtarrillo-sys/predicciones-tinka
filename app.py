
# ==============================================================
# LA TINKA AI - VERSION 4.0
# Análisis histórico, generación y backtesting
# ==============================================================
# Instalación:
# pip install streamlit pandas numpy openpyxl
#
# Ejecución:
# streamlit run app.py
#
# Archivo esperado:
# Resultados la Tinka 1994 a 2026.xlsx
#
# Hojas:
#   Resultados
#   Modificaciones
#
# Las seis bolillas principales son el objetivo de evaluación.
# El sistema es experimental: no garantiza aciertos futuros.
# ==============================================================

from __future__ import annotations

import hashlib
import io
from collections import Counter
from datetime import datetime, timedelta
from itertools import combinations
from pathlib import Path
from typing import Optional

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

APP_VERSION = "4.0 - estrategias independientes y backtesting corregido"

BALL_COLUMNS = [
    "Bolilla 1",
    "Bolilla 2",
    "Bolilla 3",
    "Bolilla 4",
    "Bolilla 5",
    "Bolilla 6",
]

DEFAULT_FILES = [
    "Resultados la Tinka 1994 a 2026.xlsx",
    "Resultados_la_Tinka_1994_2026_revisado.xlsx",
    "La_Tinka_Todos_Los_Sorteos_1994_2026.xlsx",
]

STRATEGIES = [
    "Aleatorio",
    "Modelo",
    "Frecuencia",
    "Frecuencia + Recencia",
]


# ==============================================================
# UTILIDADES GENERALES
# ==============================================================

def normalize_text(value) -> str:
    if value is None:
        return ""

    replacements = str(value).strip().lower()
    for old, new in [
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("ñ", "n"),
        ("_", " "),
    ]:
        replacements = replacements.replace(old, new)

    return replacements


def normalize_ball_column(name) -> str:
    compact = normalize_text(name).replace(" ", "")

    for i in range(1, 7):
        if compact in (f"bolilla{i}", f"bolilla0{i}"):
            return f"Bolilla {i}"

    return str(name).strip()


def stable_seed(*values) -> int:
    raw = "|".join(map(str, values))
    return int(hashlib.sha256(raw.encode("utf-8")).hexdigest()[:8], 16)


def safe_bool(value):
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    return bool(value)


def find_file() -> Optional[Path]:
    for filename in DEFAULT_FILES:
        candidates = [
            Path(filename),
            Path("datos") / filename,
            Path("/mnt/data") / filename,
        ]

        for path in candidates:
            if path.exists():
                return path

    return None


def get_file_bytes(source) -> bytes:
    if isinstance(source, (str, Path)):
        return Path(source).read_bytes()

    return source.getvalue()


@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes):
    excel = pd.ExcelFile(io.BytesIO(file_bytes))

    return {
        sheet: pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet,
        )
        for sheet in excel.sheet_names
    }


def read_excel_fresh(source):
    return load_excel_cached(get_file_bytes(source))


# ==============================================================
# PREPARACIÓN DE RESULTADOS
# ==============================================================

def detect_date_column(df):
    for column in df.columns:
        if normalize_text(column) == "fecha":
            return column

    for column in df.columns:
        if "fecha" in normalize_text(column):
            return column

    return None


def detect_draw_column(df):
    for column in df.columns:
        if normalize_text(column) == "sorteo":
            return column

    for column in df.columns:
        name = normalize_text(column)
        if "sorteo" in name and "extra" not in name:
            return column

    return None


def detect_ball_columns(df):
    found = {}

    for column in df.columns:
        normalized = normalize_ball_column(column)

        if normalized in BALL_COLUMNS:
            found[normalized] = column

    missing = [column for column in BALL_COLUMNS if column not in found]

    if missing:
        raise ValueError(
            "No se encontraron las seis bolillas principales. "
            f"Faltan: {', '.join(missing)}"
        )

    return [found[column] for column in BALL_COLUMNS]


def prepare_results(raw_df):
    df = raw_df.copy()

    date_col = detect_date_column(df)
    draw_col = detect_draw_column(df)
    ball_cols = detect_ball_columns(df)

    if date_col is None:
        raise ValueError("No se encontró la columna Fecha.")

    rename = {date_col: "Fecha"}

    if draw_col is not None:
        rename[draw_col] = "Sorteo"

    for i, column in enumerate(ball_cols, 1):
        rename[column] = f"Bolilla {i}"

    df = df.rename(columns=rename)

    df["Fecha"] = pd.to_datetime(
        df["Fecha"],
        errors="coerce",
        dayfirst=True,
    )

    if "Sorteo" not in df.columns:
        df["Sorteo"] = np.arange(1, len(df) + 1)

    df["Sorteo"] = pd.to_numeric(
        df["Sorteo"],
        errors="coerce",
    )

    for column in BALL_COLUMNS:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    original_rows = len(df)

    # Solo se excluyen filas sin fecha o sin alguna de las seis
    # bolillas principales. Las columnas auxiliares se conservan.
    df = df.dropna(
        subset=["Fecha"] + BALL_COLUMNS
    ).copy()

    rows_dropped = original_rows - len(df)

    for column in BALL_COLUMNS:
        df[column] = df[column].astype(int)

    df["Sorteo"] = df["Sorteo"].astype("Int64")

    if "Premio" in df.columns:
        df["Premio"] = pd.to_numeric(
            df["Premio"],
            errors="coerce",
        )

    df = df.sort_values(
        ["Fecha", "Sorteo"]
    ).reset_index(drop=True)

    df.attrs["rows_dropped"] = rows_dropped

    return df


# ==============================================================
# DETECCIÓN DEL RESALTADO VERDE EN PREMIO
# ==============================================================

def is_green_fill(cell):
    fill = cell.fill

    if not fill or fill.fill_type != "solid":
        return False

    fg = fill.fgColor

    if fg.type == "theme" and fg.theme == 9:
        return True

    if fg.type == "rgb" and fg.rgb:
        rgb = str(fg.rgb)[-6:].upper()

        try:
            red = int(rgb[0:2], 16)
            green = int(rgb[2:4], 16)
            blue = int(rgb[4:6], 16)

            return (
                green > red * 1.15
                and green > blue * 1.05
                and green > 90
            )
        except (ValueError, TypeError):
            return False

    return False


@st.cache_data(show_spinner=False)
def detect_green_prizes(file_bytes: bytes):
    wb = load_workbook(
        io.BytesIO(file_bytes),
        data_only=False,
    )

    if "Resultados" not in wb.sheetnames:
        return pd.DataFrame(
            columns=["Fecha", "Sorteo", "ReventoVerde"]
        )

    ws = wb["Resultados"]

    headers = {
        normalize_text(cell.value): cell.column
        for cell in ws[1]
        if cell.value is not None
    }

    date_col = headers.get("fecha")
    draw_col = headers.get("sorteo")
    prize_col = headers.get("premio")

    if not date_col or not draw_col or not prize_col:
        return pd.DataFrame(
            columns=["Fecha", "Sorteo", "ReventoVerde"]
        )

    records = []

    for row_num in range(2, ws.max_row + 1):
        date_value = ws.cell(row_num, date_col).value
        draw_value = ws.cell(row_num, draw_col).value
        prize_cell = ws.cell(row_num, prize_col)

        if draw_value is None:
            continue

        records.append({
            "Fecha": pd.to_datetime(
                date_value,
                errors="coerce",
                dayfirst=True,
            ),
            "Sorteo": pd.to_numeric(
                draw_value,
                errors="coerce",
            ),
            "ReventoVerde": is_green_fill(prize_cell),
            "PremioCeldaConDato": prize_cell.value is not None,
        })

    return pd.DataFrame(records)


def attach_prize_features(df, file_bytes):
    out = df.copy()

    if "Premio" not in out.columns:
        out["Premio"] = np.nan

    out["Premio"] = pd.to_numeric(
        out["Premio"],
        errors="coerce",
    )

    green = detect_green_prizes(file_bytes)

    out["FechaClave"] = pd.to_datetime(
        out["Fecha"],
        errors="coerce",
    ).dt.normalize()

    out["SorteoClave"] = pd.to_numeric(
        out["Sorteo"],
        errors="coerce",
    ).astype("Int64")

    if not green.empty:
        green["FechaClave"] = pd.to_datetime(
            green["Fecha"],
            errors="coerce",
        ).dt.normalize()

        green["SorteoClave"] = pd.to_numeric(
            green["Sorteo"],
            errors="coerce",
        ).astype("Int64")

        # Vinculación por ambas claves para evitar que números de
        # sorteo repetidos o inconsistentes mezclen los registros.
        green = green.drop_duplicates(
            ["FechaClave", "SorteoClave"],
            keep="last",
        )

        green = green[
            [
                "FechaClave",
                "SorteoClave",
                "ReventoVerde",
                "PremioCeldaConDato",
            ]
        ]

        out = out.merge(
            green,
            on=["FechaClave", "SorteoClave"],
            how="left",
            validate="many_to_one",
        )

        # El resaltado verde confirma el reventón.
        # Una celda sin dato de premio permanece desconocida.
        no_mark_with_prize = (
            out["ReventoVerde"].isna()
            & out["Premio"].notna()
        )

        out.loc[no_mark_with_prize, "ReventoVerde"] = False

    else:
        out["ReventoVerde"] = pd.NA

    out["ReventoVerde"] = out["ReventoVerde"].astype("boolean")

    out["ReventoConocido"] = out["ReventoVerde"].notna()
    out["PremioAnterior"] = out["Premio"].shift(1)

    out["PremioCambioPct"] = (
        (out["Premio"] - out["PremioAnterior"])
        / out["PremioAnterior"].replace(0, np.nan)
    )

    out["ReventoAnterior"] = (
        out["ReventoVerde"].shift(1).astype("boolean")
    )

    # Número de sorteos conocidos consecutivos sin reventón.
    # Si aparece un dato desconocido, la racha deja de estimarse.
    streak = []
    count = 0
    known = True

    for value in out["ReventoVerde"].tolist():
        if known:
            streak.append(count)
        else:
            streak.append(np.nan)

        if pd.isna(value):
            known = False
        elif bool(value):
            count = 0
            known = True
        else:
            count += 1
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

        value = row["ReventoVerde"]

        if pd.notna(value) and bool(value):
            last_revent_date = row["Fecha"]

    out["DiasDesdeUltimoReventonAntes"] = days_since
    out["PremioDisponible"] = out["Premio"].notna()

    return out.drop(
        columns=["FechaClave", "SorteoClave", "PremioCeldaConDato"],
        errors="ignore",
    )


# ==============================================================
# UNIVERSO HISTÓRICO: HOJA MODIFICACIONES
# ==============================================================

def prepare_universe_schedule(sheets):
    if "Modificaciones" not in sheets:
        return pd.DataFrame()

    mod = sheets["Modificaciones"].copy()

    date_col = next(
        (
            c for c in mod.columns
            if normalize_text(c) == "fecha inicio"
        ),
        None,
    )

    total_col = next(
        (
            c for c in mod.columns
            if "total bolillas" in normalize_text(c)
        ),
        None,
    )

    if date_col is None or total_col is None:
        return pd.DataFrame()

    mod["Fecha Inicio"] = pd.to_datetime(
        mod[date_col],
        errors="coerce",
        dayfirst=True,
    )

    mod["Total Bolillas"] = pd.to_numeric(
        mod[total_col],
        errors="coerce",
    )

    mod = mod.dropna(
        subset=["Fecha Inicio", "Total Bolillas"]
    ).copy()

    mod["Total Bolillas"] = mod["Total Bolillas"].astype(int)

    mod = mod.sort_values(
        "Fecha Inicio"
    ).reset_index(drop=True)

    return mod


def universe_for_date(target_date, schedule, fallback=53):
    if schedule.empty:
        return fallback

    target = pd.Timestamp(target_date).normalize()

    valid = schedule[
        schedule["Fecha Inicio"].dt.normalize() <= target
    ]

    if valid.empty:
        return int(schedule.iloc[0]["Total Bolillas"])

    return int(valid.iloc[-1]["Total Bolillas"])


def add_historical_universe(df, schedule):
    out = df.copy()

    out["UniversoVigente"] = [
        universe_for_date(value, schedule)
        for value in out["Fecha"]
    ]

    return out


# ==============================================================
# VALIDACIÓN Y RESUMEN DE DATOS
# ==============================================================

def prize_coverage(df):
    total = len(df)

    prizes = (
        int(df["Premio"].notna().sum())
        if "Premio" in df else 0
    )

    if "ReventoVerde" in df:
        known = int(df["ReventoVerde"].notna().sum())
        confirmed = int(
            df["ReventoVerde"].fillna(False).astype(bool).sum()
        )
    else:
        known = 0
        confirmed = 0

    return {
        "total": total,
        "premios": prizes,
        "cobertura_premio_pct": prizes / total * 100 if total else 0,
        "reventones_confirmados": confirmed,
        "reventones_conocidos": known,
        "cobertura_reventon_pct": known / total * 100 if total else 0,
    }


def validate_data(df):
    repeated_balls = 0
    out_of_range = 0

    for _, row in df.iterrows():
        values = [int(row[c]) for c in BALL_COLUMNS]

        if len(set(values)) != 6:
            repeated_balls += 1

        universe = int(row.get("UniversoVigente", 53))

        if any(n < 1 or n > universe for n in values):
            out_of_range += 1

    combinations_seen = {
        tuple(sorted(int(row[c]) for c in BALL_COLUMNS))
        for _, row in df.iterrows()
    }

    observed = [
        int(n)
        for c in BALL_COLUMNS
        for n in df[c].dropna().tolist()
    ]

    coverage = prize_coverage(df)

    return {
        "records": len(df),
        "min_date": df["Fecha"].min(),
        "max_date": df["Fecha"].max(),
        "duplicate_draws": int(df["Sorteo"].duplicated(keep=False).sum()),
        "exact_duplicates": int(df.duplicated().sum()),
        "repeated_balls": repeated_balls,
        "out_of_range": out_of_range,
        "unique_combinations": len(combinations_seen),
        "universe_min": min(observed) if observed else None,
        "universe_max": max(observed) if observed else None,
        "reventones": coverage["reventones_confirmados"],
        "premio_cobertura_pct": coverage["cobertura_premio_pct"],
        "reventon_cobertura_pct": coverage["cobertura_reventon_pct"],
    }


# ==============================================================
# FRECUENCIAS Y ATRASOS
# ==============================================================

def calculate_frequency(df, decay_lambda=0.0):
    freq = Counter()

    if df.empty:
        return freq

    max_date = df["Fecha"].max()

    for _, row in df.iterrows():
        days = (max_date - row["Fecha"]).days

        weight = (
            np.exp(-decay_lambda * days)
            if decay_lambda > 0 else 1.0
        )

        for column in BALL_COLUMNS:
            freq[int(row[column])] += float(weight)

    return freq


def minmax_dict(values, neutral=0.5):
    finite = [
        float(v)
        for v in values.values()
        if pd.notna(v) and np.isfinite(v)
    ]

    if not finite:
        return {key: neutral for key in values}

    lo, hi = min(finite), max(finite)

    if hi == lo:
        return {key: neutral for key in values}

    return {
        key: (
            neutral
            if pd.isna(value)
            else (float(value) - lo) / (hi - lo)
        )
        for key, value in values.items()
    }


def calculate_ball_scores(
    df,
    universe,
    historical_weight=0.65,
    recent_weight=0.35,
):
    historical = calculate_frequency(df, 0)
    recent = calculate_frequency(df, 0.003)

    hist_norm = minmax_dict({
        n: historical[n] for n in universe
    })

    recent_norm = minmax_dict({
        n: recent[n] for n in universe
    })

    return {
        n: (
            historical_weight * hist_norm[n]
            + recent_weight * recent_norm[n]
        )
        for n in universe
    }


def calculate_delay(df, universe):
    result = {}

    if df.empty:
        return {n: np.nan for n in universe}

    last_date = df["Fecha"].max()

    for n in universe:
        dates = df.loc[
            (df[BALL_COLUMNS] == n).any(axis=1),
            "Fecha",
        ]

        result[n] = (
            float((last_date - dates.max()).days)
            if not dates.empty else np.nan
        )

    return result


def frequency_table(df, universe):
    freq = calculate_frequency(df)
    delay = calculate_delay(df, universe)
    total = len(df) * 6

    rows = []

    for n in universe:
        rows.append({
            "Bolilla": n,
            "Apariciones": int(freq[n]),
            "% sobre bolillas": (
                freq[n] / total * 100 if total else 0
            ),
            "Atraso días": delay[n],
        })

    return pd.DataFrame(rows).sort_values(
        ["Apariciones", "Bolilla"],
        ascending=[False, True],
    ).reset_index(drop=True)


# ==============================================================
# PARES Y ESTRUCTURA DE COMBINACIONES
# ==============================================================

def calculate_pair_frequency(df):
    pairs = Counter()

    for _, row in df.iterrows():
        nums = sorted(int(row[c]) for c in BALL_COLUMNS)

        for pair in combinations(nums, 2):
            pairs[pair] += 1

    return pairs


def pair_score(combo, pairs):
    values = [
        np.log1p(pairs[pair])
        for pair in combinations(sorted(combo), 2)
    ]

    return float(np.mean(values)) if values else 0.0


def structure_distribution(combo, universe_size):
    nums = sorted(combo)

    return {
        "odd": sum(n % 2 for n in nums),
        "consecutive": sum(
            1 for a, b in zip(nums, nums[1:])
            if b == a + 1
        ),
        "low": sum(n <= universe_size // 2 for n in nums),
        "sum": sum(nums),
    }


def structure_statistics(df, universe_size):
    values = {
        "odd": [],
        "consecutive": [],
        "low": [],
        "sum": [],
    }

    for _, row in df.iterrows():
        combo = tuple(int(row[c]) for c in BALL_COLUMNS)

        local_universe = int(
            row.get("UniversoVigente", universe_size)
        )

        result = structure_distribution(combo, local_universe)

        for key in values:
            values[key].append(result[key])

    stats = {}

    for key, arr in values.items():
        if not arr:
            stats[key] = (0.0, 1.0)
        else:
            std = float(np.std(arr))
            stats[key] = (
                float(np.mean(arr)),
                std if std > 0 else 1.0,
            )

    return stats


def structure_score(combo, stats, universe_size):
    values = structure_distribution(combo, universe_size)
    z_scores = []

    for key in ["odd", "consecutive", "low", "sum"]:
        mean, std = stats[key]
        z_scores.append(
            abs(values[key] - mean) / std if std else 0.0
        )

    return float(np.exp(-0.5 * np.mean(z_scores)))


# ==============================================================
# VALIDACIÓN DE COMBINACIONES
# ==============================================================

def validate_combination(
    combo,
    universe,
    min_sum=None,
    max_sum=None,
    max_consecutive=None,
    exclude_historical=False,
    historical_combinations=None,
):
    combo = tuple(sorted(combo))

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
        1 for a, b in zip(combo, combo[1:])
        if b == a + 1
    )

    if (
        max_consecutive is not None
        and consecutive > max_consecutive
    ):
        return False

    if (
        exclude_historical
        and historical_combinations is not None
        and combo in historical_combinations
    ):
        return False

    return True


def weighted_sample_without_replacement(
    universe,
    scores,
    rng,
):
    available = list(universe)
    chosen = []

    for _ in range(6):
        weights = np.array([
            max(float(scores.get(n, 0.0)), 1e-9)
            for n in available
        ])

        weights /= weights.sum()

        index = int(rng.choice(len(available), p=weights))
        chosen.append(available.pop(index))

    return tuple(sorted(chosen))


# ==============================================================
# PUNTUACIÓN DE COMBINACIONES POR ESTRATEGIA
# ==============================================================

def combination_score(
    combo,
    strategy,
    ball_scores,
    pairs,
    stats,
    universe_size,
):
    ball = float(np.mean([
        ball_scores.get(n, 0.5)
        for n in combo
    ]))

    raw_pair = pair_score(combo, pairs)
    pair_scale = max(
        1.0,
        np.log1p(max(pairs.values(), default=1)),
    )

    pair = min(1.0, raw_pair / pair_scale)

    structural = structure_score(
        combo,
        stats,
        universe_size,
    )

    if strategy == "Frecuencia":
        return ball

    if strategy == "Frecuencia + Recencia":
        return ball

    # Modelo combina puntuación individual, pares y estructura.
    return 0.60 * ball + 0.25 * pair + 0.15 * structural


# ==============================================================
# GENERACIÓN DE CANDIDATOS
# ==============================================================

def generate_candidates(
    train,
    universe,
    n_candidates=5000,
    seed=12345,
    min_sum=None,
    max_sum=None,
    max_consecutive=None,
    exclude_historical=False,
    strategy="Modelo",
):
    rng = np.random.default_rng(seed)

    if strategy == "Frecuencia":
        ball_scores = calculate_ball_scores(
            train, universe, 1.0, 0.0
        )
    else:
        # Modelo y Frecuencia + Recencia usan puntuaciones
        # independientes de la estrategia correspondiente.
        ball_scores = calculate_ball_scores(
            train, universe, 0.65, 0.35
        )

    pairs = calculate_pair_frequency(train)

    universe_size = max(universe)
    stats = structure_statistics(train, universe_size)

    historical_combinations = {
        tuple(sorted(int(row[c]) for c in BALL_COLUMNS))
        for _, row in train.iterrows()
    }

    candidates = {}
    attempts = 0
    max_attempts = max(1000, int(n_candidates) * 40)

    while (
        len(candidates) < int(n_candidates)
        and attempts < max_attempts
    ):
        attempts += 1

        if strategy == "Aleatorio":
            combo = tuple(sorted(
                rng.choice(
                    universe,
                    size=6,
                    replace=False,
                ).tolist()
            ))
        else:
            combo = weighted_sample_without_replacement(
                universe,
                ball_scores,
                rng,
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

        if strategy == "Aleatorio":
            score = float(rng.random())
        else:
            score = combination_score(
                combo,
                strategy,
                ball_scores,
                pairs,
                stats,
                universe_size,
            )

        # No duplicar candidatos.
        if combo not in candidates:
            candidates[combo] = score

    result = pd.DataFrame([
        {"Combinación": combo, "Score": score}
        for combo, score in candidates.items()
    ])

    if not result.empty:
        result = result.sort_values(
            "Score",
            ascending=False,
        ).reset_index(drop=True)

    info = {
        "attempts": attempts,
        "generated": len(result),
        "acceptance_rate": (
            len(result) / attempts if attempts else 0
        ),
        "requested": int(n_candidates),
    }

    return result, info


# ==============================================================
# SELECCIÓN DIVERSIFICADA
# ==============================================================

def select_top_combinations(
    candidates,
    top_n=3,
    diversity=True,
    max_overlap=3,
):
    if candidates.empty:
        return candidates.copy()

    selected = []

    for _, row in candidates.iterrows():
        combo = tuple(row["Combinación"])

        if diversity:
            too_similar = any(
                len(set(combo) & set(prev["Combinación"]))
                > max_overlap
                for prev in selected
            )

            if too_similar:
                continue

        selected.append(row)

        if len(selected) >= int(top_n):
            break

    return pd.DataFrame(selected).reset_index(drop=True)


# ==============================================================
# REFERENCIA ALEATORIA
# ==============================================================

def random_combinations(
    universe,
    n,
    rng,
    min_sum=None,
    max_sum=None,
    max_consecutive=None,
    exclude_historical=False,
    historical_combinations=None,
    diversity=True,
    max_overlap=3,
):
    results = []
    seen = set()
    attempts = 0
    max_attempts = max(1000, int(n) * 100)

    while len(results) < int(n) and attempts < max_attempts:
        attempts += 1

        combo = tuple(sorted(
            rng.choice(
                universe,
                size=6,
                replace=False,
            ).tolist()
        ))

        if combo in seen:
            continue

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

        if diversity and any(
            len(set(combo) & set(previous)) > max_overlap
            for previous in results
        ):
            continue

        seen.add(combo)
        results.append(combo)

    return results


# ==============================================================
# BACKTESTING WALK-FORWARD
# ==============================================================

def evaluate_combination(combo, actual):
    return len(set(combo) & set(actual))


def evaluate_ticket_set(tickets, actual):
    hits = [
        evaluate_combination(combo, actual)
        for combo in tickets
    ]

    return {
        "max_hits": max(hits) if hits else 0,
        "mean_hits": float(np.mean(hits)) if hits else 0.0,
        "hits_2_plus": sum(h >= 2 for h in hits),
        "hits_3_plus": sum(h >= 3 for h in hits),
        "hits_4_plus": sum(h >= 4 for h in hits),
    }


def backtest_strategy(
    df,
    schedule,
    min_train=200,
    step=1,
    top_n=10,
    n_candidates=1000,
    min_sum=None,
    max_sum=None,
    max_consecutive=None,
    strategy="Modelo",
    exclude_historical=False,
    diversity=True,
    max_overlap=3,
    progress_callback=None,
    progress_offset=0,
    progress_total=1,
):
    data = df.sort_values(
        ["Fecha", "Sorteo"]
    ).reset_index(drop=True)

    min_train = max(1, int(min_train))
    step = max(1, int(step))

    test_indices = list(range(min_train, len(data), step))
    records = []

    for local_idx, test_index in enumerate(test_indices, 1):
        # Solo se utiliza información anterior al sorteo objetivo.
        train = data.iloc[:test_index].copy()
        actual_row = data.iloc[test_index]

        actual = tuple(
            int(actual_row[c])
            for c in BALL_COLUMNS
        )

        universe_size = universe_for_date(
            actual_row["Fecha"],
            schedule,
        )

        universe = list(range(1, universe_size + 1))

        seed = stable_seed(
            "backtest",
            strategy,
            str(actual_row["Fecha"]),
            int(test_index),
        )

        historical_combinations = {
            tuple(sorted(int(row[c]) for c in BALL_COLUMNS))
            for _, row in train.iterrows()
        }

        if strategy == "Aleatorio":
            rng = np.random.default_rng(seed)

            tickets = random_combinations(
                universe,
                top_n,
                rng,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
                exclude_historical=exclude_historical,
                historical_combinations=historical_combinations,
                diversity=diversity,
                max_overlap=max_overlap,
            )
        else:
            candidates, _ = generate_candidates(
                train=train,
                universe=universe,
                n_candidates=n_candidates,
                seed=seed,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
                exclude_historical=exclude_historical,
                strategy=strategy,
            )

            selected = select_top_combinations(
                candidates,
                top_n=top_n,
                diversity=diversity,
                max_overlap=max_overlap,
            )

            tickets = (
                [tuple(combo) for combo in selected["Combinación"]]
                if not selected.empty else []
            )

        metrics = evaluate_ticket_set(tickets, actual)

        records.append({
            "Fecha": actual_row["Fecha"],
            "Sorteo": actual_row["Sorteo"],
            "Estrategia": strategy,
            "Tickets": len(tickets),
            **metrics,
        })

        if progress_callback is not None:
            progress_callback(
                progress_offset + local_idx,
                progress_total,
                strategy,
            )

    return pd.DataFrame(records)


def summarize_backtest(results):
    if results.empty:
        return pd.DataFrame()

    summary = []

    for strategy, group in results.groupby("Estrategia"):
        summary.append({
            "Estrategia": strategy,
            "Pruebas": len(group),
            "Promedio máximo aciertos": group["max_hits"].mean(),
            "Mediana máximo aciertos": group["max_hits"].median(),
            "% pruebas con ≥2": (
                group["max_hits"].ge(2).mean() * 100
            ),
            "% pruebas con ≥3": (
                group["max_hits"].ge(3).mean() * 100
            ),
            "% pruebas con ≥4": (
                group["max_hits"].ge(4).mean() * 100
            ),
            "Máximo observado": group["max_hits"].max(),
            "Promedio de tickets": group["Tickets"].mean(),
        })

    return pd.DataFrame(summary).sort_values(
        "Promedio máximo aciertos",
        ascending=False,
    ).reset_index(drop=True)


# ==============================================================
# EXPORTACIÓN
# ==============================================================

def candidates_to_dataframe(candidates):
    if candidates.empty:
        return pd.DataFrame()

    rows = []

    for rank, (_, row) in enumerate(candidates.iterrows(), 1):
        combo = tuple(row["Combinación"])

        item = {
            "Ranking": rank,
            "Combinación": "-".join(f"{n:02d}" for n in combo),
            "Score": row["Score"],
        }

        for i, number in enumerate(combo, 1):
            item[f"Bolilla {i}"] = number

        rows.append(item)

    return pd.DataFrame(rows)


# ==============================================================
# PROGRAMACIÓN DE SORTEOS
# ==============================================================

def proximo_sorteo():
    now = datetime.now(ZoneInfo("America/Lima"))
    target = now.date()

    # Calendario utilizado por esta aplicación.
    draw_weekdays = {2: "miércoles", 6: "domingo"}

    while target.weekday() not in draw_weekdays:
        target += timedelta(days=1)

    return target, draw_weekdays[target.weekday()]


# ==============================================================
# INTERFAZ PRINCIPAL
# ==============================================================

def main():
    st.title("🎱 LA TINKA AI")

    st.caption(
        "Análisis histórico y generación de combinaciones · "
        f"{APP_VERSION}"
    )

    st.info(
        "La aplicación analiza datos históricos y compara estrategias "
        "con una referencia aleatoria. Ninguna frecuencia, puntuación "
        "o combinación garantiza un resultado futuro."
    )

    # ----------------------------------------------------------
    # ARCHIVO
    # ----------------------------------------------------------

    st.sidebar.header("📁 Datos")

    if st.sidebar.button(
        "🔄 Limpiar caché y recargar",
        use_container_width=True,
    ):
        st.cache_data.clear()
        st.rerun()

    uploaded = st.sidebar.file_uploader(
        "Cargar Excel actualizado",
        type=["xlsx", "xlsm"],
    )

    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        source_name = uploaded.name
    else:
        default_path = find_file()

        if default_path is None:
            st.error(
                "No se encontró el Excel. Cárgalo desde la barra lateral."
            )
            st.stop()

        file_bytes = default_path.read_bytes()
        source_name = default_path.name

    file_signature = hashlib.sha256(file_bytes).hexdigest()

    try:
        sheets = load_excel_cached(file_bytes)
    except Exception as exc:
        st.error(f"No se pudo leer el archivo Excel: {exc}")
        st.stop()

    if "Resultados" not in sheets:
        st.error(
            "El archivo debe contener una hoja llamada 'Resultados'."
        )
        st.stop()

    try:
        df = prepare_results(sheets["Resultados"])
        df = attach_prize_features(df, file_bytes)

        schedule = prepare_universe_schedule(sheets)
        df = add_historical_universe(df, schedule)

    except Exception as exc:
        st.error(f"Error al preparar los datos: {exc}")
        st.stop()

    if df.empty:
        st.error("No existen sorteos válidos para analizar.")
        st.stop()

    validation = validate_data(df)
    coverage = prize_coverage(df)

    latest_date = df["Fecha"].max()
    latest_rows = df[df["Fecha"] == latest_date]
    latest_draw = latest_rows["Sorteo"].max()

    st.sidebar.divider()
    st.sidebar.subheader("🔄 Estado de los datos")

    st.sidebar.success(
        f"Actualizado hasta: {latest_date.strftime('%d/%m/%Y')}"
    )

    st.sidebar.metric(
        "Último sorteo",
        int(latest_draw) if pd.notna(latest_draw) else "-",
    )

    st.sidebar.caption(f"Archivo: {source_name}")
    st.sidebar.caption(f"Registros válidos: {len(df):,}")
    st.sidebar.caption(
        f"Premios disponibles: {coverage['premios']:,}/"
        f"{coverage['total']:,} "
        f"({coverage['cobertura_premio_pct']:.1f}%)"
    )

    today = datetime.now(
        ZoneInfo("America/Lima")
    ).date()

    if latest_date.date() < today - timedelta(days=4):
        st.sidebar.warning(
            "El archivo podría estar desactualizado. "
            "Verifica la fecha del último sorteo."
        )

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Sorteos", f"{len(df):,}")
    c2.metric("Desde", df["Fecha"].min().strftime("%d/%m/%Y"))
    c3.metric("Hasta", latest_date.strftime("%d/%m/%Y"))
    c4.metric(
        "Último sorteo",
        int(latest_draw) if pd.notna(latest_draw) else "-",
    )
    c5.metric(
        "Premios con dato",
        f"{coverage['cobertura_premio_pct']:.1f}%",
    )

    # ----------------------------------------------------------
    # MENÚ PRINCIPAL
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
    # 1. INICIO
    # ==========================================================

    if menu == "Inicio":
        st.header("🏠 Resumen")

        st.write(
            "La aplicación está utilizando los resultados válidos "
            "del Excel cargado en esta sesión."
        )

        if df.attrs.get("rows_dropped", 0):
            st.warning(
                f"Se omitieron {df.attrs['rows_dropped']} filas "
                "sin fecha o sin las seis bolillas principales."
            )

        st.subheader("Últimos sorteos")

        display_cols = [
            "Fecha",
            "Sorteo",
            *BALL_COLUMNS,
        ]

        for col in ["Boliyapa", "Sí o Sí", "Premio", "ReventoVerde"]:
            if col in df.columns:
                display_cols.append(col)

        st.dataframe(
            df[display_cols].sort_values(
                ["Fecha", "Sorteo"],
                ascending=False,
            ).head(10),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Universo histórico")

        if schedule.empty:
            st.warning(
                "No se pudo leer la hoja Modificaciones. "
                "Se utilizará el universo de respaldo de 53 bolillas."
            )
        else:
            st.dataframe(
                schedule,
                use_container_width=True,
                hide_index=True,
            )

    # ==========================================================
    # 2. ESTADÍSTICAS
    # ==========================================================

    elif menu == "Estadísticas":
        st.header("📊 Estadísticas")

        universe_size = universe_for_date(
            latest_date,
            schedule,
        )

        universe = list(range(1, universe_size + 1))

        freq = frequency_table(df, universe)

        st.subheader("Frecuencia y atraso")

        st.dataframe(
            freq,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Frecuencia histórica")

        st.bar_chart(
            freq.set_index("Bolilla")["Apariciones"]
        )

        st.subheader("Estructura de los sorteos")

        stats = structure_statistics(df, universe_size)

        labels = {
            "odd": "Impares",
            "consecutive": "Pares consecutivos",
            "low": "Números bajos",
            "sum": "Suma de bolillas",
        }

        cols = st.columns(4)

        for col, key in zip(cols, stats):
            mean, std = stats[key]
            col.metric(
                labels[key],
                f"{mean:.2f}",
                f"DE {std:.2f}",
            )

        st.caption(
            "Las frecuencias y atrasos describen el pasado. "
            "No implican que una bolilla atrasada tenga más "
            "probabilidad matemática de salir."
        )

    # ==========================================================
    # 3. GENERADOR
    # ==========================================================

    elif menu == "Generador":
        st.header("🎯 Generador de combinaciones")

        universe_size = universe_for_date(
            latest_date,
            schedule,
        )

        universe = list(range(1, universe_size + 1))

        st.write(
            f"Universo vigente según Modificaciones: "
            f"**1–{universe_size}**."
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
            step=1,
        )

        diversity = c3.checkbox(
            "Diversificar combinaciones",
            value=True,
        )

        max_overlap = st.number_input(
            "Máximo de bolillas compartidas entre combinaciones",
            min_value=0,
            max_value=5,
            value=3,
            help="Por ejemplo, 3 significa que dos tickets no podrán compartir más de tres bolillas.",
        )

        strategy = st.selectbox(
            "Estrategia de generación",
            [
                "Modelo",
                "Frecuencia",
                "Frecuencia + Recencia",
                "Aleatorio",
            ],
        )

        st.subheader("Filtros opcionales")

        c1, c2, c3 = st.columns(3)

        use_sum = c1.checkbox("Limitar suma", value=False)
        use_consecutive = c2.checkbox(
            "Limitar consecutivos",
            value=False,
        )
        exclude_historical = c3.checkbox(
            "Excluir combinaciones históricas",
            value=False,
            help="Una combinación ya sorteada puede volver a salir; este filtro es solo una preferencia.",
        )

        min_sum = None
        max_sum = None
        max_consecutive = None

        if use_sum:
            c1, c2 = st.columns(2)

            min_sum = c1.number_input(
                "Suma mínima",
                min_value=0,
                max_value=600,
                value=90,
            )

            max_sum = c2.number_input(
                "Suma máxima",
                min_value=0,
                max_value=600,
                value=195,
            )

            if min_sum > max_sum:
                st.error("La suma mínima no puede superar la máxima.")

        if use_consecutive:
            max_consecutive = st.number_input(
                "Máximo de pares consecutivos",
                min_value=0,
                max_value=5,
                value=1,
            )

        seed_mode = st.radio(
            "Semilla aleatoria",
            ["Automática", "Manual"],
            horizontal=True,
        )

        if seed_mode == "Manual":
            seed = int(
                st.number_input(
                    "Seed",
                    value=12345,
                    step=1,
                )
            )
        else:
            target_date, target_day = proximo_sorteo()

            seed = stable_seed(
                "generador",
                latest_date,
                latest_draw,
                target_date,
                target_day,
                strategy,
            )

        if st.button(
            "🚀 Generar combinaciones",
            type="primary",
        ):
            if use_sum and min_sum > max_sum:
                st.stop()

            candidates, info = generate_candidates(
                train=df,
                universe=universe,
                n_candidates=int(n_candidates),
                seed=seed,
                min_sum=min_sum,
                max_sum=max_sum,
                max_consecutive=max_consecutive,
                exclude_historical=exclude_historical,
                strategy=strategy,
            )

            selected = select_top_combinations(
                candidates,
                top_n=int(top_n),
                diversity=diversity,
                max_overlap=int(max_overlap),
            )

            if selected.empty:
                st.error(
                    "No se encontraron combinaciones válidas. "
                    "Reduce los filtros o aumenta los candidatos."
                )
            else:
                st.success(
                    f"Candidatos válidos: {info['generated']:,}. "
                    f"Intentos: {info['attempts']:,}. "
                    f"Tasa de aceptación: "
                    f"{info['acceptance_rate'] * 100:.2f}%."
                )

                st.subheader("Combinaciones seleccionadas")

                for i, row in selected.iterrows():
                    combo = tuple(row["Combinación"])

                    st.markdown(
                        f"### #{i + 1} · "
                        + " - ".join(f"{n:02d}" for n in combo)
                    )

                    st.caption(
                        f"Puntuación interna: {row['Score']:.5f}"
                    )

                export_df = candidates_to_dataframe(selected)

                st.dataframe(
                    export_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "⬇️ Descargar CSV",
                    export_df.to_csv(
                        index=False
                    ).encode("utf-8-sig"),
                    file_name="combinaciones_la_tinka.csv",
                    mime="text/csv",
                )

                st.caption(
                    "La puntuación sirve para ordenar las combinaciones "
                    "según las reglas seleccionadas; no es una probabilidad "
                    "calibrada de ganar."
                )

    # ==========================================================
    # 4. BACKTESTING
    # ==========================================================

    elif menu == "Backtesting":
        st.header("🧪 Backtesting")

        st.write(
            "Cada prueba entrena con sorteos anteriores y evalúa "
            "las combinaciones frente al siguiente resultado histórico."
        )

        mode = st.radio(
            "Modo",
            ["⚡ Rápido", "🔬 Completo"],
            horizontal=True,
        )

        if mode == "⚡ Rápido":
            default_step = 10
            default_candidates = 300

            st.info(
                "Modo rápido: una prueba cada 10 sorteos y "
                "300 candidatos por estrategia."
            )
        else:
            default_step = 1
            default_candidates = 1000

            st.warning(
                "Modo completo: una prueba por sorteo evaluable. "
                "Puede tardar varios minutos."
            )

        c1, c2, c3, c4 = st.columns(4)

        min_train = c1.number_input(
            "Sorteos mínimos de entrenamiento",
            min_value=50,
            max_value=2000,
            value=200,
            step=10,
        )

        step = c2.number_input(
            "Paso entre pruebas",
            min_value=1,
            max_value=100,
            value=default_step,
            step=1,
        )

        top_n_bt = c3.number_input(
            "Tickets por estrategia",
            min_value=1,
            max_value=50,
            value=10,
            step=1,
        )

        n_candidates_bt = c4.number_input(
            "Candidatos por prueba",
            min_value=100,
            max_value=10000,
            value=default_candidates,
            step=100,
        )

        st.subheader("Restricciones")

        c1, c2, c3 = st.columns(3)

        use_sum_bt = c1.checkbox(
            "Usar rango de suma",
            value=False,
        )

        use_cons_bt = c2.checkbox(
            "Limitar consecutivos",
            value=False,
        )

        exclude_hist_bt = c3.checkbox(
            "Excluir combinaciones históricas",
            value=False,
        )

        min_sum_bt = None
        max_sum_bt = None
        max_cons_bt = None

        if use_sum_bt:
            c1, c2 = st.columns(2)

            min_sum_bt = c1.number_input(
                "Suma mínima BT",
                min_value=0,
                max_value=600,
                value=90,
            )

            max_sum_bt = c2.number_input(
                "Suma máxima BT",
                min_value=0,
                max_value=600,
                value=195,
            )

        if use_cons_bt:
            max_cons_bt = st.number_input(
                "Máximo consecutivos BT",
                min_value=0,
                max_value=5,
                value=1,
            )

        diversity_bt = st.checkbox(
            "Aplicar diversificación a todas las estrategias",
            value=True,
        )

        max_overlap_bt = st.number_input(
            "Máximo de bolillas compartidas BT",
            min_value=0,
            max_value=5,
            value=3,
        )

        if st.button(
            "▶️ Ejecutar backtesting",
            type="primary",
        ):
            if use_sum_bt and min_sum_bt > max_sum_bt:
                st.error(
                    "La suma mínima no puede superar la máxima."
                )
                st.stop()

            if len(df) <= int(min_train):
                st.error(
                    "No hay suficientes sorteos posteriores al "
                    "entrenamiento mínimo."
                )
                st.stop()

            test_count = len(
                range(
                    int(min_train),
                    len(df),
                    int(step),
                )
            )

            total_work = test_count * len(STRATEGIES)

            progress = st.progress(0.0)
            status = st.empty()
            all_results = []

            def update_progress(done, total, strategy_name):
                fraction = min(
                    1.0,
                    done / max(1, total),
                )

                progress.progress(fraction)

                status.info(
                    f"Procesando {strategy_name}: "
                    f"{done:,}/{total:,} evaluaciones "
                    f"({fraction:.1%})."
                )

            for index, strategy_name in enumerate(STRATEGIES):
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
                    strategy=strategy_name,
                    exclude_historical=exclude_hist_bt,
                    diversity=diversity_bt,
                    max_overlap=int(max_overlap_bt),
                    progress_callback=update_progress,
                    progress_offset=index * test_count,
                    progress_total=total_work,
                )

                all_results.append(result)

            progress.progress(1.0)
            status.success("Backtesting terminado.")

            results = pd.concat(
                all_results,
                ignore_index=True,
            )

            if results.empty:
                st.warning("No se obtuvieron resultados.")
            else:
                summary = summarize_backtest(results)

                st.subheader("Resumen comparativo")

                st.dataframe(
                    summary.round(2),
                    use_container_width=True,
                    hide_index=True,
                )

                st.subheader("Promedio máximo de aciertos")

                st.bar_chart(
                    summary.set_index("Estrategia")[
                        "Promedio máximo aciertos"
                    ]
                )

                st.subheader("Resultados por sorteo")

                st.dataframe(
                    results.sort_values(
                        ["Fecha", "Estrategia"]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.download_button(
                    "⬇️ Descargar resultados del backtesting",
                    results.to_csv(
                        index=False
                    ).encode("utf-8-sig"),
                    file_name="backtesting_la_tinka.csv",
                    mime="text/csv",
                )

                st.info(
                    "Para que la comparación sea interpretable, las "
                    "estrategias usan el mismo número objetivo de tickets, "
                    "los mismos filtros y la misma regla de diversificación. "
                    "Si alguna estrategia no logra producir todos los tickets "
                    "válidos, la columna Tickets muestra la diferencia."
                )

                st.warning(
                    "El rendimiento histórico no demuestra una ventaja "
                    "predictiva futura. Para evaluar diferencias pequeñas "
                    "se necesitan muchas pruebas y comparación estadística "
                    "con referencias aleatorias."
                )

    # ==========================================================
    # 5. REVENTONES
    # ==========================================================

    elif menu == "Reventones":
        st.header("💰 Reventones y premios")

        if "ReventoVerde" not in df.columns:
            st.warning(
                "No existe información de resaltado de premios."
            )
        else:
            green_df = df[
                df["ReventoVerde"].fillna(False).astype(bool)
            ].copy()

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "Reventones confirmados",
                len(green_df),
            )

            c2.metric(
                "Registros con estado conocido",
                int(df["ReventoVerde"].notna().sum()),
            )

            c3.metric(
                "Estado desconocido",
                int(df["ReventoVerde"].isna().sum()),
            )

            st.subheader("Sorteos marcados en verde")

            show_cols = [
                "Fecha",
                "Sorteo",
                "Premio",
                "SorteosSinReventarAntes",
                "DiasDesdeUltimoReventonAntes",
            ]

            show_cols = [
                c for c in show_cols if c in green_df.columns
            ]

            if green_df.empty:
                st.warning(
                    "No se detectaron celdas verdes en la columna Premio. "
                    "Comprueba el formato de relleno del Excel."
                )
            else:
                st.dataframe(
                    green_df[show_cols].sort_values(
                        ["Fecha", "Sorteo"],
                        ascending=False,
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                dates = (
                    green_df["Fecha"]
                    .sort_values()
                    .tolist()
                )

                if len(dates) >= 2:
                    intervals = [
                        (dates[i] - dates[i - 1]).days
                        for i in range(1, len(dates))
                    ]

                    st.subheader("Intervalos entre reventones")

                    c1, c2, c3 = st.columns(3)

                    c1.metric(
                        "Promedio de días",
                        f"{np.mean(intervals):.1f}",
                    )

                    c2.metric(
                        "Mediana de días",
                        f"{np.median(intervals):.1f}",
                    )

                    c3.metric(
                        "Máximo de días",
                        f"{max(intervals)}",
                    )

                    st.bar_chart(
                        pd.Series(
                            intervals,
                            name="Días entre reventones",
                        )
                    )

                prizes = green_df[
                    ["Fecha", "Sorteo", "Premio"]
                ].dropna()

                if not prizes.empty:
                    st.subheader("Premios asociados")

                    st.line_chart(
                        prizes.sort_values("Fecha").set_index("Fecha")[
                            "Premio"
                        ]
                    )

            st.caption(
                "El color verde se utiliza como señal de reventón según "
                "el archivo. La detección depende del formato real de las "
                "celdas; verifica los registros antes de extraer conclusiones."
            )

    # ==========================================================
    # 6. DIAGNÓSTICO
    # ==========================================================

    elif menu == "Diagnóstico":
        st.header("🔎 Diagnóstico de datos")

        st.write(
            "Revisa la calidad de los registros antes de generar "
            "combinaciones o ejecutar backtesting."
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Filas válidas",
            validation["records"],
        )

        c2.metric(
            "Sorteos duplicados",
            validation["duplicate_draws"],
        )

        c3.metric(
            "Filas exactamente duplicadas",
            validation["exact_duplicates"],
        )

        c4.metric(
            "Bolillas repetidas en fila",
            validation["repeated_balls"],
        )

        st.subheader("Fechas y bolillas observadas")

        st.write({
            "Primera fecha": validation["min_date"],
            "Última fecha": validation["max_date"],
            "Bolilla mínima observada": validation["universe_min"],
            "Bolilla máxima observada": validation["universe_max"],
            "Combinaciones históricas únicas": validation["unique_combinations"],
            "Bolillas fuera del universo vigente": validation["out_of_range"],
        })

        if validation["duplicate_draws"]:
            st.warning(
                "Existen números de sorteo repetidos. Comprueba si "
                "corresponden a registros distintos o duplicados."
            )

        if validation["repeated_balls"]:
            st.error(
                "Hay filas con bolillas repetidas dentro de un mismo sorteo."
            )

        if validation["out_of_range"]:
            st.warning(
                "Hay bolillas que no pertenecen al universo indicado "
                "por la hoja Modificaciones."
            )

        st.subheader("Universo vigente")

        if schedule.empty:
            st.warning(
                "No se pudo procesar la hoja Modificaciones."
            )
        else:
            st.dataframe(
                schedule,
                use_container_width=True,
                hide_index=True,
            )

        st.subheader("Cobertura de premios y reventones")

        st.write({
            "Registros": coverage["total"],
            "Premios con dato": coverage["premios"],
            "Cobertura de premios (%)": round(
                coverage["cobertura_premio_pct"], 2
            ),
            "Reventones confirmados": coverage["reventones_confirmados"],
            "Registros con estado conocido": coverage["reventones_conocidos"],
            "Cobertura de estado (%)": round(
                coverage["cobertura_reventon_pct"], 2
            ),
        })

        st.subheader("Último registro")

        st.dataframe(
            df.tail(1),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================
    # 7. DATOS
    # ==========================================================

    elif menu == "Datos":
        st.header("🗃️ Datos utilizados por la aplicación")

        st.caption(
            "Se muestran los registros después de preparar las fechas, "
            "las seis bolillas principales, el universo histórico y las "
            "variables auxiliares disponibles."
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

        csv_data = df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ Descargar datos procesados",
            data=csv_data,
            file_name="la_tinka_datos_procesados.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()

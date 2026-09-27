# ==============================================================
# LA TINKA AI - VERSION CORREGIDA Y EJECUTABLE
# ==============================================================
# IMPORTANTE:
# Este archivo es Python puro. NO contiene ```python ni ``` al inicio
# o al final. Puede ejecutarse directamente con:
#
#     streamlit run app.py
#
# ==============================================================

# ==============================================================
# LA TINKA AI - SISTEMA DE ANÁLISIS HISTÓRICO Y GENERACIÓN
# ==============================================================
#
# Autor: Sistema desarrollado para análisis estadístico
# Datos: La Tinka 1994-2026
#
# IMPORTANTE:
# Este sistema NO afirma predecir el resultado de un sorteo.
# Genera combinaciones mediante modelos heurísticos y
# probabilísticos construidos exclusivamente con información
# histórica disponible hasta el momento de análisis.
#
# ==============================================================
#
# INSTALACIÓN:
#
# pip install streamlit pandas numpy openpyxl scipy
#
# EJECUCIÓN:
#
# streamlit run app.py
#
# ==============================================================

from __future__ import annotations

import os
import hashlib
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from collections import Counter
from itertools import combinations
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st



# ==============================================================
# CONFIGURACIÓN GENERAL
# ==============================================================

st.set_page_config(
    page_title="La Tinka AI",
    page_icon="🎱",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================
# PROGRAMACIÓN AUTOMÁTICA DEL PRÓXIMO SORTEO
# ==============================================================

DIAS_SORTEO = {
    2: "miércoles",
    6: "domingo",
}


def proximo_sorteo():
    """Devuelve el próximo sorteo programado en hora de Perú.

    Si hoy es miércoles o domingo, considera el sorteo de hoy.
    En cualquier otro día, avanza hasta el siguiente miércoles/domingo.
    """
    ahora = datetime.now(ZoneInfo("America/Lima"))
    fecha = ahora.date()

    while fecha.weekday() not in DIAS_SORTEO:
        fecha += timedelta(days=1)

    return fecha, DIAS_SORTEO[fecha.weekday()]


# ==============================================================
# CONSTANTES
# ==============================================================

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

DATE_CANDIDATES = [
    "Fecha",
    "fecha",
    "FECHA",
]

DRAW_CANDIDATES = [
    "Sorteo",
    "sorteo",
    "SORTEO",
]


# ==============================================================
# UTILIDADES
# ==============================================================

def normalize_text(text):
    """Normaliza nombres de columnas."""
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
    )


def find_file():
    """
    Busca automáticamente el archivo Excel en el directorio
    actual y en la carpeta datos.
    """

    possible_paths = []

    for filename in DEFAULT_FILES:
        possible_paths.extend([
            Path(filename),
            Path("datos") / filename,
            Path("/mnt/data") / filename,
        ])

    for path in possible_paths:
        if path.exists():
            return path

    return None


def stable_seed(*values):
    """
    Genera una semilla reproducible a partir de valores.
    """

    raw = "|".join(map(str, values))

    digest = hashlib.sha256(
        raw.encode("utf-8")
    ).hexdigest()

    return int(digest[:8], 16)


# ==============================================================
# CARGA DEL EXCEL
# ==============================================================

@st.cache_data(show_spinner=False)
def load_excel(path: str) -> Dict[str, pd.DataFrame]:

    xls = pd.ExcelFile(path)

    sheets = {}

    for sheet in xls.sheet_names:
        sheets[sheet] = pd.read_excel(
            path,
            sheet_name=sheet
        )

    return sheets


# ==============================================================
# DETECCIÓN DE COLUMNAS
# ==============================================================

def detect_date_column(df):

    normalized = {
        normalize_text(c): c
        for c in df.columns
    }

    for candidate in DATE_CANDIDATES:

        key = normalize_text(candidate)

        if key in normalized:
            return normalized[key]

    raise ValueError(
        "No se encontró la columna de fecha."
    )


def detect_draw_column(df):

    normalized = {
        normalize_text(c): c
        for c in df.columns
    }

    for candidate in DRAW_CANDIDATES:

        key = normalize_text(candidate)

        if key in normalized:
            return normalized[key]

    return None


def detect_ball_columns(df):

    normalized = {
        normalize_text(c): c
        for c in df.columns
    }

    detected = []

    for i in range(1, 7):

        candidates = [
            f"bolilla {i}",
            f"bolilla{i}",
            f"bola {i}",
            f"bola{i}",
        ]

        found = None

        for candidate in candidates:

            key = normalize_text(candidate)

            if key in normalized:
                found = normalized[key]
                break

        if found is not None:
            detected.append(found)

    if len(detected) != 6:
        raise ValueError(
            "No se pudieron identificar las 6 columnas de bolillas."
        )

    return detected


# ==============================================================
# PREPARACIÓN DE DATOS
# ==============================================================

def prepare_data(df):

    date_col = detect_date_column(df)
    ball_cols = detect_ball_columns(df)
    draw_col = detect_draw_column(df)

    work = df.copy()

    work["Fecha"] = pd.to_datetime(
        work[date_col],
        errors="coerce",
        dayfirst=True
    )

    if draw_col is not None:

        work["Sorteo"] = pd.to_numeric(
            work[draw_col],
            errors="coerce"
        )

    else:

        work["Sorteo"] = np.nan

    for i, col in enumerate(ball_cols, start=1):

        work[f"B{i}"] = pd.to_numeric(
            work[col],
            errors="coerce"
        )

    required = [
        "Fecha",
        "B1",
        "B2",
        "B3",
        "B4",
        "B5",
        "B6",
    ]

    work = work.dropna(
        subset=required
    ).copy()

    for i in range(1, 7):

        work[f"B{i}"] = work[f"B{i}"].astype(int)

    work = work.sort_values(
        "Fecha"
    ).reset_index(drop=True)

    return work


# ==============================================================
# VALIDACIÓN DEL DATASET
# ==============================================================

def validate_data(df):

    report = {}

    report["registros"] = len(df)

    report["fecha_min"] = df["Fecha"].min()

    report["fecha_max"] = df["Fecha"].max()

    report["duplicados_fecha"] = int(
        df["Fecha"].duplicated().sum()
    )

    ball_matrix = df[
        [f"B{i}" for i in range(1, 7)]
    ]

    invalid_range = (
        (ball_matrix < 1)
        |
        (ball_matrix > 100)
    )

    report["valores_fuera_rango"] = int(
        invalid_range.sum().sum()
    )

    duplicate_balls = 0

    for row in ball_matrix.itertuples(index=False):

        if len(set(row)) != 6:
            duplicate_balls += 1

    report["sorteos_con_bolillas_repetidas"] = duplicate_balls

    sorted_rows = np.sort(
        ball_matrix.values,
        axis=1
    )

    historical_combinations = set(
        tuple(row)
        for row in sorted_rows
    )

    report["combinaciones_unicas"] = len(
        historical_combinations
    )

    report["universo_minimo"] = int(
        ball_matrix.min().min()
    )

    report["universo_maximo"] = int(
        ball_matrix.max().max()
    )

    return report


# ==============================================================
# EXTRACCIÓN DE SORTEOS
# ==============================================================

def get_draws(df):

    draws = []

    for _, row in df.iterrows():

        numbers = tuple(
            sorted(
                int(row[f"B{i}"])
                for i in range(1, 7)
            )
        )

        draws.append(
            {
                "fecha": row["Fecha"],
                "sorteo": row["Sorteo"],
                "numeros": numbers,
            }
        )

    return draws


# ==============================================================
# UNIVERSO DE BOLILLAS
# ==============================================================

def get_ball_universe(df):

    values = []

    for i in range(1, 7):

        values.extend(
            df[f"B{i}"].astype(int).tolist()
        )

    return sorted(
        set(values)
    )


def get_historical_universe_until(
    df,
    cutoff_date
):

    subset = df[
        df["Fecha"] <= cutoff_date
    ]

    return get_ball_universe(subset)


# ==============================================================
# FRECUENCIA
# ==============================================================

def calculate_frequency(
    df,
    decay_lambda=0.0
):

    balls = get_ball_universe(df)

    frequency = {
        ball: 0.0
        for ball in balls
    }

    if len(df) == 0:
        return frequency

    max_date = df["Fecha"].max()

    for _, row in df.iterrows():

        if decay_lambda > 0:

            days = (
                max_date - row["Fecha"]
            ).days

            weight = np.exp(
                -decay_lambda * days
            )

        else:

            weight = 1.0

        for i in range(1, 7):

            ball = int(row[f"B{i}"])

            frequency[ball] += weight

    return frequency


# ==============================================================
# ATRASO
# ==============================================================

def calculate_delay(df):

    balls = get_ball_universe(df)

    max_date = df["Fecha"].max()

    delay = {}

    for ball in balls:

        last_date = None

        for _, row in df.iloc[::-1].iterrows():

            numbers = [
                int(row[f"B{i}"])
                for i in range(1, 7)
            ]

            if ball in numbers:

                last_date = row["Fecha"]
                break

        if last_date is None:

            delay[ball] = np.inf

        else:

            delay[ball] = (
                max_date - last_date
            ).days

    return delay


# ==============================================================
# NORMALIZACIÓN
# ==============================================================

def minmax_dict(values):

    series = pd.Series(values, dtype=float)

    if series.empty:
        return {}

    minimum = series.min()
    maximum = series.max()

    if maximum == minimum:

        return {
            k: 1.0
            for k in values
        }

    return {
        k: float(
            (v - minimum)
            /
            (maximum - minimum)
        )
        for k, v in values.items()
    }


# ==============================================================
# SCORE DE BOLILLA
# ==============================================================

def calculate_ball_scores(
    df,
    frequency_weight=0.65,
    recency_weight=0.35,
):

    # ----------------------------------------------------------
    # Frecuencia histórica
    # ----------------------------------------------------------

    freq = calculate_frequency(
        df,
        decay_lambda=0
    )

    # ----------------------------------------------------------
    # Frecuencia reciente
    # ----------------------------------------------------------

    recent = calculate_frequency(
        df,
        decay_lambda=0.003
    )

    freq_norm = minmax_dict(freq)

    recent_norm = minmax_dict(
        recent
    )

    balls = sorted(
        set(freq_norm)
        |
        set(recent_norm)
    )

    scores = {}

    for ball in balls:

        scores[ball] = (
            frequency_weight
            *
            freq_norm.get(ball, 0)
            +
            recency_weight
            *
            recent_norm.get(ball, 0)
        )

    return scores


# ==============================================================
# ATRASO COMO INDICADOR DESCRIPTIVO
# ==============================================================

def calculate_delay_score(df):

    delay = calculate_delay(df)

    finite = {
        k: v
        for k, v in delay.items()
        if np.isfinite(v)
    }

    if not finite:
        return {}

    # ----------------------------------------------------------
    # IMPORTANTE:
    #
    # El atraso NO se interpreta como aumento real de
    # probabilidad.
    #
    # Solo se utiliza como característica descriptiva.
    # ----------------------------------------------------------

    return minmax_dict(
        finite
    )


# ==============================================================
# FRECUENCIA DE PARES
# ==============================================================

def calculate_pair_frequency(df):

    pair_counter = Counter()

    for _, row in df.iterrows():

        numbers = sorted(
            int(row[f"B{i}"])
            for i in range(1, 7)
        )

        for pair in combinations(
            numbers,
            2
        ):

            pair_counter[pair] += 1

    return pair_counter


# ==============================================================
# PROBABILIDAD EMPÍRICA DE ESTRUCTURA
# ==============================================================

def get_low_threshold(df):
    """
    Define dinámicamente el punto medio del universo observado.
    Esto evita asumir que La Tinka siempre utilizó 48 bolillas.
    """
    universe = get_ball_universe(df)

    if not universe:
        return 24

    return (min(universe) + max(universe)) / 2.0


def calculate_structure_distribution(df):

    structures = []
    low_threshold = get_low_threshold(df)

    for _, row in df.iterrows():

        numbers = sorted(
            int(row[f"B{i}"])
            for i in range(1, 7)
        )

        odd = sum(
            n % 2
            for n in numbers
        )

        consecutive = sum(
            1
            for a, b in zip(
                numbers,
                numbers[1:]
            )
            if b == a + 1
        )

        low = sum(
            n <= low_threshold
            for n in numbers
        )

        total = sum(numbers)

        structures.append(
            {
                "odd": odd,
                "consecutive": consecutive,
                "low": low,
                "sum": total,
            }
        )

    return pd.DataFrame(
        structures
    )


# ==============================================================
# ESTADÍSTICAS DE ESTRUCTURA
# ==============================================================

def structure_statistics(df):

    structure = calculate_structure_distribution(
        df
    )

    if structure.empty:
        return {}

    stats = {}

    for column in [
        "odd",
        "consecutive",
        "low",
        "sum",
    ]:

        stats[column] = {
            "mean": float(
                structure[column].mean()
            ),
            "std": float(
                structure[column].std()
            ),
        }

    return stats


# ==============================================================
# SCORE DE ESTRUCTURA
# ==============================================================

def structure_score(
    combination,
    structure_stats,
    low_threshold=24
):

    if not structure_stats:
        return 0.0

    numbers = sorted(
        combination
    )

    odd = sum(
        n % 2
        for n in numbers
    )

    consecutive = sum(
        1
        for a, b in zip(
            numbers,
            numbers[1:]
        )
        if b == a + 1
    )

    low = sum(
        n <= low_threshold
        for n in numbers
    )

    total = sum(numbers)

    observations = {
        "odd": odd,
        "consecutive": consecutive,
        "low": low,
        "sum": total,
    }

    distances = []

    for variable, value in observations.items():

        mean = structure_stats[
            variable
        ]["mean"]

        std = structure_stats[
            variable
        ]["std"]

        if std == 0:
            z = 0
        else:
            z = abs(
                (value - mean) / std
            )

        distances.append(z)

    mean_distance = np.mean(
        distances
    )

    # Menor distancia = mayor compatibilidad
    return float(
        np.exp(-0.5 * mean_distance)
    )


# ==============================================================
# SCORE DE PARES
# ==============================================================

def pair_score(
    combination,
    pair_frequency,
    total_draws
):

    if total_draws <= 0:
        return 0.0

    values = []

    for pair in combinations(
        sorted(combination),
        2
    ):

        count = pair_frequency.get(
            pair,
            0
        )

        values.append(
            np.log1p(count)
        )

    if not values:
        return 0.0

    return float(
        np.mean(values)
    )


# ==============================================================
# SCORE DE COMBINACIÓN
# ==============================================================

def combination_score(
    combination,
    ball_scores,
    pair_frequency,
    structure_stats,
    total_draws,
    low_threshold=24,
    w_ball=0.60,
    w_pair=0.25,
    w_structure=0.15,
):

    # ----------------------------------------------------------
    # Componente individual
    # ----------------------------------------------------------

    individual_values = [
        ball_scores.get(
            ball,
            0
        )
        for ball in combination
    ]

    individual = np.mean(
        individual_values
    )

    # ----------------------------------------------------------
    # Pares
    # ----------------------------------------------------------

    pair = pair_score(
        combination,
        pair_frequency,
        total_draws
    )

    # Normalización aproximada
    pair = pair / max(
        1.0,
        np.log1p(total_draws)
    )

    # ----------------------------------------------------------
    # Estructura
    # ----------------------------------------------------------

    structure = structure_score(
        combination,
        structure_stats,
        low_threshold=low_threshold
    )

    # ----------------------------------------------------------
    # Score final
    # ----------------------------------------------------------

    score = (
        w_ball * individual
        +
        w_pair * pair
        +
        w_structure * structure
    )

    return float(score)


# ==============================================================
# GENERACIÓN PROBABILÍSTICA
# ==============================================================

def weighted_sample_without_replacement(
    balls,
    scores,
    rng,
    k=6
):

    balls = list(balls)

    weights = np.array(
        [
            max(
                float(scores.get(ball, 0)),
                1e-9
            )
            for ball in balls
        ],
        dtype=float
    )

    selected = []

    for _ in range(k):

        if len(balls) == 0:
            break

        probabilities = (
            weights
            /
            weights.sum()
        )

        index = rng.choice(
            len(balls),
            p=probabilities
        )

        selected.append(
            balls.pop(index)
        )

        weights = np.delete(
            weights,
            index
        )

    return tuple(
        sorted(selected)
    )


# ==============================================================
# VALIDACIÓN DE UNA COMBINACIÓN
# ==============================================================

def validate_combination(
    combination,
    universe,
    historical_combinations,
    min_sum,
    max_sum,
    max_consecutive,
    exclude_historical,
):

    if len(combination) != 6:
        return False

    if len(set(combination)) != 6:
        return False

    if not all(
        n in universe
        for n in combination
    ):
        return False

    total = sum(combination)

    if total < min_sum:
        return False

    if total > max_sum:
        return False

    consecutive = sum(
        1
        for a, b in zip(
            combination,
            combination[1:]
        )
        if b == a + 1
    )

    if consecutive > max_consecutive:
        return False

    if (
        exclude_historical
        and tuple(sorted(combination))
        in historical_combinations
    ):
        return False

    return True


# ==============================================================
# GENERACIÓN DE CANDIDATOS
# ==============================================================

def generate_candidates(
    df,
    n_candidates=5000,
    seed=12345,
    min_sum=90,
    max_sum=195,
    max_consecutive=1,
    exclude_historical=True,
    w_ball=0.60,
    w_pair=0.25,
    w_structure=0.15,
):

    rng = np.random.default_rng(
        seed
    )

    universe = get_ball_universe(
        df
    )

    ball_scores = calculate_ball_scores(
        df
    )

    pair_frequency = calculate_pair_frequency(
        df
    )

    structure_stats = structure_statistics(
        df
    )

    low_threshold = get_low_threshold(df)

    historical_combinations = set(
        tuple(
            sorted(
                [
                    int(row[f"B{i}"])
                    for i in range(1, 7)
                ]
            )
        )
        for _, row in df.iterrows()
    )

    candidates = {}

    max_attempts = max(
        n_candidates * 20,
        10000
    )

    attempts = 0

    while (
        len(candidates) < n_candidates
        and attempts < max_attempts
    ):

        attempts += 1

        combination = (
            weighted_sample_without_replacement(
                universe,
                ball_scores,
                rng,
                k=6
            )
        )

        if not validate_combination(
            combination,
            universe,
            historical_combinations,
            min_sum,
            max_sum,
            max_consecutive,
            exclude_historical,
        ):
            continue

        score = combination_score(
            combination,
            ball_scores,
            pair_frequency,
            structure_stats,
            len(df),
            low_threshold=low_threshold,
            w_ball=w_ball,
            w_pair=w_pair,
            w_structure=w_structure,
        )

        candidates[combination] = score

    result = sorted(
        candidates.items(),
        key=lambda x: x[1],
        reverse=True
    )

    return result


# ==============================================================
# TOP COMBINACIONES
# ==============================================================

def select_top_combinations(
    candidates,
    top_n=10,
    diversity=True
):

    selected = []

    for combination, score in candidates:

        if not diversity:
            selected.append(
                (combination, score)
            )

        else:

            # Evitar que las 10 combinaciones
            # sean prácticamente iguales.
            overlap_ok = True

            for existing, _ in selected:

                overlap = len(
                    set(combination)
                    &
                    set(existing)
                )

                if overlap >= 5:

                    overlap_ok = False
                    break

            if overlap_ok:

                selected.append(
                    (combination, score)
                )

        if len(selected) >= top_n:
            break

    return selected


# ==============================================================
# BACKTESTING
# ==============================================================

def evaluate_combination(
    predicted,
    actual
):

    predicted = set(predicted)
    actual = set(actual)

    return len(
        predicted & actual
    )


def backtest_strategy(
    df,
    min_train=100,
    step=10,
    top_n=10,
    seed_base=1000,
):

    df = df.sort_values(
        "Fecha"
    ).reset_index(drop=True)

    results = []

    if len(df) <= min_train:
        return pd.DataFrame()

    for test_index in range(
        min_train,
        len(df),
        step
    ):

        train = df.iloc[
            :test_index
        ].copy()

        actual_row = df.iloc[
            test_index
        ]

        actual = tuple(
            sorted(
                int(
                    actual_row[f"B{i}"]
                )
                for i in range(1, 7)
            )
        )

        candidates = generate_candidates(
            train,
            n_candidates=1500,
            seed=seed_base + test_index,
            min_sum=90,
            max_sum=195,
            max_consecutive=1,
            exclude_historical=False,
        )

        top = select_top_combinations(
            candidates,
            top_n=top_n,
            diversity=True
        )

        if not top:
            continue

        hits = [
            evaluate_combination(
                combo,
                actual
            )
            for combo, _ in top
        ]

        results.append(
            {
                "fecha": actual_row["Fecha"],
                "sorteo": actual_row["Sorteo"],
                "max_aciertos": max(hits),
                "promedio_aciertos": np.mean(hits),
                "aciertos_2plus": int(
                    sum(
                        h >= 2
                        for h in hits
                    )
                ),
                "aciertos_3plus": int(
                    sum(
                        h >= 3
                        for h in hits
                    )
                ),
            }
        )

    return pd.DataFrame(
        results
    )


# ==============================================================
# RESUMEN DE BACKTEST
# ==============================================================

def summarize_backtest(
    results
):

    if results.empty:
        return {}

    return {
        "pruebas": len(results),

        "max_aciertos_promedio":
            results[
                "max_aciertos"
            ].mean(),

        "promedio_aciertos":
            results[
                "promedio_aciertos"
            ].mean(),

        "pruebas_2plus":
            (
                results[
                    "max_aciertos"
                ] >= 2
            ).mean(),

        "pruebas_3plus":
            (
                results[
                    "max_aciertos"
                ] >= 3
            ).mean(),

        "maximo_observado":
            results[
                "max_aciertos"
            ].max(),
    }


# ==============================================================
# FORMATO DE COMBINACIÓN
# ==============================================================

def format_combination(
    combination
):

    return " - ".join(
        f"{n:02d}"
        for n in combination
    )


# ==============================================================
# TABLA DE FRECUENCIAS
# ==============================================================

def frequency_table(df):

    freq = calculate_frequency(
        df
    )

    delay = calculate_delay(
        df
    )

    rows = []

    total_draws = len(df)

    for ball in sorted(freq):

        appearances = int(
            freq[ball]
        )

        percentage = (
            appearances
            /
            (total_draws * 6)
            *
            100
        )

        rows.append(
            {
                "Bolilla": ball,
                "Apariciones": appearances,
                "% sobre bolillas":
                    round(
                        percentage,
                        4
                    ),
                "Atraso días":
                    delay.get(
                        ball,
                        np.nan
                    ),
            }
        )

    return pd.DataFrame(
        rows
    ).sort_values(
        "Apariciones",
        ascending=False
    )


# ==============================================================
# EXPORTACIÓN
# ==============================================================

def candidates_to_dataframe(
    candidates
):

    rows = []

    for rank, (
        combination,
        score
    ) in enumerate(
        candidates,
        start=1
    ):

        row = {
            "Ranking": rank,
            "Combinación":
                format_combination(
                    combination
                ),
            "Score":
                round(score, 8),
        }

        for i, number in enumerate(
            combination,
            start=1
        ):
            row[
                f"Bolilla {i}"
            ] = number

        rows.append(row)

    return pd.DataFrame(
        rows
    )


# ==============================================================
# INTERFAZ
# ==============================================================

def main():

    st.title(
        "🎱 LA TINKA AI"
    )

    st.subheader(
        "Sistema de análisis estadístico y generación de combinaciones"
    )

    st.info(
        """
        El sistema utiliza información histórica para construir
        combinaciones mediante un modelo estadístico/heurístico.

        No existe garantía matemática de que una combinación tenga
        mayor probabilidad real de salir en un sorteo independiente.
        """
    )

    # ----------------------------------------------------------
    # ARCHIVO
    # ----------------------------------------------------------

    st.sidebar.header(
        "📁 Datos"
    )

    default_file = find_file()

    uploaded = st.sidebar.file_uploader(
        "Cargar Excel histórico",
        type=["xlsx"]
    )

    if uploaded is not None:

        temp_path = (
            Path(
                "uploaded_tinka.xlsx"
            )
        )

        with open(
            temp_path,
            "wb"
        ) as f:

            f.write(
                uploaded.getbuffer()
            )

        file_path = temp_path

    elif default_file is not None:

        file_path = default_file

    else:

        st.error(
            """
            No se encontró el archivo Excel.

            Coloca el archivo histórico en la carpeta
            del programa o cárgalo desde el panel lateral.
            """
        )

        st.stop()

    # ----------------------------------------------------------
    # CARGAR
    # ----------------------------------------------------------

    try:

        sheets = load_excel(
            str(file_path)
        )

    except Exception as e:

        st.error(
            f"Error leyendo Excel: {e}"
        )

        st.stop()

    if "Resultados" not in sheets:

        st.error(
            """
            El Excel debe contener una hoja denominada
            'Resultados'.
            """
        )

        st.stop()

    try:

        df = prepare_data(
            sheets["Resultados"]
        )

    except Exception as e:

        st.error(
            f"Error preparando los datos: {e}"
        )

        st.stop()

    # ----------------------------------------------------------
    # VALIDACIÓN
    # ----------------------------------------------------------

    report = validate_data(
        df
    )

    # ----------------------------------------------------------
    # MÉTRICAS
    # ----------------------------------------------------------

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric(
        "Sorteos",
        f"{len(df):,}"
    )

    c2.metric(
        "Desde",
        df["Fecha"].min().strftime(
            "%d/%m/%Y"
        )
    )

    c3.metric(
        "Hasta",
        df["Fecha"].max().strftime(
            "%d/%m/%Y"
        )
    )

    c4.metric(
        "Bolillas",
        len(
            get_ball_universe(df)
        )
    )

    c5.metric(
        "Universo",
        f"{report['universo_minimo']}–"
        f"{report['universo_maximo']}"
    )

    # ----------------------------------------------------------
    # MENÚ
    # ----------------------------------------------------------

    menu = st.sidebar.radio(
        "Módulo",
        [
            "🏠 Inicio",
            "📊 Estadísticas",
            "🎯 Generador",
            "🔬 Backtesting",
            "🧪 Diagnóstico",
            "📋 Datos",
        ]
    )

    # ==========================================================
    # INICIO
    # ==========================================================

    if menu == "🏠 Inicio":

        st.header(
            "Resumen del sistema"
        )

        st.write(
            """
            El motor ha sido diseñado para trabajar directamente
            con el historial disponible, sin asumir que el universo
            de bolillas siempre fue el mismo.
            """
        )

        st.markdown(
            """
            ### Componentes

            - Frecuencia histórica
            - Frecuencia reciente
            - Atraso descriptivo
            - Frecuencia de pares
            - Estructura de combinaciones
            - Generación probabilística
            - Diversificación
            - Backtesting temporal
            - Diagnóstico de calidad
            """
        )

        st.subheader(
            "Últimos sorteos"
        )

        display = df.tail(
            10
        ).copy()

        display["Combinación"] = (
            display[
                [
                    "B1",
                    "B2",
                    "B3",
                    "B4",
                    "B5",
                    "B6",
                ]
            ]
            .astype(int)
            .astype(str)
            .agg(
                " - ".join,
                axis=1
            )
        )

        st.dataframe(
            display[
                [
                    "Fecha",
                    "Sorteo",
                    "Combinación",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================
    # ESTADÍSTICAS
    # ==========================================================

    elif menu == "📊 Estadísticas":

        st.header(
            "Análisis estadístico"
        )

        freq_df = frequency_table(
            df
        )

        st.dataframe(
            freq_df,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader(
            "Frecuencia histórica"
        )

        chart_df = (
            freq_df
            .set_index("Bolilla")
            [["Apariciones"]]
        )

        st.bar_chart(
            chart_df
        )

        st.subheader(
            "Distribución estructural"
        )

        structure = (
            calculate_structure_distribution(
                df
            )
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Promedio suma",
            round(
                structure["sum"].mean(),
                2
            )
        )

        c2.metric(
            "Promedio impares",
            round(
                structure["odd"].mean(),
                2
            )
        )

        c3.metric(
            "Promedio consecutivas",
            round(
                structure[
                    "consecutive"
                ].mean(),
                2
            )
        )

        c4.metric(
            "Promedio ≤24",
            round(
                structure["low"].mean(),
                2
            )
        )

    # ==========================================================
    # GENERADOR
    # ==========================================================

    elif menu == "🎯 Generador":

        st.header(
            "Generador de combinaciones"
        )

        st.warning(
            """
            Las combinaciones generadas son resultados de un
            modelo estadístico/heurístico. No deben interpretarse
            como predicciones garantizadas del próximo sorteo.
            """
        )

        col1, col2 = st.columns(2)

        with col1:

            n_candidates = st.number_input(
                "Candidatos simulados",
                min_value=100,
                max_value=50000,
                value=5000,
                step=500,
            )

            # Se muestran exactamente 3 combinaciones finales.
            top_n = 3

            fecha_sorteo, dia_sorteo = proximo_sorteo()

            st.info(
                f"🎯 Sorteo objetivo: **{dia_sorteo.capitalize()} "
                f"{fecha_sorteo.strftime('%d/%m/%Y')}**\n\n"
                "El sistema genera automáticamente 3 combinaciones."
            )

            min_sum = st.number_input(
                "Suma mínima",
                min_value=21,
                max_value=300,
                value=90,
            )

            max_sum = st.number_input(
                "Suma máxima",
                min_value=21,
                max_value=400,
                value=195,
            )

        with col2:

            max_consecutive = st.number_input(
                "Máximo de pares consecutivos",
                min_value=0,
                max_value=5,
                value=1,
            )

            exclude_historical = st.checkbox(
                "Excluir combinaciones históricas",
                value=True,
            )

            diversity = st.checkbox(
                "Aplicar diversificación",
                value=True,
            )

            seed_mode = st.selectbox(
                "Semilla",
                [
                    "Automática",
                    "Manual",
                ]
            )

        if seed_mode == "Manual":

            seed = st.number_input(
                "Seed",
                min_value=1,
                max_value=999999999,
                value=12345,
            )

        else:

            # La fecha del sorteo objetivo forma parte de la semilla.
            # De esta manera, miércoles y domingo pueden generar una
            # tanda diferente aun cuando el histórico todavía no cambie.
            seed = stable_seed(
                len(df),
                str(
                    df["Fecha"].max()
                ),
                int(
                    df["Sorteo"].fillna(0).iloc[-1]
                ),
                str(fecha_sorteo),
                dia_sorteo,
            )

        if st.button(
            "🚀 Generar combinaciones",
            type="primary"
        ):

            with st.spinner(
                "Procesando modelo..."
            ):

                candidates = generate_candidates(
                    df,
                    n_candidates=int(
                        n_candidates
                    ),
                    seed=int(seed),
                    min_sum=int(min_sum),
                    max_sum=int(max_sum),
                    max_consecutive=int(
                        max_consecutive
                    ),
                    exclude_historical=
                        exclude_historical,
                )

                final = select_top_combinations(
                    candidates,
                    top_n=top_n,
                    diversity=diversity,
                )

            if not final:

                st.error(
                    """
                    No se encontraron combinaciones.
                    Relaja alguno de los filtros.
                    """
                )

            else:

                result_df = candidates_to_dataframe(
                    final
                )

                st.success(
                    f"Se generaron {len(final)} combinaciones para el "
                    f"sorteo del {dia_sorteo} ({fecha_sorteo.strftime('%d/%m/%Y')})."
                )

                st.dataframe(
                    result_df,
                    use_container_width=True,
                    hide_index=True,
                )

                csv = result_df.to_csv(
                    index=False
                ).encode(
                    "utf-8"
                )

                st.download_button(
                    "⬇️ Descargar CSV",
                    csv,
                    "combinaciones_tinka.csv",
                    "text/csv",
                )

                st.subheader(
                    "Detalle de las combinaciones"
                )

                for rank, (
                    combination,
                    score
                ) in enumerate(
                    final,
                    start=1
                ):

                    st.markdown(
                        f"""
                        ### #{rank}

                        **{format_combination(combination)}**

                        Score: `{score:.6f}`
                        """
                    )

    # ==========================================================
    # BACKTESTING
    # ==========================================================

    elif menu == "🔬 Backtesting":

        st.header(
            "Backtesting temporal"
        )

        st.write(
            """
            El backtesting simula el escenario real:

            se utiliza únicamente la información disponible antes
            de cada sorteo y luego se compara contra el resultado
            histórico real.
            """
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            min_train = st.number_input(
                "Sorteos iniciales de entrenamiento",
                min_value=50,
                max_value=1000,
                value=200,
                step=10,
            )

        with col2:

            step = st.number_input(
                "Paso",
                min_value=1,
                max_value=100,
                value=10,
            )

        with col3:

            top_n_bt = st.number_input(
                "Top combinaciones",
                min_value=1,
                max_value=30,
                value=10,
            )

        if st.button(
            "🔬 Ejecutar backtesting",
            type="primary"
        ):

            with st.spinner(
                "Ejecutando backtesting..."
            ):

                bt = backtest_strategy(
                    df,
                    min_train=int(
                        min_train
                    ),
                    step=int(step),
                    top_n=int(
                        top_n_bt
                    ),
                )

            if bt.empty:

                st.warning(
                    "No hay suficientes datos."
                )

            else:

                summary = summarize_backtest(
                    bt
                )

                c1, c2, c3, c4 = st.columns(4)

                c1.metric(
                    "Pruebas",
                    summary[
                        "pruebas"
                    ]
                )

                c2.metric(
                    "Promedio aciertos",
                    round(
                        summary[
                            "promedio_aciertos"
                        ],
                        4
                    )
                )

                c3.metric(
                    "Pruebas ≥2",
                    f"{summary['pruebas_2plus']:.2%}"
                )

                c4.metric(
                    "Máximo observado",
                    summary[
                        "maximo_observado"
                    ]
                )

                st.dataframe(
                    bt,
                    use_container_width=True,
                    hide_index=True,
                )

    # ==========================================================
    # DIAGNÓSTICO
    # ==========================================================

    elif menu == "🧪 Diagnóstico":

        st.header(
            "Diagnóstico del histórico"
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Duplicados de fecha",
            report[
                "duplicados_fecha"
            ]
        )

        c2.metric(
            "Filas con bolillas repetidas",
            report[
                "sorteos_con_bolillas_repetidas"
            ]
        )

        c3.metric(
            "Valores fuera de rango",
            report[
                "valores_fuera_rango"
            ]
        )

        st.subheader(
            "Universo observado"
        )

        universe = get_ball_universe(
            df
        )

        st.write(
            universe
        )

        st.subheader(
            "Años"
        )

        yearly = (
            df.assign(
                Año=df["Fecha"].dt.year
            )
            .groupby("Año")
            .size()
            .reset_index(
                name="Sorteos"
            )
        )

        st.dataframe(
            yearly,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader(
            "Combinaciones repetidas"
        )

        combinations_df = (
            df[
                [
                    "B1",
                    "B2",
                    "B3",
                    "B4",
                    "B5",
                    "B6",
                ]
            ]
            .apply(
                lambda row:
                tuple(
                    sorted(
                        row.astype(int)
                    )
                ),
                axis=1
            )
        )

        counts = (
            combinations_df
            .value_counts()
            .reset_index()
        )

        counts.columns = [
            "Combinación",
            "Frecuencia"
        ]

        repeated = counts[
            counts["Frecuencia"] > 1
        ]

        st.dataframe(
            repeated.head(50),
            use_container_width=True,
            hide_index=True,
        )

    # ==========================================================
    # DATOS
    # ==========================================================

    elif menu == "📋 Datos":

        st.header(
            "Datos históricos"
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )


# ==============================================================
# EJECUCIÓN
# ==============================================================

if __name__ == "__main__":

    main()

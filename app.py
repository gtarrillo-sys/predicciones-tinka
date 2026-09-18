import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from collections import Counter

st.set_page_config(page_title="Pronósticos La Tinka", page_icon="🎯", layout="centered")

st.title("🎯 Simulador y Pronósticos - La Tinka")
st.write("Sistema automatizado con rotación exacta post-sorteo.")

@st.cache_data
def cargar_datos():
    excel_file = 'La_Tinka_Todos_Los_Sorteos_1994_2026.xlsx'
    df = pd.read_excel(excel_file, sheet_name='Histórico Completo 1994-2026', skiprows=2)
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    bolilla_cols = ['Bolilla 1', 'Bolilla 2', 'Bolilla 3', 'Bolilla 4', 'Bolilla 5', 'Bolilla 6']
    for col in bolilla_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df['Fecha_dt'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y', errors='coerce')
    return df.sort_values('Fecha_dt').reset_index(drop=True)

try:
    df = cargar_datos()
    bolilla_cols = ['Bolilla 1', 'Bolilla 2', 'Bolilla 3', 'Bolilla 4', 'Bolilla 5', 'Bolilla 6']
    historical_combinations = set(tuple(sorted(row[bolilla_cols].values)) for _, row in df.iterrows())
    
    all_numbers = df[bolilla_cols].values.flatten()
    freqs = pd.Series(all_numbers).value_counts().sort_index()
    last_draw = df.iloc[-1]
    current_sorteo = int(last_draw['Sorteo N°'])
    sorteo_id = int(last_draw['Sorteo N°'])

    atrasos = {}
    for num in range(1, 49):
        appeared_rows = df[df[bolilla_cols].isin([num]).any(axis=1)]
        atrasos[num] = current_sorteo - int(appeared_rows['Sorteo N°'].astype(int).max()) if not appeared_rows.empty else len(df)

    score_dict = {num: (freqs.get(num, 1) / 428.0) * 0.5 + (atrasos[num] / 300.0) * 0.5 for num in range(1, 49)}
    probs_array = np.array([score_dict[i] for i in range(1, 49)])
    probs_array = probs_array / np.sum(probs_array)
    numbers_pool = np.arange(1, 49)

    def validate_combination(comb):
        sorted_comb = sorted(comb)
        if not (90 <= sum(sorted_comb) <= 195): return False
        if sum(1 for i in range(5) if sorted_comb[i+1] == sorted_comb[i] + 1) > 1: return False
        if tuple(sorted_comb) in historical_combinations: return False
        if sorted_comb[0] > 14 or sorted_comb[5] < 35: return False
        return True

    # 0=Lunes, 1=Martes, 2=Miércoles, 3=Jueves, 4=Viernes, 5=Sábado, 6=Domingo
    dia_actual = datetime.now().weekday()
    
    # Si es Jueves(3), Viernes(4), Sábado(5) o Domingo(6): toca el Domingo
    if dia_actual in [3, 4, 5, 6]:
        tipo_sorteo = "Domingo"
    else:  # Lunes(0), Martes(1), Miércoles(2): toca el Miércoles
        tipo_sorteo = "Miércoles"

    st.success(f"📅 El sistema detectó automáticamente que el **próximo sorteo** es el de **{tipo_sorteo}**.")

    @st.cache_data
    def calcular_jugadas_fijas(s_id, t_sorteo):
        np.random.seed(s_id + (1 if t_sorteo == "Domingo" else 0))
        valid_results = []
        M = 400000
        for _ in range(M):
            draw = np.random.choice(numbers_pool, size=6, p=probs_array, replace=False)
            if validate_combination(draw):
                valid_results.append(tuple(sorted(int(x) for x in draw)))
        return Counter(valid_results).most_common(3)

    top_recs = calcular_jugadas_fijas(sorteo_id, tipo_sorteo)
    
    st.subheader(f"Tus 3 Opciones Oficiales (Sorteo N° {sorteo_id + 1})")
    for idx, (comb, freq) in enumerate(top_recs, 1):
        pares = sum(1 for x in comb if x % 2 == 0)
        st.info(f"**Opción #{idx} (Fija)**\n* 🔢 Números: `{list(comb)}`\n* ➕ Suma: `{sum(comb)}` | Paridad: `{pares}P / {6-pares}I`")

except Exception as e:
    st.error(f"Error cargando los datos: {e}")

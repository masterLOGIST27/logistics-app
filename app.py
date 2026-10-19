import streamlit as st
import pandas as pd
import os
import json
import re
from datetime import datetime

# ==========================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И ДИЗАЙН (СЕРЫЙ ФОН)
# ==========================================
st.set_page_config(page_title="Помощник логиста", page_icon="logo.png", layout="wide")

st.markdown("""
<style>
    /* Основной фон - глубокий серый */
    .stApp { background-color: #333333; }
    
    /* Скрываем стандартную шапку Streamlit */
    header[data-testid="stHeader"] { background-color: #333333 !important; }
    
    .stMarkdown, p, h1, h2, h3, h4, h5, h6 { color: #E0E0E0 !important; font-family: 'Segoe UI', Tahoma, sans-serif; }
    
    hr { border: none; height: 4px; background-color: #F5C518; margin-top: 0px; margin-bottom: 20px; }
    
    /* ЖЕЛТЫЕ ВЫПАДАЮЩИЕ СПИСКИ */
    div[data-baseweb="select"] > div {
        background-color: #F5C518 !important;
        color: #1E1E1E !important;
        border: none !important;
        border-radius: 8px;
    }
    div[data-baseweb="select"] * { color: #1E1E1E !important; font-weight: bold !important; }
    div[data-baseweb="select"] svg { fill: #1E1E1E !important; }
    
    /* ЖЕЛТЫЕ ПОЛЯ ВВОДА */
    div[data-testid="stNumberInput"] > div > div > input {
        background-color: #F5C518 !important; color: #1E1E1E !important; font-weight: bold !important; border: none !important; border-radius: 8px !important;
    }
    div[data-testid="stNumberInput"] button {
        background-color: #F5C518 !important; color: #1E1E1E !important; border: 1px solid #D4A000 !important;
    }
    div[data-testid="stNumberInput"] button:hover { background-color: #FFD700 !important; }
    
    /* КНОПКА */
    .stButton>button {
        background-color: #F5C518 !important; color: #1E1E1E !important; font-weight: bold; border: none; border-radius: 8px; padding: 10px 24px;
    }
    .stButton>button:hover { background-color: #FFD700 !important; box-shadow: 0 0 10px rgba(245, 197, 24, 0.5); }
    
    /* ЗОНА ЗАГРУЗКИ ФАЙЛА (Темно-серая) */
    section[data-testid="stFileUploader"] {
        background-color: #2A2A2A !important; border-radius: 8px; padding: 10px;
    }
    section[data-testid="stFileUploader"] > div {
        background-color: #333333 !important; border: 2px dashed #F5C518 !important;
    }
    section[data-testid="stFileUploader"] span { color: #E0E0E0 !important; }
    section[data-testid="stFileUploader"] button {
        background-color: #F5C518 !important; color: #1E1E1E !important; font-weight: bold;
    }

    /* РАСКРЫВАЮЩИЙСЯ БЛОК (Темно-серый) */
    div[data-testid="stExpander"] {
        background-color: #2A2A2A !important; border: 1px solid #F5C518 !important; border-radius: 8px;
    }
    div[data-testid="stExpander"] summary { color: #F5C518 !important; font-weight: bold; }
    
    /* ТАБЛИЦА РЕЗУЛЬТАТОВ */
    div[data-testid="stDataFrame"] { background-color: #333333 !important; }
    
    /* БОКОВАЯ ПАНЕЛЬ (Чуть темнее основного фона) */
    [data-testid="stSidebar"] { background-color: #262626; border-right: 2px solid #F5C518; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. ЗАГРУЗКА ДАННЫХ
# ==========================================
@st.cache_data
def load_troika_data():
    if os.path.exists("troika_tariffs.json"):
        with open("troika_tariffs.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

TROIKA_DATA = load_troika_data()

if not os.path.exists("tariffs.xlsx"):
    pd.DataFrame([{
        "Перевозчик": "ПЭК", "Откуда": "Москва", "Куда": "Владивосток",
        "До_кг": 999999, "Ставка_кг": 25, "До_м3": 999999, "Ставка_м3": 3000, "Мин_цена": 1000
    }]).to_excel("tariffs.xlsx", index=False)

# ==========================================
# 3. ШАПКА (ПРОСТО КАРТИНКА БЕЗ КРАСОК)
# ==========================================
col_logo, col_title = st.columns([1, 10])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=80)
    else:
        st.warning("logo.png не найден")
        
with col_title:
    st.markdown("<h1 style='margin-bottom: 0;'>Помощник логиста</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #A0A0A0; margin-top: 0;'>Тройка ДВ (из базы) + другие ТК (из Excel)</p>", unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ==========================================
# 4. БАЗА EXCEL
# ==========================================
with st.expander("📂 База тарифов для остальных ТК (Excel)"):
    if "tariffs_df" not in st.session_state:
        st.session_state.tariffs_df = None
    try:
        st.session_state.tariffs_df = pd.read_excel("tariffs.xlsx")
        st.success("✅ Файл tariffs.xlsx загружен.")
    except Exception as e:
        st.error(f"Ошибка: {e}")
        
    uploaded_file = st.file_uploader("Обновить Excel", type=["xlsx"])
    if uploaded_file is not None:
        st.session_state.tariffs_df = pd.read_excel(uploaded_file)
        st.success("Новая база загружена!")

# ==========================================
# 5. ВВОД ДАННЫХ
# ==========================================
col1, col2 = st.columns(2)
with col1:
    st.subheader("📍 Маршрут")
    all_origins = sorted(list(TROIKA_DATA.keys()))
    if not all_origins: all_origins = ["Москва"]
    origin = st.selectbox("Откуда", options=all_origins)
    
    all_destinations = set()
    if origin in TROIKA_DATA:
        for row in TROIKA_DATA[origin]:
            dest = str(row.get("Куда", "")).strip()
            if dest and dest != "nan":
                all_destinations.add(dest)
                
    if st.session_state.tariffs_df is not None:
        excel_dests = st.session_state.tariffs_df["Куда"].dropna().astype(str).str.strip().tolist()
        all_destinations.update(excel_dests)
        
    all_destinations = sorted(list(all_destinations))
    destination = st.selectbox("Куда", options=all_destinations)

with col2:
    st.subheader("📦 Груз")
    weight = st.number_input("Вес (кг)", min_value=0.1, value=100.0)
    length = st.number_input("Длина (м)", min_value=0.01, value=1.0, step=0.1)
    width = st.number_input("Ширина (м)", min_value=0.01, value=1.0, step=0.1)
    height = st.number_input("Высота (м)", min_value=0.01, value=1.0, step=0.1)
    places = st.number_input("Количество мест", min_value=1, value=1)

volume = length * width * height * places
st.info(f"**Объём:** {volume:.2f} м³")

# ==========================================
# 6. ЛОГИКА РАСЧЕТА
# ==========================================
def calc_troika(row, weight, volume):
    price_w = 0
    price_v = 0
    for col, val in row.items():
        if "кг" in col and "м3" not in col:
            try:
                limit = float(re.search(r'\d+\.?\d*', col).group())
                if weight <= limit:
                    rate = float(str(val).replace(',', '.').replace(' ', ''))
                    price_w = rate if limit <= 40 else weight * rate
                    break
            except: pass
    for col, val in row.items():
        if "м3" in col:
            try:
                limit = float(re.search(r'\d+\.?\d*', col).group())
                if volume <= limit:
                    rate = float(str(val).replace(',', '.').replace(' ', ''))
                    price_v = volume * rate 
                    break
            except: pass
    if price_w == 0:
        for col, val in reversed(list(row.items())):
            if "кг" in col and "м3" not in col:
                try:
                    rate = float(str(val).replace(',', '.').replace(' ', ''))
                    price_w = weight * rate
                except: pass
                break
    if price_v == 0:
        for col, val in reversed(list(row.items())):
            if "м3" in col:
                try:
                    rate = float(str(val).replace(',', '.').replace(' ', ''))
                    price_v = volume * rate
                except: pass
                break
    return max(price_w, price_v)

if st.button("Рассчитать и сравнить", type="primary"):
    results = []
    
    if origin in TROIKA_DATA:
        for row in TROIKA_DATA[origin]:
            if str(row.get("Куда", "")).strip().lower() == str(destination).strip().lower():
                price = calc_troika(row, weight, volume)
                if price > 0:
                    results.append({"Перевозчик": "Тройка ДВ", "Цена, ₽": round(price, 2)})
                break
                
    if st.session_state.tariffs_df is not None:
        df = st.session_state.tariffs_df
        route_df = df[
            (df["Откуда"].astype(str).str.strip().str.lower() == str(origin).strip().lower()) & 
            (df["Куда"].astype(str).str.strip().str.lower() == str(destination).strip().lower())
        ]
        
        for tk_name, group in route_df.groupby("Перевозчик"):
            w_rate = group.iloc[0]["Ставка_кг"]
            v_rate = group.iloc[0]["Ставка_м3"]
            min_price = group.iloc[0]["Мин_цена"]
            
            pw = weight * w_rate if pd.notna(w_rate) else 0
            pv = volume * v_rate if pd.notna(v_rate) else 0
            final = max(pw, pv)
            
            if pd.notna(min_price) and final < min_price:
                final = min_price
                
            if final > 0:
                results.append({"Перевозчик": tk_name, "Цена, ₽": round(final, 2)})

    if not results:
        st.warning("Нет данных для этого маршрута.")
    else:
        res_df = pd.DataFrame(results).sort_values(by="Цена, ₽")
        st.subheader("📊 Результаты сравнения")
        
        def highlight_best(s):
            is_best = s["Цена, ₽"] == res_df["Цена, ₽"].min()
            return ['background-color: #4A4A1A; color: #F5C518; font-weight: bold;'] * len(s) if is_best else [''] * len(s)
        
        st.dataframe(res_df.style.apply(highlight_best, axis=1), use_container_width=True)
        
        if "history" not in st.session_state:
            st.session_state.history = []
        for _, row in res_df.iterrows():
            st.session_state.history.append({
                "Дата": datetime.now().strftime("%d.%m %H:%M"),
                "Маршрут": f"{origin}-{destination}",
                "Перевозчик": row["Перевозчик"],
                "Цена": row["Цена, ₽"]
            })

# ==========================================
# 7. ИСТОРИЯ
# ==========================================
with st.sidebar:
    st.header("🕒 История запросов")
    if "history" not in st.session_state or not st.session_state.history:
        st.write("История пуста.")
    else:
        for item in reversed(st.session_state.history[-20:]):
            st.write(f"{item['Дата']} | {item['Маршрут']} | {item['Перевозчик']} | {item['Цена']} ₽")
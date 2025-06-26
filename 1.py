import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder

# Очищення кешу на кожному запуску (для розробки)
st.cache_data.clear()

# Параметри сторінки
st.set_page_config(
    page_title="Аналіз постачань",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Кастомні стилі для темного фону та меню
st.markdown(
    """
    <style>
      /* Темний фон для всіх основних контейнерів */
      [data-testid="stAppViewContainer"],
      [data-testid="stMain"],
      [data-testid="stMainBlockContainer"],
      .appview-container,
      .stApp,
      .main,
      .block-container {
          background-color: #001f3f !important;
      }
      /* Темний фон бокової панелі */
      [data-testid="stSidebar"] > div:first-child {
          background-color: #001f3f !important;
      }
      /* Стилі для вкладок меню */
      div[data-baseweb="tab-list"] {
          background-color: #002957 !important;
      }
      div[data-baseweb="tab-list"] button {
          background-color: #444 !important;
          color: #fff !important;
      }
      div[data-baseweb="tab-list"] button[aria-selected="true"] {
          background-color: #0073e6 !important;
      }
      div[data-baseweb="tab-list"] button:hover {
          background-color: #555 !important;
      }
      /* Загальні тексти та посилання */
      .stText, .stMarkdown, [data-testid="stMarkdownContainer"] {
          color: #ffffff !important;
      }
      .stButton>button {
          background-color: #0073e6 !important;
          color: #ffffff !important;
      }
    </style>
    """,
    unsafe_allow_html=True
)

# Завантаження даних із кешуванням
@st.cache_data(ttl=600)
def load_data(path: str):
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.error(f"Помилка завантаження даних: {e}")
        return pd.DataFrame()

# Основний код
path = "ЄРПН (Реєстр ПН_РК товар) (59).xlsx"
df = load_data(path)

# Назви колонок
customer_col = 'Найменування/ПІБ (Покупець)'
nomenclature_col = 'Номенклатура товарів/послуг'
quantity_col = 'Кількість (об’єм , обсяг)'
supply_col = 'Обсяги постачання (база оподаткування) без урахування податку на додану вартість'
price_col = 'Ціна постачання одиниці товару / послуги без урахування податку на додану вартість'

# Sidebar: фільтри та експорт
st.sidebar.header("Фільтри та експорт")
# Фільтр за датою
date_cols = [c for c in df.columns if 'дата' in c.lower()]
if date_cols:
    dc = st.sidebar.selectbox("Стовпець з датою", date_cols)
    min_d, max_d = df[dc].min().date(), df[dc].max().date()
    start, end = st.sidebar.slider("Період", (min_d, max_d), min_d, max_d)
    df = df[df[dc].dt.date.between(start, end)]
all_customers = df[customer_col].dropna().unique()

# Функція експорту в Excel

def to_excel(df_export):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False)
    return buf.getvalue()

# Вкладки меню
tab1, tab2 = st.tabs(["Порівняння Топ-5", "Детальний аналіз"])

# Зведені дані
agg = {quantity_col: 'sum', supply_col: 'sum', price_col: 'mean', nomenclature_col: pd.Series.nunique}
summary = df.groupby(customer_col).agg(agg).reset_index()

# Вкладка 1: Топ-5 покупців
with tab1:
    st.header("Топ-5 покупців за обсягом постачання")
    top5 = summary.nlargest(5, supply_col)
    st.download_button("Експорт Топ-5 Excel", to_excel(top5), "top5.xlsx")
    gb = GridOptionsBuilder.from_dataframe(top5)
    gb.configure_pagination()
    AgGrid(top5, gridOptions=gb.build(), enable_enterprise_modules=True)
    # Метрики
    metrics = st.multiselect("Метрики", [quantity_col, supply_col, price_col, nomenclature_col], default=[supply_col])
    for metric in metrics:
        fig = px.bar(top5, x=customer_col, y=metric, title=metric, labels={customer_col: 'Покупець', metric: metric})
        st.plotly_chart(fig, use_container_width=True)

# Вкладка 2: Детальний аналіз
with tab2:
    st.header("Детальний аналіз покупки")
    buyer = st.selectbox("Оберіть покупця", sorted(all_customers))
    data_b = df[df[customer_col] == buyer]
    # KPI
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Обсяг постачання", f"{data_b[supply_col].sum():,.2f}")
    k2.metric("Кількість", f"{data_b[quantity_col].sum():,.2f}")
    k3.metric("Середня ціна", f"{data_b[price_col].mean():,.2f}")
    k4.metric("Унікальні товари", data_b[nomenclature_col].nunique())
    # Таблиця деталей
    det = data_b.groupby(nomenclature_col).agg(agg).reset_index()
    st.download_button("Експорт деталізації Excel", to_excel(det), "detail.xlsx")
    gb2 = GridOptionsBuilder.from_dataframe(det)
    gb2.configure_pagination()
    AgGrid(det, gridOptions=gb2.build(), enable_enterprise_modules=True)
    # Діаграми
    fig_pie = px.pie(det, names=nomenclature_col, values=supply_col, title="Структура обсягів постачання")
    st.plotly_chart(fig_pie, use_container_width=True)
    fig_bar = px.bar(det.nlargest(10, supply_col), x=supply_col, y=nomenclature_col, orientation='h', title="Топ-10 товарів за обсягом")
    st.plotly_chart(fig_bar, use_container_width=True)

# Експорт зведеного звіту
st.sidebar.download_button("Експорт Summary Excel", to_excel(summary), "summary.xlsx")

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# Установка широкого макета страницы
st.set_page_config(layout="wide")

# Загрузка данных из файла
file_path = '12.12 (2).xlsx'
data = pd.read_excel(file_path)

# Преобразуем столбцы к нужным типам
data['Месяц'] = pd.to_numeric(data['Месяц'], errors='coerce')
data['Сумма фактических продаж 2024'] = pd.to_numeric(data['Сумма фактических продаж 2024'], errors='coerce')
data['Сумма Д2 б/НДС'] = pd.to_numeric(data['Сумма Д2 б/НДС'], errors='coerce')
data['Сумма НПК б/НДС'] = pd.to_numeric(data['Сумма НПК б/НДС'], errors='coerce')

# 1. Расчет маржи и маржинальности
data['Маржа, руб б/НДС'] = data['Сумма фактических продаж 2024'] - data['Сумма Д2 б/НДС']
data['Маржа, %'] = (data['Маржа, руб б/НДС'] / data['Сумма Д2 б/НДС']) * 100
data['Маржинальность'] = (data['Маржа, руб б/НДС'] / data['Сумма фактических продаж 2024']) * 100

# 2. Сегментация по маржинальности
def classify_margin(row):
    if row['Маржинальность'] >= 30:
        return 'Высокая маржинальность'
    elif row['Маржинальность'] >= 15:
        return 'Средняя маржинальность'
    else:
        return 'Низкая маржинальность'

data['Сегмент'] = data.apply(classify_margin, axis=1)

# Навигация по страницам
page = st.sidebar.selectbox("Страница", ["Фильтры и таблица", "Графики"])

if page == "Фильтры и таблица":
    # Заголовок приложения
    st.title("Анализ маржинальности 2024")

    # Создаем фильтры в боковой панели
    with st.expander("Фильтры", expanded=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            # Мультивыбор для месяца
            month_options = sorted(data['Месяц'].dropna().unique())
            selected_months = st.multiselect('Месяц', month_options, default=month_options)

        with col2:
            # Выпадающий список для менеджера
            manager_options = sorted(data['Менеджер'].unique())
            selected_manager = st.selectbox('Менеджер', ["Все"] + manager_options)

        with col3:
            # Выпадающий список для имени клиента
            client_name_options = sorted(data['Наименование клиента'].unique())
            selected_client_name = st.selectbox('Наименование клиента', ["Все"] + client_name_options)

        # Фильтр по субкатегории
        subcategory_options = ["Все"] + sorted(data['Субкатегория'].dropna().unique())
        selected_subcategories = st.multiselect('Субкатегория', subcategory_options, default="Все")

    # Применяем фильтры к данным
    filtered_data = data[data['Месяц'].isin(selected_months)]

    if selected_manager != "Все":
        filtered_data = filtered_data[filtered_data['Менеджер'] == selected_manager]

    if selected_client_name != "Все":
        filtered_data = filtered_data[filtered_data['Наименование клиента'] == selected_client_name]

    # Применяем фильтр по субкатегории (если выбрано "Все", отображаем все субкатегории)
    if "Все" not in selected_subcategories:
        filtered_data = filtered_data[filtered_data['Субкатегория'].isin(selected_subcategories)]

    # Фильтр по маржинальности
    margin_level = st.selectbox("Уровень маржинальности", ["Все", "Высокая маржинальность", "Средняя маржинальность", "Низкая маржинальность"])

    if margin_level != "Все":
        filtered_data = filtered_data[filtered_data['Сегмент'] == margin_level]

    # Форматирование чисел с пробелами для отображения тысячных и миллионных значений
    for col in ['Сумма фактических продаж 2024', 'Сумма Д2 б/НДС', 'Сумма НПК б/НДС', 'Маржа, руб б/НДС', 'Маржа, %']:
        if col in filtered_data.columns:
            filtered_data[col] = filtered_data[col].apply(lambda x: f"{x:,.2f}".replace(",", " ").replace(".", ","))

    # Отображение отфильтрованных данных по продуктам
    st.write(f"Детальные данные продуктов с уровнем маржинальности: {margin_level}")
    st.dataframe(filtered_data[['Месяц', 'Менеджер', 'Наименование клиента', 'Субкатегория', 'Наименование продукта', 'Сумма фактических продаж 2024', 'Сумма Д2 б/НДС', 'Сумма НПК б/НДС', 'Маржа, руб б/НДС', 'Маржа, %', 'Маржинальность', 'Сегмент']].reset_index(drop=True), use_container_width=True)

    # Добавляем подпись о валюте
    st.caption("Цены указаны в руб б/НДС")

    # Функция для сохранения данных в Excel
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Преобразуем числовые значения обратно перед сохранением
            save_df = df.copy()
            for col in ['Сумма фактических продаж 2024', 'Сумма Д2 б/НДС', 'Сумма НПК б/НДС', 'Маржа, руб б/НДС', 'Маржа, %']:
                save_df[col] = save_df[col].replace(" ", "", regex=True).replace(",", ".", regex=True).astype(float)
            save_df.to_excel(writer, index=False, sheet_name='Детальные данные')
        output.seek(0)
        return output

    # Создаем кнопку для скачивания
    excel_data = to_excel(filtered_data)
    st.download_button(
        label="Скачать в Excel",
        data=excel_data,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif page == "Графики":
    # # Заголовок страницы с графиками
    # st.title("Графики маржинальности")

    # # Построение графика распределения маржинальности
    # fig, ax = plt.subplots(figsize=(10, 6))
    # data['Маржинальность'].hist(bins=30, ax=ax, color='skyblue', edgecolor='black')
    # ax.set_title("Распределение маржинальности", fontsize=16)
    # ax.set_xlabel("Маржинальность (%)", fontsize=14)
    # ax.set_ylabel("Частота", fontsize=14)
    # st.pyplot(fig)

    # # Построение графика средней маржинальности по сегментам
    # segment_avg_margin = data.groupby('Сегмент')['Маржинальность'].mean().sort_values()
    # fig, ax = plt.subplots(figsize=(10, 6))
    # segment_avg_margin.plot(kind='bar', ax=ax, color='lightgreen', edgecolor='black')
    # ax.set_title("Средняя маржинальность по сегментам", fontsize=16)
    # ax.set_xlabel("Сегмент", fontsize=14)
    # ax.set_ylabel("Средняя маржинальность (%)", fontsize=14)
    # st.pyplot(fig)

    # Построение графика общей маржи по месяцам
    monthly_margin = data.groupby('Месяц')['Маржа'].sum()
    fig, ax = plt.subplots(figsize=(10, 6))
    monthly_margin.plot(kind='line', ax=ax, marker='o', color='orange')
    ax.set_title("Общая маржа по месяцам", fontsize=16)
    ax.set_xlabel("Месяц", fontsize=14)
    ax.set_ylabel("Общая маржа", fontsize=14)
    ax.set_xticks(range(1, 13))  # Устанавливаем метки для всех месяцев
    ax.set_xticklabels([str(i) for i in range(1, 13)])
    st.pyplot(fig)

    # Построение графика средней маржинальности по субкатегориям
    subcategory_avg_margin = data.groupby('Субкатегория')['Маржинальность'].mean().sort_values()
    fig, ax = plt.subplots(figsize=(10, 6))
    subcategory_avg_margin.plot(kind='bar', ax=ax, color='purple', edgecolor='black')
    ax.set_title("Средняя маржинальность по субкатегориям", fontsize=16)
    ax.set_xlabel("Субкатегория", fontsize=14)
    ax.set_ylabel("Средняя маржинальность (%)", fontsize=14)
    st.pyplot(fig)


import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime
import openpyxl
from scipy.stats import mannwhitneyu, fisher_exact

# Настройка страницы
st.set_page_config(
    page_title="УЗИ Статистика + Японские свечи",
    page_icon="📊",
    layout="wide"
)

# Заголовок
st.title("📊 УЗИ Статистический калькулятор + Японские свечи")
st.markdown("---")

# Создаём вкладки
tab1, tab2 = st.tabs(["📊 Калькулятор (Фишер/Манн-Уитни)", "📈 Японские свечи (гемодинамика)"])

# ============================================
# ВКЛАДКА 1: КАЛЬКУЛЯТОР
# ============================================
with tab1:
    st.header("📊 Калькулятор статистической значимости")
    
    # Боковая панель с настройками для калькулятора
    with st.sidebar:
        st.header("⚙️ Настройки калькулятора")
        
        # Выбор типа анализа
        analysis_type = st.radio(
            "Выберите тип данных:",
            options=[
                "📊 Качественные (таблица 2x2) — критерий Фишера",
                "📈 Количественные (две группы) — Манн-Уитни"
            ],
            key="calc_type"
        )
        
        st.markdown("---")
        st.markdown("### 📋 История расчётов")
        
        # Инициализация истории в session_state
        if 'history' not in st.session_state:
            st.session_state.history = []
        
        # Показываем последние 5 расчётов
        if st.session_state.history:
            for i, calc in enumerate(st.session_state.history[-5:]):
                with st.expander(f"{calc['time']} - {calc['type']}"):
                    st.write(f"**Результат:** {calc['result']}")
                    st.write(f"**p-value:** {calc['p_value']}")
        else:
            st.info("История пуста")
    
    # Основная область калькулятора
    if "📊 Качественные" in analysis_type:
        st.subheader("📊 Точный критерий Фишера")
        st.markdown("""
        **Для каких данных:** таблицы сопряженности 2×2 (наличие/отсутствие признака в двух группах)
        
        **Пример:** сравнение частоты двустороннего поражения в группах ПОЯ и РЯ
        """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Группа 1 (например, ПОЯ)")
            a = st.number_input("Есть признак (a)", min_value=0, value=16, step=1, key='fisher_a')
            b = st.number_input("Нет признака (b)", min_value=0, value=12, step=1, key='fisher_b')
            
        with col2:
            st.subheader("Группа 2 (например, РЯ)")
            c = st.number_input("Есть признак (c)", min_value=0, value=1, step=1, key='fisher_c')
            d = st.number_input("Нет признака (d)", min_value=0, value=34, step=1, key='fisher_d')
        
        # Дополнительные настройки
        st.markdown("---")
        col3, col4 = st.columns(2)
        
        with col3:
            group1_name = st.text_input("Название группы 1", "ПОЯ", key='fisher_g1')
            group2_name = st.text_input("Название группы 2", "РЯ", key='fisher_g2')
            feature_name = st.text_input("Название признака", "Двустороннее поражение", key='fisher_feat')
        
        with col4:
            alternative = st.selectbox(
                "Альтернативная гипотеза",
                options=["two-sided", "greater", "less"],
                format_func=lambda x: {
                    "two-sided": "Двусторонняя (есть связь)",
                    "greater": "Отношение шансов > 1",
                    "less": "Отношение шансов < 1"
                }[x],
                key='fisher_alt'
            )
        
        if st.button("🧮 Рассчитать критерий Фишера", type="primary", use_container_width=True, key='fisher_btn'):
            # Проверка
            if a < 0 or b < 0 or c < 0 or d < 0:
                st.error("❌ Все значения должны быть неотрицательными!")
            else:
                # Создаем таблицу
                table = [[a, b], [c, d]]
                
                # Расчет
                odds_ratio, p_value = fisher_exact(table, alternative=alternative)
                
                # Сохраняем в историю
                timestamp = datetime.now().strftime("%H:%M:%S")
                st.session_state.history.append({
                    'time': timestamp,
                    'type': 'Критерий Фишера',
                    'result': f"{feature_name}: {group1_name} vs {group2_name}",
                    'p_value': f"{p_value:.4f}"
                })
                
                # Отображение результатов
                st.markdown("---")
                st.subheader("📊 Результаты")
                
                col_res1, col_res2, col_res3 = st.columns(3)
                
                with col_res1:
                    st.metric("Отношение шансов (OR)", f"{odds_ratio:.3f}")
                    
                    # Интерпретация OR
                    if odds_ratio > 1:
                        or_interp = f"Шанс признака выше в группе {group1_name}"
                    elif odds_ratio < 1:
                        or_interp = f"Шанс признака выше в группе {group2_name}"
                    else:
                        or_interp = "Шансы равны"
                    st.caption(or_interp)
                
                with col_res2:
                    # Форматирование p-value
                    if p_value < 0.0001:
                        p_display = "p < 0.0001"
                        p_color = "green"
                    elif p_value < 0.001:
                        p_display = "p < 0.001"
                        p_color = "green"
                    elif p_value < 0.01:
                        p_display = f"p = {p_value:.3f}"
                        p_color = "green"
                    elif p_value < 0.05:
                        p_display = f"p = {p_value:.3f}"
                        p_color = "orange"
                    else:
                        p_display = f"p = {p_value:.3f}"
                        p_color = "red"
                    
                    st.markdown(f"### :{p_color}[{p_display}]")
                
                with col_res3:
                    # Интерпретация
                    st.subheader("📋 Интерпретация")
                    if p_value < 0.05:
                        st.success("✅ **Статистически значимо**")
                        st.write(f"Различия между {group1_name} и {group2_name} по признаку **{feature_name}** статистически значимы (p < 0.05)")
                    else:
                        st.warning("⚠️ **Статистически не значимо**")
                        st.write(f"Различия между {group1_name} и {group2_name} по признаку **{feature_name}** не доказаны (p ≥ 0.05)")
                
                # Показываем таблицу
                st.markdown("---")
                st.subheader("📋 Таблица сопряженности")
                
                df_result = pd.DataFrame(
                    table,
                    index=[group1_name, group2_name],
                    columns=["Есть признак", "Нет признака"]
                )
                
                # Добавляем итоги
                df_result.loc['Итого'] = df_result.sum()
                
                st.dataframe(df_result, use_container_width=True)
                
                # Кнопка копирования
                result_text = f"{feature_name}: {group1_name} vs {group2_name}\n"
                result_text += f"OR = {odds_ratio:.3f}, p = {p_value:.4f}\n"
                result_text += f"Таблица: {a} {b} | {c} {d}"
                
                st.code(result_text, language="text")

    else:
        st.subheader("📈 U-критерий Манна-Уитни")
        st.markdown("""
        **Для каких данных:** количественные показатели в двух независимых группах
        
        **Пример:** сравнение размеров опухоли в группе ПОЯ и РЯ
        """)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Группа 1 (например, ПОЯ)")
            group1_data = st.text_area(
                "Введите значения через запятую или пробел",
                value="35, 42, 38, 41, 39, 37, 43, 40, 36, 38",
                height=150,
                key='mw_g1'
            )
            group1_name_mw = st.text_input("Название группы 1", "ПОЯ", key='mw_g1_name')
        
        with col2:
            st.subheader("Группа 2 (например, РЯ)")
            group2_data = st.text_area(
                "Введите значения через запятую или пробел",
                value="96, 102, 88, 95, 98, 91, 97, 99, 94, 93",
                height=150,
                key='mw_g2'
            )
            group2_name_mw = st.text_input("Название группы 2", "РЯ", key='mw_g2_name')
        
        feature_name_mw = st.text_input("Название показателя", "Максимальный размер, мм", key='mw_feat')
        
        if st.button("🧮 Рассчитать критерий Манна-Уитни", type="primary", use_container_width=True, key='mw_btn'):
            try:
                # Парсим данные
                g1 = [float(x.strip()) for x in group1_data.replace(',', ' ').split()]
                g2 = [float(x.strip()) for x in group2_data.replace(',', ' ').split()]
                
                if len(g1) < 2 or len(g2) < 2:
                    st.error("❌ В каждой группе должно быть минимум 2 значения")
                else:
                    # Расчет
                    statistic, p_value = mannwhitneyu(g1, g2, alternative='two-sided')
                    
                    # Сохраняем в историю
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    st.session_state.history.append({
                        'time': timestamp,
                        'type': 'Манн-Уитни',
                        'result': f"{feature_name_mw}: {group1_name_mw} vs {group2_name_mw}",
                        'p_value': f"{p_value:.4f}"
                    })
                    
                    # Отображение результатов
                    st.markdown("---")
                    st.subheader("📊 Результаты")
                    
                    col_res1, col_res2, col_res3 = st.columns(3)
                    
                    with col_res1:
                        st.metric("Статистика U", f"{statistic:.1f}")
                        st.metric(f"Медиана {group1_name_mw}", f"{np.median(g1):.1f}")
                        st.metric(f"Медиана {group2_name_mw}", f"{np.median(g2):.1f}")
                    
                    with col_res2:
                        # Форматирование p-value
                        if p_value < 0.0001:
                            p_display = "p < 0.0001"
                            p_color = "green"
                        elif p_value < 0.001:
                            p_display = "p < 0.001"
                            p_color = "green"
                        elif p_value < 0.01:
                            p_display = f"p = {p_value:.3f}"
                            p_color = "green"
                        elif p_value < 0.05:
                            p_display = f"p = {p_value:.3f}"
                            p_color = "orange"
                        else:
                            p_display = f"p = {p_value:.3f}"
                            p_color = "red"
                        
                        st.markdown(f"### :{p_color}[{p_display}]")
                    
                    with col_res3:
                        st.subheader("📋 Интерпретация")
                        if p_value < 0.05:
                            st.success("✅ **Статистически значимо**")
                            st.write(f"Различия между {group1_name_mw} и {group2_name_mw} по показателю **{feature_name_mw}** статистически значимы (p < 0.05)")
                        else:
                            st.warning("⚠️ **Статистически не значимо**")
                            st.write(f"Различия между {group1_name_mw} и {group2_name_mw} по показателю **{feature_name_mw}** не доказаны (p ≥ 0.05)")
                    
                    # Визуализация
                    st.markdown("---")
                    st.subheader("📈 Box plot")
                    
                    fig = go.Figure()
                    fig.add_trace(go.Box(y=g1, name=group1_name_mw, boxmean='sd'))
                    fig.add_trace(go.Box(y=g2, name=group2_name_mw, boxmean='sd'))
                    fig.update_layout(title=f"Сравнение: {feature_name_mw}")
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Описательная статистика
                    st.markdown("---")
                    st.subheader("📊 Описательная статистика")
                    
                    stats_df = pd.DataFrame({
                        'Показатель': ['Количество', 'Среднее', 'Медиана', 'Стд. отклонение', 'Мин', 'Макс'],
                        group1_name_mw: [len(g1), np.mean(g1), np.median(g1), np.std(g1), np.min(g1), np.max(g1)],
                        group2_name_mw: [len(g2), np.mean(g2), np.median(g2), np.std(g2), np.min(g2), np.max(g2)]
                    })
                    
                    st.dataframe(stats_df, use_container_width=True)
                    
                    # Кнопка копирования
                    result_text = f"{feature_name_mw}: {group1_name_mw} vs {group2_name_mw}\n"
                    result_text += f"U = {statistic:.1f}, p = {p_value:.4f}\n"
                    result_text += f"Медианы: {np.median(g1):.1f} vs {np.median(g2):.1f}"
                    
                    st.code(result_text, language="text")
                    
            except Exception as e:
                st.error(f"❌ Ошибка при обработке данных: {e}")
                st.info("Убедитесь, что значения введены правильно (числа через запятую или пробел)")

# ============================================
# ВКЛАДКА 2: ЯПОНСКИЕ СВЕЧИ
# ============================================
with tab2:
    st.header("📈 Японские свечи для анализа гемодинамики")
    
    # Инструкция
    with st.expander("📖 ИНСТРУКЦИЯ: Как подготовить файл Excel", expanded=True):
        st.markdown("""
        ### Формат данных (строго!)
        
        **Файл Excel должен содержать всего 5 колонок:**
        
        | **ID** | **Диагноз** | **V_капс** | **V_перег** | **V_центр** |
        |--------|-------------|------------|-------------|-------------|
        | ПОЯ-001 | ПОЯ | 35 | 12 | 15 |
        | ПОЯ-002 | ПОЯ | 38 | 10 | 18 |
        | РЯ-001 | РЯ | 32 | 35 | 41 |
        
        ### Правила:
        1. **ID** - любой текст (ПОЯ-001, РЯ-001, Пациент1 и т.д.)
        2. **Диагноз** - строго два варианта: **ПОЯ** или **РЯ**
        3. **V_капс** - скорость в капсуле (число)
        4. **V_перег** - скорость в перегородках (число)
        5. **V_центр** - скорость в центре (число)
        
        ✅ Все 5 колонок обязательны!
        """)
    
    # Загрузка файла
    uploaded_file = st.file_uploader(
        "📂 Загрузите файл Excel", 
        type=['xlsx', 'xls'],
        key='candle_upload',
        help="Файл должен содержать колонки: ID, Диагноз, V_капс, V_перег, V_центр"
    )
    
    if uploaded_file:
        try:
            # Читаем файл
            df = pd.read_excel(uploaded_file)
            
            # Проверка колонок
            required_columns = ['ID', 'Диагноз', 'V_капс', 'V_перег', 'V_центр']
            if not all(col in df.columns for col in required_columns):
                st.error(f"❌ Ошибка: В файле должны быть колонки: {', '.join(required_columns)}")
                st.write("Ваши колонки:", list(df.columns))
            else:
                st.success("✅ Файл загружен, колонки найдены")
                
                # Показываем первые строки
                st.subheader("📋 Загруженные данные (первые 5 строк)")
                st.dataframe(df.head())
                
                # Статистика по группам
                st.subheader("📊 Статистика по группам")
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"ПОЯ: {len(df[df['Диагноз'] == 'ПОЯ'])} пациентов")
                with col2:
                    st.info(f"РЯ: {len(df[df['Диагноз'] == 'РЯ'])} пациентов")
                
                # === РАСЧЁТ ПАРАМЕТРОВ СВЕЧЕЙ ===
                st.markdown("---")
                st.subheader("🧮 Расчёт параметров свечей")
                
                # Создаём новую таблицу с расчётами
                result_df = df.copy()
                
                # Расчёт параметров свечи
                result_df['Open'] = result_df['V_капс']
                result_df['Close'] = result_df['V_центр']
                result_df['High'] = result_df[['V_капс', 'V_перег', 'V_центр']].max(axis=1)
                result_df['Low'] = result_df[['V_капс', 'V_перег', 'V_центр']].min(axis=1)
                result_df['Body'] = result_df['Close'] - result_df['Open']
                result_df['Направление'] = result_df['Body'].apply(
                    lambda x: 'Бычья (рост к центру)' if x > 0 else 'Медвежья (падение к центру)'
                )
                result_df['Верхняя_тень'] = result_df['High'] - result_df[['Open', 'Close']].max(axis=1)
                result_df['Нижняя_тень'] = result_df[['Open', 'Close']].min(axis=1) - result_df['Low']
                result_df['Длина_тела'] = abs(result_df['Body'])
                result_df['Цвет'] = result_df['Body'].apply(lambda x: 'Зелёный' if x > 0 else 'Красный')
                
                # Показываем результат
                st.dataframe(result_df)
                
                # === СТАТИСТИКА ===
                st.markdown("---")
                st.subheader("📈 Статистический анализ")
                
                # Разделяем по группам
                poya = result_df[result_df['Диагноз'] == 'ПОЯ']
                rya = result_df[result_df['Диагноз'] == 'РЯ']
                
                # Проверка наличия групп
                if len(poya) == 0 or len(rya) == 0:
                    st.warning("⚠️ В файле должны быть и ПОЯ, и РЯ для сравнения")
                else:
                    col_stat1, col_stat2, col_stat3 = st.columns(3)
                    
                    with col_stat1:
                        st.markdown("**Длина тела свечи (Body)**")
                        
                        # Манн-Уитни для длины тела
                        u_stat, p_body = mannwhitneyu(
                            abs(poya['Body'].dropna()), 
                            abs(rya['Body'].dropna()),
                            alternative='two-sided'
                        )
                        
                        st.metric("ПОЯ (медиана)", f"{poya['Длина_тела'].median():.1f}")
                        st.metric("РЯ (медиана)", f"{rya['Длина_тела'].median():.1f}")
                        
                        if p_body < 0.001:
                            st.success(f"**p < 0.001**")
                        elif p_body < 0.05:
                            st.success(f"**p = {p_body:.3f}**")
                        else:
                            st.warning(f"p = {p_body:.3f}")
                    
                    with col_stat2:
                        st.markdown("**Направление свечей**")
                        
                        # Таблица для критерия Фишера
                        poya_dirs = poya['Направление'].value_counts()
                        rya_dirs = rya['Направление'].value_counts()
                        
                        # Создаём таблицу 2x2
                        poya_bull = poya_dirs.get('Бычья (рост к центру)', 0)
                        poya_bear = poya_dirs.get('Медвежья (падение к центру)', 0)
                        rya_bull = rya_dirs.get('Бычья (рост к центру)', 0)
                        rya_bear = rya_dirs.get('Медвежья (падение к центру)', 0)
                        
                        # Точный критерий Фишера
                        if poya_bull + poya_bear > 0 and rya_bull + rya_bear > 0:
                            _, p_fisher = fisher_exact([[poya_bull, poya_bear], [rya_bull, rya_bear]])
                            
                            st.write(f"ПОЯ: Бычьих {poya_bull}, Медвежьих {poya_bear}")
                            st.write(f"РЯ: Бычьих {rya_bull}, Медвежьих {rya_bear}")
                            
                            if p_fisher < 0.001:
                                st.success(f"**p < 0.001**")
                            elif p_fisher < 0.05:
                                st.success(f"**p = {p_fisher:.3f}**")
                            else:
                                st.warning(f"p = {p_fisher:.3f}")
                    
                    with col_stat3:
                        st.markdown("**Сводная таблица**")
                        summary = pd.DataFrame({
                            'Показатель': ['Количество', 'Медиана Body', 'Бычьи (%)', 'Медвежьи (%)'],
                            'ПОЯ': [
                                len(poya),
                                f"{poya['Длина_тела'].median():.1f}",
                                f"{(poya_bull/len(poya)*100):.0f}%" if len(poya)>0 else "0%",
                                f"{(poya_bear/len(poya)*100):.0f}%" if len(poya)>0 else "0%"
                            ],
                            'РЯ': [
                                len(rya),
                                f"{rya['Длина_тела'].median():.1f}",
                                f"{(rya_bull/len(rya)*100):.0f}%" if len(rya)>0 else "0%",
                                f"{(rya_bear/len(rya)*100):.0f}%" if len(rya)>0 else "0%"
                            ]
                        })
                        st.dataframe(summary)
                    
                    # === ГРАФИКИ ===
                    st.markdown("---")
                    st.subheader("📉 Визуализация: Японские свечи")
                    
                    # Подготовка данных для графика
                    plot_data = result_df.copy()
                    plot_data['Номер'] = range(1, len(plot_data) + 1)
                    
                    # Создаём график
                    fig = make_subplots(
                        rows=2, cols=1,
                        subplot_titles=('Японские свечи (все пациенты)', 'Box plot: Длина тела свечи'),
                        vertical_spacing=0.15
                    )
                    
                    # Добавляем свечи
                    for i, row in plot_data.iterrows():
                        color = 'green' if row['Body'] > 0 else 'red'
                        
                        # Тело свечи
                        fig.add_trace(
                            go.Scatter(
                                x=[row['Номер'], row['Номер']],
                                y=[row['Open'], row['Close']],
                                mode='lines',
                                line=dict(color=color, width=4),
                                showlegend=False,
                                hovertext=f"{row['ID']}<br>Open: {row['Open']}<br>Close: {row['Close']}"
                            ),
                            row=1, col=1
                        )
                        
                        # Тени (High-Low)
                        fig.add_trace(
                            go.Scatter(
                                x=[row['Номер'], row['Номер']],
                                y=[row['Low'], row['High']],
                                mode='lines',
                                line=dict(color='gray', width=1),
                                showlegend=False,
                                hovertext=f"High: {row['High']}<br>Low: {row['Low']}"
                            ),
                            row=1, col=1
                        )
                    
                    # Добавляем Box plot
                    fig.add_trace(
                        go.Box(
                            y=poya['Длина_тела'],
                            name='ПОЯ',
                            marker_color='red',
                            boxmean=True
                        ),
                        row=2, col=1
                    )
                    
                    fig.add_trace(
                        go.Box(
                            y=rya['Длина_тела'],
                            name='РЯ',
                            marker_color='green',
                            boxmean=True
                        ),
                        row=2, col=1
                    )
                    
                    # Настройка осей
                    fig.update_xaxes(title_text="Пациенты (по порядку)", row=1, col=1)
                    fig.update_yaxes(title_text="Скорость кровотока (см/с)", row=1, col=1)
                    fig.update_xaxes(title_text="Группа", row=2, col=1)
                    fig.update_yaxes(title_text="Длина тела свечи", row=2, col=1)
                    
                    fig.update_layout(height=800, showlegend=False)
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # === СОХРАНЕНИЕ РЕЗУЛЬТАТОВ ===
                    st.markdown("---")
                    st.subheader("💾 Сохранить результаты")
                    
                    col_save1, col_save2, col_save3 = st.columns(3)
                    
                    # Сохранить Excel
                    with col_save1:
                        # Подготовка Excel для скачивания
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, sheet_name='Свечи', index=False)
                            
                            # Добавляем лист со статистикой
                            stats_df = pd.DataFrame({
                                'Показатель': ['Количество ПОЯ', 'Количество РЯ', 
                                             'Медиана Body ПОЯ', 'Медиана Body РЯ',
                                             'p-value (Mann-Whitney)', 'p-value (Fisher)'],
                                'Значение': [
                                    len(poya), len(rya),
                                    f"{poya['Длина_тела'].median():.1f}",
                                    f"{rya['Длина_тела'].median():.1f}",
                                    f"{p_body:.4f}" if 'p_body' in locals() else '-',
                                    f"{p_fisher:.4f}" if 'p_fisher' in locals() else '-'
                                ]
                            })
                            stats_df.to_excel(writer, sheet_name='Статистика', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Скачать Excel (все расчёты)",
                            data=output,
                            file_name=f"uzi_candles_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key='download_excel'
                        )
                    
                    # Сохранить график PNG
                    with col_save2:
                        # Сохраняем график как PNG
                        img_bytes = fig.to_image(format="png", width=1200, height=800)
                        
                        st.download_button(
                            label="📥 Скачать график (PNG)",
                            data=img_bytes,
                            file_name=f"uzi_candles_plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                            mime="image/png",
                            key='download_png'
                        )
                    
                    # Копировать в буфер (текстовый отчёт)
                    with col_save3:
                        report = f"""УЗИ КАЛЬКУЛЯТОР - РЕЗУЛЬТАТЫ
Дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}

ВСЕГО ПАЦИЕНТОВ: {len(result_df)}
ПОЯ: {len(poya)}
РЯ: {len(rya)}

СТАТИСТИКА:
- Медиана длины тела (ПОЯ): {poya['Длина_тела'].median():.1f}
- Медиана длины тела (РЯ): {rya['Длина_тела'].median():.1f}
- p-value (Mann-Whitney): {p_body:.4f}

НАПРАВЛЕНИЯ СВЕЧЕЙ:
ПОЯ: Бычьих {poya_bull}, Медвежьих {poya_bear}
РЯ: Бычьих {rya_bull}, Медвежьих {rya_bear}
p-value (Fisher): {p_fisher:.4f}
"""
                        st.download_button(
                            label="📥 Скачать отчёт (TXT)",
                            data=report,
                            file_name=f"uzi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                            key='download_txt'
                        )
                    
                    # Инструкция по результатам
                    with st.expander("📊 Как интерпретировать результаты"):
                        st.markdown("""
                        ### Интерпретация
                        
                        **Зелёные свечи (растущие)** = Close > Open  
                        → скорость кровотока **увеличивается** к центру  
                        → характерно для **РЯ** (рак)
                        
                        **Красные свечи (падающие)** = Close < Open  
                        → скорость кровотока **уменьшается** к центру  
                        → характерно для **ПОЯ** (пограничные)
                        
                        **Статистика:**
                        - **p < 0.05** = различия значимы
                        - **p < 0.001** = различия высоко значимы
                        - **p > 0.05** = различий не обнаружено
                        """)
        
        except Exception as e:
            st.error(f"❌ Ошибка при обработке файла: {str(e)}")
            st.info("Проверьте формат файла. Должны быть колонки: ID, Диагноз, V_капс, V_перег, V_центр")
    
    else:
        st.info("👆 Загрузите файл Excel, чтобы начать расчёт")
        
        # Пример данных
        with st.expander("📋 Пример правильного файла"):
            example = pd.DataFrame({
                'ID': ['ПОЯ-001', 'ПОЯ-002', 'РЯ-001', 'РЯ-002'],
                'Диагноз': ['ПОЯ', 'ПОЯ', 'РЯ', 'РЯ'],
                'V_капс': [35, 38, 32, 34],
                'V_перег': [12, 10, 35, 33],
                'V_центр': [15, 18, 41, 40]
            })
            st.dataframe(example)

# Подвал
st.markdown("---")
st.markdown("""
**📌 Для диссертации:**
- Вкладка 1: ручной расчёт критериев Фишера и Манна-Уитни
- Вкладка 2: автоматический анализ гемодинамики методом японских свечей
""")

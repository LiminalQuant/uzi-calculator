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
                    fig.add_trace(go.Box(y=g1, name=group1_name_mw, boxmean='sd', marker_color='#e74c3c'))
                    fig.add_trace(go.Box(y=g2, name=group2_name_mw, boxmean='sd', marker_color='#2ecc71'))
                    fig.update_layout(
                        title=f"Сравнение: {feature_name_mw}",
                        yaxis_title=feature_name_mw,
                        template='plotly_white'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Описательная статистика
                    st.markdown("---")
                    st.subheader("📊 Описательная статистика")
                    
                    stats_df = pd.DataFrame({
                        'Показатель': ['Количество', 'Среднее', 'Медиана', 'Стд. отклонение', 'Мин', 'Макс'],
                        group1_name_mw: [len(g1), f"{np.mean(g1):.1f}", f"{np.median(g1):.1f}", f"{np.std(g1):.1f}", f"{np.min(g1):.1f}", f"{np.max(g1):.1f}"],
                        group2_name_mw: [len(g2), f"{np.mean(g2):.1f}", f"{np.median(g2):.1f}", f"{np.std(g2):.1f}", f"{np.min(g2):.1f}", f"{np.max(g2):.1f}"]
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
# ВКЛАДКА 2: ЯПОНСКИЕ СВЕЧИ (ПРАВИЛЬНЫЕ!)
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
        1. **ID** - любой текст
        2. **Диагноз** - строго **ПОЯ** или **РЯ**
        3. **V_капс** - скорость в капсуле (периферия)
        4. **V_перег** - скорость в перегородках
        5. **V_центр** - скорость в центре
        """)
    
    # Загрузка файла
    uploaded_file = st.file_uploader(
        "📂 Загрузите файл Excel", 
        type=['xlsx', 'xls'],
        key='candle_upload'
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
                st.success("✅ Файл загружен")
                
                # Показываем первые строки
                st.subheader("📋 Загруженные данные")
                st.dataframe(df.head(10))
                
                # Статистика по группам
                poya_count = len(df[df['Диагноз'] == 'ПОЯ'])
                rya_count = len(df[df['Диагноз'] == 'РЯ'])
                
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"📊 ПОЯ: {poya_count} пациентов")
                with col2:
                    st.info(f"📊 РЯ: {rya_count} пациентов")
                
                if poya_count == 0 or rya_count == 0:
                    st.warning("⚠️ В файле должны быть и ПОЯ, и РЯ для сравнения")
                else:
                    # === РАСЧЁТ ПАРАМЕТРОВ СВЕЧЕЙ ===
                    st.markdown("---")
                    st.subheader("🧮 Расчёт параметров свечей")
                    
                    result_df = df.copy()
                    
                    # Расчёт параметров
                    result_df['Open'] = result_df['V_капс']
                    result_df['Close'] = result_df['V_центр']
                    result_df['High'] = result_df[['V_капс', 'V_перег', 'V_центр']].max(axis=1)
                    result_df['Low'] = result_df[['V_капс', 'V_перег', 'V_центр']].min(axis=1)
                    result_df['Body'] = result_df['Close'] - result_df['Open']
                    result_df['Направление'] = result_df['Body'].apply(
                        lambda x: 'Бычья (рост к центру)' if x > 0 else 'Медвежья (падение к центру)'
                    )
                    result_df['Длина_тела'] = abs(result_df['Body'])
                    result_df['Цвет'] = result_df['Body'].apply(lambda x: 'Зелёный' if x > 0 else 'Красный')
                    
                    # Показываем результат
                    st.dataframe(result_df)
                    
                    # === СТАТИСТИКА ===
                    st.markdown("---")
                    st.subheader("📈 Статистический анализ")
                    
                    poya = result_df[result_df['Диагноз'] == 'ПОЯ']
                    rya = result_df[result_df['Диагноз'] == 'РЯ']
                    
                    # Манн-Уитни для длины тела
                    u_stat, p_body = mannwhitneyu(
                        poya['Длина_тела'].dropna(), 
                        rya['Длина_тела'].dropna(),
                        alternative='two-sided'
                    )
                    
                    # Таблица для критерия Фишера (направление свечей)
                    poya_bull = len(poya[poya['Body'] > 0])
                    poya_bear = len(poya[poya['Body'] < 0])
                    rya_bull = len(rya[rya['Body'] > 0])
                    rya_bear = len(rya[rya['Body'] < 0])
                    
                    _, p_fisher = fisher_exact([[poya_bull, poya_bear], [rya_bull, rya_bear]])
                    
                    # Три колонки со статистикой
                    col_stat1, col_stat2, col_stat3 = st.columns(3)
                    
                    with col_stat1:
                        st.markdown("### 📊 Длина тела")
                        st.metric("ПОЯ (медиана)", f"{poya['Длина_тела'].median():.1f}")
                        st.metric("РЯ (медиана)", f"{rya['Длина_тела'].median():.1f}")
                        if p_body < 0.001:
                            st.success("**p < 0.001**")
                        else:
                            st.success(f"**p = {p_body:.3f}**")
                    
                    with col_stat2:
                        st.markdown("### 🎯 Направление свечей")
                        st.write(f"ПОЯ: Бычьих {poya_bull}, Медвежьих {poya_bear}")
                        st.write(f"РЯ: Бычьих {rya_bull}, Медвежьих {rya_bear}")
                        if p_fisher < 0.001:
                            st.success("**p < 0.001**")
                        else:
                            st.success(f"**p = {p_fisher:.3f}**")
                    
                    with col_stat3:
                        st.markdown("### 📋 Сводка")
                        st.write(f"Всего пациентов: {len(result_df)}")
                        st.write(f"ПОЯ: {len(poya)} ({len(poya)/len(result_df)*100:.0f}%)")
                        st.write(f"РЯ: {len(rya)} ({len(rya)/len(result_df)*100:.0f}%)")
                    
                    # === ПРАВИЛЬНЫЕ ЯПОНСКИЕ СВЕЧИ ===
                    st.markdown("---")
                    st.subheader("📈 Японские свечи")
                    st.markdown("*Наведите на свечу для деталей • Нажмите на иконку камеры, чтобы скачать PNG*")
                    
                    # Подготовка данных
                    plot_df = result_df.copy()
                    plot_df['x'] = range(len(plot_df))
                    
                    # Создаём график
                    fig = go.Figure()
                    
                    # Рисуем каждую свечу
                    for _, row in plot_df.iterrows():
                        if row['Body'] > 0:  # Бычья свеча (зелёная)
                            body_color = '#2ecc71'
                            body_low = row['Open']
                            body_high = row['Close']
                        else:  # Медвежья свеча (красная)
                            body_color = '#e74c3c'
                            body_low = row['Close']
                            body_high = row['Open']
                        
                        # Фитиль (High-Low) - тонкая чёрная линия
                        fig.add_trace(go.Scatter(
                            x=[row['x'], row['x']],
                            y=[row['Low'], row['High']],
                            mode='lines',
                            line=dict(color='#000000', width=1),
                            showlegend=False,
                            hoverinfo='skip'
                        ))
                        
                        # Тело свечи (Open-Close) - толстый прямоугольник
                        fig.add_trace(go.Scatter(
                            x=[row['x'], row['x']],
                            y=[body_low, body_high],
                            mode='lines',
                            line=dict(color=body_color, width=12),
                            name=row['ID'],
                            text=f"ID: {row['ID']}<br>"
                                 f"Диагноз: {row['Диагноз']}<br>"
                                 f"Open (капсула): {row['Open']:.1f}<br>"
                                 f"Close (центр): {row['Close']:.1f}<br>"
                                 f"High: {row['High']:.1f}<br>"
                                 f"Low: {row['Low']:.1f}<br>"
                                 f"Body: {row['Body']:.1f}<br>"
                                 f"{'↑ Рост к центру' if row['Body'] > 0 else '↓ Падение к центру'}",
                            hoverinfo='text',
                            showlegend=False
                        ))
                    
                    # Разделительная линия между группами
                    fig.add_vline(x=poya_count - 0.5, line_dash="dash", line_color="gray", line_width=2)
                    
                    # Подписи групп
                    fig.add_annotation(
                        x=poya_count/2,
                        y=plot_df['High'].max() * 1.1,
                        text="ПОЯ",
                        showarrow=False,
                        font=dict(size=18, color="#e74c3c", family="Arial Black")
                    )
                    fig.add_annotation(
                        x=poya_count + rya_count/2,
                        y=plot_df['High'].max() * 1.1,
                        text="РЯ",
                        showarrow=False,
                        font=dict(size=18, color="#2ecc71", family="Arial Black")
                    )
                    
                    # Настройка графика
                    fig.update_layout(
                        title=dict(
                            text="Японские свечи: гемодинамический профиль (периферия → центр)",
                            font=dict(size=18)
                        ),
                        xaxis_title="Пациенты",
                        yaxis_title="Скорость кровотока (см/с)",
                        template="plotly_white",
                        height=600,
                        hovermode='closest',
                        showlegend=False
                    )
                    
                    fig.update_xaxes(tickmode='linear', tick0=0, dtick=2, gridcolor='lightgray')
                    fig.update_yaxes(gridcolor='lightgray')
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # === BOX PLOT ДЛЯ ДЛИНЫ ТЕЛА ===
                    st.subheader("📊 Сравнение длины тела свечей")
                    st.markdown("*Нажмите на иконку камеры, чтобы скачать график*")
                    
                    fig2 = go.Figure()
                    
                    fig2.add_trace(go.Box(
                        y=poya['Длина_тела'],
                        name='ПОЯ',
                        marker_color='#e74c3c',
                        boxmean='sd',
                        boxpoints='all',
                        jitter=0.3,
                        pointpos=-1.8,
                        line=dict(color='#c0392b', width=2),
                        fillcolor='rgba(231,76,60,0.3)'
                    ))
                    
                    fig2.add_trace(go.Box(
                        y=rya['Длина_тела'],
                        name='РЯ',
                        marker_color='#2ecc71',
                        boxmean='sd',
                        boxpoints='all',
                        jitter=0.3,
                        pointpos=-1.8,
                        line=dict(color='#27ae60', width=2),
                        fillcolor='rgba(46,204,113,0.3)'
                    ))
                    
                    fig2.update_layout(
                        title=dict(
                            text="Распределение длины тела свечи |Close - Open|",
                            font=dict(size=16)
                        ),
                        yaxis_title="Длина тела (см/с)",
                        template="plotly_white",
                        height=500,
                        showlegend=False
                    )
                    
                    st.plotly_chart(fig2, use_container_width=True)
                    
                    # === СОХРАНЕНИЕ РЕЗУЛЬТАТОВ ===
                    st.markdown("---")
                    st.subheader("💾 Сохранить результаты")
                    
                    col_save1, col_save2 = st.columns(2)
                    
                    # Excel
                    with col_save1:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, sheet_name='Свечи', index=False)
                            
                            stats_df = pd.DataFrame({
                                'Показатель': [
                                    'ПОЯ (n)', 'РЯ (n)', 
                                    'Медиана Body ПОЯ', 'Медиана Body РЯ',
                                    'Среднее Body ПОЯ', 'Среднее Body РЯ',
                                    'Бычьих ПОЯ', 'Медвежьих ПОЯ',
                                    'Бычьих РЯ', 'Медвежьих РЯ',
                                    'p-value (Mann-Whitney)', 'p-value (Fisher)'
                                ],
                                'Значение': [
                                    len(poya), len(rya),
                                    f"{poya['Длина_тела'].median():.1f}",
                                    f"{rya['Длина_тела'].median():.1f}",
                                    f"{poya['Длина_тела'].mean():.1f}",
                                    f"{rya['Длина_тела'].mean():.1f}",
                                    poya_bull, poya_bear,
                                    rya_bull, rya_bear,
                                    f"{p_body:.4f}",
                                    f"{p_fisher:.4f}"
                                ]
                            })
                            stats_df.to_excel(writer, sheet_name='Статистика', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Скачать Excel (все расчёты)",
                            data=output,
                            file_name=f"uzi_candles_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # TXT отчёт
                    with col_save2:
                        report = f"""УЗИ КАЛЬКУЛЯТОР - ЯПОНСКИЕ СВЕЧИ
================================
Дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}
Файл: {uploaded_file.name}

ОБЩАЯ ИНФОРМАЦИЯ
================================
Всего пациентов: {len(result_df)}
ПОЯ: {len(poya)} ({len(poya)/len(result_df)*100:.0f}%)
РЯ: {len(rya)} ({len(rya)/len(result_df)*100:.0f}%)

СТАТИСТИКА ПО ДЛИНЕ ТЕЛА СВЕЧИ
================================
ПОЯ (медиана): {poya['Длина_тела'].median():.1f}
ПОЯ (среднее): {poya['Длина_тела'].mean():.1f}
ПОЯ (мин-макс): {poya['Длина_тела'].min():.1f} - {poya['Длина_тела'].max():.1f}

РЯ (медиана): {rya['Длина_тела'].median():.1f}
РЯ (среднее): {rya['Длина_тела'].mean():.1f}
РЯ (мин-макс): {rya['Длина_тела'].min():.1f} - {rya['Длина_тела'].max():.1f}

p-value (Mann-Whitney): {p_body:.4f}

НАПРАВЛЕНИЕ СВЕЧЕЙ
================================
ПОЯ: Бычьих {poya_bull}, Медвежьих {poya_bear}
РЯ: Бычьих {rya_bull}, Медвежьих {rya_bear}

p-value (точный критерий Фишера): {p_fisher:.4f}

ИНТЕРПРЕТАЦИЯ
================================
{'✅ Различия в длине тела статистически значимы (p < 0.05)' if p_body < 0.05 else '❌ Различия в длине тела не доказаны (p ≥ 0.05)'}
{'✅ Различия в направлении свечей значимы (p < 0.05)' if p_fisher < 0.05 else '❌ Различия в направлении свечей не доказаны (p ≥ 0.05)'}
"""
                        
                        st.download_button(
                            label="📥 Скачать отчёт (TXT)",
                            data=report,
                            file_name=f"uzi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                        )
                    
                    # Интерпретация
                    with st.expander("📊 Как интерпретировать результаты"):
                        st.markdown("""
                        ### Как читать график японских свечей
                        
                        **Каждая свеча — одна пациентка:**
                        - **Тонкая чёрная линия (фитиль)** — разброс скоростей (High - Low)
                        - **Толстое тело** — изменение скорости от периферии к центру
                        
                        **Цвет тела:**
                        - **🟢 Зелёное тело** (Close > Open) — скорость растёт к центру → **РЯ**
                        - **🔴 Красное тело** (Close < Open) — скорость падает к центру → **ПОЯ**
                        
                        **Длина тела:** |Close - Open| — насколько сильно изменилась скорость
                        
                        **Статистика:**
                        - **p < 0.05** — различия статистически значимы
                        - **p < 0.001** — различия высоко значимы
                        
                        **Как сохранить графики:**
                        1. Наведите на график
                        2. Нажмите на иконку **камеры** в правом верхнем углу
                        3. Выберите формат PNG
                        """)
        
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            st.info("Проверьте формат файла. Должны быть колонки: ID, Диагноз, V_капс, V_перег, V_центр")

# Подвал
st.markdown("---")
st.markdown("**📌 Для диссертации: статистический калькулятор + визуализация японскими свечами**")

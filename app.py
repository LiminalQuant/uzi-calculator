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
st.title("📊 УЗИ Статистический калькулятор")
st.markdown("---")

# ============================================
# ВКЛАДКА 1: РУЧНОЙ ВВОД (ФИШЕР/МАНН-УИТНИ)
# ============================================
tab1, tab2 = st.tabs(["📊 Ручной ввод (Фишер/Манн-Уитни)", "📈 Загрузка Excel (Японские свечи)"])

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
        
        # Инициализация истории
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
            if a < 0 or b < 0 or c < 0 or d < 0:
                st.error("❌ Все значения должны быть неотрицательными!")
            else:
                table = [[a, b], [c, d]]
                odds_ratio, p_value = fisher_exact(table, alternative=alternative)
                
                timestamp = datetime.now().strftime("%H:%M:%S")
                st.session_state.history.append({
                    'time': timestamp,
                    'type': 'Критерий Фишера',
                    'result': f"{feature_name}: {group1_name} vs {group2_name}",
                    'p_value': f"{p_value:.4f}"
                })
                
                st.markdown("---")
                st.subheader("📊 Результаты")
                
                col_res1, col_res2, col_res3 = st.columns(3)
                
                with col_res1:
                    st.metric("Отношение шансов (OR)", f"{odds_ratio:.3f}")
                    if odds_ratio > 1:
                        st.caption(f"Шанс признака выше в группе {group1_name}")
                    elif odds_ratio < 1:
                        st.caption(f"Шанс признака выше в группе {group2_name}")
                    else:
                        st.caption("Шансы равны")
                
                with col_res2:
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
                        st.write(f"Различия между {group1_name} и {group2_name} по признаку **{feature_name}** статистически значимы (p < 0.05)")
                    else:
                        st.warning("⚠️ **Статистически не значимо**")
                        st.write(f"Различия между {group1_name} и {group2_name} по признаку **{feature_name}** не доказаны (p ≥ 0.05)")
                
                st.markdown("---")
                st.subheader("📋 Таблица сопряженности")
                
                df_result = pd.DataFrame(
                    table,
                    index=[group1_name, group2_name],
                    columns=["Есть признак", "Нет признака"]
                )
                df_result.loc['Итого'] = df_result.sum()
                st.dataframe(df_result, use_container_width=True)
                
                result_text = f"{feature_name}: {group1_name} vs {group2_name}\nOR = {odds_ratio:.3f}, p = {p_value:.4f}\nТаблица: {a} {b} | {c} {d}"
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
                g1 = [float(x.strip()) for x in group1_data.replace(',', ' ').split()]
                g2 = [float(x.strip()) for x in group2_data.replace(',', ' ').split()]
                
                if len(g1) < 2 or len(g2) < 2:
                    st.error("❌ В каждой группе должно быть минимум 2 значения")
                else:
                    statistic, p_value = mannwhitneyu(g1, g2, alternative='two-sided')
                    
                    timestamp = datetime.now().strftime("%H:%M:%S")
                    st.session_state.history.append({
                        'time': timestamp,
                        'type': 'Манн-Уитни',
                        'result': f"{feature_name_mw}: {group1_name_mw} vs {group2_name_mw}",
                        'p_value': f"{p_value:.4f}"
                    })
                    
                    st.markdown("---")
                    st.subheader("📊 Результаты")
                    
                    col_res1, col_res2, col_res3 = st.columns(3)
                    
                    with col_res1:
                        st.metric("Статистика U", f"{statistic:.1f}")
                        st.metric(f"Медиана {group1_name_mw}", f"{np.median(g1):.1f}")
                        st.metric(f"Медиана {group2_name_mw}", f"{np.median(g2):.1f}")
                    
                    with col_res2:
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
                    
                    st.markdown("---")
                    st.subheader("📊 Описательная статистика")
                    
                    stats_df = pd.DataFrame({
                        'Показатель': ['Количество', 'Среднее', 'Медиана', 'Стд. отклонение', 'Мин', 'Макс'],
                        group1_name_mw: [len(g1), f"{np.mean(g1):.1f}", f"{np.median(g1):.1f}", f"{np.std(g1):.1f}", f"{np.min(g1):.1f}", f"{np.max(g1):.1f}"],
                        group2_name_mw: [len(g2), f"{np.mean(g2):.1f}", f"{np.median(g2):.1f}", f"{np.std(g2):.1f}", f"{np.min(g2):.1f}", f"{np.max(g2):.1f}"]
                    })
                    
                    st.dataframe(stats_df, use_container_width=True)
                    
                    result_text = f"{feature_name_mw}: {group1_name_mw} vs {group2_name_mw}\nU = {statistic:.1f}, p = {p_value:.4f}\nМедианы: {np.median(g1):.1f} vs {np.median(g2):.1f}"
                    st.code(result_text, language="text")
                    
            except Exception as e:
                st.error(f"❌ Ошибка при обработке данных: {e}")
                st.info("Убедитесь, что значения введены правильно (числа через запятую или пробел)")

# ============================================
# ВКЛАДКА 2: ЗАГРУЗКА EXCEL - ЯПОНСКИЕ СВЕЧИ (ИСПРАВЛЕНО)
# ============================================
with tab2:
    st.header("📈 Японские свечи: анализ гемодинамики")
    st.markdown("---")
    
    with st.expander("📖 ИНСТРУКЦИЯ: Формат Excel файла", expanded=True):
        st.markdown("""
        ### Формат данных
        
        **Excel файл должен содержать 2 листа:**
        - **Лист 1:** Рак яичников (название любое)
        - **Лист 2:** Пограничные опухоли (название любое)
        
        **Структура каждого листа:**
        |   | Капсула |   |   | Центр |   |   |
        |---|---------|---|---|-------|---|---|
        |   | МСС | ИР | ПИ | МСС | ИР | ПИ |
        | 1 | 18.07 | 0.48 | 0.66 | 29.03 | 0.4 | 0.55 |
        
        **Программа автоматически:**
        1. Определит бычьи 🟢 (рост) и медвежьи 🔴 (падение) свечи
        2. Построит графики для МСС, ИР и ПИ
        3. Рассчитает статистику (Манн-Уитни + Фишер)
        """)
    
    uploaded_file = st.file_uploader(
        "📂 Загрузите Excel файл (2 листа)", 
        type=['xlsx', 'xls'],
        key='excel_upload'
    )
    
    if uploaded_file:
        try:
            # Читаем листы
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.success(f"✅ Найдено листов: {len(sheet_names)}")
            st.info(f"Листы: {', '.join(sheet_names)}")
            
            # Определяем листы (первые два)
            cancer_sheet = sheet_names[0]
            border_sheet = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
            
            # Читаем данные (пропускаем 2 строки заголовков)
            df_cancer = pd.read_excel(uploaded_file, sheet_name=cancer_sheet, header=None, skiprows=2)
            df_border = pd.read_excel(uploaded_file, sheet_name=border_sheet, header=None, skiprows=2)
            
            # Проверяем и удаляем лишний первый столбец (индекс)
            if df_cancer.shape[1] > 6:
                df_cancer = df_cancer.iloc[:, 1:7]
            if df_border.shape[1] > 6:
                df_border = df_border.iloc[:, 1:7]
            
            # Назначаем имена колонок
            col_names = ['МСС_капс', 'ИР_капс', 'ПИ_капс', 'МСС_центр', 'ИР_центр', 'ПИ_центр']
            df_cancer.columns = col_names
            df_border.columns = col_names
            
            # Конвертируем в числа
            for col in col_names:
                df_cancer[col] = pd.to_numeric(df_cancer[col], errors='coerce')
                df_border[col] = pd.to_numeric(df_border[col], errors='coerce')
            
            # Удаляем строки с NaN
            df_cancer = df_cancer.dropna()
            before = len(df_border)
            df_border = df_border.dropna()
            after = len(df_border)
            
            if before > after:
                st.warning(f"⚠️ В ПОЯ исключено {before - after} пациентов с пропусками")
            
            # Добавляем диагноз
            df_cancer['Диагноз'] = 'РЯ'
            df_border['Диагноз'] = 'ПОЯ'
            
            # Объединяем
            df_all = pd.concat([df_cancer, df_border], ignore_index=True)
            df_all['ID'] = range(1, len(df_all) + 1)
            
            # Показываем статистику
            st.markdown("---")
            st.subheader("📊 Загруженные данные")
            
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"**РЯ:** {len(df_cancer)} пациентов")
                st.dataframe(df_cancer.head(3))
            with col2:
                st.info(f"**ПОЯ:** {len(df_border)} пациентов")
                st.dataframe(df_border.head(3))
            
            # Выбор параметра
            st.markdown("---")
            st.subheader("🎯 Выберите параметр для анализа")
            
            param_choice = st.radio(
                "Какой показатель анализировать?",
                options=["📈 МСС (скорость)", "📊 ИР (индекс резистентности)", "📉 ПИ (пульсационный индекс)"],
                horizontal=True
            )
            
            # Настройка параметров
            if param_choice == "📈 МСС (скорость)":
                cap_col = 'МСС_капс'
                cent_col = 'МСС_центр'
                param_name = "МСС"
                y_title = "Скорость кровотока (см/с)"
            elif param_choice == "📊 ИР (индекс резистентности)":
                cap_col = 'ИР_капс'
                cent_col = 'ИР_центр'
                param_name = "ИР"
                y_title = "Индекс резистентности"
            else:
                cap_col = 'ПИ_капс'
                cent_col = 'ПИ_центр'
                param_name = "ПИ"
                y_title = "Пульсационный индекс"
            
            # === РАСЧЁТ ПАРАМЕТРОВ СВЕЧЕЙ ===
            df_plot = df_all.copy()
            df_plot['Open'] = df_plot[cap_col]
            df_plot['Close'] = df_plot[cent_col]
            df_plot['High'] = df_plot[[cap_col, cent_col]].max(axis=1)
            df_plot['Low'] = df_plot[[cap_col, cent_col]].min(axis=1)
            df_plot['Body'] = df_plot['Close'] - df_plot['Open']
            df_plot['Длина_тела'] = abs(df_plot['Body'])
            
            # ПРАВИЛЬНОЕ ОПРЕДЕЛЕНИЕ:
            # Быки 🟢 = рост = Close > Open
            # Медведи 🔴 = падение = Close < Open
            df_plot['Тип'] = df_plot['Body'].apply(
                lambda x: 'БЫЧЬЯ 🟢 (рост к центру)' if x > 0 else 'МЕДВЕЖЬЯ 🔴 (падение к центру)'
            )
            df_plot['Цвет'] = df_plot['Body'].apply(lambda x: 'Зелёная' if x > 0 else 'Красная')
            
            # Показываем рассчитанные параметры
            st.markdown("---")
            st.subheader(f"🧮 Рассчитанные параметры для {param_name}")
            display_cols = ['ID', 'Диагноз', 'Open', 'Close', 'High', 'Low', 'Body', 'Тип']
            
            # Стилизация
            def color_type(val):
                if 'БЫЧЬЯ' in str(val):
                    return 'color: green; font-weight: bold'
                elif 'МЕДВЕЖЬЯ' in str(val):
                    return 'color: red; font-weight: bold'
                return ''
            
            styled_df = df_plot[display_cols].style.applymap(color_type, subset=['Тип'])
            st.dataframe(styled_df, use_container_width=True)
            
            # === СТАТИСТИКА ===
            st.markdown("---")
            st.subheader(f"📊 Статистический анализ для {param_name}")
            
            poya = df_plot[df_plot['Диагноз'] == 'ПОЯ']
            rya = df_plot[df_plot['Диагноз'] == 'РЯ']
            
            # Манн-Уитни для длины тела
            u_stat, p_body = mannwhitneyu(
                poya['Длина_тела'].dropna(),
                rya['Длина_тела'].dropna(),
                alternative='two-sided'
            )
            
            # Таблица для Фишера (направление)
            poya_bull = len(poya[poya['Body'] > 0])   # Быки в ПОЯ (рост)
            poya_bear = len(poya[poya['Body'] < 0])   # Медведи в ПОЯ (падение)
            rya_bull = len(rya[rya['Body'] > 0])      # Быки в РЯ (рост)
            rya_bear = len(rya[rya['Body'] < 0])      # Медведи в РЯ (падение)
            
            _, p_fisher = fisher_exact([[poya_bull, poya_bear], [rya_bull, rya_bear]])
            
            col_s1, col_s2, col_s3 = st.columns(3)
            
            with col_s1:
                st.markdown("**📏 Длина тела свечи**")
                st.metric("ПОЯ (медиана)", f"{poya['Длина_тела'].median():.3f}")
                st.metric("РЯ (медиана)", f"{rya['Длина_тела'].median():.3f}")
                if p_body < 0.001:
                    st.success("**p < 0.001**")
                else:
                    st.success(f"**p = {p_body:.3f}**")
            
            with col_s2:
                st.markdown("**🎯 Направление свечей**")
                st.write(f"**ПОЯ:** Бычьих 🟢 {poya_bull}, Медвежьих 🔴 {poya_bear}")
                st.write(f"**РЯ:** Бычьих 🟢 {rya_bull}, Медвежьих 🔴 {rya_bear}")
                if p_fisher < 0.001:
                    st.success("**p < 0.001**")
                else:
                    st.success(f"**p = {p_fisher:.3f}**")
            
            with col_s3:
                st.markdown("**📋 Сводка**")
                st.write(f"ПОЯ: {len(poya)} пациентов")
                st.write(f"РЯ: {len(rya)} пациентов")
                st.write(f"Всего: {len(df_plot)} пациентов")
            
            # === ГРАФИК 1: ЯПОНСКИЕ СВЕЧИ ===
            st.markdown("---")
            st.subheader(f"📈 Японские свечи для {param_name}")
            st.markdown("*🟢 Бычьи (рост к центру) = РЯ | 🔴 Медвежьи (падение к центру) = ПОЯ*")
            
            fig1 = go.Figure()
            
            plot_df = df_plot.copy()
            plot_df['x'] = range(len(plot_df))
            
            # Разделительная линия
            poya_count = len(poya)
            fig1.add_vline(x=poya_count - 0.5, line_dash="dash", line_color="gray", line_width=2)
            
            # Подписи групп (ИСПРАВЛЕНО)
            fig1.add_annotation(
                x=poya_count/2,
                y=plot_df['High'].max() * 1.1,
                text="🔴 ПОЯ (МЕДВЕДИ)",
                showarrow=False,
                font=dict(size=16, color="#e74c3c", family="Arial Black")
            )
            fig1.add_annotation(
                x=poya_count + len(rya)/2,
                y=plot_df['High'].max() * 1.1,
                text="🟢 РЯ (БЫКИ)",
                showarrow=False,
                font=dict(size=16, color="#2ecc71", family="Arial Black")
            )
            
            # Рисуем свечи
            for _, row in plot_df.iterrows():
                if row['Body'] > 0:  # БЫЧЬЯ (рост)
                    body_color = '#2ecc71'
                    body_low = row['Open']
                    body_high = row['Close']
                    trend_text = '🟢 Рост к центру (БЫЧЬЯ)'
                else:  # МЕДВЕЖЬЯ (падение)
                    body_color = '#e74c3c'
                    body_low = row['Close']
                    body_high = row['Open']
                    trend_text = '🔴 Падение к центру (МЕДВЕЖЬЯ)'
                
                # Фитиль
                fig1.add_trace(go.Scatter(
                    x=[row['x'], row['x']],
                    y=[row['Low'], row['High']],
                    mode='lines',
                    line=dict(color='black', width=1),
                    showlegend=False,
                    hoverinfo='skip'
                ))
                
                # Тело
                fig1.add_trace(go.Scatter(
                    x=[row['x'], row['x']],
                    y=[body_low, body_high],
                    mode='lines',
                    line=dict(color=body_color, width=12),
                    name=row['ID'],
                    text=f"ID: {row['ID']}<br>"
                         f"Диагноз: {row['Диагноз']}<br>"
                         f"Open: {row['Open']:.3f}<br>"
                         f"Close: {row['Close']:.3f}<br>"
                         f"Body: {row['Body']:.3f}<br>"
                         f"{trend_text}",
                    hoverinfo='text',
                    showlegend=False
                ))
            
            fig1.update_layout(
                title=f"{param_name}: периферия → центр",
                xaxis_title="Пациенты",
                yaxis_title=y_title,
                template="plotly_white",
                height=600,
                showlegend=False
            )
            
            st.plotly_chart(fig1, use_container_width=True)
            
            # === ГРАФИК 2: BOX PLOT ===
            st.subheader(f"📊 Сравнение |Body| для {param_name}")
            
            fig2 = go.Figure()
            
            fig2.add_trace(go.Box(
                y=poya['Длина_тела'],
                name='🔴 ПОЯ (Медведи)',
                marker_color='#e74c3c',
                boxmean='sd',
                boxpoints='all',
                jitter=0.3,
                pointpos=-1.8
            ))
            
            fig2.add_trace(go.Box(
                y=rya['Длина_тела'],
                name='🟢 РЯ (Быки)',
                marker_color='#2ecc71',
                boxmean='sd',
                boxpoints='all',
                jitter=0.3,
                pointpos=-1.8
            ))
            
            fig2.update_layout(
                title=f"Распределение |{param_name}_центр - {param_name}_капсула|",
                yaxis_title=f"|Δ{param_name}|",
                template="plotly_white",
                height=500
            )
            
            st.plotly_chart(fig2, use_container_width=True)
            
            # === ТАБЛИЦА НАПРАВЛЕНИЙ ===
            st.subheader("📋 Таблица направлений свечей")
            
            dir_table = pd.DataFrame({
                'Группа': ['ПОЯ (Медведи 🔴)', 'РЯ (Быки 🟢)', 'p-value (Fisher)'],
                'Бычьи 🟢 (рост)': [poya_bull, rya_bull, f"{p_fisher:.4f}"],
                'Медвежьи 🔴 (падение)': [poya_bear, rya_bear, ""],
                'Всего': [len(poya), len(rya), ""]
            })
            
            st.dataframe(dir_table, use_container_width=True)
            
            # === СОХРАНЕНИЕ ===
            st.markdown("---")
            st.subheader("💾 Сохранить результаты")
            
            # Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_plot.to_excel(writer, sheet_name=f'Свечи_{param_name}', index=False)
                
                stats_df = pd.DataFrame({
                    'Показатель': ['Группа', 'n', 'Медиана |Body|', 'Бычьи 🟢', 'Медвежьи 🔴'],
                    'ПОЯ': ['ПОЯ', len(poya), f"{poya['Длина_тела'].median():.3f}", poya_bull, poya_bear],
                    'РЯ': ['РЯ', len(rya), f"{rya['Длина_тела'].median():.3f}", rya_bull, rya_bear],
                    'p-value': ['', '', f"{p_body:.4f}", f"{p_fisher:.4f}", '']
                })
                stats_df.to_excel(writer, sheet_name='Статистика', index=False)
            
            output.seek(0)
            
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.download_button(
                    label=f"📥 Скачать Excel ({param_name})",
                    data=output,
                    file_name=f"uzi_candles_{param_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
            
            with col_d2:
                report = f"""УЗИ КАЛЬКУЛЯТОР - АНАЛИЗ {param_name}
================================
Дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}

ДАННЫЕ:
РЯ (Быки 🟢): {len(rya)} пациентов
ПОЯ (Медведи 🔴): {len(poya)} пациентов

СТАТИСТИКА ДЛИНЫ ТЕЛА:
ПОЯ (медиана |Body|): {poya['Длина_тела'].median():.3f}
РЯ (медиана |Body|): {rya['Длина_тела'].median():.3f}
p-value (Mann-Whitney): {p_body:.4f}

НАПРАВЛЕНИЕ СВЕЧЕЙ:
ПОЯ: Бычьих 🟢 {poya_bull}, Медвежьих 🔴 {poya_bear}
РЯ: Бычьих 🟢 {rya_bull}, Медвежьих 🔴 {rya_bear}
p-value (Fisher): {p_fisher:.4f}

ВЫВОД:
{'✅ Есть статистически значимые различия' if p_body < 0.05 else '❌ Различия не доказаны'}
"""
                
                st.download_button(
                    label="📥 Скачать отчёт (TXT)",
                    data=report,
                    file_name=f"uzi_report_{param_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                )
            
            # Интерпретация
            with st.expander("📊 Интерпретация результатов"):
                st.markdown("""
                ### Как читать график японских свечей
                
                **Правила трейдинга:**
                - **🟢 Быки** = покупают = **цена растёт** = Close > Open
                - **🔴 Медведи** = продают = **цена падает** = Close < Open
                
                **В твоих данных:**
                - **🟢 РЯ (рак)** = скорость растёт к центру = **БЫЧЬЯ свеча**
                - **🔴 ПОЯ (пограничные)** = скорость падает к центру = **МЕДВЕЖЬЯ свеча**
                
                **Статистика:**
                - **p < 0.05** — различия статистически значимы
                - **p < 0.001** — различия высоко значимы
                """)
            
        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            st.exception(e)

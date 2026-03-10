import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import fisher_exact, mannwhitneyu, norm
import plotly.graph_objects as go
from datetime import datetime
import io

# Настройка страницы
st.set_page_config(
    page_title="Медицинский статистический калькулятор",
    page_icon="🏥",
    layout="wide"
)

# Заголовок
st.title("🏥 Медицинский статистический калькулятор")
st.markdown("### ")
st.markdown("---")

# Боковая панель с настройками
with st.sidebar:
    st.header("⚙️ Настройки")
    
    # Выбор типа анализа
    analysis_type = st.radio(
        "Выберите тип данных:",
        options=[
            "📊 Качественные (таблица 2x2) — критерий Фишера",
            "📈 Количественные (две группы) — Манн-Уитни"
        ]
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

# Основная область
if "📊 Качественные" in analysis_type:
    st.header("📊 Точный критерий Фишера")
    st.markdown("""
    **Для каких данных:** таблицы сопряженности 2×2 (наличие/отсутствие признака в двух группах)
    
    **Пример:** сравнение частоты выявления патологии в группе ПОЯ и РЯ
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Группа 1 (например, ПОЯ)")
        a = st.number_input("Есть признак (a)", min_value=0, value=16, step=1, key='a')
        b = st.number_input("Нет признака (b)", min_value=0, value=12, step=1, key='b')
        
    with col2:
        st.subheader("Группа 2 (например, РЯ)")
        c = st.number_input("Есть признак (c)", min_value=0, value=1, step=1, key='c')
        d = st.number_input("Нет признака (d)", min_value=0, value=34, step=1, key='d')
    
    # Дополнительные настройки
    st.markdown("---")
    col3, col4 = st.columns(2)
    
    with col3:
        group1_name = st.text_input("Название группы 1", "ПОЯ")
        group2_name = st.text_input("Название группы 2", "РЯ")
        feature_name = st.text_input("Название признака", "Двухстороннее поражение")
    
    with col4:
        alternative = st.selectbox(
            "Альтернативная гипотеза",
            options=["two-sided", "greater", "less"],
            format_func=lambda x: {
                "two-sided": "Двусторонняя (есть связь)",
                "greater": "Отношение шансов > 1",
                "less": "Отношение шансов < 1"
            }[x]
        )
    
    if st.button("🧮 Рассчитать критерий Фишера", type="primary", use_container_width=True):
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
    st.header("📈 U-критерий Манна-Уитни")
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
            key='group1'
        )
        group1_name_mw = st.text_input("Название группы 1", "ПОЯ", key='g1_name')
    
    with col2:
        st.subheader("Группа 2 (например, РЯ)")
        group2_data = st.text_area(
            "Введите значения через запятую или пробел",
            value="96, 102, 88, 95, 98, 91, 97, 99, 94, 93",
            height=150,
            key='group2'
        )
        group2_name_mw = st.text_input("Название группы 2", "РЯ", key='g2_name')
    
    feature_name_mw = st.text_input("Название показателя", "Максимальный размер, мм")
    
    if st.button("🧮 Рассчитать критерий Манна-Уитни", type="primary", use_container_width=True):
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

# Подвал
st.markdown("---")
st.markdown("""
**📌 Как использовать:**
1. Выберите тип анализа в боковой панели
2. Введите данные (качественные или количественные)
3. Нажмите кнопку расчета
4. Получите p-value и интерпретацию

**📚 Для диссертации:**
- Критерий Фишера — для таблиц 2×2
- Манн-Уитни — для количественных данных
""")

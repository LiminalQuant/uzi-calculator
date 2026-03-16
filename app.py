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
    page_title="УЗИ Калькулятор: Японские свечи",
    page_icon="📊",
    layout="wide"
)

# Заголовок
st.title("📊 УЗИ Калькулятор: Японские свечи для гемодинамики")
st.markdown("---")

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
            
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            
            with col_stat1:
                st.markdown("**Длина тела свечи (Body)**")
                
                # Манн-Уитни для длины тела
                if len(poya) > 0 and len(rya) > 0:
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
                            poya['Длина_тела'].median() if len(poya)>0 else '-',
                            rya['Длина_тела'].median() if len(rya)>0 else '-',
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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Сохранить график PNG
            with col_save2:
                # Сохраняем график как PNG
                img_bytes = fig.to_image(format="png", width=1200, height=800)
                
                st.download_button(
                    label="📥 Скачать график (PNG)",
                    data=img_bytes,
                    file_name=f"uzi_candles_plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png"
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
                    file_name=f"uzi_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
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

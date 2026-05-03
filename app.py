import base64
import io
import os
import traceback
from datetime import datetime
from typing import Any, Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from building_project_predictor import BuildingPredictor

APP_VERSION = "V9.0 - Final Premium UI + Hero Building Background + Language/Currency + Reset Button"
EXCEL_PATH = "D_Building_2000_prediction_dataset.xlsx"
UNIVERSITY_LOGO = "university_logo.png"
HERO_IMAGE = "hero_building.webp"

STUDENTS = ["Hashemi Sayed Baset", "Salehy Sayed Moh Meraj"]
SUPERVISOR = "Mikheev Pavel Yurievich"
UNIVERSITY = "Peter the Great St. Petersburg Polytechnic University (SPbSTU)"

GREEN = "#2EAD4B"
GREEN_DARK = "#1E8E3E"
NAVY = "#10233F"
TEAL = "#0E7A70"
MUTED = "#5B6B7F"
CARD = "rgba(255,255,255,0.96)"

COMPUTED_COLS = {"avg_floor_plate_m2", "planned_unit_cost_per_m2", "schedule_days_per_1000_m2"}
CORE_INPUTS = ["gross_floor_area_m2", "floors_count", "planned_cost", "planned_duration", "complexity_level", "building_use_type"]
SITE_INPUTS = ["climate_zone", "structural_system", "foundation_type", "soil_type", "seismic_zone_or_pga", "load_bearing_capacity", "temperature", "humidity", "weather_condition", "air_quality_index"]
RESOURCE_INPUTS = ["equipment_utilization", "material_intensity_per_m2", "labor_intensity_hr_per_m2"]

TRANSLATIONS = {
    "en": {
        "language": "Language", "currency": "Currency", "exchange": "1 USD equals how many RUB?", "nav": "Navigation", "home": "Home", "prediction": "New Prediction", "about": "About",
        "overview": "Dashboard Overview", "overview_text": "This application supports early-stage building project decision-making by estimating cost performance, schedule performance, and risk level.",
        "developed_by": "Developed by", "supervisor": "Supervisor", "university": "University", "version": "Version",
        "badge_academic": "Academic", "badge_professional": "Professional", "badge_modern": "Modern Engineering UI", "title": "Building Project Prediction System",
        "subtitle": "A modern engineering decision-support platform for forecasting actual cost, actual duration, cost and schedule overrun, risk level, risk scores, and practical project recommendations.",
        "open_workspace": "Open prediction workspace", "read_system": "Read about the system", "sidebar_hint": "Use the sidebar to move between pages anytime.",
        "highlights": "Final dashboard highlights", "highlights_text": "This final interface combines a light, modern look with academic structure. It starts with a clear landing page, gives direct access to prediction, and produces a graphical PDF report for thesis presentation and decision support.",
        "decision_dashboard": "Decision dashboard", "decision_dashboard_text": "KPI cards, risk gauges, bar charts, and interpretation panels presented in one workflow.",
        "engineering_oriented": "Engineering oriented", "engineering_oriented_text": "Input sections are organized around project scale, site context, and productivity assumptions.",
        "exact_logic": "Exact uploaded logic", "exact_logic_text": "The app uses the uploaded BuildingPredictor logic, not a simplified demo calculator.",
        "pdf_report": "Graphical PDF report", "pdf_report_text": "Download a branded PDF report with green styling, charts, recommendations, and the university logo.",
        "workflow": "Engineering workflow", "workflow_text": "A clear final structure for professors and users: start with context, enter project data, review analytical outputs, then export the report.",
        "how_to_use": "How to use this application", "how_to_use_body": "1. Open New Prediction from the sidebar or the button above.\n2. Enter the new building project's initial data.\n3. Click Run project prediction.\n4. Review the KPI summary, graphs, warnings, and recommendations.\n5. Download the graphical PDF report.\n6. Use the output together with engineering judgment.",
        "about_system": "About this system", "about_system_text": "An academic-professional interface designed to present the building project prediction model in a cleaner, lighter, and more modern format.",
        "methodology": "Methodology", "methodology_text": "The application loads the provided workbook, trains the uploaded prediction engine, and applies the same logic to each new submitted project.",
        "intended_use": "Intended use", "intended_use_text": "Useful for thesis presentation, academic demonstration, and early-stage engineering decision support.",
        "included_outputs": "Included outputs", "included_outputs_text": "KPI dashboard, risk charts, output table, explanations, warnings, recommendations, and a graphical PDF report.",
        "decision_support_note": "This system should be used as a decision-support tool. Final decisions should also consider engineering review, site conditions, updated market data, and project-specific constraints.",
        "workspace": "Prediction workspace", "workspace_text": "Use this page to enter a new project and generate a complete engineering dashboard.",
        "structured_input": "Structured input", "structured_input_text": "The form is divided into project scale, site context, and productivity assumptions for clarity.",
        "analytical_output": "Analytical output", "analytical_output_text": "The dashboard includes charts for risk probability, scores, planned vs predicted performance, and overrun profile.",
        "ready_report": "Ready-to-present report", "ready_report_text": "A branded PDF report is generated for presentation, thesis use, or professional documentation.",
        "new_project_form": "New project input form", "new_project_form_text": "Enter the project information below. Three consistency indicators are calculated automatically from the main project inputs.",
        "project_scale": "1. Project scale and planning", "site_context": "2. Site and structural context", "resource_assumptions": "3. Resource and productivity assumptions", "additional_inputs": "4. Additional inputs",
        "run_prediction": "Run project prediction", "complete_form": "Complete the form and click Run project prediction to generate the dashboard and graphical PDF report.",
        "results": "Prediction results", "results_text": "Review the executive summary below, then open the dashboard and PDF report tabs for deeper analysis.",
        "start_new_project": "Start New Project / Clear Form", "start_new_help": "Clear the current prediction and prepare the form for another project.",
        "predicted_cost": "Predicted actual cost", "predicted_duration": "Predicted duration", "overrun": "Overrun", "deviation": "Deviation", "risk_level": "Risk level", "risk_score": "Risk score", "scenarios": "Scenarios", "cost_time_scenario": "Cost scenario / Time scenario",
        "tab_dashboard": "Dashboard", "tab_interpretation": "Interpretation", "tab_input": "Input profile", "tab_pdf": "Graphical PDF report",
        "interpretation": "Prediction interpretation", "interpretation_text": "Explanation causes, model summary, and management recommendations based on the uploaded prediction engine.",
        "summary": "Summary", "cost_explanation": "Cost explanation", "schedule_explanation": "Schedule explanation", "primary_cause": "Primary cause", "secondary_cause": "Secondary cause", "input_warnings": "Input warnings", "recommendations": "Project management recommendations",
        "submitted_profile": "Submitted project profile", "output_table": "Output table", "pdf_success": "Graphical PDF report generated successfully.", "download_pdf": "Download graphical PDF report", "pdf_info": "The PDF includes the university logo, green titles, summary tables, explanation, recommendations, and embedded graphs.",
        "loading": "Loading prediction engine and workbook...", "running": "Running prediction using the uploaded BuildingPredictor logic...", "load_error": "The prediction system could not be loaded. Make sure app.py, the predictor file, the workbook, and the image files are all uploaded together.", "prediction_error": "Prediction failed for the submitted input. Please check the values and try again.",
        "technical_details": "Technical error details", "internal_note": "Currency changes only affect display. The model still calculates internally using the original workbook logic.",
        "risk_probability": "Risk probability distribution", "probability": "Probability (%)", "score_comparison": "Risk score comparison", "score_axis": "Score (0-100)", "cost_chart": "Planned vs predicted cost", "duration_chart": "Planned vs predicted duration", "overrun_chart": "Overrun percentage profile",
        "planned": "Planned", "predicted": "Predicted", "cost_overrun": "Cost overrun", "schedule_overrun": "Schedule overrun",
        "report_title": "Building Project Prediction Report", "report_subtitle": "A Data-Driven Integrated Risk Intelligence Framework for Predictive and Sustainable Construction Project Management", "generated": "Generated", "exec_summary": "1. Executive prediction summary", "likely_explanation": "2. Likely explanation", "graphical_dashboard": "3. Graphical dashboard", "warnings_flags": "4. Input warnings and flags", "management_recs": "5. Project management recommendations", "appendix": "Appendix A. Submitted project input profile", "method_note": "Methodological note",
        "no_warning": "No input warning was generated for this submitted profile.", "report_method_note": "This report is based on the uploaded BuildingPredictor logic and workbook. The prediction should be interpreted as a decision-support estimate rather than a final contractual or engineering value. Final decisions should consider updated market conditions, site investigation, contract requirements, and engineering judgment.",
        "indicator": "Indicator", "value": "Value", "prediction_col": "Prediction", "actual_cost": "Actual cost", "actual_duration": "Actual duration", "cost_overrun_pct": "Cost overrun percentage", "schedule_overrun_pct": "Schedule overrun percentage", "cost_scenario": "Cost scenario", "time_scenario": "Time scenario", "cost_risk_score": "Cost risk score", "schedule_risk_score": "Schedule risk score",
    },
    "ru": {
        "language": "Язык", "currency": "Валюта", "exchange": "Сколько рублей в 1 USD?", "nav": "Навигация", "home": "Главная", "prediction": "Новый прогноз", "about": "О системе",
        "overview": "Обзор панели", "overview_text": "Приложение поддерживает принятие решений на ранней стадии строительного проекта, оценивая стоимость, сроки и уровень риска.",
        "developed_by": "Разработали", "supervisor": "Научный руководитель", "university": "Университет", "version": "Версия",
        "badge_academic": "Академический", "badge_professional": "Профессиональный", "badge_modern": "Современный инженерный интерфейс", "title": "Система прогнозирования строительных проектов",
        "subtitle": "Современная инженерная система поддержки решений для прогнозирования фактической стоимости, длительности, перерасхода, задержек, уровня риска и управленческих рекомендаций.",
        "open_workspace": "Открыть прогнозирование", "read_system": "О системе", "sidebar_hint": "Используйте боковую панель для перехода между разделами.",
        "highlights": "Ключевые возможности панели", "highlights_text": "Финальный интерфейс объединяет легкий современный дизайн и академическую структуру. Он содержит стартовую страницу, доступ к прогнозу и PDF-отчет с графиками.",
        "decision_dashboard": "Аналитическая панель", "decision_dashboard_text": "KPI-карточки, индикаторы риска, диаграммы и интерпретация в одном рабочем процессе.",
        "engineering_oriented": "Инженерная структура", "engineering_oriented_text": "Входные данные организованы по масштабу проекта, условиям площадки и ресурсным предположениям.",
        "exact_logic": "Исходная логика модели", "exact_logic_text": "Сайт использует загруженную логику BuildingPredictor, а не упрощенный демонстрационный расчет.",
        "pdf_report": "PDF-отчет с графиками", "pdf_report_text": "Скачайте брендированный PDF-отчет с зеленым стилем, графиками, рекомендациями и логотипом университета.",
        "workflow": "Инженерный процесс", "workflow_text": "Понятная структура: изучить контекст, ввести данные проекта, просмотреть результаты и экспортировать отчет.",
        "how_to_use": "Как пользоваться приложением", "how_to_use_body": "1. Откройте Новый прогноз в боковой панели или кнопкой выше.\n2. Введите исходные данные нового строительного проекта.\n3. Нажмите Выполнить прогноз.\n4. Просмотрите KPI, графики, предупреждения и рекомендации.\n5. Скачайте PDF-отчет с графиками.\n6. Используйте результат вместе с инженерной оценкой.",
        "about_system": "О системе", "about_system_text": "Академико-профессиональный интерфейс для представления модели прогнозирования строительных проектов в чистом, светлом и современном формате.",
        "methodology": "Методология", "methodology_text": "Приложение загружает предоставленную рабочую книгу, обучает загруженный механизм прогнозирования и применяет ту же логику к каждому новому проекту.",
        "intended_use": "Назначение", "intended_use_text": "Подходит для презентации диссертации, академической демонстрации и поддержки инженерных решений на ранней стадии.",
        "included_outputs": "Выходные данные", "included_outputs_text": "KPI-панель, графики риска, таблица результатов, объяснения, предупреждения, рекомендации и PDF-отчет с графиками.",
        "decision_support_note": "Система должна использоваться как инструмент поддержки решений. Окончательные решения должны учитывать инженерную проверку, условия площадки, рыночные данные и особенности проекта.",
        "workspace": "Рабочее пространство прогноза", "workspace_text": "Используйте эту страницу для ввода нового проекта и создания полной инженерной панели.",
        "structured_input": "Структурированный ввод", "structured_input_text": "Форма разделена на масштаб проекта, условия площадки и ресурсные предположения.",
        "analytical_output": "Аналитический результат", "analytical_output_text": "Панель содержит графики вероятности риска, оценок, плановых и прогнозных значений, а также профиля перерасхода.",
        "ready_report": "Готовый отчет", "ready_report_text": "Брендированный PDF-отчет создается для презентации, диссертации или профессиональной документации.",
        "new_project_form": "Форма ввода нового проекта", "new_project_form_text": "Введите данные проекта ниже. Три проверочных показателя рассчитываются автоматически из основных входных данных.",
        "project_scale": "1. Масштаб проекта и планирование", "site_context": "2. Условия площадки и конструктивный контекст", "resource_assumptions": "3. Ресурсы и производительность", "additional_inputs": "4. Дополнительные входные данные",
        "run_prediction": "Выполнить прогноз проекта", "complete_form": "Заполните форму и нажмите Выполнить прогноз проекта, чтобы создать панель и PDF-отчет.",
        "results": "Результаты прогнозирования", "results_text": "Просмотрите сводку, затем откройте вкладки панели и PDF-отчета для детального анализа.",
        "start_new_project": "Начать новый проект / Очистить форму", "start_new_help": "Очистить текущий прогноз и подготовить форму для другого проекта.",
        "predicted_cost": "Прогнозная фактическая стоимость", "predicted_duration": "Прогнозная длительность", "overrun": "Перерасход", "deviation": "Отклонение", "risk_level": "Уровень риска", "risk_score": "Оценка риска", "scenarios": "Сценарии", "cost_time_scenario": "Сценарий стоимости / времени",
        "tab_dashboard": "Панель", "tab_interpretation": "Интерпретация", "tab_input": "Профиль ввода", "tab_pdf": "PDF-отчет",
        "interpretation": "Интерпретация прогноза", "interpretation_text": "Причины, сводка модели и управленческие рекомендации на основе загруженного механизма прогнозирования.",
        "summary": "Сводка", "cost_explanation": "Объяснение стоимости", "schedule_explanation": "Объяснение графика", "primary_cause": "Основная причина", "secondary_cause": "Вторичная причина", "input_warnings": "Предупреждения ввода", "recommendations": "Рекомендации по управлению проектом",
        "submitted_profile": "Профиль введенного проекта", "output_table": "Таблица результатов", "pdf_success": "PDF-отчет с графиками успешно создан.", "download_pdf": "Скачать PDF-отчет", "pdf_info": "PDF включает логотип университета, зеленые заголовки, таблицы, объяснение, рекомендации и графики.",
        "loading": "Загрузка модели и рабочей книги...", "running": "Выполняется прогноз с использованием загруженной логики BuildingPredictor...", "load_error": "Система прогнозирования не загрузилась. Убедитесь, что app.py, файл модели, рабочая книга и изображения загружены вместе.", "prediction_error": "Прогноз не выполнен. Проверьте значения и попробуйте снова.",
        "technical_details": "Технические детали ошибки", "internal_note": "Валюта изменяет только отображение. Модель по-прежнему рассчитывает внутри по исходной логике рабочей книги.",
        "risk_probability": "Распределение вероятности риска", "probability": "Вероятность (%)", "score_comparison": "Сравнение оценок риска", "score_axis": "Оценка (0-100)", "cost_chart": "Плановая и прогнозная стоимость", "duration_chart": "Плановая и прогнозная длительность", "overrun_chart": "Профиль перерасхода",
        "planned": "План", "predicted": "Прогноз", "cost_overrun": "Перерасход стоимости", "schedule_overrun": "Отклонение по срокам",
        "report_title": "Отчет по прогнозированию строительного проекта", "report_subtitle": "Интегрированная система риск-аналитики для прогнозного и устойчивого управления строительными проектами", "generated": "Создано", "exec_summary": "1. Краткая сводка прогноза", "likely_explanation": "2. Вероятное объяснение", "graphical_dashboard": "3. Графическая панель", "warnings_flags": "4. Предупреждения и факторы", "management_recs": "5. Рекомендации по управлению проектом", "appendix": "Приложение A. Профиль введенного проекта", "method_note": "Методологическое примечание",
        "no_warning": "Для данного профиля проекта предупреждения не сформированы.", "report_method_note": "Этот отчет основан на загруженной логике BuildingPredictor и рабочей книге. Прогноз следует интерпретировать как оценку поддержки решений, а не как окончательное договорное или инженерное значение. Окончательные решения должны учитывать рыночные условия, обследование площадки, требования договора и инженерную оценку.",
        "indicator": "Показатель", "value": "Значение", "prediction_col": "Прогноз", "actual_cost": "Фактическая стоимость", "actual_duration": "Фактическая длительность", "cost_overrun_pct": "Процент перерасхода стоимости", "schedule_overrun_pct": "Процент отклонения графика", "cost_scenario": "Сценарий стоимости", "time_scenario": "Сценарий времени", "cost_risk_score": "Оценка риска стоимости", "schedule_risk_score": "Оценка риска графика",
    }
}

LABEL_RU = {
    "gross_floor_area_m2": "Общая площадь (м²)", "floors_count": "Количество этажей", "planned_cost": "Плановая стоимость", "planned_duration": "Плановая длительность", "complexity_level": "Уровень сложности", "building_use_type": "Назначение здания", "climate_zone": "Климатическая зона", "structural_system": "Конструктивная система", "foundation_type": "Тип фундамента", "soil_type": "Тип грунта", "seismic_zone_or_pga": "Сейсмическая зона / PGA", "load_bearing_capacity": "Несущая способность", "temperature": "Температура", "humidity": "Влажность", "weather_condition": "Погодные условия", "air_quality_index": "Индекс качества воздуха", "equipment_utilization": "Использование оборудования", "material_intensity_per_m2": "Материалоемкость на м²", "labor_intensity_hr_per_m2": "Трудоемкость (ч/м²)", "avg_floor_plate_m2": "Средняя площадь этажа (м²)", "planned_unit_cost_per_m2": "Плановая стоимость за м²", "schedule_days_per_1000_m2": "Дней на 1000 м²"
}


def lang_code() -> str:
    return "ru" if st.session_state.get("language") == "Русский" else "en"


def tr(key: str) -> str:
    code = lang_code()
    return TRANSLATIONS.get(code, {}).get(key, TRANSLATIONS["en"].get(key, key))


def labelize(name: str) -> str:
    if lang_code() == "ru" and name in LABEL_RU:
        return LABEL_RU[name]
    special = {"gross_floor_area_m2": "Gross Floor Area (m²)", "material_intensity_per_m2": "Material Intensity per m²", "labor_intensity_hr_per_m2": "Labor Intensity (hr/m²)", "planned_unit_cost_per_m2": "Planned Unit Cost per m²", "schedule_days_per_1000_m2": "Schedule Days per 1,000 m²", "avg_floor_plate_m2": "Average Floor Plate (m²)", "seismic_zone_or_pga": "Seismic Zone / PGA", "air_quality_index": "Air Quality Index"}
    return special.get(name, name.replace("_", " ").title())


def display_term(v: Any) -> str:
    if lang_code() != "ru":
        return str(v)
    mp = {"Low": "Низкий", "Medium": "Средний", "High": "Высокий", "Optimistic": "Оптимистичный", "Most Likely": "Наиболее вероятный", "Pessimistic": "Пессимистичный"}
    return mp.get(str(v), str(v))


def image_to_b64(path: str) -> str:
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def money_factor() -> float:
    return float(st.session_state.get("exchange_rate", 1.0)) if st.session_state.get("currency", "USD") == "RUB" else 1.0


def money_symbol() -> str:
    return "₽" if st.session_state.get("currency", "USD") == "RUB" else "$"


def fmt_money(v: float) -> str:
    converted = float(v) * money_factor()
    sym = money_symbol()
    return f"{sym}{converted:,.0f}"


def fmt_money_plain(v: float, factor: float, symbol: str) -> str:
    return f"{symbol}{float(v) * factor:,.0f}"


def fmt_days(v: float) -> str:
    return f"{float(v):,.0f} days" if lang_code() == "en" else f"{float(v):,.0f} дней"


def fmt_pct(v: float) -> str:
    return f"{float(v) * 100:.2f}%"


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def risk_color(risk: str) -> str:
    r = str(risk).lower()
    if "low" in r or "низ" in r:
        return "#15803d"
    if "high" in r or "выс" in r:
        return "#b91c1c"
    return "#b45309"


def inject_css() -> None:
    hero_b64 = image_to_b64(HERO_IMAGE)
    hero_bg = f"linear-gradient(90deg, rgba(255,255,255,0.92) 0%, rgba(255,255,255,0.84) 50%, rgba(255,255,255,0.66) 100%), url(data:image/webp;base64,{hero_b64})" if hero_b64 else "linear-gradient(135deg, #ffffff 0%, #f0fbf3 100%)"
    st.markdown(f"""
        <style>
        .stApp {{ background: radial-gradient(circle at 10% 10%, rgba(46,173,75,0.08), transparent 28%), linear-gradient(180deg, #f7fbf8 0%, #f3fbf6 46%, #f8fbff 100%); }}
        section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, #fbfffd 0%, #f2fbf4 56%, #ecf5fb 100%); border-right: 1px solid rgba(16,35,63,0.10); }}
        section[data-testid="stSidebar"] * {{ color: {NAVY} !important; }}
        .sidebar-soft-card {{ background: rgba(255,255,255,0.76); border: 1px solid rgba(16,35,63,0.08); border-radius: 18px; padding: 0.9rem 1rem; margin-bottom: 0.85rem; box-shadow: 0 8px 24px rgba(16,35,63,0.05); }}
        .hero {{ padding: 3rem 3rem; min-height: 335px; border-radius: 32px; background-image: {hero_bg}; background-size: cover; background-position: center; border: 1px solid rgba(46,173,75,0.16); box-shadow: 0 22px 50px rgba(16,35,63,0.10); margin-bottom: 1rem; display: flex; flex-direction: column; justify-content: center; }}
        .hero h1 {{ color: {NAVY}; font-size: 2.8rem; margin: 0 0 0.55rem 0; font-weight: 900; letter-spacing: 0.01em; max-width: 950px; }}
        .hero p {{ color: #31445f; font-size: 1.08rem; line-height: 1.72; max-width: 900px; font-weight: 500; }}
        .version-pill {{ display: inline-block; width: fit-content; background: rgba(46,173,75,0.12); color: {GREEN_DARK}; border: 1px solid rgba(46,173,75,0.26); padding: 0.36rem 0.78rem; border-radius: 999px; font-size: 0.84rem; font-weight: 800; margin-top: 0.55rem; }}
        .mini-badge {{ display: inline-block; background: rgba(255,255,255,0.76); color: {NAVY}; border-radius: 999px; padding: 0.30rem 0.66rem; font-size: 0.82rem; font-weight: 800; margin-right: 0.35rem; margin-bottom: 0.45rem; border: 1px solid rgba(16,35,63,0.07); }}
        .section-card, .feature-card, .stat-card, .step-card, .footer-card {{ background: {CARD}; border: 1px solid rgba(16,35,63,0.08); border-radius: 22px; box-shadow: 0 12px 32px rgba(16,35,63,0.05); }}
        .section-card {{ padding: 1.2rem 1.3rem; margin-bottom: 1rem; }} .section-title {{ color: {NAVY}; font-size: 1.25rem; font-weight: 900; margin-bottom: 0.35rem; }} .section-sub {{ color: {MUTED}; font-size: 0.95rem; line-height: 1.62; }}
        .feature-card {{ padding: 1rem 1.05rem; min-height: 172px; }} .feature-icon {{ font-size: 1.6rem; margin-bottom: 0.4rem; }} .feature-title {{ color: {NAVY}; font-size: 1rem; font-weight: 900; margin-bottom: 0.35rem; }} .feature-text {{ color: {MUTED}; font-size: 0.92rem; line-height: 1.58; }}
        .step-card {{ padding: 1rem 1.05rem; min-height: 150px; }} .step-index {{ width: 34px; height: 34px; border-radius: 50%; display:flex; align-items:center; justify-content:center; background: linear-gradient(135deg, {GREEN} 0%, {GREEN_DARK} 100%); color: white; font-weight: 900; margin-bottom: 0.55rem; }}
        .stat-card {{ padding: 1rem; }} .kpi-label {{ color: {MUTED}; font-size: 0.83rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; }} .kpi-value {{ color: {NAVY}; font-size: 1.45rem; font-weight: 900; margin-top: 0.25rem; }} .kpi-sub {{ color: {MUTED}; font-size: 0.9rem; margin-top: 0.12rem; }}
        .green-callout {{ border-left: 5px solid {GREEN}; background: #eefaf0; color: {NAVY}; padding: 1rem; border-radius: 16px; margin-bottom: 0.6rem; }} .warn-box {{ border-left: 5px solid #ea580c; background: #fff7ed; color: #9a3412; padding: 0.95rem 1rem; border-radius: 16px; margin-bottom: 0.55rem; }} .footer-card {{ padding: 1rem 1.2rem; margin-top: 1rem; }}
        .stButton > button, .stDownloadButton > button {{ border-radius: 14px; border: 0; background: linear-gradient(135deg, {GREEN} 0%, {GREEN_DARK} 100%); color: white; font-weight: 800; padding: 0.72rem 1.08rem; box-shadow: 0 10px 22px rgba(46,173,75,0.22); }}
        .secondary-button button {{ background: linear-gradient(135deg, #10233F 0%, #1f3b5f 100%) !important; }}
        .stTabs [data-baseweb="tab-list"] {{ gap: 0.55rem; }} .stTabs [data-baseweb="tab"] {{ border-radius: 999px; padding: 0.55rem 1rem; background: white; border: 1px solid rgba(16,35,63,0.08); }} .stTabs [aria-selected="true"] {{ background: rgba(46,173,75,0.12) !important; border-color: rgba(46,173,75,0.24) !important; }}
        </style>
    """, unsafe_allow_html=True)


@st.cache_resource(show_spinner=False)
def load_system() -> BuildingPredictor:
    system = BuildingPredictor(EXCEL_PATH)
    system.load_data()
    system.train(system.data.copy())
    return system


def init_state() -> None:
    if "language" not in st.session_state:
        st.session_state.language = "English"
    if "currency" not in st.session_state:
        st.session_state.currency = "USD"
    if "exchange_rate" not in st.session_state:
        st.session_state.exchange_rate = 80.0
    if "page" not in st.session_state:
        st.session_state.page = "Home"
    if "prediction_result" not in st.session_state:
        st.session_state.prediction_result = None
    if "raw_df" not in st.session_state:
        st.session_state.raw_df = None
    if "form_serial" not in st.session_state:
        st.session_state.form_serial = 0


def reset_project() -> None:
    st.session_state.prediction_result = None
    st.session_state.raw_df = None
    st.session_state.form_serial += 1
    st.session_state.page = "New Prediction"


def go_to(page: str) -> None:
    st.session_state.page = page


def render_sidebar() -> None:
    with st.sidebar:
        if os.path.exists(UNIVERSITY_LOGO):
            st.image(UNIVERSITY_LOGO, use_container_width=True)
        st.markdown('<div class="sidebar-soft-card"><b>Settings</b></div>', unsafe_allow_html=True)
        st.selectbox(tr("language"), ["English", "Русский"], key="language")
        st.selectbox(tr("currency"), ["USD", "RUB"], key="currency")
        if st.session_state.currency == "RUB":
            st.number_input(tr("exchange"), min_value=1.0, max_value=500.0, value=float(st.session_state.exchange_rate), step=1.0, key="exchange_rate")
        st.caption(tr("internal_note"))
        st.markdown(f'<div class="sidebar-soft-card"><b>{tr("nav")}</b></div>', unsafe_allow_html=True)
        page_labels = {"Home": tr("home"), "New Prediction": tr("prediction"), "About": tr("about")}
        current = page_labels.get(st.session_state.page, page_labels["Home"])
        selected_label = st.radio("Go to", list(page_labels.values()), index=list(page_labels.values()).index(current), label_visibility="collapsed")
        for key, label in page_labels.items():
            if label == selected_label:
                st.session_state.page = key
                break
        st.markdown(f'<div class="sidebar-soft-card"><b>{tr("overview")}</b><br><br>{tr("overview_text")}</div>', unsafe_allow_html=True)
        st.markdown('<div class="sidebar-soft-card"><b>' + tr("developed_by") + '</b><br><br>' + '<br>'.join(STUDENTS) + f'<br><br><b>{tr("supervisor")}</b><br><br>{SUPERVISOR}<br><br><b>{tr("university")}</b><br><br>{UNIVERSITY}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sidebar-soft-card"><b>{tr("version")}</b><br><br>{APP_VERSION}</div>', unsafe_allow_html=True)


def numeric_stats(system: BuildingPredictor, col_name: str) -> Tuple[float, float, float, bool]:
    vals = pd.to_numeric(system.inputs_df[col_name], errors="coerce").dropna()
    lo = float(vals.min()) if not vals.empty else 0.0
    hi = float(vals.max()) if not vals.empty else 0.0
    med = float(vals.median()) if not vals.empty else 0.0
    int_like = bool(np.allclose(vals, np.round(vals))) if not vals.empty else False
    return lo, hi, med, int_like


def broad_min_for_col(col_name: str, lo: float) -> float:
    if col_name == "temperature":
        return min(-50.0, lo)
    if col_name == "floors_count":
        return 1.0
    return 0.0


def numeric_input_for(system: BuildingPredictor, col_name: str, key: str) -> float:
    lo, hi, median, int_like = numeric_stats(system, col_name)
    display_factor = money_factor() if col_name == "planned_cost" else 1.0
    label = labelize(col_name) + (" (RUB)" if col_name == "planned_cost" and st.session_state.currency == "RUB" else "")
    min_value = float(broad_min_for_col(col_name, lo) * display_factor)
    value = float(median * display_factor)
    display_lo, display_hi = lo * display_factor, hi * display_factor
    if col_name in {"humidity", "equipment_utilization"}:
        max_value = 100.0
        step = 1.0 if int_like else 0.1
    elif col_name == "air_quality_index":
        max_value = 500.0
        step = 1.0
    else:
        max_value = None
        span = max(abs(display_hi - display_lo), 1.0)
        if int_like:
            step = max(1.0, round(span / 100.0))
        elif span > 1000:
            step = 100.0
        elif span > 100:
            step = 10.0
        elif span > 10:
            step = 1.0
        else:
            step = 0.1
    if col_name == "planned_cost" and st.session_state.currency == "RUB":
        step = max(100.0, step)
    fmt = "%.0f" if int_like or col_name == "planned_cost" else "%.4f"
    val_display = st.number_input(label, min_value=float(min_value), max_value=(None if max_value is None else float(max_value)), value=float(value), step=float(step), format=fmt, key=key, help=f"Workbook reference range: {display_lo:,.2f} to {display_hi:,.2f}.")
    st.caption(f"Reference range: {display_lo:,.2f} – {display_hi:,.2f}" if lang_code() == "en" else f"Ориентировочный диапазон: {display_lo:,.2f} – {display_hi:,.2f}")
    if col_name == "planned_cost" and st.session_state.currency == "RUB":
        return float(val_display) / money_factor()
    return float(val_display)


def categorical_input_for(system: BuildingPredictor, col_name: str, key: str) -> str:
    options = sorted([str(x) for x in system.inputs_df[col_name].dropna().astype(str).unique().tolist()], key=lambda x: x.lower())
    if not options:
        return st.text_input(labelize(col_name), value="", key=key)
    mode_series = system.inputs_df[col_name].mode(dropna=True)
    mode = str(mode_series.iloc[0]) if not mode_series.empty else options[0]
    idx = options.index(mode) if mode in options else 0
    return st.selectbox(labelize(col_name), options=options, index=idx, key=key)


def render_field(system: BuildingPredictor, col_name: str, raw: Dict[str, Any], container) -> None:
    with container:
        key = f"field_{st.session_state.form_serial}_{col_name}"
        if col_name in system.numeric_columns:
            raw[col_name] = numeric_input_for(system, col_name, key)
        else:
            raw[col_name] = categorical_input_for(system, col_name, key)


def add_computed_fields(raw: Dict[str, Any]) -> Dict[str, Any]:
    area = max(safe_float(raw.get("gross_floor_area_m2"), 0.0), 0.0001)
    floors = max(safe_float(raw.get("floors_count"), 1.0), 1.0)
    planned_cost = safe_float(raw.get("planned_cost"), 0.0)
    planned_duration = safe_float(raw.get("planned_duration"), 0.0)
    raw["avg_floor_plate_m2"] = area / floors
    raw["planned_unit_cost_per_m2"] = planned_cost / area
    raw["schedule_days_per_1000_m2"] = planned_duration / (area / 1000.0)
    return raw


def build_input_form(system: BuildingPredictor) -> Tuple[pd.DataFrame, bool]:
    st.markdown(f'<div class="section-card"><div class="section-title">{tr("new_project_form")}</div><div class="section-sub">{tr("new_project_form_text")}</div></div>', unsafe_allow_html=True)
    raw: Dict[str, Any] = {"project_id": f"NEW_{datetime.now().strftime('%Y%m%d_%H%M%S')}"}
    editable_cols = [c for c in system.input_columns if c not in COMPUTED_COLS]
    with st.form(f"prediction_form_{st.session_state.form_serial}", clear_on_submit=False):
        st.subheader(tr("project_scale"))
        a, b = st.columns(2)
        for i, col_name in enumerate([c for c in CORE_INPUTS if c in editable_cols]):
            render_field(system, col_name, raw, a if i % 2 == 0 else b)
        st.subheader(tr("site_context"))
        c, d = st.columns(2)
        for i, col_name in enumerate([cname for cname in SITE_INPUTS if cname in editable_cols]):
            render_field(system, col_name, raw, c if i % 2 == 0 else d)
        st.subheader(tr("resource_assumptions"))
        e, f = st.columns(2)
        for i, col_name in enumerate([cname for cname in RESOURCE_INPUTS if cname in editable_cols]):
            render_field(system, col_name, raw, e if i % 2 == 0 else f)
        remaining = [c for c in editable_cols if c not in set(CORE_INPUTS + SITE_INPUTS + RESOURCE_INPUTS)]
        if remaining:
            st.subheader(tr("additional_inputs"))
            g, h = st.columns(2)
            for i, col_name in enumerate(remaining):
                render_field(system, col_name, raw, g if i % 2 == 0 else h)
        submitted = st.form_submit_button(tr("run_prediction"), type="primary", use_container_width=True)
    raw = add_computed_fields(raw)
    ordered = {"project_id": raw.get("project_id")}
    for col in system.input_columns:
        if col in raw:
            ordered[col] = raw[col]
        elif col in system.numeric_columns:
            _, _, med, _ = numeric_stats(system, col)
            ordered[col] = float(med)
        else:
            mode_series = system.inputs_df[col].mode(dropna=True)
            ordered[col] = str(mode_series.iloc[0]) if not mode_series.empty else ""
    raw_df = pd.DataFrame([ordered], columns=["project_id"] + system.input_columns)
    for col in system.numeric_columns:
        vals = pd.to_numeric(system.inputs_df[col], errors="coerce").dropna()
        if len(vals) and np.allclose(vals, np.round(vals)):
            raw_df[col] = pd.to_numeric(raw_df[col], errors="coerce").round().astype(int)
        else:
            raw_df[col] = pd.to_numeric(raw_df[col], errors="coerce").astype(float)
    return raw_df, submitted


def render_hero() -> None:
    st.markdown(f"""
        <div class="hero">
            <div>
                <span class="mini-badge">{tr('badge_academic')}</span>
                <span class="mini-badge">{tr('badge_professional')}</span>
                <span class="mini-badge">{tr('badge_modern')}</span>
                <h1>{tr('title')}</h1>
                <p>{tr('subtitle')}</p>
                <span class="version-pill">{APP_VERSION}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)


def render_home_page() -> None:
    render_hero()
    nav1, nav2, nav3 = st.columns(3)
    with nav1:
        if st.button(tr("open_workspace"), use_container_width=True):
            go_to("New Prediction")
            st.rerun()
    with nav2:
        if st.button(tr("read_system"), use_container_width=True):
            go_to("About")
            st.rerun()
    with nav3:
        st.caption(tr("sidebar_hint"))
    st.markdown(f'<div class="section-card"><div class="section-title">{tr("highlights")}</div><div class="section-sub">{tr("highlights_text")}</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    cards = [("📊", tr("decision_dashboard"), tr("decision_dashboard_text")), ("🏗️", tr("engineering_oriented"), tr("engineering_oriented_text")), ("🧠", tr("exact_logic"), tr("exact_logic_text")), ("📝", tr("pdf_report"), tr("pdf_report_text"))]
    for col, (icon, title, text) in zip([c1, c2, c3, c4], cards):
        col.markdown(f'<div class="feature-card"><div class="feature-icon">{icon}</div><div class="feature-title">{title}</div><div class="feature-text">{text}</div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="section-card"><div class="section-title">{tr("workflow")}</div><div class="section-sub">{tr("workflow_text")}</div></div>', unsafe_allow_html=True)
    with st.expander(tr("how_to_use"), expanded=True):
        st.markdown(tr("how_to_use_body"))
    st.markdown(f'<div class="footer-card"><b>{tr("developed_by")}:</b> {STUDENTS[0]} and {STUDENTS[1]} &nbsp;&nbsp;|&nbsp;&nbsp; <b>{tr("supervisor")}:</b> {SUPERVISOR} &nbsp;&nbsp;|&nbsp;&nbsp; <b>{tr("university")}:</b> {UNIVERSITY}</div>', unsafe_allow_html=True)


def render_about_page() -> None:
    st.markdown(f'<div class="section-card"><div class="section-title">{tr("about_system")}</div><div class="section-sub">{tr("about_system_text")}</div></div>', unsafe_allow_html=True)
    l, r = st.columns([1.2, 2.3], vertical_alignment="top")
    with l:
        if os.path.exists(UNIVERSITY_LOGO):
            st.image(UNIVERSITY_LOGO, use_container_width=True)
    with r:
        st.markdown(f"**{tr('developed_by')}:** {STUDENTS[0]} and {STUDENTS[1]}  \n**{tr('supervisor')}:** {SUPERVISOR}  \n**{tr('university')}:** {UNIVERSITY}")
        st.markdown(tr("about_system_text"))
    a1, a2, a3 = st.columns(3)
    a1.markdown(f'<div class="feature-card"><div class="feature-icon">⚙️</div><div class="feature-title">{tr("methodology")}</div><div class="feature-text">{tr("methodology_text")}</div></div>', unsafe_allow_html=True)
    a2.markdown(f'<div class="feature-card"><div class="feature-icon">🎯</div><div class="feature-title">{tr("intended_use")}</div><div class="feature-text">{tr("intended_use_text")}</div></div>', unsafe_allow_html=True)
    a3.markdown(f'<div class="feature-card"><div class="feature-icon">📦</div><div class="feature-title">{tr("included_outputs")}</div><div class="feature-text">{tr("included_outputs_text")}</div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="footer-card">{tr("decision_support_note")}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, caption: str = "") -> str:
    return f'<div class="stat-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-sub">{caption}</div></div>'


def make_gauge(title: str, value: float, max_value: float = 100.0, bar_color: str = GREEN) -> go.Figure:
    fig = go.Figure(go.Indicator(mode="gauge+number", value=float(value), number={"font": {"size": 34}}, title={"text": title, "font": {"size": 16}}, gauge={"axis": {"range": [0, max_value]}, "bar": {"color": bar_color}, "bgcolor": "white", "borderwidth": 1, "bordercolor": "#d8e1ea", "steps": [{"range": [0, max_value*0.33], "color": "#dcfce7"}, {"range": [max_value*0.33, max_value*0.66], "color": "#fef3c7"}, {"range": [max_value*0.66, max_value], "color": "#fee2e2"}] }))
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=58, b=20), paper_bgcolor="rgba(0,0,0,0)")
    return fig


def risk_probability_chart(probs: Dict[str, float]) -> go.Figure:
    items = [(k, 100 * float(v)) for k, v in sorted(probs.items(), key=lambda kv: kv[0])]
    x = [display_term(k) for k, _ in items]
    y = [v for _, v in items]
    colors_list = []
    for label, _ in items:
        ll = label.lower()
        colors_list.append(GREEN if "low" in ll else "#ef4444" if "high" in ll else "#eab308")
    fig = go.Figure(go.Bar(x=x, y=y, text=[f"{v:.1f}%" for v in y], textposition="outside", marker_color=colors_list))
    fig.update_layout(title=tr("risk_probability"), yaxis_title=tr("probability"), height=340, margin=dict(l=20, r=20, t=60, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.7)", yaxis=dict(range=[0, max(100, max(y) + 10 if y else 100)]))
    return fig


def score_bar_chart(out: Dict[str, Any]) -> go.Figure:
    labels = [tr("risk_score"), tr("cost_risk_score"), tr("schedule_risk_score")]
    values = [safe_float(out.get("risk_score")), safe_float(out.get("cost_risk_score")), safe_float(out.get("schedule_risk_score"))]
    fig = go.Figure(go.Bar(x=labels, y=values, text=[f"{v:.1f}" for v in values], textposition="outside", marker_color=[GREEN, "#2563eb", TEAL]))
    fig.update_layout(title=tr("score_comparison"), yaxis_title=tr("score_axis"), yaxis=dict(range=[0, 105]), height=340, margin=dict(l=20, r=20, t=60, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.7)")
    return fig


def overrun_chart(out: Dict[str, Any]) -> go.Figure:
    labels = [tr("cost_overrun"), tr("schedule_overrun")]
    vals = [100 * safe_float(out.get("cost_overrun_percentage")), 100 * safe_float(out.get("schedule_overrun_percentage"))]
    fig = go.Figure(go.Bar(x=labels, y=vals, text=[f"{v:.2f}%" for v in vals], textposition="outside", marker_color=[GREEN, TEAL]))
    fig.update_layout(title=tr("overrun_chart"), yaxis_title="%", height=340, margin=dict(l=20, r=20, t=60, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.7)")
    return fig


def cost_chart(raw_df: pd.DataFrame, out: Dict[str, Any]) -> go.Figure:
    planned_cost = safe_float(raw_df.iloc[0]["planned_cost"]) * money_factor()
    predicted_cost = safe_float(out["actual_cost"]) * money_factor()
    fig = go.Figure()
    fig.add_trace(go.Bar(name=tr("planned"), x=[tr("planned")], y=[planned_cost], text=[fmt_money(safe_float(raw_df.iloc[0]["planned_cost"]))], textposition="outside", marker_color="#bdc9d7"))
    fig.add_trace(go.Bar(name=tr("predicted"), x=[tr("predicted")], y=[predicted_cost], text=[fmt_money(out["actual_cost"])], textposition="outside", marker_color=GREEN))
    fig.update_layout(title=tr("cost_chart"), height=330, margin=dict(l=20, r=20, t=60, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.7)", showlegend=False)
    return fig


def duration_chart(raw_df: pd.DataFrame, out: Dict[str, Any]) -> go.Figure:
    planned_duration = safe_float(raw_df.iloc[0]["planned_duration"])
    predicted_duration = safe_float(out["actual_duration"])
    fig = go.Figure()
    fig.add_trace(go.Bar(name=tr("planned"), x=[tr("planned")], y=[planned_duration], text=[fmt_days(planned_duration)], textposition="outside", marker_color="#bdc9d7"))
    fig.add_trace(go.Bar(name=tr("predicted"), x=[tr("predicted")], y=[predicted_duration], text=[fmt_days(predicted_duration)], textposition="outside", marker_color=TEAL))
    fig.update_layout(title=tr("duration_chart"), height=330, margin=dict(l=20, r=20, t=60, b=40), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(255,255,255,0.7)", showlegend=False)
    return fig


def build_recommendations(result: Dict[str, Any]) -> List[str]:
    out = result["predicted_outputs"]
    # Recommendations remain in English/Russian interface only partly translated; model flags stay original.
    if lang_code() == "ru":
        base = ["Используйте еженедельный мониторинг, чтобы контролировать умеренный уровень проектного риска.", "Обновляйте стоимостную базу при изменении объемов, цен или условий поставок.", "Используйте планирование ближайших работ и регулярный контроль прогресса для управления отклонениями сроков."]
        base.extend(result.get("input_driven_flags", [])[:4])
        base.append("Используйте результат как поддержку решений вместе с инженерной оценкой и актуальными рыночными данными.")
        return list(dict.fromkeys(base))[:8]
    recs: List[str] = []
    risk = str(out.get("risk_level", "Medium"))
    cost_scenario = str(out.get("cost_scenario", "Most Likely"))
    time_scenario = str(out.get("time_scenario", "Most Likely"))
    if risk.lower() == "high":
        recs.append("Treat this project as a high-priority control case and establish an explicit risk register before execution.")
    elif risk.lower() == "medium":
        recs.append("Use weekly monitoring because the project shows moderate risk exposure and needs structured follow-up.")
    else:
        recs.append("Maintain standard controls but keep cost and schedule indicators under regular review.")
    recs.append("Prepare a cost contingency plan and monitor material prices, procurement timing, and high-value packages closely." if cost_scenario == "Pessimistic" else "Keep the cost baseline updated whenever scope, quantities, or market prices change.")
    recs.append("Strengthen schedule control by checking the critical path, labor availability, and weather-sensitive work sequences." if time_scenario == "Pessimistic" else "Use look-ahead planning and weekly progress checks to keep schedule deviation manageable.")
    for flag in result.get("input_driven_flags", [])[:4]:
        recs.append(flag)
    recs.append("Use the result as decision support together with engineering judgment, site review, and current market information.")
    return list(dict.fromkeys(recs))[:8]


def input_profile_table(raw_df: pd.DataFrame) -> pd.DataFrame:
    row = raw_df.iloc[0].to_dict()
    selected = CORE_INPUTS + SITE_INPUTS + RESOURCE_INPUTS + ["avg_floor_plate_m2", "planned_unit_cost_per_m2", "schedule_days_per_1000_m2"]
    rows = []
    for c in selected:
        if c in row:
            val = row[c]
            if isinstance(val, (int, float, np.number)):
                if c in {"planned_cost", "planned_unit_cost_per_m2"}:
                    val = fmt_money(float(val))
                elif c == "planned_duration":
                    val = fmt_days(float(val))
                else:
                    val = f"{float(val):,.2f}"
            rows.append({tr("indicator"): labelize(c), tr("value"): val})
    return pd.DataFrame(rows)


def result_table(result: Dict[str, Any]) -> pd.DataFrame:
    out = result["predicted_outputs"]
    return pd.DataFrame([
        {tr("indicator"): tr("predicted_cost"), tr("value"): fmt_money(out["actual_cost"])},
        {tr("indicator"): tr("predicted_duration"), tr("value"): fmt_days(out["actual_duration"])},
        {tr("indicator"): tr("cost_overrun"), tr("value"): fmt_money(out["cost_overrun"])},
        {tr("indicator"): tr("deviation"), tr("value"): fmt_days(out["schedule_deviation"])},
        {tr("indicator"): tr("cost_overrun_pct"), tr("value"): fmt_pct(out["cost_overrun_percentage"])},
        {tr("indicator"): tr("schedule_overrun_pct"), tr("value"): fmt_pct(out["schedule_overrun_percentage"])},
        {tr("indicator"): tr("cost_scenario"), tr("value"): display_term(out["cost_scenario"])},
        {tr("indicator"): tr("time_scenario"), tr("value"): display_term(out["time_scenario"])},
        {tr("indicator"): tr("risk_level"), tr("value"): display_term(out["risk_level"])},
        {tr("indicator"): tr("risk_score"), tr("value"): f"{safe_float(out['risk_score']):.2f}/100"},
        {tr("indicator"): tr("cost_risk_score"), tr("value"): f"{safe_float(out['cost_risk_score']):.2f}/100"},
        {tr("indicator"): tr("schedule_risk_score"), tr("value"): f"{safe_float(out['schedule_risk_score']):.2f}/100"},
    ])


def _fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def create_pdf_chart_images(raw_df: pd.DataFrame, result: Dict[str, Any]) -> List[bytes]:
    out = result["predicted_outputs"]
    charts: List[bytes] = []
    probs = result.get("risk_probabilities", {})
    labels = list(probs.keys()) if probs else ["Low", "Medium", "High"]
    values = [100 * float(probs[k]) for k in labels] if probs else [0, 0, 0]
    fig1, ax1 = plt.subplots(figsize=(5.8, 3.2))
    bar_colors = []
    for label in labels:
        ll = label.lower()
        bar_colors.append(GREEN if "low" in ll else "#ef4444" if "high" in ll else "#eab308")
    ax1.bar([display_term(x) for x in labels], values, color=bar_colors)
    ax1.set_title(tr("risk_probability"), color=NAVY, fontsize=12, fontweight="bold")
    ax1.set_ylabel(tr("probability"))
    ax1.set_ylim(0, max(100, max(values) + 10 if values else 100))
    for i, v in enumerate(values):
        ax1.text(i, v + 1, f"{v:.1f}%", ha="center", fontsize=9)
    ax1.grid(axis="y", alpha=0.2)
    charts.append(_fig_to_png_bytes(fig1))

    fig2, ax2 = plt.subplots(figsize=(5.8, 3.2))
    score_labels = [tr("risk_score"), tr("cost_risk_score"), tr("schedule_risk_score")]
    score_vals = [safe_float(out.get("risk_score")), safe_float(out.get("cost_risk_score")), safe_float(out.get("schedule_risk_score"))]
    bars = ax2.bar(score_labels, score_vals, color=[GREEN, "#2563eb", TEAL])
    ax2.set_title(tr("score_comparison"), color=NAVY, fontsize=12, fontweight="bold")
    ax2.set_ylabel(tr("score_axis"))
    ax2.set_ylim(0, 105)
    for bar, val in zip(bars, score_vals):
        ax2.text(bar.get_x() + bar.get_width() / 2, val + 1, f"{val:.1f}", ha="center", fontsize=9)
    ax2.grid(axis="y", alpha=0.2)
    charts.append(_fig_to_png_bytes(fig2))

    fig3, ax3 = plt.subplots(figsize=(5.8, 3.2))
    cost_vals = [safe_float(raw_df.iloc[0]["planned_cost"]) * money_factor(), safe_float(out["actual_cost"]) * money_factor()]
    bars3 = ax3.bar([tr("planned"), tr("predicted")], cost_vals, color=["#bdc9d7", GREEN])
    ax3.set_title(tr("cost_chart"), color=NAVY, fontsize=12, fontweight="bold")
    for bar, label in zip(bars3, [fmt_money(raw_df.iloc[0]["planned_cost"]), fmt_money(out["actual_cost"]) ]):
        ax3.text(bar.get_x() + bar.get_width() / 2, bar.get_height() * 1.01, label, ha="center", fontsize=8)
    ax3.grid(axis="y", alpha=0.2)
    charts.append(_fig_to_png_bytes(fig3))

    fig4, ax4 = plt.subplots(figsize=(5.8, 3.2))
    olabels = [tr("cost_overrun"), tr("schedule_overrun")]
    ovals = [100 * safe_float(out.get("cost_overrun_percentage")), 100 * safe_float(out.get("schedule_overrun_percentage"))]
    bars4 = ax4.bar(olabels, ovals, color=[GREEN, TEAL])
    ax4.set_title(tr("overrun_chart"), color=NAVY, fontsize=12, fontweight="bold")
    ax4.set_ylabel("%")
    for bar, val in zip(bars4, ovals):
        ax4.text(bar.get_x() + bar.get_width() / 2, val + 0.2, f"{val:.2f}%", ha="center", fontsize=9)
    ax4.grid(axis="y", alpha=0.2)
    charts.append(_fig_to_png_bytes(fig4))
    return charts


def generate_pdf(raw_df: pd.DataFrame, result: Dict[str, Any], recommendations: List[str]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallNote", fontSize=8.5, leading=11, textColor=colors.HexColor(MUTED)))
    styles.add(ParagraphStyle(name="GreenTitle", parent=styles["Title"], textColor=colors.HexColor(GREEN), fontSize=20, leading=24))
    styles.add(ParagraphStyle(name="GreenHeader", parent=styles["Heading2"], textColor=colors.HexColor(GREEN), fontSize=13, leading=16))
    styles.add(ParagraphStyle(name="BodySmall", parent=styles["BodyText"], fontSize=9.3, leading=12.8, textColor=colors.HexColor(NAVY)))
    story = []
    if os.path.exists(UNIVERSITY_LOGO):
        try:
            story.extend([Image(UNIVERSITY_LOGO, width=3.7 * inch, height=2.07 * inch, kind="proportional"), Spacer(1, 6)])
        except Exception:
            pass
    story.extend([Paragraph(tr("report_title"), styles["GreenTitle"]), Paragraph(tr("report_subtitle"), styles["SmallNote"]), Spacer(1, 8), Paragraph(f"<b>{tr('developed_by')}:</b> {', '.join(STUDENTS)}", styles["BodySmall"]), Paragraph(f"<b>{tr('supervisor')}:</b> {SUPERVISOR}", styles["BodySmall"]), Paragraph(f"<b>{tr('university')}:</b> {UNIVERSITY}", styles["BodySmall"]), Paragraph(f"<b>{tr('currency')}:</b> {st.session_state.currency}" + (f" ({tr('exchange')} {money_factor():.2f})" if st.session_state.currency == "RUB" else ""), styles["BodySmall"]), Paragraph(f"<b>{tr('generated')}:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["BodySmall"]), Spacer(1, 10)])
    out = result["predicted_outputs"]
    summary_data = [[tr("indicator"), tr("prediction_col")], [tr("actual_cost"), fmt_money(out["actual_cost"])], [tr("actual_duration"), fmt_days(out["actual_duration"])], [tr("cost_overrun_pct"), fmt_pct(out["cost_overrun_percentage"])], [tr("schedule_overrun_pct"), fmt_pct(out["schedule_overrun_percentage"])], [tr("cost_scenario"), display_term(out["cost_scenario"])], [tr("time_scenario"), display_term(out["time_scenario"])], [tr("risk_level"), display_term(out["risk_level"])], [tr("risk_score"), f"{safe_float(out['risk_score']):.2f}/100"], [tr("cost_risk_score"), f"{safe_float(out['cost_risk_score']):.2f}/100"], [tr("schedule_risk_score"), f"{safe_float(out['schedule_risk_score']):.2f}/100"]]
    story.append(Paragraph(tr("exec_summary"), styles["GreenHeader"]))
    table = Table(summary_data, colWidths=[2.45 * inch, 3.65 * inch])
    table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(GREEN)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7fbf7")]), ("FONTSIZE", (0, 0), (-1, -1), 9)]))
    story.extend([table, Spacer(1, 12), Paragraph(tr("likely_explanation"), styles["GreenHeader"]), Paragraph(f"<b>{tr('cost_explanation')}:</b> {result.get('cost_primary_cause', '')}; {result.get('cost_secondary_cause', '')}", styles["BodySmall"]), Paragraph(f"<b>{tr('schedule_explanation')}:</b> {result.get('schedule_primary_cause', '')}; {result.get('schedule_secondary_cause', '')}", styles["BodySmall"]), Paragraph(f"<b>{tr('summary')}:</b> {result.get('result_reason_summary', '')}", styles["BodySmall"]), Spacer(1, 12), Paragraph(tr("graphical_dashboard"), styles["GreenHeader"])])
    for idx in range(0, len(create_pdf_chart_images(raw_df, result)), 2):
        imgs = create_pdf_chart_images(raw_df, result)[idx:idx+2]
        row = [Image(io.BytesIO(img), width=3.0 * inch, height=1.75 * inch) for img in imgs]
        if len(row) == 1:
            row.append(Spacer(1, 1.75 * inch))
        grid = Table([row], colWidths=[3.1 * inch, 3.1 * inch])
        grid.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.extend([grid, Spacer(1, 8)])
    story.append(Paragraph(tr("warnings_flags"), styles["GreenHeader"]))
    warnings = result.get("input_warnings", [])
    if warnings:
        for item in warnings:
            story.append(Paragraph(f"• {item}", styles["BodySmall"]))
    else:
        story.append(Paragraph(tr("no_warning"), styles["BodySmall"]))
    for item in result.get("input_driven_flags", []):
        story.append(Paragraph(f"• {item}", styles["BodySmall"]))
    story.extend([Spacer(1, 10), Paragraph(tr("management_recs"), styles["GreenHeader"])])
    for item in recommendations:
        story.append(Paragraph(f"• {item}", styles["BodySmall"]))
    story.append(PageBreak())
    story.append(Paragraph(tr("appendix"), styles["GreenHeader"]))
    rows = [list(input_profile_table(raw_df).columns)] + input_profile_table(raw_df).astype(str).values.tolist()
    tab = Table(rows, colWidths=[2.75 * inch, 3.35 * inch])
    tab.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7fbf7")]), ("FONTSIZE", (0, 0), (-1, -1), 8.5), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.extend([tab, Spacer(1, 10), Paragraph(tr("method_note"), styles["GreenHeader"]), Paragraph(tr("report_method_note"), styles["BodySmall"])])
    doc.build(story)
    return buffer.getvalue()


def render_prediction_results(raw_df: pd.DataFrame, result: Dict[str, Any]) -> None:
    out = result["predicted_outputs"]
    recommendations = build_recommendations(result)
    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.markdown(f'<div class="section-card"><div class="section-title">{tr("results")}</div><div class="section-sub">{tr("results_text")}</div></div>', unsafe_allow_html=True)
    with top_right:
        if st.button(tr("start_new_project"), help=tr("start_new_help"), use_container_width=True):
            reset_project()
            st.rerun()
    a, b, c, d = st.columns(4)
    a.markdown(kpi_card(tr("predicted_cost"), fmt_money(out["actual_cost"]), f"{tr('overrun')}: {fmt_money(out['cost_overrun'])}"), unsafe_allow_html=True)
    b.markdown(kpi_card(tr("predicted_duration"), fmt_days(out["actual_duration"]), f"{tr('deviation')}: {fmt_days(out['schedule_deviation'])}"), unsafe_allow_html=True)
    c.markdown(f'<div class="stat-card"><div class="kpi-label">{tr("risk_level")}</div><div class="kpi-value" style="color:{risk_color(out["risk_level"])}">{display_term(out["risk_level"])}</div><div class="kpi-sub">{tr("risk_score")}: {safe_float(out["risk_score"]):.2f}/100</div></div>', unsafe_allow_html=True)
    d.markdown(kpi_card(tr("scenarios"), f"{display_term(out['cost_scenario'])} / {display_term(out['time_scenario'])}", tr("cost_time_scenario")), unsafe_allow_html=True)
    tabs = st.tabs([tr("tab_dashboard"), tr("tab_interpretation"), tr("tab_input"), tr("tab_pdf")])
    with tabs[0]:
        r1, r2 = st.columns([1, 1])
        r1.plotly_chart(make_gauge(tr("risk_score"), safe_float(out.get("risk_score"))), use_container_width=True)
        r2.plotly_chart(risk_probability_chart(result.get("risk_probabilities", {})), use_container_width=True)
        r3, r4 = st.columns([1, 1])
        r3.plotly_chart(cost_chart(raw_df, out), use_container_width=True)
        r4.plotly_chart(duration_chart(raw_df, out), use_container_width=True)
        r5, r6 = st.columns([1, 1])
        r5.plotly_chart(score_bar_chart(out), use_container_width=True)
        r6.plotly_chart(overrun_chart(out), use_container_width=True)
    with tabs[1]:
        st.markdown(f'<div class="section-card"><div class="section-title">{tr("interpretation")}</div><div class="section-sub">{tr("interpretation_text")}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="green-callout"><b>{tr("summary")}:</b> {result.get("result_reason_summary", "")}</div>', unsafe_allow_html=True)
        e1, e2 = st.columns(2)
        with e1:
            st.subheader(tr("cost_explanation"))
            st.write(f"**{tr('primary_cause')}:** {result.get('cost_primary_cause', '')}")
            st.write(f"**{tr('secondary_cause')}:** {result.get('cost_secondary_cause', '')}")
        with e2:
            st.subheader(tr("schedule_explanation"))
            st.write(f"**{tr('primary_cause')}:** {result.get('schedule_primary_cause', '')}")
            st.write(f"**{tr('secondary_cause')}:** {result.get('schedule_secondary_cause', '')}")
        warnings = result.get("input_warnings", [])
        if warnings:
            st.subheader(tr("input_warnings"))
            for item in warnings:
                st.markdown(f'<div class="warn-box">{item}</div>', unsafe_allow_html=True)
        st.subheader(tr("recommendations"))
        for rec in recommendations:
            st.markdown(f'<div class="green-callout">{rec}</div>', unsafe_allow_html=True)
    with tabs[2]:
        st.subheader(tr("submitted_profile"))
        st.dataframe(input_profile_table(raw_df), use_container_width=True, hide_index=True)
        st.subheader(tr("output_table"))
        st.dataframe(result_table(result), use_container_width=True, hide_index=True)
    with tabs[3]:
        pdf_bytes = generate_pdf(raw_df, result, recommendations)
        st.success(tr("pdf_success"))
        st.download_button(tr("download_pdf"), data=pdf_bytes, file_name=f"building_prediction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", use_container_width=True)
        st.info(tr("pdf_info"))


def render_prediction_page(system: BuildingPredictor) -> None:
    st.markdown(f'<div class="section-card"><div class="section-title">{tr("workspace")}</div><div class="section-sub">{tr("workspace_text")}</div></div>', unsafe_allow_html=True)
    info1, info2, info3 = st.columns(3)
    info1.markdown(f'<div class="feature-card"><div class="feature-icon">🧮</div><div class="feature-title">{tr("structured_input")}</div><div class="feature-text">{tr("structured_input_text")}</div></div>', unsafe_allow_html=True)
    info2.markdown(f'<div class="feature-card"><div class="feature-icon">📈</div><div class="feature-title">{tr("analytical_output")}</div><div class="feature-text">{tr("analytical_output_text")}</div></div>', unsafe_allow_html=True)
    info3.markdown(f'<div class="feature-card"><div class="feature-icon">📄</div><div class="feature-title">{tr("ready_report")}</div><div class="feature-text">{tr("ready_report_text")}</div></div>', unsafe_allow_html=True)
    raw_df, submitted = build_input_form(system)
    if submitted:
        with st.spinner(tr("running")):
            result = system.predict(raw_df)
        st.session_state.raw_df = raw_df
        st.session_state.prediction_result = result
    if st.session_state.prediction_result is None:
        st.info(tr("complete_form"))
        return
    render_prediction_results(st.session_state.raw_df, st.session_state.prediction_result)


def main() -> None:
    init_state()
    inject_css()
    with st.spinner(tr("loading")):
        try:
            system = load_system()
        except Exception:
            st.error(tr("load_error"))
            with st.expander(tr("technical_details")):
                st.code(traceback.format_exc())
            return
    render_sidebar()
    if st.session_state.page == "Home":
        render_home_page()
    elif st.session_state.page == "About":
        render_about_page()
    else:
        try:
            render_prediction_page(system)
        except Exception:
            st.error(tr("prediction_error"))
            with st.expander(tr("technical_details")):
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()

# AI Business Combine — однофайловое приложение.
# Установка:  pip install streamlit pandas openpyxl python-dotenv
# Секреты:    скопируйте .env.example в .env и заполните ключи
# Запуск:     streamlit run app.py

from __future__ import annotations

import hmac
import io
import json
import os
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------------------------------------------------------------------
# Конфиг (секреты только из окружения / .env, никогда из виджетов)
# ---------------------------------------------------------------------------

load_dotenv(Path(__file__).resolve().parent / ".env")

APP_NAME = "AI Business Combine"
KILL_SWITCH_LABELS = {
    "ru": {"cogs": "себестоимость (COGS)", "cac": "реклама (CAC)", "price": "цена"},
    "en": {"cogs": "COGS", "cac": "CAC / ads", "price": "price"},
}

SYSTEM_PROMPT = """Ты — Генеральный AI-ментор комбайна AI Business Combine.
Помогаешь предпринимателю разобрать идею, юнит-экономику, маркетинг и риски.
Отвечай на языке пользователя. Будь конкретным: цифры, формулы, следующие шаги.
Если в контексте есть расчёт калькулятора или данные Excel — опирайся на них.
Не выдумывай факты о рынке без оговорки, что это гипотеза."""


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def llm_api_key() -> str:
    return _env("LLM_API_KEY") or _env("OPENAI_API_KEY")


def llm_base_url() -> str:
    return _env("LLM_BASE_URL", "https://api.openai.com/v1")


def llm_model() -> str:
    return _env("LLM_MODEL", "gpt-4o-mini")


def database_url() -> str:
    return _env("DATABASE_URL")


def app_username() -> str:
    return _env("APP_USERNAME")


def app_password() -> str:
    return _env("APP_PASSWORD")


def secrets_configured() -> bool:
    return bool(app_username() and app_password())


def _const_eq(left: str, right: str) -> bool:
    digest_left = hmac.new(b"combine-auth", left.encode("utf-8"), "sha256").digest()
    digest_right = hmac.new(b"combine-auth", right.encode("utf-8"), "sha256").digest()
    return hmac.compare_digest(digest_left, digest_right)


def credentials_match(user: str, password: str) -> bool:
    expected_user = app_username()
    expected_pass = app_password()
    if not expected_user or not expected_pass:
        return False
    return _const_eq(user, expected_user) and _const_eq(password, expected_pass)


# ---------------------------------------------------------------------------
# Юнит-экономика (реальные формулы, без заглушек)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class UnitInputs:
    cogs: float
    price: float
    cac: float
    fixed_costs: float
    units_per_month: float


@dataclass(frozen=True)
class UnitResult:
    revenue: float
    contribution_per_unit: float
    contribution_total: float
    gross_profit: float
    net_profit: float
    gross_margin: float
    net_margin: float
    roi: float
    break_even_units: float | None
    break_even_revenue: float | None
    kill_switch: str
    kill_switch_label: str
    profitable: bool


def calc_unit(inputs: UnitInputs, locale: str = "ru") -> UnitResult:
    price = float(inputs.price)
    cogs = float(inputs.cogs)
    cac = float(inputs.cac)
    units = float(inputs.units_per_month)
    fixed_costs = float(inputs.fixed_costs)

    contribution = price - cogs - cac
    revenue = price * units
    gross_profit = (price - cogs) * units
    contribution_total = contribution * units
    net_profit = contribution_total - fixed_costs
    gross_margin = ((price - cogs) / price) if price else 0.0
    net_margin = (net_profit / revenue) if revenue else 0.0
    invested = cac * units + fixed_costs
    roi = (net_profit / invested) if invested else 0.0

    if contribution > 0:
        break_even_units = fixed_costs / contribution
        break_even_revenue = break_even_units * price
    else:
        break_even_units = None
        break_even_revenue = None

    if price <= 0:
        kill = "price"
    else:
        shares = {"cogs": cogs / price, "cac": cac / price}
        kill = max(shares, key=shares.get)

    labels = KILL_SWITCH_LABELS.get(locale, KILL_SWITCH_LABELS["ru"])
    return UnitResult(
        revenue=round(revenue, 2),
        contribution_per_unit=round(contribution, 4),
        contribution_total=round(contribution_total, 2),
        gross_profit=round(gross_profit, 2),
        net_profit=round(net_profit, 2),
        gross_margin=round(gross_margin, 6),
        net_margin=round(net_margin, 6),
        roi=round(roi, 6),
        break_even_units=round(break_even_units, 2) if break_even_units is not None else None,
        break_even_revenue=round(break_even_revenue, 2) if break_even_revenue is not None else None,
        kill_switch=kill,
        kill_switch_label=labels[kill],
        profitable=net_profit > 0,
    )


def scenario_inputs(base: UnitInputs) -> dict[str, UnitInputs]:
    return {
        "pessimistic": UnitInputs(
            cogs=base.cogs * 1.15,
            price=base.price * 0.95,
            cac=base.cac * 1.25,
            fixed_costs=base.fixed_costs,
            units_per_month=base.units_per_month * 0.7,
        ),
        "base": base,
        "optimistic": UnitInputs(
            cogs=base.cogs * 0.9,
            price=base.price * 1.05,
            cac=base.cac * 0.8,
            fixed_costs=base.fixed_costs,
            units_per_month=base.units_per_month * 1.3,
        ),
    }


def calc_scenarios(base: UnitInputs, locale: str = "ru") -> dict[str, dict[str, Any]]:
    out: dict[str, dict[str, Any]] = {}
    for name, inp in scenario_inputs(base).items():
        result = calc_unit(inp, locale=locale)
        out[name] = {"inputs": asdict(inp), "result": asdict(result)}
    return out


# ---------------------------------------------------------------------------
# LLM: OpenAI-compatible Chat Completions (реальный HTTP, без заглушек)
# ---------------------------------------------------------------------------

def call_llm(
    messages: list[dict[str, str]],
    *,
    api_key: str,
    base_url: str,
    model: str,
    timeout: int = 90,
) -> str:
    if not api_key.strip():
        raise RuntimeError(
            "LLM API-ключ не задан. Пропишите LLM_API_KEY или OPENAI_API_KEY в файле .env."
        )

    url = base_url.rstrip("/") + "/chat/completions"
    payload = json.dumps(
        {
            "model": model,
            "messages": messages,
            "temperature": 0.4,
        }
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=payload,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key.strip()}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"LLM HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Не удалось связаться с LLM API: {exc.reason}") from exc

    choices = body.get("choices") or []
    if not choices:
        raise RuntimeError(f"Пустой ответ LLM: {body}")
    content = choices[0].get("message", {}).get("content")
    if not content:
        raise RuntimeError(f"В ответе LLM нет текста: {body}")
    return str(content).strip()


def context_block() -> str:
    parts: list[str] = []
    calc = st.session_state.get("last_calc")
    if calc:
        parts.append("Последний расчёт юнит-экономики:\n" + json.dumps(calc, ensure_ascii=False, indent=2))
    excel = st.session_state.get("excel_brief")
    if excel:
        parts.append("Краткая сводка загруженного Excel:\n" + excel)
    return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# Excel: отчёт калькулятора + разбор загруженного файла
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
GOOD_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")


def build_calculator_xlsx(calc: dict[str, Any]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Сценарии"

    headers = [
        "Сценарий",
        "Цена",
        "COGS",
        "CAC",
        "Постоянные",
        "Ед./мес",
        "Выручка",
        "Маржа/ед",
        "Валовая прибыль",
        "Чистая прибыль",
        "Gross margin",
        "Net margin",
        "ROI",
        "Точка безубыточности, ед",
        "Точка безубыточности, выручка",
        "Kill-switch",
        "Прибыльный",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    names = {"pessimistic": "Пессимистичный", "base": "Базовый", "optimistic": "Оптимистичный"}
    for key in ("pessimistic", "base", "optimistic"):
        block = calc[key]
        inp, res = block["inputs"], block["result"]
        row = [
            names[key],
            inp["price"],
            inp["cogs"],
            inp["cac"],
            inp["fixed_costs"],
            inp["units_per_month"],
            res["revenue"],
            res["contribution_per_unit"],
            res["gross_profit"],
            res["net_profit"],
            res["gross_margin"],
            res["net_margin"],
            res["roi"],
            res["break_even_units"] if res["break_even_units"] is not None else "не достигается",
            res["break_even_revenue"] if res["break_even_revenue"] is not None else "не достигается",
            res["kill_switch_label"],
            "да" if res["profitable"] else "нет",
        ]
        ws.append(row)
        profit_cell = ws.cell(ws.max_row, 17)
        profit_cell.fill = GOOD_FILL if res["profitable"] else BAD_FILL

    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 36)

    chart = BarChart()
    chart.title = "Чистая прибыль по сценариям"
    chart.y_axis.title = "Деньги"
    data = Reference(ws, min_col=10, min_row=1, max_row=4)
    cats = Reference(ws, min_col=1, min_row=2, max_row=4)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, "A8")

    formulas = wb.create_sheet("Формулы")
    formulas.append(["Метрика", "Формула"])
    for cell in formulas[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    rows = [
        ("Выручка", "price × units_per_month"),
        ("Маржа с единицы (contribution)", "price − COGS − CAC"),
        ("Валовая прибыль", "(price − COGS) × units_per_month"),
        ("Чистая прибыль", "contribution × units − fixed_costs"),
        ("Gross margin", "(price − COGS) / price"),
        ("Net margin", "net_profit / revenue"),
        ("ROI", "net_profit / (CAC × units + fixed_costs)"),
        ("Точка безубыточности, ед", "fixed_costs / contribution  (если contribution > 0)"),
        ("Пессимистичный сценарий", "COGS ×1.15, цена ×0.95, CAC ×1.25, объём ×0.70"),
        ("Оптимистичный сценарий", "COGS ×0.90, цена ×1.05, CAC ×0.80, объём ×1.30"),
        ("Kill-switch", "наибольшая доля в цене: COGS или CAC (если цена ≤ 0 — цена)"),
    ]
    for r in rows:
        formulas.append(list(r))
    formulas.column_dimensions["A"].width = 34
    formulas.column_dimensions["B"].width = 70

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def analyze_excel(file_bytes: bytes, filename: str) -> tuple[dict[str, pd.DataFrame], str, bytes]:
    bio = io.BytesIO(file_bytes)
    xl = pd.ExcelFile(bio, engine="openpyxl")
    sheets: dict[str, pd.DataFrame] = {}
    brief_lines = [f"Файл: {filename}", f"Листы: {', '.join(xl.sheet_names)}"]

    report = Workbook()
    summary = report.active
    summary.title = "Сводка"
    summary.append(["Лист", "Строк", "Столбцов", "Числовых колонок", "Пропусков", "Дубликатов"])
    for cell in summary[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        df.columns = [str(c) for c in df.columns]
        sheets[sheet_name] = df
        numeric = df.select_dtypes(include="number")
        missing = int(df.isna().sum().sum())
        dups = int(df.duplicated().sum())
        summary.append([sheet_name, len(df), df.shape[1], numeric.shape[1], missing, dups])

        brief_lines.append(
            f"— {sheet_name}: {len(df)}×{df.shape[1]}, "
            f"числовых={numeric.shape[1]}, пропусков={missing}, дублей={dups}"
        )
        if not numeric.empty:
            describe = numeric.describe().T
            for col, row in describe.iterrows():
                brief_lines.append(
                    f"   {col}: mean={row.get('mean', float('nan')):.4g}, "
                    f"min={row.get('min', float('nan')):.4g}, max={row.get('max', float('nan')):.4g}"
                )

        ws = report.create_sheet(sheet_name[:31] or "Sheet")
        preview = df.head(200)
        for r in dataframe_to_rows(preview, index=False, header=True):
            ws.append(list(r))
        for cell in ws[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT

        if not numeric.empty:
            stats_ws = report.create_sheet((sheet_name[:24] + "_stats")[:31])
            stats = numeric.describe().T.reset_index().rename(columns={"index": "column"})
            for r in dataframe_to_rows(stats, index=False, header=True):
                stats_ws.append(list(r))
            for cell in stats_ws[1]:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT

    for col in summary.columns:
        summary.column_dimensions[col[0].column_letter].width = 22

    out = io.BytesIO()
    report.save(out)
    return sheets, "\n".join(brief_lines), out.getvalue()


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

def money(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v:,.2f}".replace(",", " ")


def pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{v * 100:.1f}%"


def init_state() -> None:
    st.session_state.setdefault("authenticated", False)
    st.session_state.setdefault("messages", [])
    st.session_state.setdefault("last_calc", None)
    st.session_state.setdefault("excel_brief", None)
    st.session_state.setdefault("excel_sheets", None)
    st.session_state.setdefault("excel_report", None)
    st.session_state.setdefault("excel_name", None)


def render_login() -> bool:
    """Стартовый экран. Пока нет входа — чат и API недоступны."""
    st.title(APP_NAME)
    st.subheader("Вход")
    st.caption("Доступ только для авторизованных пользователей. Это защищает баланс нейросети.")

    if not secrets_configured():
        st.error(
            "Логин и пароль не заданы. Добавьте APP_USERNAME и APP_PASSWORD в файл .env "
            "и перезапустите приложение."
        )
        st.stop()

    with st.form("login_form", clear_on_submit=False):
        user = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти", type="primary", use_container_width=True)

    if submitted:
        if credentials_match(user.strip(), password):
            st.session_state.authenticated = True
            st.rerun()
        st.error("Неверный логин или пароль.")

    return False


def render_metrics(result: dict[str, Any]) -> None:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Выручка / мес", money(result["revenue"]))
    c2.metric("Чистая прибыль", money(result["net_profit"]))
    c3.metric("ROI", pct(result["roi"]))
    c4.metric("Gross margin", pct(result["gross_margin"]))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Маржа с единицы", money(result["contribution_per_unit"]))
    c6.metric("Net margin", pct(result["net_margin"]))
    be_u = result["break_even_units"]
    be_r = result["break_even_revenue"]
    c7.metric("Безубыточность, ед", "∞" if be_u is None else money(be_u))
    c8.metric("Kill-switch", result["kill_switch_label"])
    if result["profitable"]:
        st.success("Модель прибыльна на горизонте месяца при заданных допущениях.")
    else:
        st.error("Модель убыточна. Смотрите kill-switch — это главный рычаг.")
    if be_u is None:
        st.warning(
            "Точка безубыточности не достигается: маржа с единицы ≤ 0 "
            "(цена не покрывает COGS + CAC). Сначала чините юнит, потом масштаб."
        )
    else:
        st.caption(f"Выручка в точке безубыточности: {money(be_r)}")


# ---------------------------------------------------------------------------
# Страницы
# ---------------------------------------------------------------------------

def page_chat() -> None:
    st.subheader("Генеральный Чат")
    st.caption("Ментор по идее, экономике и рискам. Подключён к реальному LLM API.")

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    prompt = st.chat_input("Опишите идею, задайте вопрос по цифрам или попросите разбор…")
    if not prompt:
        return

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    history = [{"role": "system", "content": SYSTEM_PROMPT}]
    ctx = context_block()
    if ctx:
        history.append({"role": "system", "content": "Контекст сессии:\n" + ctx})
    for msg in st.session_state.messages:
        history.append({"role": msg["role"], "content": msg["content"]})

    with st.chat_message("assistant"):
        with st.spinner("Думаю…"):
            try:
                reply = call_llm(
                    history,
                    api_key=llm_api_key(),
                    base_url=llm_base_url(),
                    model=llm_model(),
                )
            except Exception as exc:
                reply = f"**Ошибка API:** {exc}"
        st.markdown(reply)

    st.session_state.messages.append({"role": "assistant", "content": reply})


def page_calculator() -> None:
    st.subheader("Бизнес-Калькулятор")
    st.caption(
        "Юнит-экономика: contribution = цена − COGS − CAC; "
        "чистая прибыль = contribution × объём − постоянные."
    )

    left, right = st.columns([1, 1.2])
    with left:
        st.markdown("**Допущения (базовый сценарий)**")
        price = st.number_input("Цена за единицу", min_value=0.01, value=100.0, step=1.0)
        cogs = st.number_input("Себестоимость (COGS)", min_value=0.0, value=40.0, step=1.0)
        cac = st.number_input("CAC / привлечение на продажу", min_value=0.0, value=15.0, step=1.0)
        fixed_costs = st.number_input("Постоянные расходы в месяц", min_value=0.0, value=50000.0, step=1000.0)
        units = st.number_input("Ожидаемый объём, ед./мес", min_value=0.01, value=2000.0, step=10.0)
        run = st.button("Рассчитать три сценария", type="primary", use_container_width=True)

    if run or st.session_state.last_calc is None:
        inputs = UnitInputs(
            cogs=cogs,
            price=price,
            cac=cac,
            fixed_costs=fixed_costs,
            units_per_month=units,
        )
        st.session_state.last_calc = {
            "calculated_at": datetime.now().isoformat(timespec="seconds"),
            **calc_scenarios(inputs),
        }

    calc = st.session_state.last_calc
    with right:
        tabs = st.tabs(["Базовый", "Оптимистичный", "Пессимистичный"])
        mapping = {"Базовый": "base", "Оптимистичный": "optimistic", "Пессимистичный": "pessimistic"}
        for tab, key in zip(tabs, mapping):
            with tab:
                render_metrics(calc[mapping[key]]["result"])

    chart_df = pd.DataFrame(
        {
            "Сценарий": ["Пессимистичный", "Базовый", "Оптимистичный"],
            "Выручка": [
                calc["pessimistic"]["result"]["revenue"],
                calc["base"]["result"]["revenue"],
                calc["optimistic"]["result"]["revenue"],
            ],
            "Чистая прибыль": [
                calc["pessimistic"]["result"]["net_profit"],
                calc["base"]["result"]["net_profit"],
                calc["optimistic"]["result"]["net_profit"],
            ],
        }
    )
    st.bar_chart(chart_df.set_index("Сценарий"))

    xlsx = build_calculator_xlsx(calc)
    st.download_button(
        "Скачать Excel-отчёт по сценариям",
        data=xlsx,
        file_name="unit_economics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def page_excel() -> None:
    st.subheader("Аналитика Excel")
    st.caption("Загрузите .xlsx / .xls — разбор листов, статистика, отчёт и AI-комментарий.")

    uploaded = st.file_uploader("Файл Excel", type=["xlsx", "xlsm"])
    if uploaded is not None:
        raw = uploaded.getvalue()
        try:
            sheets, brief, report_bytes = analyze_excel(raw, uploaded.name)
        except Exception as exc:
            st.error(f"Не удалось прочитать файл: {exc}")
            return
        st.session_state.excel_sheets = sheets
        st.session_state.excel_brief = brief
        st.session_state.excel_report = report_bytes
        st.session_state.excel_name = uploaded.name

    sheets: dict[str, pd.DataFrame] | None = st.session_state.excel_sheets
    if not sheets:
        st.info("Загрузите таблицу продаж, P&L, юниты или любой рабочий Excel.")
        return

    st.success(f"Разобран файл: {st.session_state.excel_name}")
    st.code(st.session_state.excel_brief, language=None)

    names = list(sheets.keys())
    chosen = st.selectbox("Лист", names)
    df = sheets[chosen]
    st.dataframe(df, use_container_width=True, height=360)

    numeric = df.select_dtypes(include="number")
    if not numeric.empty:
        st.markdown("**Описательная статистика**")
        st.dataframe(numeric.describe().T, use_container_width=True)
        st.markdown("**Пропуски по колонкам**")
        miss = df.isna().sum().to_frame("пропусков")
        st.dataframe(miss[miss["пропусков"] > 0] if miss["пропусков"].sum() else miss, use_container_width=True)
        if numeric.shape[1] >= 2:
            st.markdown("**Корреляции**")
            st.dataframe(numeric.corr(numeric_only=True), use_container_width=True)
        chart_cols = st.multiselect(
            "Колонки для графика",
            list(numeric.columns),
            default=list(numeric.columns)[: min(3, numeric.shape[1])],
        )
        if chart_cols:
            st.line_chart(numeric[chart_cols])
    else:
        st.warning("На этом листе нет числовых колонок — показан только предпросмотр.")

    if st.session_state.excel_report:
        st.download_button(
            "Скачать аналитический отчёт (.xlsx)",
            data=st.session_state.excel_report,
            file_name="excel_analytics_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if st.button("Спросить AI по этой таблице", type="primary"):
        sample = df.head(30).to_csv(index=False)
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    "Проанализируй Excel. Дай: что это за данные, аномалии, "
                    "3 вывода для бизнеса и 3 следующих действия.\n\n"
                    f"{st.session_state.excel_brief}\n\nПервые строки листа «{chosen}»:\n{sample}"
                ),
            },
        ]
        with st.spinner("AI разбирает таблицу…"):
            try:
                comment = call_llm(
                    messages,
                    api_key=llm_api_key(),
                    base_url=llm_base_url(),
                    model=llm_model(),
                )
                st.markdown(comment)
            except Exception as exc:
                st.error(str(exc))


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="📊", layout="wide")
    init_state()

    if not st.session_state.authenticated:
        render_login()
        return

    st.sidebar.title(APP_NAME)
    st.sidebar.caption("Чат · Калькулятор · Excel")
    if llm_api_key():
        st.sidebar.caption("Нейросеть подключена из окружения.")
    else:
        st.sidebar.warning("LLM_API_KEY не найден в .env — чат недоступен.")

    if st.sidebar.button("Очистить чат"):
        st.session_state.messages = []
        st.rerun()
    if st.sidebar.button("Выйти"):
        st.session_state.authenticated = False
        st.session_state.messages = []
        st.rerun()

    page = st.sidebar.radio(
        "Раздел",
        ("Генеральный Чат", "Бизнес-Калькулятор", "Аналитика Excel"),
    )

    st.title(APP_NAME)
    if page == "Генеральный Чат":
        page_chat()
    elif page == "Бизнес-Калькулятор":
        page_calculator()
    else:
        page_excel()


if __name__ == "__main__":
    main()

import streamlit as st
import requests
import json
import os
from openpyxl import Workbook

# Намертво вшитые настройки ИИ, чтобы обойти баги облака
LLM_API_KEY = "hf_jSebPXvGfbvRwDLIGwDKZqrmXMCkpGvsErLLM"
LLM_BASE_URL = "https://huggingface.co"
LLM_MODEL = "Qwen/Qwen2.5-72B-Instruct"

# Инициализация сессии
if "messages" not in st.session_state:
    st.session_state.messages = []
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# Функция авторизации
def check_password():
    if st.session_state.authenticated:
        return True
    
    st.title("Вход в AI Business Combine")
    username = st.text_input("Логин")
    password = st.text_input("Пароль", type="password")
    
    if st.button("Войти"):
        if username == "admin" and password == "change-me-now":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Неверный логин или пароль")
    return False

if check_password():
    # Боковая панель
    with st.sidebar:
        st.title("AI Business Combine")
        st.write("---")
        page = st.sidebar.radio("Раздел", ["Генеральный Чат", "Бизнес-Калькулятор", "Аналитика Excel"])
        st.write("---")
        if st.sidebar.button("Очистить чат"):
            st.session_state.messages = []
            st.rerun()

    # Главный раздел: Чат
    if page == "Генеральный Чат":
        st.title("Генеральный Чат")
        st.caption("Ментор по идее, экономике и рискам. Подключён к стабильному LLM API.")
        
        # Отображение истории сообщений
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        
        # Поле ввода нового сообщения
        if prompt := st.chat_input("Опишите идею, задайте вопрос по цифрам или попросите разбор..."):
            with st.chat_message("user"):
                st.write(prompt)
            st.session_state.messages.append({"role": "user", "content": prompt})
            
            # Отправка запроса к ИИ
            headers = {"Authorization": f"Bearer {LLM_API_KEY}", "Content-Type": "application/json"}
            payload = {
                "model": LLM_MODEL,
                "messages": st.session_state.messages,
                "max_tokens": 1000,
                "temperature": 0.7
            }
            
            try:
                response = requests.post(f"{LLM_BASE_URL}/chat/completions", headers=headers, json=payload, timeout=30)
                if response.status_code == 200:
                    ai_response = response.json()['choices']['message']['content']
                    with st.chat_message("assistant"):
                        st.write(ai_response)
                    st.session_state.messages.append({"role": "assistant", "content": ai_response})
                else:
                    st.error(f"Ошибка API: Сервер вернул код {response.status_code}")
            except Exception as e:
                st.error(f"Ошибка связи с ИИ: {str(e)}")

    # Раздел: Бизнес-Калькулятор
    elif page == "Бизнес-Калькулятор":
        st.title("Финансовый Калькулятор")
        st.info("Раздел в разработке. Скоро здесь появится расчет юнит-экономики.")

    # Раздел: Аналитика Excel
    elif page == "Аналитика Excel":
        st.title("Аналитика Excel")
        st.info("Раздел в разработке. Скоро здесь появится генерация сводных таблиц.")

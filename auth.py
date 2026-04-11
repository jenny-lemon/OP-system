import os
import streamlit as st
import requests
from bs4 import BeautifulSoup

LOGIN_URL = "https://backend.lemonclean.com.tw/login"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded",
}

def get_account(city: str):
    mapping = {
        "台北": "taipei",
        "台中": "taichung",
    }

    key = mapping.get(city)
    if not key:
        raise ValueError(f"未知城市: {city}")

    return {
        "email": st.secrets["accounts"][key]["email"],
        "password": st.secrets["accounts"][key]["password"],
    }


def login(email: str, password: str) -> requests.Session:
    session = requests.Session()

    res = session.get(LOGIN_URL, headers=HEADERS)
    soup = BeautifulSoup(res.text, "html.parser")

    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise RuntimeError("找不到 _token")

    csrf_token = token_input.get("value")

    payload = {
        "_token": csrf_token,
        "email": email,
        "password": password,
    }

    login_res = session.post(LOGIN_URL, data=payload, headers=HEADERS)

    if "login" in login_res.url.lower():
        raise RuntimeError("登入失敗")

    return session

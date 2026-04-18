import os

try:
    import streamlit as st
except Exception:
    st = None


def _get_secret(path, default=None):
    if st is not None:
        try:
            value = st.secrets
            for key in path:
                value = value[key]
            return value
        except Exception:
            pass
    return default


def _get_env(name, default=None):
    return os.getenv(name, default)


def _pick(secret_path, env_name):
    return _get_secret(secret_path) or _get_env(env_name)


ACCOUNTS = {}

taipei_email = _pick(["accounts", "taipei", "email"], "TAIPEI_EMAIL")
taipei_password = _pick(["accounts", "taipei", "password"], "TAIPEI_PASSWORD")
if taipei_email and taipei_password:
    ACCOUNTS["台北"] = {"email": taipei_email, "password": taipei_password}

taichung_email = _pick(["accounts", "taichung", "email"], "TAICHUNG_EMAIL")
taichung_password = _pick(["accounts", "taichung", "password"], "TAICHUNG_PASSWORD")
if taichung_email and taichung_password:
    ACCOUNTS["台中"] = {"email": taichung_email, "password": taichung_password}

taoyuan_email = _pick(["accounts", "taoyuan", "email"], "TAOYUAN_EMAIL")
taoyuan_password = _pick(["accounts", "taoyuan", "password"], "TAOYUAN_PASSWORD")
if taoyuan_email and taoyuan_password:
    ACCOUNTS["桃園"] = {"email": taoyuan_email, "password": taoyuan_password}

hsinchu_email = _pick(["accounts", "hsinchu", "email"], "HSINCHU_EMAIL")
hsinchu_password = _pick(["accounts", "hsinchu", "password"], "HSINCHU_PASSWORD")
if hsinchu_email and hsinchu_password:
    ACCOUNTS["新竹"] = {"email": hsinchu_email, "password": hsinchu_password}

kaohsiung_email = _pick(["accounts", "kaohsiung", "email"], "KAOHSIUNG_EMAIL")
kaohsiung_password = _pick(["accounts", "kaohsiung", "password"], "KAOHSIUNG_PASSWORD")
if kaohsiung_email and kaohsiung_password:
    ACCOUNTS["高雄"] = {"email": kaohsiung_email, "password": kaohsiung_password}

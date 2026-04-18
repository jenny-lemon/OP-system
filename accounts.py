import streamlit as st

def _get_secret(path, default=None):
    try:
        value = st.secrets
        for key in path:
            value = value[key]
        return value
    except Exception:
        return default


ACCOUNTS = {}

# 台北
taipei_email = _get_secret(["accounts", "taipei", "email"])
taipei_password = _get_secret(["accounts", "taipei", "password"])
if taipei_email and taipei_password:
    ACCOUNTS["台北"] = {
        "email": taipei_email,
        "password": taipei_password,
    }

# 台中
taichung_email = _get_secret(["accounts", "taichung", "email"])
taichung_password = _get_secret(["accounts", "taichung", "password"])
if taichung_email and taichung_password:
    ACCOUNTS["台中"] = {
        "email": taichung_email,
        "password": taichung_password,
    }

# 桃園
taoyuan_email = _get_secret(["accounts", "taoyuan", "email"])
taoyuan_password = _get_secret(["accounts", "taoyuan", "password"])
if taoyuan_email and taoyuan_password:
    ACCOUNTS["桃園"] = {
        "email": taoyuan_email,
        "password": taoyuan_password,
    }

# 新竹
hsinchu_email = _get_secret(["accounts", "hsinchu", "email"])
hsinchu_password = _get_secret(["accounts", "hsinchu", "password"])
if hsinchu_email and hsinchu_password:
    ACCOUNTS["新竹"] = {
        "email": hsinchu_email,
        "password": hsinchu_password,
    }

# 高雄
kaohsiung_email = _get_secret(["accounts", "kaohsiung", "email"])
kaohsiung_password = _get_secret(["accounts", "kaohsiung", "password"])
if kaohsiung_email and kaohsiung_password:
    ACCOUNTS["高雄"] = {
        "email": kaohsiung_email,
        "password": kaohsiung_password,
    }

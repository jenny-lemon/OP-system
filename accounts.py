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
            return default
    return default


def _get_env(name, default=None):
    return os.getenv(name, default)


def _first_secret(paths):
    for path in paths:
        value = _get_secret(path)
        if value not in (None, ""):
            return value
    return None


def _pick(city_key, field, env_name):
    """Read both the current nested secrets and legacy flat/city formats."""
    upper_city = city_key.upper()
    upper_field = field.upper()
    return (
        _first_secret([
            ["accounts", city_key, field],
            [city_key, field],
            [f"{upper_city}_{upper_field}"],
            [f"{city_key}_{field}"],
        ])
        or _get_env(env_name)
    )


ACCOUNTS = {}

_CITY_CONFIG = {
    "台北": ("taipei", "TAIPEI_EMAIL", "TAIPEI_PASSWORD"),
    "台中": ("taichung", "TAICHUNG_EMAIL", "TAICHUNG_PASSWORD"),
    "桃園": ("taoyuan", "TAOYUAN_EMAIL", "TAOYUAN_PASSWORD"),
    "新竹": ("hsinchu", "HSINCHU_EMAIL", "HSINCHU_PASSWORD"),
    "高雄": ("kaohsiung", "KAOHSIUNG_EMAIL", "KAOHSIUNG_PASSWORD"),
}

for city, (city_key, email_env, password_env) in _CITY_CONFIG.items():
    email = _pick(city_key, "email", email_env)
    password = _pick(city_key, "password", password_env)
    if email and password:
        ACCOUNTS[city] = {
            "email": str(email).strip(),
            "password": str(password),
        }

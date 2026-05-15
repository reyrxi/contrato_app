import re

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def validar_cpf(cpf):
    digits = re.sub(r"\D", "", cpf)
    if len(digits) != 11 or digits == digits[0] * 11:
        return False
    for i in range(9, 11):
        total = sum(int(digits[j]) * (i + 1 - j) for j in range(i))
        if (total * 10 % 11) % 10 != int(digits[i]):
            return False
    return True


def fmt_cpf(value):
    digits = re.sub(r"\D", "", value)[:11]
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return value


def fmt_cep(value):
    digits = re.sub(r"\D", "", value)[:8]
    if len(digits) == 8:
        return f"{digits[:5]}-{digits[5:]}"
    return value


def fmt_fone(value):
    digits = re.sub(r"\D", "", value)[:11]
    if len(digits) == 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return value


def fmt_cur(value):
    cleaned = value.strip().replace("R$", "").replace(".", "").replace(",", ".")
    try:
        amount = float(cleaned)
        return f"R$ {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return value


def buscar_cep(cep):
    if not HAS_REQUESTS:
        return None
    digits = re.sub(r"\D", "", cep)
    if len(digits) != 8:
        return None
    try:
        response = requests.get(
            f"https://viacep.com.br/ws/{digits}/json/",
            timeout=5,
        )
        data = response.json()
        return None if "erro" in data else data
    except Exception:
        return None

import json, sys
from pathlib import Path

def hs_mapping_path() -> Path:
    # файл рядом с exe/скриптом
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).with_name("hs_mapping.json")
    return Path(__file__).with_name("hs_mapping.json")

def load_hs_mapping() -> dict[str, dict[str, str]]:
    """
    Возвращает { 'Имя ХС': {'uuid': '...', 'schema': '...'}, ... }.
    Если в json элемент был строкой — считаем это UUID, а schema потребуется в другом элементе
    или будет отсутствовать (тогда упадём с понятной ошибкой при запуске).
    """

    p = hs_mapping_path()

    data = json.loads(p.read_text(encoding="utf-8"))
    norm: dict[str, dict[str, str]] = {}
    for name, v in data.items():
        if isinstance(v, str):
            norm[str(name)] = {"uuid": v, "schema": data.get("_schema_default", "")}  # опц. дефолт из json
        elif isinstance(v, dict):
            uuid = str(v.get("uuid") or v.get("id") or v.get("HS_ID") or "").strip()
            schema = str(v.get("schema") or v.get("schema_value") or "").strip()
            if uuid:
                norm[str(name)] = {"uuid": uuid, "schema": schema}
    return norm

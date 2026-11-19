from hashlib import file_digest
from functools import lru_cache
import csv
import sys
from collections import Counter
from copy import deepcopy
import json
import re
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory
import os
import subprocess
import shutil

import pdfplumber
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
CONFIG = json.loads((ROOT / "settings.json").read_text(encoding="utf-8"))
CLIENTS = ROOT / CONFIG.get("clients_folder", ".")
REGISTRY = ROOT / CONFIG["registry"]

COMPANY = "Название компании"
NUMBER = "№Заявки"
DATE = "Дата заявки"
KIND = "Название договора"
CONTRACT = "№Договора"
SIGNER = "Исполняющий обязанности"
ROLE = "Должность"
SIGNER_GEN = "ФИО подписанта (родительный падеж)"
ROLE_GEN = "Должность (родительный падеж)"
CITY = "Город выдачи"
RETURN_COUNTRY = "Страна сдачи"
SIZE = "Тип контейнера"
COUNT = "Количество"
STATUS = "Статус заявки"

REQUEST = re.compile(
    r"заявк[аиу]\s*№?\s*([\w/-]+)\s+от\s+(\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2}))(?!\d)", re.I
)
BOX = re.compile(r"[A-Z]{4}\d{7}")
TYPE = re.compile(r"(?<!\d)(20|40)\s*['’′-]?\s*(DC|HC)\b", re.I)


def clean(value):
    return " ".join(str(value or "").split())


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    value = clean(value)
    try:
        return date.fromisoformat(value)
    except ValueError:
        return datetime.strptime(value, "%d.%m.%Y").date()


def ru_date(value):
    return as_date(value).strftime("%d.%m.%Y")


def request_id(value):
    match = REQUEST.search(clean(value))
    if not match:
        raise ValueError("Не найдены номер и дата заявки")
    day = match[2]
    if len(day.rsplit(".", 1)[1]) == 2:
        prefix, year = day.rsplit(".", 1)
        day = f"{prefix}.20{year}"
    return match[1].casefold(), as_date(day)


def key(row):
    return clean(row[COMPANY]).casefold(), clean(row[NUMBER]).casefold(), as_date(row[DATE])


def safe_component(value):
    return re.sub(r'[<>:"/\\|?*]', "_", str(value)).rstrip(". ")


def company_slug(value):
    value = re.sub(r"^(?:ООО|АО|ПАО|ИП)\s+", "", clean(value), flags=re.I)
    return "".join(ch for ch in value.casefold().replace("ё", "е") if ch.isalnum())


def client_folder(company):
    aliases = {**CONFIG["clients"], **CONFIG.get("client_folders", {})}
    alias = next((name for full, name in aliases.items() if full.casefold() == company.casefold()),
                 company)
    for candidate in (alias, company):
        path = CLIENTS / safe_component(candidate)
        if path.is_dir():
            return path
    matches = [path for path in CLIENTS.glob("*")
               if path.is_dir() and company_slug(path.name) == company_slug(company)]
    return matches[0] if len(matches) == 1 else CLIENTS / safe_component(alias)


def request_folders(folder):
    for path in sorted(folder.iterdir()):
        if path.is_dir() and REQUEST.match(path.name):
            try:
                identifier = request_id(path.name)
            except ValueError:
                continue
            yield identifier, path


def act_folder(row):
    folder = client_folder(row[COMPANY])
    identifier = key(row)[1:]
    matches = [path for current, path in request_folders(folder) if current == identifier] if folder.is_dir() else []
    if len(matches) > 1:
        raise ValueError(f"Несколько папок для заявки № {row[NUMBER]} в {folder.name}")
    if matches:
        return matches[0]
    return folder / f"Заявка {safe_component(row[NUMBER])} от {ru_date(row[DATE])}"


def request_pdfs(folder):
    grouped = {}
    for identifier, path in request_folders(folder):
        grouped.setdefault(identifier, []).append(path)
    for identifier, paths in grouped.items():
        if len(paths) > 1:
            yield identifier, paths[0], "Несколько папок с одним номером и датой"
            continue
        path = paths[0]
        pdfs = [file for file in sorted(path.iterdir()) if file.is_file()
                and file.suffix.casefold() == ".pdf"
                and re.match(r"заявка\b", file.stem.replace("_", " "), re.I)]
        if len(pdfs) != 1:
            yield identifier, path, "Ожидается один PDF с именем, начинающимся на «Заявка»"
        else:
            yield identifier, pdfs[0], None


ISSUE_ACT = "Статус АПП выдачи"
RETURN_ACT = "Статус АПП сдачи"
ACT_FIELDS = {1: ISSUE_ACT, 2: RETURN_ACT}

# SHA-256 exact matches for the six documents reviewed visually.
VERIFIED_ACTS = {
    "11b14ee1838b90a8806233f49f4bdb914b2036759215c74ffb85fd23e9fb13be": (2, 2),
    "f992885d4899163718a1b8521558eb874af8d60dee59f43c0f5231765c8cad6b": (1, 1),
    "2546dd44005b692c79b937592c2d5c8316808ff77c4338e80e2b82f6a7ab6037": (1, 1),
    "12214cee5714a3a050174b2b98d8b82e9bbd2696281f053d8c6764c32e5abb02": (2, 2),
    "d75ca70028e05d50b4e340fda8dc50dcd93365a153ed0e978a1dae439105de34": (1, 2),
    "4ca263b32354a7664ca7f847b91ad40e689e9c4493a1721f10be6009a4192f8c": (1, 2),
}


def inspect_acts(folder):
    results = {}
    issues = []
    for path in sorted(folder.glob("*")):
        if not path.is_file() or path.suffix.casefold() != ".pdf":
            continue
        name = path.stem.casefold().replace("_", " ")
        is_act = bool(re.match(r"(?:апп|акт)\b|апп[12](?!\d)", name))
        is_scan = bool(re.fullmatch(r"doc\d+", name))
        if not (is_act or is_scan):
            continue
        try:
            with path.open("rb") as handle:
                digest = file_digest(handle, "sha256").hexdigest()
        except OSError as error:
            issues.append(f"{path.name} — {error}")
            continue
        verified = VERIFIED_ACTS.get(digest)
        if verified:
            kind, seals = verified
        else:
            number = re.search(r"(?:^апп\s*|^акт.*?№\s*)([12])(?!\d)", name)
            kind = 1 if "выдач" in name else 2 if "сдач" in name else int(number[1]) if number else None
            if kind and re.search(r"\bподписан(?:о)?$", name):
                seals = 2
            else:
                issues.append(f"{path.name} — проверьте печати" +
                              ("" if kind else " и вид акта"))
                continue
        results[kind] = max(results.get(kind, 0), seals)
    for kind in (1, 2):
        if (folder / f"АПП{kind}_отправлен.txt").is_file():
            results.setdefault(kind, 1)
    return results, issues


def status(row):
    issue = row.get(ISSUE_ACT)
    returned = row.get(RETURN_ACT)
    if returned == "Подписан":
        return "Заявка закрыта" if issue == "Подписан" else "АПП2 подписан"
    if returned == "Ожидается подписание":
        return "АПП2 ожидает подписания"
    if issue == "Подписан":
        return "АПП1 подписан"
    if issue == "Ожидается подписание":
        return "АПП1 ожидает подписания"
    try:
        folder = act_folder(row)
    except ValueError:
        return "Новая заявка"
    for name, label in (
        ("АПП2_подписан", "Заявка закрыта"),
        ("АПП2_отправлен.txt", "АПП2 ожидает подписания"),
        ("АПП1_подписан", "АПП1 подписан"),
        ("АПП1_отправлен.txt", "АПП1 ожидает подписания"),
    ):
        if name.endswith(".txt"):
            if (folder / name).is_file():
                return label
        elif any((folder / f"{name}.{ext}").is_file() for ext in ("pdf", "docx")):
            return label
    return "Новая заявка"


def read_pdf(path):
    with pdfplumber.open(path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


class Registry:
    def __init__(self):
        self.book = load_workbook(REGISTRY)
        self.sheet = self.book.worksheets[0]
        self.columns = {cell.value: cell.column for cell in self.sheet[1] if cell.value}
        required = {COMPANY, NUMBER, DATE, KIND, CONTRACT, ROLE, SIGNER}
        missing = sorted(required - self.columns.keys())
        if missing:
            raise ValueError("Нет столбцов Excel: " + ", ".join(missing))
        for field in (SIZE, COUNT, CITY, RETURN_COUNTRY, STATUS, ISSUE_ACT, RETURN_ACT, SIGNER_GEN, ROLE_GEN):
            if field not in self.columns:
                column = self.sheet.max_column + 1
                self.sheet.cell(1, column, field)
                self.columns[field] = column
        self.rows = {}
        for index in range(2, self.sheet.max_row + 1):
            row = {field: self.sheet.cell(index, col).value for field, col in self.columns.items()}
            if not any(row[field] for field in (COMPANY, NUMBER, DATE)):
                continue
            identifier = key(row)
            if identifier in self.rows:
                raise ValueError(f"Дублирующий ключ в Excel, строка {index}")
            self.rows[identifier] = index, row

    def write(self, index, row):
        for field, value in row.items():
            self.sheet.cell(index, self.columns[field]).value = value

    def save(self):
        with NamedTemporaryFile(dir=REGISTRY.parent, suffix=".xlsx", delete=False) as handle:
            temporary = Path(handle.name)
        try:
            self.book.save(temporary)
            temporary.replace(REGISTRY)
        finally:
            temporary.unlink(missing_ok=True)

    def update(self, *, report=print):
        added = duplicates = 0
        issues = []
        if not CLIENTS.is_dir():
            raise ValueError(f"Папка CABO не найдена: {CLIENTS}")
        aliases = {**CONFIG["clients"], **CONFIG.get("client_folders", {})}
        names = {name.casefold(): company for company, name in aliases.items()}
        known_companies = set(aliases) | {row[COMPANY] for _, row in self.rows.values()}
        known_by_slug = {}
        for company in known_companies:
            known_by_slug.setdefault(company_slug(company), set()).add(company)
        for folder in sorted(CLIENTS.iterdir()):
            if not folder.is_dir():
                continue
            for (number, day), path, problem in request_pdfs(folder):
                try:
                    if problem:
                        raise ValueError(problem)
                    company = names.get(folder.name.casefold())
                    if not company:
                        matches = known_by_slug.get(company_slug(folder.name), set())
                        if len(matches) > 1:
                            raise ValueError("Укажите соответствие папки и клиента в client_folders")
                        company = next(iter(matches), folder.name)
                    stem = path.stem.replace("_", " ")
                    if REQUEST.match(stem) and request_id(stem) != (number, day):
                        raise ValueError("Номер или дата PDF отличаются от названия папки")
                    previous = self.rows.get(key({COMPANY: company, NUMBER: number, DATE: day}))
                    if previous:
                        duplicates += 1
                        continue
                    source = read_pdf(path)
                    if REQUEST.search(clean(source)) and request_id(source) != (number, day):
                        raise ValueError("Номер или дата в PDF отличаются от имени файла")
                    sizes = {length + kind.upper() for length, kind in TYPE.findall(source + " " + path.stem)}
                    quantities = set(re.findall(
                        r"(?:количество(?: контейнеров)?|кол-во)\s*[:–-]?\s*(\d+)|"
                        r"(\d+)\s*[xх×]\s*(?:20|40)", source + " " + path.stem, re.I
                    ))
                    counts = {int(left or right) for left, right in quantities}
                    row = {COMPANY: company, NUMBER: number, DATE: ru_date(day)}
                    if len(sizes) == 1:
                        row[SIZE] = sizes.pop()
                    if len(counts) == 1 and min(counts) > 0:
                        row[COUNT] = counts.pop()
                    if SIZE not in row or COUNT not in row:
                        issues.append(f"{path.relative_to(CLIENTS)} — добавлена без однозначного типа или количества")
                    contract = re.search(r"договор[ау]?\s*№\s*([\w/-]+)", source, re.I)
                    if contract:
                        row[CONTRACT] = contract[1]
                        for _, known in self.rows.values():
                            if (known[COMPANY] == company
                                    and clean(known[CONTRACT]).casefold() == contract[1].casefold()):
                                for field in (KIND, SIGNER, ROLE, SIGNER_GEN, ROLE_GEN, "Дата договора", "ИНН компании"):
                                    if field in self.columns and known.get(field):
                                        row[field] = known[field]
                                break
                    identifier = key(row)
                    index = self.sheet.max_row + 1
                    self.write(index, row)
                    self.rows[identifier] = index, row
                    added += 1
                except Exception as error:
                    issues.append(f"{path.relative_to(CLIENTS)} — {error}")
        for index, row in self.rows.values():
            try:
                updates = signer_fields(row, self.rows.values())
                self.write(index, updates)
                row.update(updates)
            except ValueError as error:
                issues.append(f"{row[COMPANY]} / № {row[NUMBER]} — {error}")
            try:
                folder = act_folder(row)
                acts, act_issues = inspect_acts(folder)
                issues.extend(f"{folder.relative_to(CLIENTS)} / {problem}" for problem in act_issues)
                updates = {ACT_FIELDS[kind]: "Подписан" if seals == 2 else "Ожидается подписание"
                           for kind, seals in acts.items()
                           if seals == 2 or row.get(ACT_FIELDS[kind]) != "Подписан"}
                self.write(index, updates)
                row.update(updates)
            except (OSError, ValueError) as error:
                issues.append(f"{row[COMPANY]} / № {row[NUMBER]} — {error}")
            self.write(index, {STATUS: status(row)})
        self.save()
        report(f"Добавлено {added} · Уже в реестре {duplicates} · На проверку {len(issues)}")
        for issue in issues:
            report(issue)


def positional_records(rows):
    positions = CONFIG["workbook_columns"]
    records = []
    for values in rows:
        if len(values) < positions["request"] or not values[positions["request"] - 1]:
            continue
        record = {}
        for field, position in positions.items():
            value = values[position - 1] if len(values) >= position else None
            if field in {"issue_date", "return_date"} and isinstance(value, (int, float)):
                value = from_excel(value)
            record[field] = value
        records.append(record)
    return records


def sheet_rows(row):
    folder = client_folder(row[COMPANY])
    export = folder / "контейнеры.csv"
    google_url = CONFIG.get("google_sheet_url", "").strip()
    if google_url:
        match = re.fullmatch(r"https://docs\.google\.com/spreadsheets/d/([\w-]+)(?:[/?#].*)?", google_url)
        if not match:
            raise ValueError("Для google_sheet_url нужна ссылка docs.google.com/spreadsheets/d/…")
        tab = CONFIG["clients"].get(row[COMPANY])
        if not tab:
            raise ValueError("Для клиента не настроено имя вкладки в clients")
        import gspread

        credentials = ROOT / CONFIG["credentials"]
        if not credentials.is_file():
            raise ValueError("Не найден файл доступа Google Sheets, указанный в credentials")
        client = gspread.service_account(
            filename=str(credentials),
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
        )
        sheet = client.open_by_key(match[1]).worksheet(tab)
        return positional_records(sheet.get_all_values(value_render_option="UNFORMATTED_VALUE")[1:])
    if export.is_file():
        with export.open(encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            return list(csv.DictReader(handle, dialect=csv.Sniffer().sniff(sample, delimiters=",;")))
    workbook_path = CONFIG.get("container_workbook")
    if workbook_path:
        workbook = load_workbook(ROOT / workbook_path, read_only=True, data_only=True)
        try:
            tab = CONFIG["clients"].get(row[COMPANY])
            if tab not in workbook.sheetnames:
                raise ValueError(f"В книге контейнеров нет вкладки {tab or row[COMPANY]}")
            return positional_records(workbook[tab].iter_rows(
                min_row=2, max_col=max(CONFIG["workbook_columns"].values()), values_only=True
            ))
        finally:
            workbook.close()
    if not CONFIG["sheet_id"]:
        raise ValueError("Укажите sheet_id или положите контейнеры.csv в папку клиента")
    import gspread

    tab = CONFIG["clients"].get(row[COMPANY])
    if not tab:
        raise ValueError("Для клиента не настроено название вкладки")
    client = gspread.service_account(
        filename=str(ROOT / CONFIG["credentials"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return client.open_by_key(CONFIG["sheet_id"]).worksheet(tab).get_all_records()


def containers(row, act):
    use_workbook = CONFIG.get("container_workbook") and not (
        client_folder(row[COMPANY]) / "контейнеры.csv"
    ).is_file()
    columns = (
        {field: field for field in CONFIG["workbook_columns"]}
        if use_workbook or CONFIG.get("google_sheet_url", "").strip() else CONFIG["columns"]
    )
    fields = ["request", "quantity", "container", "issue_date"]
    if act == 2:
        fields += ["return_date", "return_city"]
    if any(not columns[field] for field in fields):
        raise ValueError("Укажите точные заголовки колонок в settings.json")
    records = sheet_rows(row)
    if not records:
        raise ValueError("Таблица контейнеров пуста")
    if any(columns[field] not in records[0] for field in fields):
        raise ValueError("Заголовки таблицы контейнеров отличаются от settings.json")
    expected = key(row)[1:]
    matched = []
    counts = set()
    seen = set()
    sizes = set()
    for record in records:
        try:
            current = request_id(record[columns["request"]])
        except ValueError:
            continue
        if current != expected:
            continue
        counts.add(int(clean(record[columns["quantity"]])))
        number = re.sub(r"\s+", "", clean(record[columns["container"]])).upper()
        if not number:
            continue
        if not BOX.fullmatch(number) or number in seen:
            raise ValueError(f"Некорректный или повторный номер контейнера: {number}")
        seen.add(number)
        type_column = columns.get("type")
        if type_column and clean(record.get(type_column)):
            sizes.add(clean(record[type_column]).upper())
        if not clean(record[columns["issue_date"]]):
            raise ValueError(f"{number}: нет даты выдачи")
        issued = as_date(record[columns["issue_date"]])
        if act == 1:
            matched.append((number, issued, clean(record.get(columns.get("issue_city"))) or clean(row[CITY])))
        elif clean(record[columns["return_date"]]):
            returned = as_date(record[columns["return_date"]])
            city = clean(record[columns["return_city"]])
            if returned < issued or not city:
                raise ValueError(f"{number}: проверьте дату и город сдачи")
            matched.append((number, returned, city))
    expected_count = int(row[COUNT]) if clean(row[COUNT]) else next(iter(counts)) if len(counts) == 1 else 0
    if expected_count < 1 or counts != {expected_count}:
        raise ValueError("Проверьте количество контейнеров по заявке")
    if len(matched) != expected_count:
        raise ValueError(f"Данные для АПП{act} заполнены для {len(matched)} из {expected_count} контейнеров")
    size = clean(row[SIZE]).upper() or (next(iter(sizes)) if len(sizes) == 1 else "")
    if not size or sizes - {size}:
        raise ValueError("Проверьте тип контейнера в реестре или таблице")
    return matched, size


def contract_file(row):
    number = clean(row[CONTRACT])
    if not number:
        raise ValueError("Укажите номер договора в строке заявки в Excel")
    folder = client_folder(row[COMPANY])
    pattern = re.compile(r"(?<!\w)" + re.escape(safe_component(number)) + r"(?!\w)", re.I)
    found = []
    for path in folder.glob("*"):
        if path.suffix.lower() not in (".pdf", ".docx") or "договор" not in path.stem.casefold():
            continue
        if pattern.search(path.stem):
            found.append(path)
    if len(found) != 1:
        raise ValueError(f"Ожидается один файл «Договор {safe_component(number)}» в папке {folder.name}")
    return found[0]


def replace_text(paragraph, original, replacement):
    for match in reversed(list(re.finditer(re.escape(original), paragraph.text))):
        start, end = match.span()
        offset = 0
        for run in paragraph.runs:
            length = len(run.text)
            if offset < end and offset + length > start:
                left = max(0, start - offset)
                right = min(length, end - offset)
                run.text = run.text[:left] + (replacement if offset <= start else "") + run.text[right:]
                run.font.highlight_color = None
            offset += length


@lru_cache(maxsize=1)
def morphology():
    import pymorphy3
    return pymorphy3.MorphAnalyzer()


def genitive(value, *, person=False):
    def convert(match):
        word = match[0]
        variants = morphology().parse(word)
        if person:
            variants = [parse for parse in variants
                        if any(tag in parse.tag for tag in ("Name", "Surn", "Patr"))]
            if not variants:
                raise ValueError(f"Проверьте форму ФИО в реестре: {value}")
        parsed = variants[0]
        if not person and parsed.tag.POS not in {"NOUN", "ADJF", "PRTF"}:
            return word
        inflected = parsed.inflect({"gent"})
        if inflected is None:
            raise ValueError(f"Заполните родительный падеж в реестре: {value}")
        return inflected.word.capitalize() if word[0].isupper() else inflected.word
    return re.sub(r"[А-Яа-яЁё]+", convert, value)


def signer_fields(row, known_rows=()):
    updates = {}
    candidates = [known for _, known in known_rows
                  if known.get(COMPANY) == row.get(COMPANY)
                  and clean(known.get(CONTRACT)) == clean(row.get(CONTRACT))]
    for field in (SIGNER, ROLE, SIGNER_GEN, ROLE_GEN):
        if clean(row.get(field)):
            continue
        # Genitive forms belong to the specific person and role.
        values = {clean(known.get(field)) for known in candidates
                  if clean(known.get(field)) and (
                      field not in (SIGNER_GEN, ROLE_GEN)
                      or (clean(known.get(SIGNER)) == clean(row.get(SIGNER) or updates.get(SIGNER))
                          and clean(known.get(ROLE)) == clean(row.get(ROLE) or updates.get(ROLE))))}
        if len(values) == 1:
            updates[field] = values.pop()
    effective = {**row, **updates}
    for field in (SIGNER, ROLE):
        if not clean(effective.get(field)):
            raise ValueError("Не заполнены ФИО или должность подписанта в реестре")
    for base, target in ((SIGNER, SIGNER_GEN), (ROLE, ROLE_GEN)):
        if not clean(effective.get(target)):
            updates[target] = genitive(clean(effective[base]), person=base == SIGNER)
    return updates


def build_act(row, act, *, signer=None, role=None, output_folder=None):
    contract_file(row)
    title = clean(row[KIND])
    if "транспортно-экспедицион" in title.casefold():
        kind = "teu"
    elif "предоставлен" in title.casefold():
        kind = "supply"
    else:
        raise ValueError("Укажите вид договора в Excel")
    name = clean(row[SIGNER]).split()
    if len(name) < 2 or not clean(row[ROLE]):
        raise ValueError("Заполните ФИО и должность подписанта в Excel")
    initials = name[0] + " " + " ".join(part[0] + "." for part in name[1:])
    signer = signer or clean(row.get(SIGNER_GEN))
    role = role or clean(row.get(ROLE_GEN))
    if not signer or not role:
        raise ValueError("Обновите реестр для заполнения реквизитов подписанта")
    boxes, size = containers(row, act)
    last_date = max(day for _, day, _ in boxes)
    cities = {city for _, day, city in boxes if day == last_date}
    if len(cities) != 1 or not next(iter(cities)):
        raise ValueError("Укажите один город для даты последней операции")
    doc = Document(ROOT / "templates" / f"{act}-{kind}.docx")
    values = {
        "№Заявки": clean(row[NUMBER]), "дата заявки": ru_date(row[DATE]),
        "Название компании": row[COMPANY], "Должность1": role,
        "Исполняющий обязанности": signer, "наименование договора": title,
        "№Договора": clean(row[CONTRACT]), "Должность2": clean(row[ROLE]),
        "Инициалы": initials,
    }
    for paragraph in doc.paragraphs:
        if paragraph.text.startswith("г.") and "ктк по заявке" in paragraph.text.lower():
            replace_text(paragraph, paragraph.text.strip(),
                         f"г. {next(iter(cities))}\t{ru_date(last_date)}")
        for original, replacement in values.items():
            replace_text(paragraph, original, replacement)
    for table in doc.tables:
        for table_row in table.rows:
            for cell in table_row.cells:
                for paragraph in cell.paragraphs:
                    for original, replacement in values.items():
                        replace_text(paragraph, original, replacement)
    for table in doc.tables:
        if "Префикс, номер" not in [cell.text for cell in table.rows[1].cells]:
            continue
        prototype = deepcopy(table.rows[2]._tr)
        while len(table.rows) > 2:
            table._tbl.remove(table.rows[-1]._tr)
        for index, (number, day, _) in enumerate(boxes, 1):
            table._tbl.append(deepcopy(prototype))
            cells = table.rows[-1].cells
            for cell, value in zip(cells, (index, size, number, ru_date(day), "-"), strict=True):
                cell.text = str(value)
        break
    else:
        raise ValueError("В шаблоне не найдена таблица контейнеров")
    folder = Path(output_folder) if output_folder else act_folder(row)
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"АПП{act}.docx"
    with target.open("xb") as handle:
        try:
            doc.save(handle)
        except Exception:
            handle.close()
            target.unlink(missing_ok=True)
            raise
    if output_folder is None:
        print(f"Сформирован {target}")
    return target


def convert_act_pdf(source, target):
    if sys.platform != "win32":
        raise ValueError("Для создания PDF нужен Windows с установленным Microsoft Word")
    script = """
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($env:NECO_DOCX_PATH, $false, $true)
    $document.ExportAsFixedFormat($env:NECO_PDF_PATH, 17)
} finally {
    if ($null -ne $document) {
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
"""
    env = {**os.environ, "NECO_DOCX_PATH": str(source.resolve()),
           "NECO_PDF_PATH": str(target.resolve())}
    try:
        subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
                       env=env, check=True, capture_output=True, timeout=120,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    except (OSError, subprocess.SubprocessError) as error:
        raise ValueError("Не удалось создать PDF. Проверьте установку Microsoft Word "
                         "и закройте открытые диалоговые окна Word") from error
    if not target.is_file() or target.read_bytes()[:5] != b"%PDF-":
        raise ValueError("Microsoft Word не создал корректный PDF")


def build_pdf_act(row, act, *, report=print):
    folder = act_folder(row)
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"АПП{act} - {'выдача' if act == 1 else 'сдача'}.pdf"
    if target.exists():
        raise ValueError(f"Акт уже существует: {target}. Переместите его перед повторным формированием")
    with TemporaryDirectory(prefix="neco-act-") as temporary:
        temporary = Path(temporary)
        source = build_act(row, act, output_folder=temporary)
        pdf = temporary / "act.pdf"
        convert_act_pdf(source, pdf)
        # Exclusive creation protects existing documents, including concurrent attempts.
        with target.open("xb") as output:
            try:
                with pdf.open("rb") as incoming:
                    shutil.copyfileobj(incoming, output)
            except Exception:
                output.close()
                target.unlink(missing_ok=True)
                raise
    try:
        os.startfile(target)
    except OSError:
        report(f"PDF сохранён: {target}. Откройте его вручную — просмотрщик не запустился")
    else:
        report(f"PDF сохранён и открыт: {target}")
    return target


STATUS_STYLES = {
    "Новая заявка": "cyan",
    "АПП1 ожидает подписания": "yellow",
    "АПП1 подписан": "green",
    "АПП2 ожидает подписания": "yellow",
    "АПП2 подписан": "green",
    "Заявка закрыта": "green",
}

MENU = {
    "1": "Загрузить заявки и обновить статусы",
    "2": "Показать реестр",
    "3": "Сформировать АПП1",
    "4": "Сформировать АПП2",
    "5": "Отметить акт отправленным",
    "0": "Выход",
}


def dashboard():
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.text import Text

    console = Console()
    try:
        registry = Registry()
        totals = Counter(status(row) for _, row in registry.rows.values())
    except Exception as error:
        summary = Text(f"Реестр недоступен. {error}", style="red")
    else:
        summary = Text(f"Всего заявок: {totals.total()}\n", style="bold")
        for label, style in STATUS_STYLES.items():
            summary.append(f"\n{totals[label]:>3}  {label}", style=style)
    console.print(Panel(summary, title="Реестр контейнерных заявок", border_style="cyan",
                        expand=False))
    menu = Table.grid(padding=(0, 2))
    menu.add_column(style="bold cyan", justify="right")
    menu.add_column()
    for number, label in MENU.items():
        menu.add_row(f"[{number}]", label)
    console.print(Panel(menu, title="Действия", border_style="blue", expand=False))


def show(registry):
    from rich.console import Console
    from rich.table import Table
    from rich.text import Text

    rows = [row for _, row in registry.rows.values()]
    table = Table(title="Реестр заявок", header_style="bold cyan", row_styles=["", "on grey11"])
    table.add_column("Строка", justify="right", style="bold cyan")
    table.add_column("Клиент", ratio=2)
    table.add_column("Заявка")
    table.add_column("Дата")
    table.add_column("Тип контейнера")
    table.add_column("Кол-во", justify="right")
    table.add_column("Город выдачи")
    table.add_column("Страна сдачи")
    table.add_column("Статус", ratio=2)
    for index, row in enumerate(rows, 1):
        current_status = status(row)
        table.add_row(
            str(index), Text(clean(row[COMPANY])), Text(clean(row[NUMBER])),
            ru_date(row[DATE]), Text(clean(row[SIZE]) or "—"),
            Text(clean(row[COUNT]) or "—"), Text(clean(row.get(CITY)) or "—"),
            Text(clean(row.get(RETURN_COUNTRY)) or "—"),
            Text(current_status, STATUS_STYLES[current_status]),
        )
    console = Console()
    if rows:
        console.print(table)
    else:
        console.print("В реестре пока нет заявок.")
    return rows


def select(registry):
    rows = show(registry)
    index = int(input("Номер строки: "))
    if not 1 <= index <= len(rows):
        raise ValueError("Такой строки нет")
    return rows[index - 1]


def mark_sent(registry, row, act):
    folder = act_folder(row)
    if act not in {"1", "2"}:
        raise ValueError("Выберите АПП1 или АПП2")
    pdf = folder / f"АПП{act} - {'выдача' if act == '1' else 'сдача'}.pdf"
    if not pdf.is_file() and not (folder / f"АПП{act}.docx").is_file():
        raise ValueError("Сначала сформируйте выбранный акт")
    (folder / f"АПП{act}_отправлен.txt").touch()
    field = ACT_FIELDS[int(act)]
    if row.get(field) != "Подписан":
        row[field] = "Ожидается подписание"
    registry.write(registry.rows[key(row)][0], {field: row[field], STATUS: status(row)})
    registry.save()


def main():
    while True:
        dashboard()
        try:
            choice = input("Выберите действие: ").strip()
        except (EOFError, KeyboardInterrupt):
            return
        if choice == "0":
            return
        if choice not in MENU:
            print("Выберите пункт меню")
            continue
        try:
            registry = Registry()
            if choice == "1":
                registry.update()
            elif choice == "2":
                show(registry)
            else:
                row = select(registry)
                if choice in {"3", "4"}:
                    build_pdf_act(row, int(choice) - 2)
                else:
                    act = input("Вид акта (1 или 2): ").strip()
                    mark_sent(registry, row, act)
                    print(status(row))
        except (EOFError, KeyboardInterrupt):
            return
        except Exception as error:
            print(f"Требуется проверка. {error}")


if __name__ == "__main__":
    main()

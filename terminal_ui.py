"""Light terminal interface for the CABO registry."""
import io
from collections import Counter
from contextlib import redirect_stdout

from rich.text import Text
from textual import on, work
from textual.app import App, ComposeResult
from textual.containers import Horizontal, Vertical
from textual.widgets import (
    Button, DataTable, Footer, Input, Label, RichLog, Select, Static, TabbedContent, TabPane,
)

import registry as core


class RegistryApp(App):
    TITLE = "NECO LINE / Реестр заявок"
    CSS = """
    Screen { background: #ffffff; color: #203a5b; }
    #brand { height: 3; padding: 1 2 0 2; background: white; color: #244d7f; text-style: bold; }
    #toolbar { height: 5; padding: 0 1; background: white; }
    #search { width: 1fr; }
    #filter { width: 32; }
    #refresh { width: 23; }
    TabbedContent { height: 1fr; }
    TabPane { padding: 1; }
    DataTable { height: 1fr; background: white; color: #203a5b; }
    DataTable > .datatable--header { background: #edf2f8; color: #142f55; text-style: bold; }
    DataTable > .datatable--cursor { background: #62e6e7; color: #08223d; }
    #master { width: 3fr; }
    #card { width: 2fr; background: white; border: solid #183458; padding: 1 2; overflow-y: auto; }
    #card-content { height: auto; margin-bottom: 1; }
    #card Button { width: 100%; margin-bottom: 1; }
    #counts { width: 25; height: 1fr; padding: 1; background: white; border: solid #183458; }
    #work-table { width: 3fr; }
    #journal { width: 2fr; border: solid #183458; background: white; }
    #feedback { height: 2; background: white; color: #526479; padding: 0 2; }
    Button { background: #e4edf7; color: #203a5b; border: none; height: 3; content-align: center middle; }
    Button.-primary { background: #315f94; color: white; }
    Footer { background: #10294b; color: white; }
    Footer > .footer--key { background: #10294b; color: #62e6e7; }
    Footer > .footer--description { color: white; }
    DataTable { border: solid #183458; }
    Tabs { background: #edf2f8; }
    Tab.-active { color: #10294b; background: #62e6e7; }

    Input, Select { background: white; color: #203a5b; }
    Input { border: solid #183458; }
    Input:focus { border: solid #315f94; }
    Select > SelectCurrent { border: solid #183458; background: white; }
    Select:focus > SelectCurrent { border: solid #315f94; }

    """
    BINDINGS = [
        ("f2", "refresh", "Обновить реестр"),
        ("f3", "search", "Поиск"),
        ("f5", "issue", "АПП1"),
        ("f6", "return_act", "АПП2"),
        ("escape", "quit", "Выход"),
    ]

    def __init__(self):
        super().__init__()
        self.registry = None
        self.records = {}
        self.selected_key = None
        self.busy = False

    def compose(self) -> ComposeResult:
        yield Static("NECO LINE  /  Реестр заявок", id="brand")
        with Horizontal(id="toolbar"):
            yield Input(placeholder="Поиск по клиенту или номеру заявки", id="search")
            yield Select([("Все статусы", "*")] + [(s, s) for s in core.STATUS_STYLES],
                         value="*", allow_blank=False, id="filter")
            yield Button("Обновить реестр", id="refresh", variant="primary")
        with TabbedContent():
            with TabPane("Реестр", id="registry-tab"):
                yield DataTable(id="registry-table", cursor_type="row")
            with TabPane("Список и карточка", id="card-tab"):
                with Horizontal():
                    yield DataTable(id="master", cursor_type="row")
                    with Vertical(id="card"):
                        yield Static("Выбери заявку слева", id="card-content")
                        yield Button("Сформировать АПП1", id="issue")
                        yield Button("Сформировать АПП2", id="return")
                        yield Button("АПП1 отправлен", id="sent1")
                        yield Button("АПП2 отправлен", id="sent2")
            with TabPane("Статусы", id="status-tab"):
                yield DataTable(id="status-table", cursor_type="row")
            with TabPane("Рабочая панель", id="work-tab"):
                with Horizontal():
                    yield Static(id="counts")
                    yield DataTable(id="work-table", cursor_type="row")
                    yield RichLog(id="journal", wrap=True, markup=False)
        yield Static("Готово", id="feedback")
        yield Footer()

    def on_mount(self):
        self.theme = "textual-light"
        for name in ("registry-table", "work-table"):
            self.query_one("#" + name, DataTable).add_columns("Клиент", "Заявка", "Дата", "Тип контейнера", "Количество", "Город выдачи", "Страна сдачи", "Статус")
        self.query_one("#master", DataTable).add_columns("Клиент / заявка", "Дата", "Статус")
        self.query_one("#status-table", DataTable).add_columns("Статус", "Количество")
        self.reload()
        self.call_after_refresh(self.action_refresh)

    def feedback(self, message):
        self.query_one("#feedback", Static).update(Text(message))
        self.query_one("#journal", RichLog).write(message)

    def reload(self):
        try:
            self.registry = core.Registry()
        except Exception as error:
            self.registry = None
            self.feedback(f"Не удалось открыть реестр. {error}")
        self.render_rows()

    @on(Input.Changed, "#search")
    @on(Select.Changed, "#filter")
    def filters_changed(self):
        if self.is_mounted:
            self.render_rows()

    def render_rows(self):
        query = self.query_one("#search", Input).value.casefold().strip()
        selected_status = self.query_one("#filter", Select).value
        self.records = {}
        totals = Counter()
        for table_id in ("registry-table", "master", "work-table"):
            self.query_one("#" + table_id, DataTable).clear()
        for index, row in sorted(self.registry.rows.values(), key=lambda pair: (core.clean(pair[1][core.COMPANY]).casefold(), core.as_date(pair[1][core.DATE]))) if self.registry else []:
            state = core.status(row)
            totals[state] += 1
            if query not in f"{row[core.COMPANY]} {row[core.NUMBER]}".casefold():
                continue
            if selected_status != "*" and selected_status != state:
                continue
            key = str(index)
            self.records[key] = row
            values = (
                core.clean(row[core.COMPANY]), core.clean(row[core.NUMBER]),
                core.ru_date(row[core.DATE]), core.clean(row.get(core.SIZE)) or "—",
                core.clean(row.get(core.COUNT)) or "—",
                core.clean(row.get(core.CITY)) or "—",
                core.clean(row.get(core.RETURN_COUNTRY)) or "—", state,
            )
            for table_id in ("registry-table", "work-table"):
                self.query_one("#" + table_id, DataTable).add_row(
                    *(Text(value) for value in values), key=key)
        company_totals = Counter(row[core.COMPANY] for row in self.records.values())
        master = self.query_one("#master", DataTable)
        previous_company = None
        for key, row in self.records.items():
            company = row[core.COMPANY]
            if company != previous_company:
                master.add_row(Text(f"{company} ({company_totals[company]})", style="bold"),
                               "", "", key="company-" + key)
                previous_company = company
            master.add_row(Text(f"  № {row[core.NUMBER]}"), core.ru_date(row[core.DATE]),
                           Text(core.status(row)), key=key)
        counts = self.query_one("#status-table", DataTable)
        counts.clear()
        for state in core.STATUS_STYLES:
            counts.add_row(Text(state), str(totals[state]), key=state)
        self.query_one("#counts", Static).update(Text(
            f"ОБЗОР\n\nВсего заявок: {totals.total()}\nВидно: {len(self.records)}\n\n" +
            "\n\n".join(f"{name}: {count}" for name, count in totals.items())
        ))
        if self.selected_key not in self.records:
            self.selected_key = next(iter(self.records), None)
        self.render_card()

    @on(DataTable.RowHighlighted)
    def highlight(self, event):
        if event.data_table.id == "status-table":
            return
        key = event.row_key.value
        if key in self.records:
            self.selected_key = key
            self.render_card()

    @on(DataTable.RowSelected, "#status-table")
    def choose_status(self, event):
        self.query_one("#filter", Select).value = event.row_key.value

    def render_card(self):
        row = self.records.get(self.selected_key)
        if row:
            text = f"{row[core.COMPANY]}\nЗаявка № {row[core.NUMBER]} от {core.ru_date(row[core.DATE])}\n\n"
            for field, label in (
                (core.CONTRACT, "Договор"), (core.CITY, "Город выдачи"),
                (core.SIZE, "Тип контейнера"), (core.COUNT, "Количество"),
                (core.RETURN_COUNTRY, "Страна сдачи"),
                (core.ISSUE_ACT, "АПП выдачи"), (core.RETURN_ACT, "АПП сдачи"),
            ):
                text += f"{label}: {core.clean(row.get(field)) or '—'}\n"
            text += f"\nСтатус: {core.status(row)}"
        else:
            text = "Заявки не найдены"
        self.query_one("#card-content", Static).update(Text(text))
        for button in self.query("#card Button"):
            button.disabled = self.busy or not row
        self.query_one("#refresh", Button).disabled = self.busy or self.registry is None

    def action_search(self):
        self.query_one("#search", Input).focus()

    @on(Button.Pressed)
    def button_pressed(self, event):
        actions = {"refresh": self.action_refresh, "issue": self.action_issue,
                   "return": self.action_return_act,
                   "sent1": lambda: self.send(1), "sent2": lambda: self.send(2)}
        callback = actions.get(event.button.id)
        if callback:
            callback()

    def action_refresh(self):
        if not self.busy and self.registry:
            self.operation(lambda: core.Registry().update(), "Реестр обновлён")

    def action_issue(self):
        self.create_act(1)

    def action_return_act(self):
        self.create_act(2)

    def create_act(self, act):
        row = self.records.get(self.selected_key)
        if row and not self.busy:
            self.operation(lambda: core.build_pdf_act(row, act), f"АПП{act} сформирован")

    def send(self, act):
        row = self.records.get(self.selected_key)
        if row and not self.busy:
            self.operation(lambda: core.mark_sent(core.Registry(), row, str(act)),
                           f"АПП{act} отмечен отправленным")

    @work(exclusive=True)
    async def operation(self, callback, message):
        # Run in the app loop so stdout redirection cannot intercept terminal rendering.
        self.busy = True
        self.render_card()
        try:
            with redirect_stdout(io.StringIO()) as capture:
                callback()
            self.reload()
            self.feedback(capture.getvalue().strip() or message)
        except Exception as error:
            self.feedback(f"Требуется проверка. {error}")
        finally:
            self.busy = False
            self.render_card()


def main():
    RegistryApp().run()


if __name__ == "__main__":
    main()

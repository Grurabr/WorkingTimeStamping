import math
import sys
import csv
import traceback
from collections import defaultdict
from datetime import datetime, timedelta

from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QAction, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QPushButton, QLabel, QDialog, QCalendarWidget, QLineEdit, QDialogButtonBox, QDateEdit, QRadioButton,
    QButtonGroup
)


class DayOffDialog(QDialog):
    def __init__(self, persons, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Lisää arkipyhäpäivä")
        self.setFixedSize(500, 600)

        layout = QVBoxLayout(self)

        self.person_menu = QComboBox()
        self.person_menu.addItem("Kaikki")
        self.person_menu.addItems(persons)
        layout.addWidget(QLabel("Työntekijä:"))
        layout.addWidget(self.person_menu)

        self.calendar = QCalendarWidget()
        layout.addWidget(QLabel("Pvm:"))
        layout.addWidget(self.calendar)

        self.dayoff_or_pekkanen = QComboBox()
        self.dayoff_or_pekkanen.addItem("Arkipyhäpäivä", 14)
        self.dayoff_or_pekkanen.addItem("½ Pekkanen", 93)
        self.dayoff_or_pekkanen.addItem("Pekkanen", 1)
        layout.addWidget(self.dayoff_or_pekkanen)


        self.comment_input = QLineEdit()
        self.comment_input.setPlaceholderText("Comment: ")
        layout.addWidget(self.comment_input)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_data(self):
        return (
            self.person_menu.currentText(),
            self.calendar.selectedDate().toPyDate(),
            self.dayoff_or_pekkanen.currentData(),
            self.comment_input.text()
        )


class NewVuoroDialog(QDialog):
    def __init__(self, persons, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Lisää työvuoro")
        self.setFixedSize(500, 600)

        layout = QVBoxLayout(self)

        self.person_menu = QComboBox()
        self.person_menu.addItems(persons)
        layout.addWidget(QLabel("Työntekijä:"))
        layout.addWidget(self.person_menu)

        self.calendar = QCalendarWidget()
        layout.addWidget(QLabel("Pvm:"))
        layout.addWidget(self.calendar)

        self.vuoro = QComboBox()
        self.vuoro.addItem("Aamuvuoro", 4)
        self.vuoro.addItem("Päivävuoro", 8)
        self.vuoro.addItem("Iltavuoro", 5)
        self.vuoro.addItem("Yövuoro", 16)
        self.vuoro.addItem("Kuukausi palkka", 10)
        layout.addWidget(self.vuoro)


        self.comment_input = QLineEdit()
        self.comment_input.setPlaceholderText("Comment: ")
        layout.addWidget(self.comment_input)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_data(self):
        return (
            self.person_menu.currentText(),
            self.calendar.selectedDate().toPyDate(),
            self.vuoro.currentData(),
            self.comment_input.text()
        )

class YhteinenDialog(QDialog):
    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Korjaus")
        self.setFixedSize(400, 100)

        layout = QVBoxLayout(self)

        self.person_label = QLabel(f"{message}")
        self.person_layout = QHBoxLayout()
        self.person_layout.addWidget(self.person_label)
        layout.addLayout(self.person_layout)


        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

class AnalysointiDialog(QDialog):
    def __init__(self, persons, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Analysointi")
        self.setFixedSize(600, 400)

        layout = QVBoxLayout(self)

        self.person_men = QComboBox()
        self.person_men.addItem("Kaikki")
        self.person_men.addItems(persons)
        self.person_label = QLabel("Työntekijä:")
        self.person_layout = QHBoxLayout()
        self.person_layout.addWidget(self.person_label)
        self.person_layout.addWidget(self.person_men)
        layout.addLayout(self.person_layout)

        # Создаем два поля для даты
        self.start_dat_input = QDateEdit()
        self.start_dat_input.setCalendarPopup(True)  # Позволяет выбрать дату через календарь
        self.start_dat_input.setDate(QDate.currentDate())  # Устанавливаем текущую дату по умолчанию
        self.start_dat_label = QLabel("Alkupäivämäärä:")
        self.start_dat_layout = QHBoxLayout()
        self.start_dat_layout.addWidget(self.start_dat_label)
        self.start_dat_layout.addWidget(self.start_dat_input)  # Исправили, добавили виджет напрямую
        layout.addLayout(self.start_dat_layout)

        self.end_dat_input = QDateEdit()
        self.end_dat_input.setCalendarPopup(True)
        self.end_dat_input.setDate(QDate.currentDate())  # Устанавливаем текущую дату по умолчанию
        self.end_dat_label = QLabel("Loppupäivämäärä:")  # Исправляем текст на "Loppupäivämäärä"
        self.end_dat_layout = QHBoxLayout()
        self.end_dat_layout.addWidget(self.end_dat_label)
        self.end_dat_layout.addWidget(self.end_dat_input)  # Исправили, добавили виджет напрямую
        layout.addLayout(self.end_dat_layout)


        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_data(self):
        return (
            self.person_men.currentText(),
            self.start_dat_input.date().toPyDate(),
            self.end_dat_input.date().toPyDate(),
        )

class LomaDialog(QDialog):
    def __init__(self, persons, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Lomat")
        self.setFixedSize(600, 400)

        layout = QVBoxLayout(self)

        self.person_menu = QComboBox()
        self.person_menu.addItems(persons)
        self.person_label = QLabel("Työntekijä:")
        self.person_layout = QHBoxLayout()
        self.person_layout.addWidget(self.person_label)
        self.person_layout.addWidget(self.person_menu)
        layout.addLayout(self.person_layout)

        # Создаем два поля для даты
        self.start_date_input = QDateEdit()
        self.start_date_input.setCalendarPopup(True)  # Позволяет выбрать дату через календарь
        self.start_date_input.setDate(QDate.currentDate())  # Устанавливаем текущую дату по умолчанию
        self.start_date_label = QLabel("Alkupäivämäärä:")
        self.start_date_layout = QHBoxLayout()
        self.start_date_layout.addWidget(self.start_date_label)
        self.start_date_layout.addWidget(self.start_date_input)  # Исправили, добавили виджет напрямую
        layout.addLayout(self.start_date_layout)

        self.end_date_input = QDateEdit()
        self.end_date_input.setCalendarPopup(True)
        self.end_date_input.setDate(QDate.currentDate())  # Устанавливаем текущую дату по умолчанию
        self.end_date_label = QLabel("Loppupäivämäärä:")  # Исправляем текст на "Loppupäivämäärä"
        self.end_date_layout = QHBoxLayout()
        self.end_date_layout.addWidget(self.end_date_label)
        self.end_date_layout.addWidget(self.end_date_input)  # Исправили, добавили виджет напрямую
        layout.addLayout(self.end_date_layout)

        # Radio button
        self.radiobutton_layout = QVBoxLayout()
        self.radio_button_1 = QRadioButton("Maanantai - Lauantai")
        self.radio_button_2 = QRadioButton("Maanantai - Perjantai")
        self.radio_button_1.setChecked(True)
        self.radiobutton_group = QButtonGroup(self)
        self.radiobutton_group.addButton(self.radio_button_1, 1)
        self.radiobutton_group.addButton(self.radio_button_2, 2)
        self.radiobutton_layout.addWidget(self.radio_button_1)
        self.radiobutton_layout.addWidget(self.radio_button_2)
        layout.addLayout(self.radiobutton_layout)


        self.comment_input = QLineEdit()
        self.comment_input.setPlaceholderText("Kommentti: ")
        layout.addWidget(self.comment_input)

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_data(self):
        return (
            self.person_menu.currentText(),
            self.start_date_input.date().toPyDate(),
            self.end_date_input.date().toPyDate(),
            self.radiobutton_group.checkedId(),
            self.comment_input.text()
        )

class AnalyysiTulosDialog(QDialog):
    def __init__(self, results, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Analyysitulokset")
        self.setMinimumSize(650, 800)

        self.results = results

        layout = QVBoxLayout(self)

        # Фильтр по работнику
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Työntekijä:")
        self.person_filter = QComboBox()
        self.person_filter.addItem("Kaikki")
        self.person_filter.addItems(sorted(set(r["person"] for r in results)))
        self.person_filter.currentIndexChanged.connect(self.update_table)

        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.person_filter)
        layout.addLayout(filter_layout)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Työntekijä", "Viikko", "Ajanjakso", "Viikon tila", "Tunnit"])
        layout.addWidget(self.table)

        self.update_table()

    def update_table(self):
        selected = self.person_filter.currentText()
        filtered = self.results if selected == "Kaikki" else [r for r in self.results if r["person"] == selected]

        self.table.setRowCount(len(filtered))
        for row, item in enumerate(filtered):
            self.table.setItem(row, 0, QTableWidgetItem(item["person"]))
            self.table.setItem(row, 1, QTableWidgetItem(str(item["week_num"])))
            self.table.setItem(row, 2, QTableWidgetItem(f"{item['start']} – {item['end']}"))
            self.table.setItem(row, 3, QTableWidgetItem(item["completeness"]))
            self.table.setItem(row, 4, QTableWidgetItem(item["status"]))

        self.table.resizeColumnsToContents()

class CSVExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Valitse tiedot")

        layout = QVBoxLayout()

        label = QLabel("Haluatko lisätä tiedostoon kaikki tiedot vai vain avoimesta taulukosta?")
        layout.addWidget(label)

        self.kaikki_btn = QPushButton("Kaikki")
        self.taulukosta_btn = QPushButton("Taulukosta")
        self.cancel_btn = QPushButton("Peruuta")

        self.kaikki_btn.clicked.connect(self.accept_all)
        self.taulukosta_btn.clicked.connect(self.accept_table)
        self.cancel_btn.clicked.connect(self.reject)

        layout.addWidget(self.kaikki_btn)
        layout.addWidget(self.taulukosta_btn)
        layout.addWidget(self.cancel_btn)

        self.setLayout(layout)
        self.selection = None

    def accept_all(self):
        self.selection = "all"
        self.accept()

    def accept_table(self):
        self.selection = "table"
        self.accept()

    def get_selection(self):
        return self.selection

class CSVViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.username = ""

        self.setWindowTitle("Working time stamping")
        self.setGeometry(100, 100, 1600, 700)

        # main layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)

        # left layout
        self.select_layout = QVBoxLayout()
        self.select_menu = QComboBox()
        self.select_menu.setFixedWidth(300)
        self.select_menu.addItem("Kaikki työntekijät")
        self.select_menu.currentTextChanged.connect(self.filter_by_person)
        self.select_layout.addWidget(self.select_menu, alignment=Qt.AlignTop)

        self.name = QLineEdit()
        self.name.setPlaceholderText("Käyttäjän nimi")
        self.name.setText(self.username)
        self.name.setFixedWidth(200)
        self.name.textChanged.connect(self.update_username)
        self.name_label = QLabel("Käyttäjä: ")
        self.name_label.setFixedWidth(100)
        self.name_layout = QHBoxLayout()
        self.name_layout.addWidget(self.name_label)
        self.name_layout.addWidget(self.name)
        self.select_layout.addLayout(self.name_layout)


        self.add_day_off_button = QPushButton("Lisää arkipyhäpäivä \\ Yhteinen pekkanen")
        self.add_day_off_button.clicked.connect(self.open_day_off_dialog)
        self.add_day_off_button.setEnabled(False)
        self.select_layout.addWidget(self.add_day_off_button)

        self.add_new_vuoro_button = QPushButton("Lisää työvuoro")
        self.add_new_vuoro_button.clicked.connect(self.open_new_vuoro_dialog)
        self.add_new_vuoro_button.setEnabled(False)
        self.select_layout.addWidget(self.add_new_vuoro_button)

        self.loma_button = QPushButton("Lisää lomapäivät")
        self.loma_button.clicked.connect(self.open_loma_dialog)
        self.loma_button.setEnabled(False)
        self.select_layout.addWidget(self.loma_button)

        self.korjaus_button = QPushButton("Työaikojen korjaus")
        self.korjaus_button.clicked.connect(self.open_korjaus_dialog)
        self.korjaus_button.setEnabled(False)
        self.select_layout.addWidget(self.korjaus_button)

        self.analysointi_button = QPushButton("Analysointi")
        self.analysointi_button.clicked.connect(self.analysointi_dialog)
        self.analysointi_button.setEnabled(False)
        self.select_layout.addWidget(self.analysointi_button)

        self.print_button = QPushButton("Tulosta CSV")
        self.print_button.clicked.connect(self.dialog_tulosta_asiakirja)
        self.print_button.setEnabled(False)
        self.select_layout.addWidget(self.print_button)


        self.main_layout.addLayout(self.select_layout)

        # navigointi
        self.date_nav_layout = QHBoxLayout()
        self.prev_button = QPushButton("←")
        self.next_button = QPushButton("→")
        self.prev_button.setEnabled(False)
        self.next_button.setEnabled(False)
        self.date_filter_menu = QComboBox()
        self.date_filter_menu.setFixedWidth(300)
        self.date_filter_menu.currentTextChanged.connect(self.filter_by_date)
        self.prev_button.clicked.connect(self.navigate_prev)
        self.next_button.clicked.connect(self.navigate_next)

        self.date_nav_layout.addWidget(self.prev_button)
        self.date_nav_layout.addWidget(self.date_filter_menu)
        self.date_nav_layout.addWidget(self.next_button)

        # table
        self.table = QTableWidget()
        self.table_container = QVBoxLayout()
        self.table_container.addLayout(self.date_nav_layout)
        self.table_container.addWidget(self.table)
        self.main_layout.addLayout(self.select_layout)
        self.main_layout.addLayout(self.table_container)



        # data
        self.data = []
        self.filtered_data = []
        self.init_menu()
        self.date_filter_list = ["Kaikki ajanjaksot"]
        self.date_filter_index = 0
        self.kaikki_tyontekijat = ()

        self.months = {'1': 'Tammikuu',
                       '2': 'Helmikuu',
                       '3': 'Maaliskuu',
                       '4': 'Huhtikuu',
                       '5': 'Toukokuu',
                       '6': 'Kesäkuu',
                       '7': 'Heinäkuu',
                       '8': 'Elokuu',
                       '9': 'Syyskuu',
                       '10': 'Lokakuu',
                       '11': 'Marraskuu',
                       '12': 'Joulukuu'}

    def update_username(self, text):
        self.username = text

    def dialog_tulosta_asiakirja(self):
        dialog = CSVExportDialog(self)
        if dialog.exec_():
            valinta = dialog.get_selection()
            if valinta == "all":
                self.tulosta_asiakirja()
            elif valinta == "table":
                self.tulosta_taulusta()




    def get_column_index(self, header_name):
        for i in range(self.table.columnCount()):
            if self.table.horizontalHeaderItem(i).text() == header_name:
                return i
        return -1


    def analysointi_dialog(self):
        persons = []
        for index in range(self.select_menu.count()):
            person_text = self.select_menu.itemText(index)
            persons.append(person_text)
        dialog = AnalysointiDialog(persons, self)
        if dialog.exec_():
            person, start_dat, end_dat = dialog.get_data()
            self.analysointi(person, start_dat, end_dat)


    def analysointi(self, persons, start_date, end_date):
        results = []

        def kasittely(person):
            data = [record for record in self.filtered_data if record.get("Person_number") == person]

            weekly_hours = defaultdict(float)

            for day in data:
                date_str = day['Timestamp_date']

                try:
                    dt = datetime.strptime(date_str, "%d.%m.%Y %H.%M.%S")
                except ValueError:
                    dt = datetime.strptime(date_str, "%d.%m.%Y")

                if not (start <= dt.date() <= end):
                    continue

                hours = float(day['TimeStamp_hours'].replace(',', '.'))

                year, week_num, _ = dt.isocalendar()
                key = (year, week_num)
                weekly_hours[key] += hours



            for week in week_data:
                key = (week["year"], week["week_num"])
                hours = weekly_hours.get(key, 0.0)
                status = f"✔️ OK {hours:.1f} h" if hours >= 40 else f"⚠️ Vain {hours:.1f} h"
                completeness = f"täysi viikko - {len(week['days'])} päivää" if week["is_full_week"] else f"vajaa viikko - {len(week['days'])} päivää"
                results.append({
                    "person": person,
                    "week_num": week["week_num"],
                    "start": week["start"].strftime("%d.%m.%Y"),
                    "end": week["end"].strftime("%d.%m.%Y"),
                    "completeness": completeness,
                    "status": status
                })




        start = start_date
        end = end_date

        week_data = []
        current = start

        while current <= end:
            week_start = current - timedelta(days=current.weekday())
            week_end = week_start + timedelta(days=6)

            period_start = max(start, week_start)
            period_end = min(end, week_end)

            weekdays = [d for d in (period_start + timedelta(days=i)
                                    for i in range((period_end - period_start).days + 1))
                        if d.weekday() < 5]

            is_full_week = len(weekdays) == 5
            year, week_num, _ = period_start.isocalendar()

            week_data.append({
                "year": year,
                "week_num": week_num,
                "start": period_start,
                "end": period_end,
                "days": weekdays,
                "is_full_week": is_full_week
            })

            current = week_end + timedelta(days=1)


        if persons == "Kaikki":
            added_persons = self.all_persons_in_data()

            for person in added_persons:
                kasittely(person)
        else:
            kasittely(persons)

        dialog = AnalyysiTulosDialog(results, self)
        dialog.exec_()


    def add_day_off(self, person, day_off_date, dayoff_or_pekkanen, comment=""):
        timestamp_now = datetime.now().strftime("%d.%m.%Y %H.%M.%S")
        start = day_off_date.strftime("%d.%m.%Y 6.00.00")
        end = day_off_date.strftime("%d.%m.%Y 14.00.00")
        description= {14: "Arkipyhäkorvaus", 93: "½ Pekkanen", 1: "Pekkanen"}

        timeofreturn = ""
        reason = "1"
        if description[dayoff_or_pekkanen] == "½ Pekkanen":
            reason = "6"

            start_dt = datetime.strptime(start, "%d.%m.%Y %H.%M.%S")
            return_dt = start_dt + timedelta(hours=4)
            timeofreturn = return_dt.strftime("%d.%m.%Y %H.%M.%S")

        elif description[dayoff_or_pekkanen] == "Pekkanen":
            reason = "5"
            timeofreturn = end
        elif description[dayoff_or_pekkanen] == "Arkipyhäkorvaus":
            reason = "4"

        def create_entry(name):
            return {
                "Person_number": name,
                "Timestamp_date": start,
                "Timestamp_enddate": end,
                "TimeStamp_hours": "8,0000",
                "TimeStamp_description": description[dayoff_or_pekkanen],
                "TimeStamp_comments": comment,
                "Timestamp_workshift": "4",
                "Formofsalary_number": str(dayoff_or_pekkanen),
                "Timestamp_editdate": timestamp_now,
                "TimeStamp_editor": self.username,
                "Timestamp_reason": reason,
                "Timestamp_Salary_id": "0",
                "Timestamp_state": "1",
                "Timestamp_timeofreturn": timeofreturn,
                "Timestamp_Useminlunch_bit": "False"
            }

        if person == "Kaikki":
            for name in self.kaikki_tyontekijat:
                entry = create_entry(name)
                self.filtered_data.append(entry)
        else:
            entry = create_entry(person)
            self.filtered_data.append(entry)

        # Обновляем таблицу и фильтры
        #self.update_person_list()
        self.update_date_list()
        self.filter_by_person(self.select_menu.currentText())

    def all_persons_in_data(self):
        added_persons = set()

        for person in self.filtered_data:
            person_number = person['Person_number']

            if person_number not in added_persons:
                added_persons.add(person_number)

        return sorted(added_persons, key=lambda x: int(x))


    def add_loma(self, person, loma_start_date, loma_end_date, loman_pituus, comment=""):
        timestamp_now = datetime.now().strftime("%d.%m.%Y %H.%M.%S")

        loma_paivat = []

        current_date = loma_start_date
        while current_date <= loma_end_date:
            if loman_pituus == 1:
                if current_date.weekday() != 6:
                    loma_paivat.append(current_date)
            else:
                if current_date.weekday() != 5 and current_date.weekday() != 6:
                    loma_paivat.append(current_date)

            current_date += timedelta(days=1)

        def create_entry(name, start, end):
            return {
                "Person_number": name,
                "Timestamp_date": start,
                "Timestamp_enddate": end,
                "TimeStamp_hours": "8,0000",
                "TimeStamp_description": "Lomapäivä",
                "TimeStamp_comments": comment,
                "Timestamp_workshift": "4",
                "Formofsalary_number": "90",
                "Timestamp_editdate": timestamp_now,
                "TimeStamp_editor": self.username,
                "Timestamp_reason": "2",
                "Timestamp_Salary_id": "0",
                "Timestamp_state": "1",
                "Timestamp_timeofreturn": "",
                "Timestamp_Useminlunch_bit": "False"
            }

        for day in loma_paivat:
            start = day.strftime("%d.%m.%Y 6.00.00")
            end = day.strftime("%d.%m.%Y 14.00.00")
            entry = create_entry(person, start, end)
            self.filtered_data.append(entry)


        # Обновляем таблицу и фильтры
        #self.update_person_list()
        self.update_date_list()
        self.filter_by_person(self.select_menu.currentText())

    def open_day_off_dialog(self):

        dialog = DayOffDialog(self.kaikki_tyontekijat, self)
        if dialog.exec_():
            person, day_off_date, dayoff_or_pekkanen, comment = dialog.get_data()
            self.add_day_off(person, day_off_date, dayoff_or_pekkanen, comment)

    def open_new_vuoro_dialog(self):

        dialog = NewVuoroDialog(self.kaikki_tyontekijat, self)
        if dialog.exec_():
            person, day_off_date, vuoro, comment = dialog.get_data()
            self.add_new_vuoro(person, day_off_date, vuoro, comment)

    def add_new_vuoro(self, person, day_off_date, vuoro, comment=""):
        timestamp_now = datetime.now().strftime("%d.%m.%Y %H.%M.%S")

        shifts = {
            4: ("06:00:00", "14:00:00"),  # Aamuvuoro
            8: ("07:00:00", "15:00:00"),  # Päivävuoro
            5: ("14:00:00", "22:00:00"),  # Iltavuoro
            16: ("22:00:00", "06:00:00"),  # Yövuoro
            10: ("06:00:00", "14:00:00"),  # Kuukausi palkka
        }

        shift_times = shifts.get(vuoro, ("08:00:00", "16:00:00"))  # default
        start_time_str, end_time_str = shift_times

        start = datetime.strptime(f"{day_off_date.strftime('%d.%m.%Y')} {start_time_str}", "%d.%m.%Y %H:%M:%S")

        end_date = day_off_date + timedelta(days=1) if end_time_str < start_time_str else day_off_date
        end = datetime.strptime(f"{end_date.strftime('%d.%m.%Y')} {end_time_str}", "%d.%m.%Y %H:%M:%S")

        start_str = start.strftime("%d.%m.%Y %H.%M.%S")
        end_str = end.strftime("%d.%m.%Y %H.%M.%S")

        workshift = "2" if vuoro == 10 else "1"

        def create_entry(name):
            return {
                "Person_number": name,
                "Timestamp_date": start_str,
                "Timestamp_enddate": end_str,
                "TimeStamp_hours": "8,0000",
                "TimeStamp_description": "Ulos",
                "TimeStamp_comments": comment,
                "Timestamp_workshift": str(vuoro),
                "Formofsalary_number": workshift,
                "Timestamp_editdate": timestamp_now,
                "TimeStamp_editor": self.username,
                "Timestamp_reason": "1",
                "Timestamp_Salary_id": "0",
                "Timestamp_state": "1",
                "Timestamp_timeofreturn": "",
                "Timestamp_Useminlunch_bit": "False"
            }

        if person == "Kaikki":
            for name in self.filtered_data:
                entry = create_entry(name)
                self.filtered_data.append(entry)
        else:
            entry = create_entry(person)
            self.filtered_data.append(entry)

        # Обновляем таблицу и фильтры
        #self.update_person_list()
        self.update_date_list()
        self.filter_by_person(self.select_menu.currentText())

    def open_loma_dialog(self):

        dialog = LomaDialog(self.kaikki_tyontekijat, self)
        if dialog.exec_():
            person, loma_start_date, loma_end_date, loman_pituus, comment = dialog.get_data()
            self.add_loma(person, loma_start_date, loma_end_date, loman_pituus, comment)


    def open_korjaus_dialog(self):
        persons = self.select_menu.currentText()
        message = f"Tehdäänkö työaikojen korjaus käyttäjälle {persons}"
        dialog = YhteinenDialog(message, self)
        if dialog.exec_():
            self.add_korjaus(persons)

    def add_korjaus(self, persons):
        #print(self.filtered_data)

        def round_to_half(value):
            try:
                number = float(value.replace(',', '.'))
            except ValueError:
                return value  # вернуть как есть, если не число

            # Округляем до ближайшего 0.5
            if 8.5 > number > 8.0:
                rounded = math.floor(number)
            elif 10.2 > number > 10.0:
                rounded = math.floor(number)
            else:
                rounded = number

            return f"{rounded:.4f}".replace('.', ',')

        def korjaustyot(person):
            data = self.filtered_data
            timestamp_now = datetime.now().strftime("%d.%m.%Y %H.%M.%S")
            for index, rivi in enumerate(data):
                if rivi['Person_number'] == person:
                    if rivi['Timestamp_workshift'] == '10':
                        continue
                    else:
                        new_stamp = round_to_half(rivi['TimeStamp_hours'])
                        if rivi['TimeStamp_hours'] != new_stamp:
                            rivi['TimeStamp_comments'] = f'Vanhat tiedot: {rivi['TimeStamp_hours']}'
                            rivi['Timestamp_editdate'] = timestamp_now
                            rivi['TimeStamp_editor'] = self.username
                            rivi['TimeStamp_hours'] = new_stamp

            self.filtered_data = data



        if persons == "Kaikki":
            for person in self.kaikki_tyontekijat:
                korjaustyot(person)
        else:
            korjaustyot(persons)

        self.filter_by_date("Kaikki ajanjaksot")


    def update_person_list(self):
        self.select_menu.blockSignals(True)
        self.select_menu.clear()
        self.select_menu.addItem("Kaikki")

        added_persons = self.all_persons_in_data()
        self.select_menu.addItems(added_persons)

        self.select_menu.blockSignals(False)

    def filter_by_person(self, person):
        if person == "Kaikki":
            all_records = self.filtered_data
        else:
            all_records = [record for record in self.filtered_data if record.get("Person_number") == person]

        all_records.sort(key=lambda record: int(record.get("Person_number", 0)))

        self.display_data(all_records)

    def filter_by_date(self, value):
        person = self.select_menu.currentText()

        if person == "Kaikki":
            all_records = self.filtered_data
        else:
            all_records = [record for record in self.filtered_data if record.get("Person_number") == person]

        all_records.sort(key=lambda record: int(record.get("Person_number", 0)))

        if value == "Kaikki ajanjaksot":
            self.display_data(all_records)
            return


        filtered = []
        for record in all_records:
            date_str = record.get("Timestamp_date", "")
            try:
                dt = datetime.strptime(date_str, "%d.%m.%Y %H.%M.%S")
                if value.startswith("Viikko"):
                    parts = value.split()
                    if len(parts) == 3:
                        week_val = int(parts[1])
                        year_val = int(parts[2])
                        if dt.isocalendar()[1] == week_val and dt.year == year_val:
                            filtered.append(record)
                else:
                    # value is something like 'Toukokuu 2024'
                    for month_number, name in self.months.items():
                        if value.startswith(name):
                            year_val = int(value.split()[-1])
                            if dt.month == int(month_number) and dt.year == year_val:
                                filtered.append(record)
                            break
            except:
                continue

        self.display_data(filtered)

    def update_date_list(self):
        self.date_filter_menu.blockSignals(True)
        all_records = self.filtered_data

        months = set()
        weeks = set()

        for record in all_records:
            try:
                dt = datetime.strptime(record.get("Timestamp_date", ""), "%d.%m.%Y %H.%M.%S")
                months.add((dt.year, dt.month))
                weeks.add((dt.year, dt.isocalendar()[1]))
            except Exception:
                continue

        self.date_filter_list = ["Kaikki ajanjaksot"]

        for year, m in sorted(months):
            month_name = self.months.get(str(m), str(m))
            self.date_filter_list.append(f"{month_name} {year}")

        for year, w in sorted(weeks):
            self.date_filter_list.append(f"Viikko {w} {year}")

        self.date_filter_index = 0

        self.date_filter_menu.clear()
        self.date_filter_menu.addItems(self.date_filter_list)
        self.date_filter_menu.setCurrentIndex(self.date_filter_index)

        self.date_filter_menu.blockSignals(False)

    def navigate_prev(self):
        if self.date_filter_index > 0:
            self.date_filter_index -= 1
            selected_text = self.date_filter_list[self.date_filter_index]
            self.date_filter_menu.setCurrentText(selected_text)
            self.filter_by_date(self.date_filter_list[self.date_filter_index])

    def navigate_next(self):
        if self.date_filter_index < len(self.date_filter_list) - 1:
            self.date_filter_index += 1
            selected_text = self.date_filter_list[self.date_filter_index]
            self.date_filter_menu.setCurrentText(selected_text)
            self.filter_by_date(self.date_filter_list[self.date_filter_index])

    def init_menu(self):
        menu = self.menuBar()
        file_menu = menu.addMenu("File")

        open_action = QAction("Open", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv);;All Files (*)"
        )

        if file_path:
            try:

                with open(file_path, encoding='utf-8-sig') as file:
                    reader = csv.DictReader(file, delimiter=';')
                    self.data = list(reader)

                #print(self.data)
                self.filtered_data = self.data

                def parse_datetime(date_str):
                    try:
                        return datetime.strptime(date_str, "%d.%m.%Y %H.%M.%S")
                    except ValueError:
                        return None

                self.filtered_data = sorted(self.filtered_data, key=lambda x: (int(x['Person_number']), parse_datetime(x['Timestamp_date']) or datetime.min))

                self.kaikki_tyontekijat = self.all_persons_in_data()
                self.display_data(self.filtered_data)
                self.update_person_list()
                self.update_date_list()
                self.add_day_off_button.setEnabled(True)
                self.loma_button.setEnabled(True)
                self.korjaus_button.setEnabled(True)
                self.analysointi_button.setEnabled(True)
                self.prev_button.setEnabled(True)
                self.next_button.setEnabled(True)
                self.print_button.setEnabled(True)
                self.add_new_vuoro_button.setEnabled(True)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open file:\n{e}")
                print(f"Error opening file: {e}")


    def display_data(self, records):

        if not records:
            self.table.clear()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        records = sorted(records, key=lambda x: int(x['Person_number']))

        headers = list(records[0].keys())
        #print(headers)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(records))

        for row_idx, row in enumerate(records):
            #print(row)
            for col_idx, key in enumerate(headers):
                item = QTableWidgetItem(row.get(key, ""))
                self.table.setItem(row_idx, col_idx, item)


    def tulosta_asiakirja(self):
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save CSV File", "", "CSV Files (*.csv);;All Files (*)"
            )

            if file_path:
                all_records = self.filtered_data

                if all_records:
                    headers = list(all_records[0].keys())

                    with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                        writer = csv.DictWriter(file, fieldnames=headers, delimiter=';')
                        writer.writeheader()
                        writer.writerows(all_records)

                    QMessageBox.information(self, "Success", "CSV file saved successfully.")
                else:
                    QMessageBox.warning(self, "No Data", "There is no data to export.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

    def tulosta_taulusta(self):
        try:
            path, _ = QFileDialog.getSaveFileName(self, "Tallenna CSV", "", "CSV files (*.csv)")
            if not path:
                return

            with open(path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file, delimiter=';')


                headers = []
                for column in range(self.table.columnCount()):
                    item = self.table.horizontalHeaderItem(column)
                    headers.append(item.text() if item else "")
                writer.writerow(headers)

                for row in range(self.table.rowCount()):
                    row_data = []
                    for column in range(self.table.columnCount()):
                        item = self.table.item(row, column)
                        row_data.append(item.text() if item else "")
                    writer.writerow(row_data)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CSVViewer()
    window.show()
    try:
        sys.exit(app.exec_())
    except Exception as e:
        print("Unexpected error occurred:")
        traceback.print_exc()

import sys
from datetime import datetime
from collections import defaultdict

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QGroupBox, QScrollArea,
    QMessageBox, QFileDialog, QToolTip, QDialog, QCheckBox, QDialogButtonBox,
    QSizePolicy
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# Для экспорта в Excel
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ----- КОНФИГУРАЦИЯ -----
ALL_DIRECTIONS = ["N", "S", "E", "W"]
DIRECTION_NAMES = {"N": "Север (N)", "S": "Юг (S)", "E": "Восток (E)", "W": "Запад (W)"}

VEHICLE_TYPES = {
    "car":        {"label": "Легковой",          "desc": "Легковые автомобили всех типов, включая седаны, хэтчбеки, универсалы."},
    "mini_bus":   {"label": "Микроавтобус",      "desc": "Маршрутное такси, пассажирские микроавтобусы (Газель, Ford Transit малый). В том числе скорая помощь."},
    "middle_bus": {"label": "Средний автобус",   "desc": "Средние автобусы (ПАЗ, ЛИАЗ малый, городские автобусы средней вместимости)."},
    "bus":        {"label": "Большой автобус",   "desc": "Большие автобусы (ЛиАЗ, МАЗ, Mercedes Citaro и аналоги)."},
    "mini_truck": {"label": "Малый грузовик",    "desc": "Небольшие грузовики грузоподъёмностью до 2 тонн (Газель, УАЗ с кузовом, японские/китайские аналоги)."},
    "middle_truck":{"label": "Средний грузовик", "desc": "Грузовики грузоподъёмностью 2–6 тонн, 2 оси (ГАЗ, ЗИЛ, «Газон NEXT»)."},
    "truck":      {"label": "Тяжёлый грузовик",  "desc": "Грузовики грузоподъёмностью более 6 тонн, 3 и более осей (КАМАЗ, МАЗ, Volvo, MAN)."},
    "road_train": {"label": "Автопоезд",         "desc": "Грузовик с прицепом/полуприцепом, 4 и более осей."},
    "trol":       {"label": "Троллейбус",        "desc": "Троллейбус с рогами, контактной сетью."},
    "tram":       {"label": "Трамвай",           "desc": "Трамвай – состав, двигающийся по рельсам."}
}

PUBLIC_TRANSPORT = {"mini_bus", "middle_bus", "bus", "trol", "tram"}

def get_ordered_exits(entry):
    if entry == "N": return ["E", "S", "W", "N"]
    elif entry == "S": return ["W", "N", "E", "S"]
    elif entry == "E": return ["S", "W", "N", "E"]
    elif entry == "W": return ["N", "E", "S", "W"]
    else: return []

class DirectionSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор направлений перекрёстка")
        self.setModal(True)
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)

        info_label = QLabel("Выберите, откуда могут въезжать машины (въезды) и куда они могут поворачивать (выезды).\n"
                            "Будут созданы все возможные комбинации въезд → выезд.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        entry_group = QGroupBox("Въезды (откуда едут)")
        entry_layout = QHBoxLayout(entry_group)
        self.entry_checkboxes = {}
        for d in ALL_DIRECTIONS:
            cb = QCheckBox(DIRECTION_NAMES[d])
            cb.setChecked(True)
            self.entry_checkboxes[d] = cb
            entry_layout.addWidget(cb)
        layout.addWidget(entry_group)

        exit_group = QGroupBox("Выезды (куда могут направляться)")
        exit_layout = QHBoxLayout(exit_group)
        self.exit_checkboxes = {}
        for d in ALL_DIRECTIONS:
            cb = QCheckBox(DIRECTION_NAMES[d])
            cb.setChecked(True)
            self.exit_checkboxes[d] = cb
            exit_layout.addWidget(cb)
        layout.addWidget(exit_group)

        btn_layout = QHBoxLayout()
        def select_all_entries(checked):
            for cb in self.entry_checkboxes.values(): cb.setChecked(checked)
        def select_all_exits(checked):
            for cb in self.exit_checkboxes.values(): cb.setChecked(checked)

        btn_all_entries = QPushButton("Въезды: все")
        btn_all_entries.clicked.connect(lambda: select_all_entries(True))
        btn_none_entries = QPushButton("Въезды: снять все")
        btn_none_entries.clicked.connect(lambda: select_all_entries(False))
        btn_all_exits = QPushButton("Выезды: все")
        btn_all_exits.clicked.connect(lambda: select_all_exits(True))
        btn_none_exits = QPushButton("Выезды: снять все")
        btn_none_exits.clicked.connect(lambda: select_all_exits(False))

        btn_layout.addWidget(btn_all_entries)
        btn_layout.addWidget(btn_none_entries)
        btn_layout.addWidget(btn_all_exits)
        btn_layout.addWidget(btn_none_exits)
        layout.addLayout(btn_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_selected_directions(self):
        entries = [d for d, cb in self.entry_checkboxes.items() if cb.isChecked()]
        exits = [d for d, cb in self.exit_checkboxes.items() if cb.isChecked()]
        return entries, exits

class TrafficCounterApp(QMainWindow):
    def __init__(self, entries, exits):
        super().__init__()
        self.entries = entries
        self.exits = exits
        self.directions = {}
        for entry in self.entries:
            ordered = get_ordered_exits(entry)
            filtered = [ex for ex in ordered if ex in self.exits]
            for ex in filtered:
                code = entry + ex
                display = f"{entry} → {ex}"
                self.directions[code] = display

        if not self.directions:
            QMessageBox.critical(self, "Ошибка", "Не выбрано ни одного направления. Приложение закроется.")
            sys.exit(1)

        self.setWindowTitle("Счётчик транспортных средств на перекрёстке")
        self.setMinimumSize(800, 600)
        self.counters = defaultdict(int)
        self.buttons = {}

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Верхняя панель
        top_layout = QHBoxLayout()
        self.cross_name_edit = QLineEdit()
        self.cross_name_edit.setPlaceholderText("Название перекрёстка")
        self.cross_name_edit.setText("Перекрёсток ул. Ленина - ул. Советская")
        top_layout.addWidget(QLabel("Перекрёсток:"))
        top_layout.addWidget(self.cross_name_edit)

        self.date_edit = QLineEdit()
        self.date_edit.setPlaceholderText("Дата записи камеры")
        self.date_edit.setText(datetime.now().strftime("%Y-%m-%d %H:%M"))
        top_layout.addWidget(QLabel("Дата записи:"))
        top_layout.addWidget(self.date_edit)

        self.export_btn = QPushButton("Экспорт в Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        top_layout.addWidget(self.export_btn)

        main_layout.addLayout(top_layout)

        # Прокручиваемая область
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.table_widget = QWidget()
        self.table_layout = QVBoxLayout(self.table_widget)
        self.table_layout.setContentsMargins(10, 10, 10, 10)
        self.table_layout.setSpacing(15)

        for entry in self.entries:
            group = QGroupBox(f"Направления из {DIRECTION_NAMES[entry]}")
            group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            group_layout = QGridLayout(group)
            group_layout.setColumnStretch(0, 0)

            headers = list(VEHICLE_TYPES.keys())
            for col, vtype in enumerate(headers):
                header_widget = QWidget()
                header_layout = QHBoxLayout(header_widget)
                header_layout.setContentsMargins(0, 0, 0, 0)
                label = QLabel(VEHICLE_TYPES[vtype]["label"])
                label.setWordWrap(True)
                label.setAlignment(Qt.AlignCenter)
                label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                info_btn = QPushButton("i")
                info_btn.setFixedSize(22, 22)
                info_btn.setToolTip(VEHICLE_TYPES[vtype]["desc"])
                header_layout.addWidget(label, 1)
                header_layout.addWidget(info_btn, 0)
                group_layout.addWidget(header_widget, 0, col+1)
                group_layout.setColumnStretch(col+1, 1)

            ordered_exits = get_ordered_exits(entry)
            filtered_exits = [ex for ex in ordered_exits if ex in self.exits]
            for i, ex in enumerate(filtered_exits):
                dir_code = entry + ex
                dir_display = self.directions[dir_code]
                dir_label = QLabel(dir_display)
                dir_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                dir_label.setStyleSheet("font-weight: bold;")
                dir_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
                group_layout.addWidget(dir_label, i+1, 0)

                for col, vtype in enumerate(headers):
                    key = (dir_code, vtype)
                    btn = QPushButton(str(self.counters[key]))
                    btn.setMinimumSize(70, 50)
                    btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                    btn.setStyleSheet("""
                        QPushButton {
                            background-color: #e0e0e0;
                            border: 1px solid #aaa;
                            border-radius: 5px;
                            font-size: 14px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #c0c0c0;
                        }
                    """)
                    btn.clicked.connect(lambda checked, d=dir_code, t=vtype: self.increment_counter(d, t))
                    btn.setContextMenuPolicy(Qt.CustomContextMenu)
                    btn.customContextMenuRequested.connect(lambda pos, d=dir_code, t=vtype: self.decrement_counter(d, t))
                    self.buttons[key] = btn
                    group_layout.addWidget(btn, i+1, col+1)

            self.table_layout.addWidget(group)

        scroll.setWidget(self.table_widget)
        main_layout.addWidget(scroll)
        self.statusBar().showMessage("Левая кнопка +1, правая -1 | Окно можно менять размер")

    def increment_counter(self, direction, vtype):
        key = (direction, vtype)
        self.counters[key] += 1
        self.update_button(key)

    def decrement_counter(self, direction, vtype):
        key = (direction, vtype)
        if self.counters[key] > 0:
            self.counters[key] -= 1
            self.update_button(key)

    def update_button(self, key):
        if key in self.buttons:
            self.buttons[key].setText(str(self.counters[key]))

    def export_to_excel(self):
        if not OPENPYXL_AVAILABLE:
            QMessageBox.critical(self, "Ошибка", "Библиотека openpyxl не установлена.\nУстановите её командой: pip install openpyxl")
            return

        cross_name = self.cross_name_edit.text().strip() or "Без названия"
        camera_date = self.date_edit.text().strip() or datetime.now().strftime("%Y-%m-%d %H:%M")

        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить таблицу",
                                                   f"{cross_name}_{camera_date.replace(' ', '_').replace(':', '-')}.xlsx",
                                                   "Excel files (*.xlsx)")
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Перекрёсток"

        ws.merge_cells('A1:I1')
        ws['A1'] = f"Перекрёсток: {cross_name}  |  Дата записи: {camera_date}"
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center')

        headers = ["Направление"] + [VEHICLE_TYPES[v]["label"] for v in VEHICLE_TYPES.keys()]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = h
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[chr(64+col)].width = 20 if col == 1 else 15

        row = 4
        direction_data = {}
        for entry in self.entries:
            ordered_exits = get_ordered_exits(entry)
            for ex in ordered_exits:
                if ex not in self.exits:
                    continue
                dir_code = entry + ex
                dir_display = self.directions[dir_code]
                ws.cell(row=row, column=1, value=dir_display)
                for col, vtype in enumerate(VEHICLE_TYPES.keys(), 2):
                    cnt = self.counters[(dir_code, vtype)]
                    ws.cell(row=row, column=col, value=cnt)
                    direction_data.setdefault(dir_code, {})[vtype] = cnt
                row += 1

        total_all = 0
        total_no_public = 0
        for types_dict in direction_data.values():
            for vtype, cnt in types_dict.items():
                total_all += cnt
                if vtype not in PUBLIC_TRANSPORT:
                    total_no_public += cnt

        public_percent = (total_no_public / total_all * 100) if total_all > 0 else 0

        entry_totals = defaultdict(int)
        entry_exit_counts = defaultdict(lambda: defaultdict(int))
        for dir_code, types_dict in direction_data.items():
            entry = dir_code[0]
            exit_ = dir_code[1]
            total_dir = sum(types_dict.values())
            if total_dir > 0:
                entry_totals[entry] += total_dir
                entry_exit_counts[entry][exit_] += total_dir

        row += 1
        ws.cell(row=row, column=1, value="")
        row += 1

        ws.cell(row=row, column=1, value="Общее количество ТС:")
        ws.cell(row=row, column=2, value=total_all)
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 1

        ws.cell(row=row, column=1, value="Количество ТС без общественного транспорта:")
        ws.cell(row=row, column=2, value=total_no_public)
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 1

        ws.cell(row=row, column=1, value="Доля транспорта без общественного (%):")
        ws.cell(row=row, column=2, value=f"{public_percent:.2f}%")
        ws.cell(row=row, column=1).font = Font(bold=True)
        row += 2

        ws.cell(row=row, column=1, value="ДОЛИ ПОВОРОТОВ ПО ВЪЕЗДАМ")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1

        for entry in self.entries:
            total_entry = entry_totals.get(entry, 0)
            ws.cell(row=row, column=1, value=f"Въезд: {DIRECTION_NAMES[entry]}")
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            ordered_exits = get_ordered_exits(entry)
            filtered_exits = [ex for ex in ordered_exits if ex in self.exits]
            for col, ex in enumerate(filtered_exits, 2):
                ws.cell(row=row, column=col, value=f"→ {ex}")
                ws.cell(row=row, column=col).font = Font(bold=True)
            row += 1
            for ex in filtered_exits:
                cnt = entry_exit_counts[entry].get(ex, 0)
                percent = (cnt / total_entry * 100) if total_entry > 0 else 0
                col = filtered_exits.index(ex) + 2
                ws.cell(row=row, column=col, value=f"{percent:.1f}%")
            row += 1
            row += 1

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for r in range(3, row+1):
            for c in range(1, len(headers)+1):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    cell.border = thin_border

        wb.save(file_path)
        QMessageBox.information(self, "Экспорт завершён", f"Таблица сохранена в:\n{file_path}")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    dialog = DirectionSelectionDialog()
    if dialog.exec() != QDialog.Accepted:
        sys.exit(0)
    entries, exits = dialog.get_selected_directions()
    if not entries or not exits:
        QMessageBox.critical(None, "Ошибка", "Не выбрано ни одного въезда или выезда. Приложение закроется.")
        sys.exit(1)
    window = TrafficCounterApp(entries, exits)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
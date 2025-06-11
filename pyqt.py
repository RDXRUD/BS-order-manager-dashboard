import sys
import os
import pandas as pd
from datetime import date
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QComboBox, QDateEdit,
    QTableView, QSplitter, QCompleter
)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem


# 📌 Custom Table View to override Enter key
class CustomTableView(QTableView):
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            index = self.currentIndex()
            row = index.row() + 1
            column = index.column()
            model = self.model()
            if row < model.rowCount():
                next_index = model.index(row, column)
                self.setCurrentIndex(next_index)
            return  # prevent default behavior
        else:
            super().keyPressEvent(event)


# 🧾 Replace this with your PDF logic
def fetch_products(file_path, selected_date, selected_company, text_company, text_location, order_no, emails):
    print("Generating PDF with:")
    print("File Path:", file_path)
    print("Order No:", order_no)
    print("Company:", text_company)
    print("Location:", text_location)
    print("Emails:", emails)
    print("Date:", selected_date.toString("yyyy-MM-dd"))
    return file_path


class OrderManagementApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ORDER MANAGEMENT")
        self.resize(1000, 500)
        self.setup_ui()

    def setup_ui(self):
        # Layout setup
        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)

        # LEFT Panel: Form
        form_widget = QWidget()
        form_layout = QVBoxLayout()

        # Set base directory
        if getattr(sys, 'frozen', False):
            BASE_DIR = os.path.dirname(sys.executable)
        else:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        self.directory = os.path.join(BASE_DIR, "data files", "Purchase Order")
        self.file_list = [f for f in os.listdir(self.directory)
                          if f.endswith('.xlsx') or f.endswith('.xls')]

        # Order Number
        self.order_no_input = QLineEdit("1")
        form_layout.addLayout(self._make_form_row("Order Number:", self.order_no_input))

        # Company Select (Dropdown + Search)
        self.company_combo = QComboBox()
        self.company_combo.setEditable(True)
        self.company_combo.addItems(self.file_list)

        completer = QCompleter(self.file_list)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.company_combo.setCompleter(completer)

        self.company_combo.lineEdit().editingFinished.connect(
            lambda: self.handle_manual_entry(self.company_combo.currentText())
        )
        self.company_combo.currentTextChanged.connect(self.update_company_name)
        self.company_combo.currentTextChanged.connect(self.load_excel_to_table)

        form_layout.addLayout(self._make_form_row("Select a Company:", self.company_combo))

        # Order Date
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        form_layout.addLayout(self._make_form_row("Order Date:", self.date_edit))

        # Company Name
        self.company_name_input = QLineEdit()
        form_layout.addLayout(self._make_form_row("Company Name:", self.company_name_input))

        # Location
        self.location_input = QLineEdit()
        form_layout.addLayout(self._make_form_row("Company Location:", self.location_input))

        # Emails
        self.email_input = QLineEdit()
        form_layout.addLayout(self._make_form_row("Email IDs:", self.email_input))

        # Generate PDF
        self.submit_btn = QPushButton("Generate PDF")
        self.submit_btn.clicked.connect(self.handle_submit)
        form_layout.addWidget(self.submit_btn)
        form_layout.addStretch()

        form_widget.setLayout(form_layout)
        splitter.addWidget(form_widget)

        # RIGHT Panel: Excel Table Viewer
        self.table_view = CustomTableView()
        splitter.addWidget(self.table_view)
        splitter.setSizes([350, 650])  # Width ratio

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # Initialize view with the first file
        if self.file_list:
            self.update_company_name(self.file_list[0])
            self.load_excel_to_table(self.file_list[0])

    def _make_form_row(self, label_text, widget):
        layout = QHBoxLayout()
        label = QLabel(label_text)
        label.setMinimumWidth(140)
        layout.addWidget(label)
        layout.addWidget(widget)
        return layout

    def handle_manual_entry(self, text):
        if text in self.file_list:
            self.company_combo.setCurrentText(text)
            self.update_company_name(text)
            self.load_excel_to_table(text)

    def update_company_name(self, selected_file):
        self.company_name_input.setText(selected_file.split(".")[0])

    def load_excel_to_table(self, selected_file):
        file_path = os.path.join(self.directory, selected_file)
        try:
            df = pd.read_excel(file_path)
            model = QStandardItemModel()
            model.setColumnCount(len(df.columns))
            model.setHorizontalHeaderLabels(df.columns.astype(str))

            for row in df.itertuples(index=False):
                items = [QStandardItem(str(field)) for field in row]
                model.appendRow(items)

            self.table_view.setModel(model)
            self.table_view.resizeColumnsToContents()

        except Exception as e:
            print(f"Error loading Excel file: {e}")

    def handle_submit(self):
        selected_file = self.company_combo.currentText()
        file_path = os.path.join(self.directory, selected_file)
        fetch_products(
            file_path,
            self.date_edit.date(),
            selected_file,
            self.company_name_input.text(),
            self.location_input.text(),
            self.order_no_input.text(),
            self.email_input.text()
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OrderManagementApp()
    window.show()
    sys.exit(app.exec_())
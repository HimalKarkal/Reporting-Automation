# app.py
import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QComboBox,
    QPushButton,
    QLabel,
    QFileDialog,
    QTextEdit,
    QLineEdit,
    QGroupBox,
    QFormLayout,
    QMessageBox,
    QDateEdit,
)
from PyQt5.QtCore import QDate, Qt

# Assuming functions.py is in the same directory
import functions


class ReportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.df = None

        self.club_list = [
            "DeakinACTIVE Waurn Ponds",  # Default/example first for easy selection
            "DeakinACTIVE Burwood",
            "DeakinACTIVE Waterfront",
            "DeakinACTIVE Warrnambool",
        ]
        self.example_target_club = "DeakinACTIVE Waurn Ponds"  # Matches one in the list
        # self.example_payment_plan removed as it's fixed in functions.py
        self.example_end_date = "2025-04-30"  # Current date is 2025-05-14

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Excel Reporting Automation")
        self.setGeometry(100, 100, 800, 700)

        main_layout = QVBoxLayout()

        # Report Selection
        report_selection_layout = QHBoxLayout()
        report_label = QLabel("Select Report Type:")
        self.report_combo = QComboBox()
        self.report_options = {
            "--Select Report--": None,
            "Current Members": functions.current_members,
            "New Members": functions.new_members,
            "Technogym Reporting (Consults/PT)": functions.technogym_reporting,
            "Group Fitness Summary": functions.groupFitness,
            "Booking Zones Analysis": functions.booking_zones,
        }
        self.report_combo.addItems(self.report_options.keys())
        self.report_combo.currentTextChanged.connect(self.on_report_type_change)
        report_selection_layout.addWidget(report_label)
        report_selection_layout.addWidget(self.report_combo)
        main_layout.addLayout(report_selection_layout)

        # File Upload
        file_upload_layout = QHBoxLayout()
        self.upload_button = QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.upload_file)
        self.file_label = QLabel("No file selected.")
        self.file_label.setWordWrap(True)
        file_upload_layout.addWidget(self.upload_button)
        file_upload_layout.addWidget(self.file_label, 1)
        main_layout.addLayout(file_upload_layout)

        # Parameter GroupBox
        self.params_groupbox = QGroupBox("Report Parameters")
        self.params_form_layout = QFormLayout()
        self.params_groupbox.setLayout(self.params_form_layout)
        self.params_groupbox.setVisible(False)

        self.target_club_combo = QComboBox()  # Changed from QLineEdit
        self.target_club_combo.addItems(self.club_list)

        # self.payment_plan_edit QLineEdit removed

        self.end_date_edit = QDateEdit(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")

        self.params_form_layout.addRow(QLabel("Target Club:"), self.target_club_combo)
        # Row for payment_plan_edit removed
        self.params_form_layout.addRow(QLabel("End Date:"), self.end_date_edit)
        main_layout.addWidget(self.params_groupbox)

        # Generate Report Button
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.clicked.connect(self.generate_report)
        main_layout.addWidget(self.generate_button)

        # Output Display
        self.output_display = QTextEdit()
        self.output_display.setReadOnly(True)
        self.output_display.setLineWrapMode(QTextEdit.NoWrap)
        main_layout.addWidget(self.output_display)

        self.setLayout(main_layout)
        self.show()

    def on_report_type_change(self, report_name):
        if report_name in ["Current Members", "New Members"]:
            # Set example club in the ComboBox
            if self.example_target_club in self.club_list:
                self.target_club_combo.setCurrentText(self.example_target_club)
            # Payment plan input is removed
            q_end_date = QDate.fromString(self.example_end_date, "yyyy-MM-dd")
            if q_end_date.isValid():
                self.end_date_edit.setDate(q_end_date)
            else:
                self.end_date_edit.setDate(QDate.currentDate())
            self.params_groupbox.setVisible(True)
        else:
            self.params_groupbox.setVisible(False)

    def upload_file(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(
            self,
            "Upload Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
            options=options,
        )
        if filePath:
            self.file_path = filePath
            self.file_label.setText(filePath.split("/")[-1])
            self.output_display.setText(
                f"Selected file: {self.file_path}\nReady to load data."
            )

    def generate_report(self):
        if not self.file_path:
            QMessageBox.warning(self, "Warning", "Please upload an Excel file first.")
            return

        selected_report_name = self.report_combo.currentText()
        if selected_report_name == "--Select Report--":
            QMessageBox.warning(self, "Warning", "Please select a report type.")
            return

        report_function = self.report_options[selected_report_name]

        try:
            self.output_display.setText(f"Loading file: {self.file_label.text()}...")
            QApplication.processEvents()

            if selected_report_name == "Group Fitness Summary":
                self.df = pd.read_excel(self.file_path, skiprows=1)
            else:
                self.df = pd.read_excel(self.file_path, engine="openpyxl")

            self.output_display.append(
                f"File loaded successfully.\nProcessing report: {selected_report_name}..."
            )
            QApplication.processEvents()

            result = None
            if selected_report_name in ["Current Members", "New Members"]:
                target_club = self.target_club_combo.currentText()  # Get from ComboBox
                end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd")

                # Payment plan list is no longer taken from GUI as it's fixed in functions.py
                result = report_function(self.df, target_club, end_date_str)

            elif report_function:
                result = report_function(self.df)

            if result is not None:
                if isinstance(result, pd.DataFrame):
                    self.output_display.setText(
                        f"--- {selected_report_name} Results ---\n\n{result.to_string()}"
                    )
                elif isinstance(result, dict):
                    output_str = f"--- {selected_report_name} Results ---\n\n"
                    for key, value in result.items():
                        output_str += f"{key}: {value}\n"
                    self.output_display.setText(output_str)
                elif isinstance(result, list):
                    self.output_display.setText(
                        f"--- {selected_report_name} Results ---\n\n{str(result)}"
                    )
                else:
                    self.output_display.setText(
                        f"--- {selected_report_name} Results ---\n\n{str(result)}"
                    )
            else:
                self.output_display.setText(
                    f"No result generated for {selected_report_name}, or the report function returned None."
                )

        except FileNotFoundError:
            QMessageBox.critical(self, "Error", f"File not found: {self.file_path}")
            self.output_display.setText(f"Error: File not found {self.file_path}")
        except ValueError as ve:
            QMessageBox.critical(self, "Error", f"Value error during processing: {ve}")
            self.output_display.setText(
                f"Value error: {ve}\nPlease check your input file and parameters."
            )
        except KeyError as ke:
            QMessageBox.critical(
                self, "Error", f"Missing column or data in Excel: {ke}"
            )
            self.output_display.setText(
                f"Key error: {ke}\nPlease ensure the Excel file has the correct columns."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {e}")
            self.output_display.setText(
                f"An unexpected error occurred: {type(e).__name__} - {e}"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ReportApp()
    sys.exit(app.exec_())

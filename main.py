import sys
import pandas as pd
import datetime as dt
import os
import webbrowser
import subprocess

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QComboBox,
    QPushButton,
    QLabel,
    QFileDialog,
    QTextEdit,
    QGroupBox,
    QFormLayout,
    QMessageBox,
    QDateEdit,
)
from PySide6.QtCore import QDate
from PySide6.QtGui import QIcon

import functions


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")  # In development, use current directory
    return os.path.join(base_path, relative_path)


class ReportingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.df = None  # General DataFrame, loaded per report generation

        self.club_list = [
            "DeakinACTIVE Waurn Ponds",
            "DeakinACTIVE Burwood",
            "DeakinACTIVE Waterfront",
            "DeakinACTIVE Warrnambool",
        ]
        # Default for example_target_club. If club_list is empty, it becomes "Select Target Club"
        # Otherwise, it's the first actual club name for pre-selection demonstration.
        self.example_target_club = (
            self.club_list[0] if self.club_list else "Select Target Club"
        )
        self.example_end_date = str(dt.date.today())

        self.initUI()
        self.setAppIcon()  # Call to set the window icon

    def setAppIcon(self):
        """Sets the window icon."""
        icon_filename = "icon.png"  # Your PNG icon file
        try:
            icon_path_resolved = resource_path(icon_filename)
            if os.path.exists(icon_path_resolved):
                self.setWindowIcon(QIcon(icon_path_resolved))
            else:
                print(
                    f"Warning: Window icon file '{icon_filename}' not found at resolved path: {icon_path_resolved}"
                )
        except Exception as e:
            print(f"Error setting window icon: {e}")

    def initUI(self):
        self.setWindowTitle("DeakinACTIVE Reporting Tool")
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

        self.target_club_combo = QComboBox()
        self.target_club_combo.addItems(
            ["--Select Club--"] + self.club_list
        )  # Add a placeholder

        self.end_date_edit = QDateEdit(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")

        self.params_form_layout.addRow(QLabel("Target Club:"), self.target_club_combo)
        self.params_form_layout.addRow(QLabel("End Date:"), self.end_date_edit)
        main_layout.addWidget(self.params_groupbox)

        # Generate Report Button
        self.generate_button = QPushButton("Generate Report")
        self.generate_button.clicked.connect(self.generate_report)
        main_layout.addWidget(self.generate_button)

        # Output Display
        self.output_display = QTextEdit()
        self.output_display.setReadOnly(True)
        self.output_display.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        main_layout.addWidget(self.output_display)

        self.setLayout(main_layout)
        self.show()

    def on_report_type_change(self, report_name):
        self.output_display.clear()
        # Optionally reset file label if you want to force re-upload for each report type change
        # self.file_label.setText("No file selected.")
        # self.file_path = None

        if report_name in ["Current Members", "New Members"]:
            if self.example_target_club in self.club_list:
                self.target_club_combo.setCurrentText(self.example_target_club)
            else:
                self.target_club_combo.setCurrentIndex(
                    0
                )  # Default to "--Select Club--"

            q_end_date = QDate.fromString(self.example_end_date, "yyyy-MM-dd")
            if q_end_date.isValid():
                self.end_date_edit.setDate(q_end_date)
            else:
                self.end_date_edit.setDate(QDate.currentDate())
            self.params_groupbox.setVisible(True)
        else:
            self.params_groupbox.setVisible(False)

    def upload_file(self):
        filePath, _ = QFileDialog.getOpenFileName(
            self,
            "Upload Excel File",
            "",  # Start in the last-used directory or a default one
            "Excel Files (*.xlsx *.xls);;All Files (*)",
        )
        if filePath:
            self.file_path = filePath
            self.file_label.setText(
                os.path.basename(filePath)
            )  # Use basename for cleaner filename
            self.output_display.setText(
                f"Selected file: {self.file_path}\nReady to load data."
            )

    def _open_file_externally(self, filepath: str):
        """Attempts to open the given file with its default system application."""
        try:
            abs_filepath = os.path.abspath(filepath)
            if not os.path.exists(abs_filepath):
                self.output_display.append(
                    f"\nError: File not found at {abs_filepath} for opening."
                )
                QMessageBox.warning(
                    self, "File Open Error", f"File not found: {abs_filepath}"
                )
                return

            if sys.platform == "win32":
                os.startfile(abs_filepath)
            elif sys.platform == "darwin":
                subprocess.run(["open", abs_filepath], check=True)
            else:  # Linux and other POSIX
                subprocess.run(["xdg-open", abs_filepath], check=True)
            self.output_display.append(f"\nAttempted to open: {abs_filepath}")
        except AttributeError:  # os.startfile might not be available
            try:
                webbrowser.open("file://" + abs_filepath)
                self.output_display.append(
                    f"\nAttempted to open via webbrowser: {abs_filepath}"
                )
            except Exception as e_wb:
                self.output_display.append(
                    f"\nCould not automatically open file (webbrowser fallback failed): {str(e_wb)}"
                )
                QMessageBox.warning(
                    self,
                    "File Open Failed",
                    f"Could not automatically open file.\nPlease open it manually from:\n{abs_filepath}",
                )
        except FileNotFoundError as e_fnf:  # For open/xdg-open if command not found
            self.output_display.append(
                f"\nCould not automatically open file (command not found): {str(e_fnf)}\nPlease open it manually."
            )
            QMessageBox.warning(
                self,
                "File Open Failed",
                f"Could not automatically open file: Required command not found.\nPlease open it manually from:\n{abs_filepath}",
            )
        except subprocess.CalledProcessError as e_cpe:  # For errors from open/xdg-open
            self.output_display.append(
                f"\nError opening file with system command: {str(e_cpe)}\nPlease open it manually."
            )
            QMessageBox.warning(
                self,
                "File Open Failed",
                f"Error opening file: {str(e_cpe)}\nPlease open it manually from:\n{abs_filepath}",
            )
        except Exception as e:  # Catch-all for other errors
            self.output_display.append(
                f"\nCould not automatically open file: {str(e)}\nPlease open it manually."
            )
            QMessageBox.warning(
                self,
                "File Open Failed",
                f"Could not automatically open file: {str(e)}\nPlease open it manually from:\n{abs_filepath}",
            )

    def generate_report(self):
        if not self.file_path:
            QMessageBox.warning(self, "Warning", "Please upload an Excel file first.")
            self.output_display.setText("Operation cancelled: No Excel file uploaded.")
            return

        selected_report_name = self.report_combo.currentText()
        if selected_report_name == "--Select Report--":
            QMessageBox.warning(self, "Warning", "Please select a report type.")
            self.output_display.setText("Operation cancelled: No report type selected.")
            return

        report_function = self.report_options[selected_report_name]
        if not report_function:
            self.output_display.setText(
                f"No function associated with report: {selected_report_name}"
            )
            return

        try:
            self.output_display.setText(
                f"Loading file: {os.path.basename(self.file_path)}..."
            )
            QApplication.processEvents()

            if selected_report_name == "Group Fitness Summary":
                self.df = pd.read_excel(self.file_path, skiprows=1)
            else:
                self.df = pd.read_excel(self.file_path, engine="openpyxl")

            self.output_display.append(
                f"File loaded successfully.\nProcessing report: '{selected_report_name}'..."
            )
            QApplication.processEvents()

            if selected_report_name == "Booking Zones Analysis":
                result_df = functions.booking_zones(self.df)
                if isinstance(result_df, pd.DataFrame):
                    report_name_slug = "Booking_Zones_Analysis"
                    suggested_filename = f"{report_name_slug}_report_{dt.date.today().strftime('%Y%m%d')}.csv"
                    filePath, _ = QFileDialog.getSaveFileName(
                        self,
                        "Save Booking Zones Report",
                        suggested_filename,
                        "CSV Files (*.csv);;All Files (*)",
                    )
                    if filePath:
                        try:
                            result_df.to_csv(filePath, index=False)
                            self.output_display.setText(
                                f"Report saved to:\n{filePath}\nAttempting to open..."
                            )
                            QApplication.processEvents()
                            self._open_file_externally(filePath)
                        except Exception as e_save_open:
                            QMessageBox.critical(
                                self,
                                "Error",
                                f"Failed to save or open report: {str(e_save_open)}",
                            )
                            self.output_display.setHtml(
                                f"<b><font color='red'>Error during Booking Zones report save/open:</font></b><br>{str(e_save_open)}"
                            )
                    else:
                        self.output_display.setText(
                            "Booking Zones Analysis report generation cancelled (save dialogue cancelled)."
                        )
                else:
                    self.output_display.setHtml(
                        f"<b><font color='red'>Error:</font></b> '{selected_report_name}' did not produce valid data for CSV export."
                    )
                return

            result = None
            if selected_report_name in ["Current Members", "New Members"]:
                target_club = self.target_club_combo.currentText()
                if target_club == "--Select Club--":
                    QMessageBox.warning(
                        self,
                        "Parameter Missing",
                        "Please select a Target Club for this report.",
                    )
                    self.output_display.setText(
                        "Report generation cancelled: Target Club not selected."
                    )
                    return
                end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd")
                result = report_function(self.df, target_club, end_date_str)
            else:
                result = report_function(self.df)

            if result is not None:
                html_title = f"<h3>--- {selected_report_name} Results ---</h3>"
                plain_text_title = f"--- {selected_report_name} Results ---\n\n"

                if isinstance(result, pd.DataFrame):
                    formatted_string = result.to_string(
                        index=False, justify="left", float_format="{:.2f}".format
                    )
                    self.output_display.setHtml(
                        f"{html_title}<pre>{formatted_string}</pre>"
                    )
                elif isinstance(result, str):
                    is_html_result = any(
                        tag in result.lower()
                        for tag in ["<table", "<p>", "<h3>", "<b>", "<br>", "<font"]
                    )
                    if is_html_result:
                        self.output_display.setHtml(html_title + result)
                    else:
                        self.output_display.setText(plain_text_title + result)
                else:
                    self.output_display.setText(plain_text_title + str(result))
            else:
                self.output_display.setText(
                    f"No result generated for '{selected_report_name}', or the report function returned None."
                )

        except FileNotFoundError:
            QMessageBox.critical(self, "Error", f"File not found: {self.file_path}")
            self.output_display.setText(f"Error: File not found {self.file_path}")
        except ValueError as ve:
            QMessageBox.critical(
                self, "Data Error", f"Error processing file or data: {ve}"
            )
            self.output_display.setHtml(
                f"<b><font color='red'>Data Error:</font></b> {ve}<br>Please check if the Excel file is corrupted, has an unexpected format, or if report parameters are valid."
            )
        except KeyError as ke:
            error_message_for_display = (
                f"<b><font color='red'>Error: Missing expected column or data.</font></b><br><br>"
                f"This often means the selected Excel file (<b>{os.path.basename(self.file_path)}</b>) "
                f"is not the correct one for the '<b>{selected_report_name}</b>' report, or the sheet structure is wrong.<br><br>"
                f"Please verify the file and its contents.<br><br>"
                f"<i>Technical Detail (missing column/key): {ke}</i>"
            )
            QMessageBox.critical(
                self,
                "Data Structure Error",
                f"Missing column/data: {ke}. Is this the correct Excel file and sheet for '{selected_report_name}'?",
            )
            self.output_display.setHtml(error_message_for_display)
        except Exception as e:
            error_message_for_display = (
                f"<b><font color='red'>An unexpected error occurred:</font></b><br>"
                f"Report: '{selected_report_name}'<br>"
                f"File: {os.path.basename(self.file_path)}<br>"
                f"Error Type: {type(e).__name__}<br>"
                f"Details: {str(e)}"
            )
            QMessageBox.critical(
                self, "Unexpected Error", f"An unexpected error occurred: {e}"
            )
            self.output_display.setHtml(error_message_for_display)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    try:
        # Use icon.png for runtime QIcon, PyInstaller --icon will use icon.ico for the .exe
        app_icon_filename = "icon.png"
        resolved_app_icon_path = resource_path(app_icon_filename)
        if os.path.exists(resolved_app_icon_path):
            app.setWindowIcon(QIcon(resolved_app_icon_path))
        else:
            print(
                f"Warning: Global application icon file '{app_icon_filename}' not found at resolved path: {resolved_app_icon_path}"
            )
    except Exception as e:
        print(f"Error setting global application icon: {e}")

    ex = ReportingApp()
    sys.exit(app.exec())

import sys
import pandas as pd  # Changed from polars

# Removed Polars specific exception imports if not used elsewhere
# from polars.exceptions import NoDataError, ColumnNotFoundError, ComputeError
from pandas.errors import EmptyDataError  # Pandas equivalent for empty data
import datetime as dt
import os
import webbrowser
import subprocess
import warnings
import inspect

warnings.filterwarnings(
    "ignore",
    message="Could not determine dtype for column",  # This warning is often from pandas read_excel
    category=UserWarning,
)
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",  # Common openpyxl warning
    category=UserWarning,
)


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

import functions  # Your functions.py file (now all pandas)


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ReportingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.df_pandas: pd.DataFrame | None = None  # Changed from df_polars

        self.club_list = [
            "DeakinACTIVE Waurn Ponds",
            "DeakinACTIVE Burwood",
            "DeakinACTIVE Waterfront",
            "DeakinACTIVE Warrnambool",
        ]
        self.example_target_club = (
            self.club_list[0] if self.club_list else "Select Target Club"
        )
        self.example_end_date = str(dt.date.today())

        self.initUI()
        self.setAppIcon()

    def setAppIcon(self):
        icon_filename = "icon.png"
        try:
            icon_path_resolved = resource_path(icon_filename)
            if os.path.exists(icon_path_resolved):
                self.setWindowIcon(QIcon(icon_path_resolved))
            else:
                print(
                    f"Warning: icon file '{icon_filename}' not found: {icon_path_resolved}"
                )
        except Exception as e:
            print(f"Error setting window icon: {e}")

    def initUI(self):
        self.setWindowTitle("DeakinACTIVE Reporting Tool")  # Updated title
        self.setGeometry(100, 100, 850, 750)
        main_layout = QVBoxLayout()

        report_selection_layout = QHBoxLayout()
        report_label = QLabel("Select Report Type:")
        self.report_combo = QComboBox()
        # Ensure these function names match your all-pandas functions.py
        self.report_options = {
            "--Select Report--": None,
            "Current Members": functions.current_members,
            "New Members": functions.new_members,  # Matches your provided functions.py
            "Technogym Reporting (Consults/PT)": functions.technogym_reporting,
            "Group Fitness Summary": functions.groupFitness,
            "Booking Zones Analysis": functions.booking_zones,
            "Ending Members Report": functions.generate_ending_members_report,  # New function
        }
        self.report_combo.addItems(self.report_options.keys())
        self.report_combo.currentTextChanged.connect(self.on_report_type_change)
        report_selection_layout.addWidget(report_label)
        report_selection_layout.addWidget(self.report_combo)
        main_layout.addLayout(report_selection_layout)

        file_upload_layout = QHBoxLayout()
        self.upload_button = QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.upload_file)
        self.file_label = QLabel("No file selected.")
        self.file_label.setWordWrap(True)
        file_upload_layout.addWidget(self.upload_button)
        file_upload_layout.addWidget(self.file_label, 1)
        main_layout.addLayout(file_upload_layout)

        self.params_groupbox = QGroupBox("Report Parameters")
        self.params_form_layout = QFormLayout()
        self.params_groupbox.setLayout(self.params_form_layout)
        self.params_groupbox.setVisible(False)
        self.target_club_combo = QComboBox()
        self.target_club_combo.addItems(["--Select Club--"] + self.club_list)
        self.date_param_label = QLabel("Target Date:")
        self.end_date_edit = QDateEdit(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.params_form_layout.addRow(QLabel("Target Club:"), self.target_club_combo)
        self.params_form_layout.addRow(self.date_param_label, self.end_date_edit)
        main_layout.addWidget(self.params_groupbox)

        self.generate_button = QPushButton("Generate Report")
        self.generate_button.clicked.connect(self.generate_report)
        main_layout.addWidget(self.generate_button)

        self.output_display = QTextEdit()
        self.output_display.setReadOnly(True)
        main_layout.addWidget(self.output_display)

        self.setLayout(main_layout)
        self.show()

    def on_report_type_change(self, report_name: str):
        self.output_display.clear()
        reports_needing_params = ["Current Members", "New Members"]
        if report_name in reports_needing_params:
            if self.example_target_club in self.club_list:
                self.target_club_combo.setCurrentText(self.example_target_club)
            else:
                self.target_club_combo.setCurrentIndex(0)
            q_end_date = QDate.fromString(self.example_end_date, "yyyy-MM-dd")
            self.end_date_edit.setDate(
                q_end_date if q_end_date.isValid() else QDate.currentDate()
            )
            if report_name == "New Members":  # Uses end_date to determine month/year
                self.date_param_label.setText("Target Month (select any day in month):")
            else:
                self.date_param_label.setText("End Date:")
            self.params_groupbox.setVisible(True)
        else:
            self.params_groupbox.setVisible(False)

    def upload_file(self):
        filePath, _ = QFileDialog.getOpenFileName(
            self,
            "Upload Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)",
        )
        if filePath:
            self.file_path = filePath
            self.file_label.setText(os.path.basename(filePath))
            self.output_display.setText(
                f"Selected file: {self.file_path}\nReady to load data."
            )
            self.df_pandas = None  # Reset cached DataFrame

    def _open_file_externally(self, filepath: str):
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
            else:
                subprocess.run(["xdg-open", abs_filepath], check=True)
            self.output_display.append(f"\nAttempted to open: {abs_filepath}")
        except AttributeError:
            try:
                webbrowser.open("file://" + abs_filepath)
            except Exception as e_wb:
                QMessageBox.warning(
                    self,
                    "File Open Failed",
                    f"Could not open file (webbrowser fallback failed): {e_wb}\nPlease open manually:\n{abs_filepath}",
                )
                self.output_display.append(f"\nCould not open via webbrowser: {e_wb}")
        except Exception as e:
            QMessageBox.warning(
                self,
                "File Open Failed",
                f"Could not open file: {str(e)}\nPlease open manually:\n{abs_filepath}",
            )
            self.output_display.append(f"\nError opening file: {str(e)}")

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

        report_function = self.report_options.get(selected_report_name)
        if not callable(report_function):
            self.output_display.setHtml(
                f"<b><font color='red'>Configuration Error:</font></b><br>No valid function for report: {selected_report_name}"
            )
            return

        try:
            self.output_display.setText(
                f"Loading file: {os.path.basename(self.file_path)}..."
            )
            QApplication.processEvents()

            # --- Data Loading: All reports now use pandas DataFrame ---
            df_loaded_pandas = None
            if selected_report_name == "Group Fitness Summary":
                try:
                    df_loaded_pandas = pd.read_excel(self.file_path, skiprows=1)
                except Exception as e_load_gf:
                    msg = (
                        f"Could not load Group Fitness sheet using pandas: {e_load_gf}. "
                        "Ensure Excel is valid, 1st row is skippable, 2nd row has headers."
                    )
                    QMessageBox.critical(self, "Load Error (Group Fitness)", msg)
                    self.output_display.setHtml(
                        f"<b><font color='red'>Load Error:</font></b><br>{msg}"
                    )
                    return
            else:  # Standard pandas load for all other reports
                try:
                    df_loaded_pandas = pd.read_excel(self.file_path)
                except (
                    EmptyDataError
                ):  # Catch pandas specific error for empty file/sheet
                    msg = f"<b><font color='red'>Data Error:</font></b><br>No data in Excel file/sheet: {self.file_path}. File might be empty."
                    QMessageBox.critical(
                        self,
                        "Data Error",
                        f"No data in Excel file/sheet: {self.file_path}.",
                    )
                    self.output_display.setHtml(msg)
                    return
                except Exception as e_load:  # General pandas load error
                    msg = f"<b><font color='red'>File Load Error:</font></b><br>Could not load Excel file with pandas: {e_load}"
                    QMessageBox.critical(
                        self, "Load Error", f"Failed to load Excel file: {e_load}"
                    )
                    self.output_display.setHtml(msg)
                    return

            # Cache the loaded pandas DataFrame if needed by multiple calls, or just use df_loaded_pandas
            self.df_pandas = df_loaded_pandas

            self.output_display.append(
                f"File loaded.\nProcessing: '{selected_report_name}'..."
            )
            QApplication.processEvents()

            # --- Report Function Execution ---
            # All functions now expect a pandas DataFrame

            if (
                selected_report_name == "Booking Zones Analysis"
                or selected_report_name == "Ending Members Report"
            ):
                # These functions return a pandas DataFrame to be saved as CSV
                returned_df = report_function(self.df_pandas)

                if isinstance(returned_df, pd.DataFrame):
                    slug = selected_report_name.replace(
                        " ", "_"
                    )  # Generate slug from report name
                    sugg_fname = f"{slug}_{dt.date.today().strftime('%Y%m%d')}.csv"
                    filePath, _ = QFileDialog.getSaveFileName(
                        self,
                        f"Save {selected_report_name}",
                        sugg_fname,
                        "CSV (*.csv);;All (*)",
                    )
                    if filePath:
                        try:
                            returned_df.to_csv(
                                filePath, index=False
                            )  # Pandas save to CSV
                            self.output_display.setText(
                                f"{selected_report_name} saved: {filePath}"
                            )
                            self._open_file_externally(filePath)
                        except Exception as e_save:
                            QMessageBox.critical(
                                self, "Error", f"Failed to save/open report: {e_save}"
                            )
                            self.output_display.setHtml(
                                f"<b><font color='red'>Save/Open Error:</font></b><br>{e_save}"
                            )
                    else:
                        self.output_display.setText(
                            f"{selected_report_name} save cancelled."
                        )
                else:
                    self.output_display.setHtml(
                        f"<b><font color='red'>Report Error:</font></b><br>{selected_report_name} did not return a pandas DataFrame as expected. Got: {type(returned_df).__name__}"
                    )
                return  # Specific handling done

            # --- For reports that return HTML directly ---
            result_display_data = None
            if selected_report_name == "Current Members":
                target_club = self.target_club_combo.currentText()
                if target_club == "--Select Club--":
                    QMessageBox.warning(
                        self, "Input Missing", "Please select a Target Club."
                    )
                    self.output_display.setText("Cancelled: Target Club not selected.")
                    return
                end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd")
                result_display_data = report_function(
                    self.df_pandas, target_club, end_date_str
                )

            elif selected_report_name == "New Members":
                target_club = self.target_club_combo.currentText()
                if target_club == "--Select Club--":
                    QMessageBox.warning(
                        self, "Input Missing", "Please select a Target Club."
                    )
                    self.output_display.setText("Cancelled: Target Club not selected.")
                    return
                end_date_str = self.end_date_edit.date().toString(
                    "yyyy-MM-dd"
                )  # function new_members takes this
                result_display_data = report_function(
                    self.df_pandas, target_club, end_date_str
                )

            else:  # Generic path for other reports returning HTML (Technogym, Group Fitness)
                try:
                    result_display_data = report_function(self.df_pandas)
                except TypeError as te:
                    sig = inspect.signature(report_function)
                    num_expected_args = len(
                        [
                            p
                            for p_name, p in sig.parameters.items()
                            if p.default == inspect.Parameter.empty
                            and p_name not in ("df", "df_input", "df_pandas_input")
                        ]
                    )
                    if num_expected_args > 0:
                        self.output_display.setHtml(
                            f"<b><font color='red'>Parameter Error:</font></b><br>Report '{selected_report_name}' requires additional parameters.<br>Function signature: {str(sig)}<br>Error: {te}"
                        )
                    else:
                        self.output_display.setHtml(
                            f"<b><font color='red'>Execution Error:</font></b><br>Report '{selected_report_name}' encountered a TypeError.<br>Details: {te}"
                        )
                    return

            # --- Displaying results (HTML strings) ---
            if result_display_data is not None and isinstance(result_display_data, str):
                title = f"<h3>--- {selected_report_name} Results ---</h3>"
                self.output_display.setHtml(title + result_display_data)
            elif result_display_data is None:
                self.output_display.setHtml(
                    f"<h3>--- {selected_report_name} Results ---</h3><p>Report generated, but no specific data was returned (None).</p>"
                )
            elif not isinstance(result_display_data, str):
                self.output_display.setHtml(
                    f"<b><font color='red'>Display Error:</font></b><br>Report '{selected_report_name}' did not return an HTML string. Type: {type(result_display_data).__name__}"
                )

        except (
            FileNotFoundError
        ):  # Should be caught by pandas read_excel if path is wrong
            msg = f"<b><font color='red'>File Error:</font></b><br>Input Excel file not found: {self.file_path}"
            QMessageBox.critical(
                self, "Error", f"Input Excel file not found: {self.file_path}"
            )
            self.output_display.setHtml(msg)
        # Removed Polars specific NoDataError and ComputeError
        except pd.errors.EmptyDataError:  # Pandas specific for empty data
            msg = f"<b><font color='red'>Data Error (Pandas):</font></b><br>No data in Excel file/sheet (Pandas): {self.file_path}."
            QMessageBox.critical(
                self,
                "Data Error",
                f"No data in Excel file/sheet (Pandas): {self.file_path}.",
            )
            self.output_display.setHtml(msg)
        except (KeyError, AttributeError) as ke:  # Common pandas errors
            col_name = ke.args[0] if ke.args else str(ke)
            msg = (
                f"<b><font color='red'>Error:</font></b><br>"
                f"File: <b>{os.path.basename(self.file_path) if self.file_path else 'N/A'}</b>, Report: '<b>{selected_report_name}</b>'<br>"
                f"Details: Error with column/key/attribute '<b>{col_name}</b>'.<br>"
                f"Ensure the Excel sheet has correct headers/data."
            )
            QMessageBox.critical(
                self,
                "Data/Attribute Error",
                f"Error with '{col_name}' for '{selected_report_name}'.",
            )
            self.output_display.setHtml(msg)
        except (
            ValueError
        ) as ve:  # For ValueErrors raised by report functions or data validation
            msg = f"<b><font color='red'>Input/Value Error:</font></b><br>{ve}"  # ValueError message is usually descriptive
            QMessageBox.critical(
                self, "Data/Parameter Error", f"Error with data or parameters: {ve}"
            )
            self.output_display.setHtml(msg)
        except Exception as e:
            msg = (
                f"<b><font color='red'>An Unexpected Error Occurred:</font></b><br>"
                f"Report: '{selected_report_name}'<br>"
                f"File: {os.path.basename(self.file_path) if self.file_path else 'N/A'}<br>"
                f"Type: {type(e).__name__}<br>Details: {str(e)}"
            )
            QMessageBox.critical(self, "Unexpected Error", f"Unexpected error: {e}")
            self.output_display.setHtml(msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    try:
        app_icon_filename = "icon.png"
        resolved_app_icon_path = resource_path(app_icon_filename)
        if os.path.exists(resolved_app_icon_path):
            app.setWindowIcon(QIcon(resolved_app_icon_path))
        else:
            print(
                f"Warning: App icon '{app_icon_filename}' not found at: {resolved_app_icon_path}"
            )
    except Exception as e:
        print(f"Error setting app icon: {e}")

    ex = ReportingApp()
    sys.exit(app.exec())

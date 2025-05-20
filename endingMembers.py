import polars as pl
import datetime as dt
import warnings
import os  # Import the os module
import subprocess  # For more cross-platform opening (alternative)
import sys  # To check the platform

# Suppress specific fastexcel dtype warnings if they appear and are not critical
warnings.filterwarnings(
    "ignore",
    message="Could not determine dtype for column",
    category=UserWarning,
    module="fastexcel.types.dtype",
)

# Get today's date
today = dt.date.today()

# Ask the user for input
path = input("Enter the path to the Excel file then hit enter: ")
path = path[
    1, -1
]  # Remove the first and last character of the path. Preventing double quotes.

try:
    # Read the Excel file
    df = pl.read_excel(source=path)

    # Retain necessary columns
    required_columns = [
        "Name",
        "Last name",
        "Club",
        "Payment Plan Name",
        "End date",
        "Email",
        "Mobile number",
    ]
    df = df.select(required_columns)

    # Ensure 'End date' is a date type for comparison
    df = df.with_columns(pl.col("End date").cast(pl.Date))

    # Filter those ending today
    df_filtered = df.filter(pl.col("End date") == today)

    # Download to excel
    output_filename = "ending_members.xlsx"
    # Get the absolute path to the output file
    # This ensures it works correctly regardless of where the script/exe is run from
    absolute_output_path = os.path.abspath(output_filename)

    df_filtered.write_excel(absolute_output_path)
    print(f"File saved as {absolute_output_path}")

    # --- Add code to open the file ---
    print(f"Attempting to open {absolute_output_path}...")
    try:
        if sys.platform == "win32":  # Check if running on Windows
            os.startfile(absolute_output_path)
        elif sys.platform == "darwin":  # Check if running on macOS
            subprocess.call(["open", absolute_output_path])
        else:  # Assume Linux or other Unix-like
            subprocess.call(["xdg-open", absolute_output_path])
        print("File should now be opening in your default Excel application.")
    except FileNotFoundError:
        # This might happen if xdg-open or open is not found, or if the file itself wasn't created properly
        print(f"Error: Could not find the file {absolute_output_path} to open it.")
    except OSError as e:
        # This can happen if there's no default application for .xlsx files
        # or other OS-level issues.
        print(
            f"Error: Could not open the file. Your system might not have a default application for .xlsx files, or another issue occurred: {e}"
        )
    except Exception as e:
        print(f"An unexpected error occurred while trying to open the file: {e}")
    # --- End of code to open the file ---

except pl.exceptions.ColumnNotFoundError as e:
    print(f"Error: A required column was not found in the Excel file. Details: {e}")
    print(
        f"Please ensure your Excel file contains the following columns: {', '.join(required_columns)}"
    )
except pl.exceptions.ComputeError as e:
    print(
        f"Error: Could not process the 'End date' column. Is it in a recognizable date format? Details: {e}"
    )
except FileNotFoundError:
    print(f"Error: The input file was not found at the specified path: {path}")
except Exception as e:
    print(f"An unexpected error occurred during processing: {e}")

input("Press Enter to exit.")

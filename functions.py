import pandas as pd
import datetime as dt

# --- Report Generation Functions ---


def current_members(df: pd.DataFrame, target_club: str, end_date: str) -> str:
    """
    Counts current members for a target club and end date.
    Returns an HTML formatted table string.
    """
    try:
        required_cols = ["Club", "Payment plan type", "End date"]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            raise ValueError(
                f"Current Members: Missing required columns: {', '.join(missing)}"
            )

        df_copy = df.copy()  # Work on a copy to avoid SettingWithCopyWarning
        target_payment_plan = ["Fortnightly-Fixed", "Upfront"]
        end_date_dt = pd.to_datetime(end_date, format="%Y-%m-%d")
        df_copy["End date"] = pd.to_datetime(
            df_copy["End date"],
            format="%Y-%m-%d",
            errors="coerce",  # Adapt format if needed
        )

        df_ff_filtered = df_copy.loc[
            (df_copy["Club"] == target_club)
            & (df_copy["Payment plan type"] == target_payment_plan[0])
            & ((df_copy["End date"].isnull()) | (df_copy["End date"] > end_date_dt))
        ]
        fortnightly_fixed_members = len(df_ff_filtered)

        df_total_filtered = df_copy.loc[
            (df_copy["Club"] == target_club)
            & (df_copy["Payment plan type"].isin(target_payment_plan))
            & ((df_copy["End date"].isnull()) | (df_copy["End date"] > end_date_dt))
        ]
        total_members = len(df_total_filtered)

        html_output = f"""
        <table width='95%' style='font-family: Monospace; border-collapse: collapse; margin-top: 10px;'>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>FORTNIGHTLY-FIXED MEMBERS:</b></td>
                <td align='right' style='padding: 5px 0;'>{fortnightly_fixed_members}</td>
            </tr>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>TOTAL MEMBERS:</b></td>
                <td align='right' style='padding: 5px 0;'>{total_members}</td>
            </tr>
        </table>
        """
        return html_output
    except KeyError as e:
        return f"<p><b><font color='red'>Error (Current Members):</font></b><br>Missing expected column: {e}</p>"
    except ValueError as ve:
        return f"<p><b><font color='red'>Error (Current Members):</font></b><br>Data error: {str(ve)}</p>"
    except Exception as e:
        return f"<p><b><font color='red'>An error occurred in Current Members report:</font></b><br>{str(e)}</p>"


def new_members(df: pd.DataFrame, target_club: str, end_date: str) -> str:
    """
    Counts new members for a target club within the month of the given end_date.
    Returns an HTML formatted table string.
    """
    try:
        required_cols = ["Club", "Payment plan type", "End date", "Join date"]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            raise ValueError(
                f"New Members: Missing required columns: {', '.join(missing)}"
            )

        df_copy = df.copy()
        target_payment_plan = ["Fortnightly-Fixed", "Upfront"]
        end_date_dt = pd.to_datetime(end_date, format="%Y-%m-%d")
        month = end_date_dt.month
        year = end_date_dt.year
        start_date_month = pd.to_datetime(f"{year}-{month:02d}-01", format="%Y-%m-%d")

        df_copy["End date"] = pd.to_datetime(
            df_copy["End date"], format="%Y-%m-%d", errors="coerce"
        )
        df_copy["Join date"] = pd.to_datetime(
            df_copy["Join date"], errors="coerce"
        )  # Flexible parsing

        df_ff_new = df_copy.loc[
            (df_copy["Club"] == target_club)
            & (df_copy["Payment plan type"] == target_payment_plan[0])
            & ((df_copy["End date"].isnull()) | (df_copy["End date"] > end_date_dt))
            & (df_copy["Join date"].notnull())
            & (df_copy["Join date"] >= start_date_month)
            & (df_copy["Join date"] <= end_date_dt)
        ]
        new_fortnightly_fixed_members = len(df_ff_new)

        df_total_new = df_copy.loc[
            (df_copy["Club"] == target_club)
            & (df_copy["Payment plan type"].isin(target_payment_plan))
            & ((df_copy["End date"].isnull()) | (df_copy["End date"] > end_date_dt))
            & (df_copy["Join date"].notnull())
            & (df_copy["Join date"] >= start_date_month)
            & (df_copy["Join date"] <= end_date_dt)
        ]
        new_total_members = len(df_total_new)

        html_output = f"""
        <table width='95%' style='font-family: Monospace; border-collapse: collapse; margin-top: 10px;'>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>NEW FORTNIGHTLY-FIXED MEMBERS:</b></td>
                <td align='right' style='padding: 5px 0;'>{new_fortnightly_fixed_members}</td>
            </tr>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>TOTAL NEW MEMBERS:</b></td>
                <td align='right' style='padding: 5px 0;'>{new_total_members}</td>
            </tr>
            <tr>
                <td colspan='2' style='padding-top: 8px; font-size: smaller;'><i>(For month of {end_date_dt.strftime("%B %Y")}, from {start_date_month.strftime("%d-%m-%Y")} to {end_date_dt.strftime("%d-%m-%Y")})</i></td>
            </tr>
        </table>
        """
        return html_output
    except KeyError as e:
        return f"<p><b><font color='red'>Error (New Members):</font></b><br>Missing expected column: {e}</p>"
    except ValueError as ve:
        return f"<p><b><font color='red'>Error (New Members):</font></b><br>Data error: {str(ve)}</p>"
    except Exception as e:
        return f"<p><b><font color='red'>An error occurred in New Members report:</font></b><br>{str(e)}</p>"


def technogym_reporting(df: pd.DataFrame) -> str:
    """
    Calculates number of PT and health consult sessions using pandas.
    Returns an HTML formatted table string.
    """
    try:
        if "Activity" not in df.columns:
            raise ValueError(
                "DataFrame missing 'Activity' column for Technogym report."
            )

        consults_list = [
            "Body Scan",
            "Exercise Program Check-in",
            "Follow-Up Health Consultation",
            "Follow-Up Health Consultation and Program Update",
            "Initial Health Consultation",
            "Initial Program Introduction",
        ]
        pt_list = [
            "Group Training",
            "Personal Training 30 Minutes",
            "Personal Training 45 Minutes",
            "Personal Training 60 Minutes",
        ]

        df_consults = df.loc[df["Activity"].isin(consults_list)]
        df_pts = df.loc[df["Activity"].isin(pt_list)]

        consults_no = len(df_consults)
        pts_no = len(df_pts)

        html_output = f"""
        <table width='95%' style='font-family: Monospace; border-collapse: collapse; margin-top: 10px;'>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>NUMBER OF HEALTH CONSULTS:</b></td>
                <td align='right' style='padding: 5px 0;'>{consults_no}</td>
            </tr>
            <tr>
                <td style='padding: 5px 10px 5px 0;'><b>NUMBER OF PERSONAL TRAINING SESSIONS:</b></td>
                <td align='right' style='padding: 5px 0;'>{pts_no}</td>
            </tr>
        </table>
        """
        return html_output
    except KeyError as e:
        return f"<p><b><font color='red'>Error (Technogym):</font></b><br>Missing expected column 'Activity': {e}</p>"
    except ValueError as ve:
        return f"<p><b><font color='red'>Error (Technogym):</font></b><br>Data error: {str(ve)}</p>"
    except Exception as e:
        return f"<p><b><font color='red'>An error occurred in Technogym report:</font></b><br>{str(e)}</p>"


def groupFitness(df_pandas_input: pd.DataFrame) -> str:
    """
    Generates an HTML table string for Group Fitness Summary using pandas.
    Input DataFrame should have headers from the second row of the Excel.
    """
    try:
        required_cols = ["Club", "UserActive"]
        if not all(col in df_pandas_input.columns for col in required_cols):
            missing = [
                col for col in required_cols if col not in df_pandas_input.columns
            ]
            raise ValueError(
                f"Group Fitness: Missing required columns: {', '.join(missing)}"
            )

        df_filtered = df_pandas_input[required_cols].copy()

        clubs_to_report = [
            "DeakinACTIVE Burwood",
            "DeakinACTIVE Waterfront",
            "DeakinACTIVE Waurn Ponds",
            "DeakinACTIVE Warrnambool",
        ]

        html_output_lines = [
            "<table width='95%' style='font-family: Monospace; border-collapse: collapse; margin-top: 10px;'>",
            "<tr>",
            "<td style='padding: 5px 10px 5px 0;'><b>Club Name</b></td>",
            "<td align='right' style='padding-right: 20px; padding: 5px 0;'><b>Classes Run</b></td>",
            "<td align='right' style='padding: 5px 0;'><b>Total Attendees</b></td>",
            "</tr>",
            "<tr><td colspan='3' style='line-height: 0.5em;'><hr></td></tr>",
        ]

        for club_name in clubs_to_report:
            club_df = df_filtered[df_filtered["Club"] == club_name]
            # User's logic for duplicate rows
            num_classes = len(club_df) / 2
            total_attendees = (
                pd.to_numeric(club_df["UserActive"], errors="coerce").fillna(0).sum()
                / 2
            )

            html_output_lines.append("<tr>")
            html_output_lines.append(
                f"<td style='padding: 5px 10px 5px 0;'>{club_name}</td>"
            )
            html_output_lines.append(
                f"<td align='right' style='padding-right: 20px; padding-top: 5px; padding-bottom: 5px;'>{int(num_classes)}</td>"
            )
            html_output_lines.append(
                f"<td align='right' style='padding-top: 5px; padding-bottom: 5px;'>{int(total_attendees)}</td>"
            )
            html_output_lines.append("</tr>")

        html_output_lines.append("</table>")
        return "".join(html_output_lines)
    except KeyError as e:
        return f"<p><b><font color='red'>Error (Group Fitness):</font></b><br>Missing column: {e}</p>"
    except ValueError as ve:
        return f"<p><b><font color='red'>Error (Group Fitness):</font></b><br>Data error: {str(ve)}</p>"
    except Exception as e:
        return f"<p><b><font color='red'>An error occurred in Group Fitness report:</font></b><br>{str(e)}</p>"


def booking_zones(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processes booking data using pandas, applies weights, and returns a summary DataFrame.
    """
    try:
        weights_dict = {
            "BUR - Badminton Court": 1 / 6,
            "WP - Badminton Court": 1 / 6,
            "BUR - Court": 1 / 2,
            "WP - Court": 1 / 2,
            "WP - Athletic Track Lane": 1 / 4,
        }
        required_cols = [
            "Facility Booking Definition",
            "Club",
            "Club Zone Type Name",
            "Length of Booking",
        ]
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            raise ValueError(
                f"Booking Zones: Missing required columns: {', '.join(missing_cols)}"
            )

        filtered_df = df.loc[
            (
                df["Facility Booking Definition"].str.contains(
                    "Unavailable", case=False, na=False
                )
                == False
            )
            & (
                df["Facility Booking Definition"].str.contains(
                    "University Class", case=False, na=False
                )
                == False
            ),
            required_cols,
        ].copy()

        try:
            filtered_df["Length of Booking"] = pd.to_timedelta(
                filtered_df["Length of Booking"]
            )
        except (ValueError, TypeError):
            if pd.api.types.is_numeric_dtype(filtered_df["Length of Booking"]):
                filtered_df["Length of Booking"] = pd.to_timedelta(
                    filtered_df["Length of Booking"], unit="m"
                )
            else:
                filtered_df["Length of Booking"] = pd.NaT

        df_sum = filtered_df.groupby(["Club", "Club Zone Type Name"], as_index=False)[
            "Length of Booking"
        ].sum()

        df_sum["Adjusted Time"] = df_sum.apply(
            lambda row: row["Length of Booking"]
            * weights_dict.get(row["Club Zone Type Name"], 1.0),
            axis=1,
        )

        df_sum["Length of Booking (Hours)"] = (
            df_sum["Length of Booking"].dt.total_seconds().fillna(0) / 3600
        )
        df_sum["Adjusted Time (Hours)"] = (
            df_sum["Adjusted Time"].dt.total_seconds().fillna(0) / 3600
        )

        df_final = df_sum[
            [
                "Club",
                "Club Zone Type Name",
                "Length of Booking (Hours)",
                "Adjusted Time (Hours)",
            ]
        ].copy()
        return df_final
    except KeyError as e:
        raise KeyError(f"Booking Zones: Missing column: {e}")
    except ValueError as ve:
        raise ValueError(f"Booking Zones: Data error: {ve}")
    except Exception as e:
        raise Exception(f"Booking Zones: An unexpected error occurred: {e}")


# --- NEW PANDAS-BASED FUNCTION ---
def generate_ending_members_report(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Filters members whose contract 'End date' is today using pandas and returns
    a pandas DataFrame containing their details.

    Args:
        df_input: A pandas DataFrame containing member data. Expected columns include:
                  "Name", "Last name", "Club", "Payment Plan Name", "End date",
                  "Email", "Mobile number".

    Returns:
        pd.DataFrame: A pandas DataFrame of members whose contracts end today.
                      May be empty if no such members are found.

    Raises:
        ValueError: If required columns are missing or 'End date' conversion fails.
        KeyError: If essential columns are not found during pandas operations.
    """
    today_date_obj = dt.date.today()  # Get today as a date object

    required_columns = [
        "Name",
        "Last name",
        "Club",
        "Payment Plan Name",
        "End date",
        "Email",
        "Mobile number",
    ]

    missing_cols = [col for col in required_columns if col not in df_input.columns]
    if missing_cols:
        raise ValueError(
            f"Ending Members Report: Missing required columns: {', '.join(missing_cols)}"
        )

    df_processed = df_input[required_columns].copy()

    try:
        # Convert 'End date' to datetime objects, coercing errors to NaT
        df_processed["End date"] = pd.to_datetime(
            df_processed["End date"], errors="coerce"
        )
    except Exception as e:
        raise ValueError(
            f"Ending Members Report: Error converting 'End date' column to datetime: {str(e)}"
        )

    # Drop rows where 'End date' could not be parsed (is NaT)
    df_processed_valid_dates = df_processed.dropna(subset=["End date"])

    # Filter: Compare the date part of 'End date' (which is datetime64[ns]) with today's date object
    df_filtered = df_processed_valid_dates[
        df_processed_valid_dates["End date"].dt.date == today_date_obj
    ]

    return df_filtered

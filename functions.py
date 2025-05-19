import pandas as pd
import datetime as dt

# --- Report Generation Functions ---


def current_members(df: pd.DataFrame, target_club: str, end_date: str) -> str:
    """
    Counts current members for a target club and end date.
    Returns an HTML formatted table string.
    """
    try:
        target_payment_plan = ["Fortnightly-Fixed", "Upfront"]
        end_date_dt = pd.to_datetime(end_date, format="%Y-%m-%d")
        df["End date"] = pd.to_datetime(
            df["End date"], format="%Y-%m-%d", errors="coerce"
        )

        df_ff_filtered = df.loc[
            (df["Club"] == target_club)
            & (df["Payment plan type"] == target_payment_plan[0])
            & ((df["End date"].isnull()) | (df["End date"] > end_date_dt))
        ]
        fortnightly_fixed_members = len(df_ff_filtered)

        df_total_filtered = df.loc[
            (df["Club"] == target_club)
            & (df["Payment plan type"].isin(target_payment_plan))
            & ((df["End date"].isnull()) | (df["End date"] > end_date_dt))
        ]
        total_members = len(df_total_filtered)

        # HTML Table Output
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
        return f"<font color='red'>Error: Missing expected column in data for Current Members: {e}</font>"
    except Exception as e:
        return f"<font color='red'>An error occurred in Current Members report: {str(e)}</font>"


def new_members(df: pd.DataFrame, target_club: str, end_date: str) -> str:
    """
    Counts new members for a target club within the month of the given end_date.
    Returns an HTML formatted table string.
    """
    try:
        target_payment_plan = ["Fortnightly-Fixed", "Upfront"]
        end_date_dt = pd.to_datetime(end_date, format="%Y-%m-%d")
        month = end_date_dt.month
        year = end_date_dt.year
        start_date_dt = pd.to_datetime(f"{year}-{month}-01", format="%Y-%m-%d")

        df["End date"] = pd.to_datetime(
            df["End date"], format="%Y-%m-%d", errors="coerce"
        )
        df["Join date"] = pd.to_datetime(
            df["Join date"], format="%Y-%m-%d", errors="coerce"
        )

        df_ff_new = df.loc[
            (df["Club"] == target_club)
            & (df["Payment plan type"] == target_payment_plan[0])
            & ((df["End date"].isnull()) | (df["End date"] > end_date_dt))
            & (df["Join date"] >= start_date_dt)
            & (df["Join date"] <= end_date_dt)
        ]
        new_fortnightly_fixed_members = len(df_ff_new)

        df_total_new = df.loc[
            (df["Club"] == target_club)
            & (df["Payment plan type"].isin(target_payment_plan))
            & ((df["End date"].isnull()) | (df["End date"] > end_date_dt))
            & (df["Join date"] >= start_date_dt)
            & (df["Join date"] <= end_date_dt)
        ]
        new_total_members = len(df_total_new)

        # HTML Table Output
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
                <td colspan='2' style='padding-top: 8px; font-size: smaller;'><i>(For month starting {start_date_dt.strftime("%d-%m-%Y")} up to {end_date_dt.strftime("%d-%m-%Y")})</i></td>
            </tr>
        </table>
        """
        return html_output
    except KeyError as e:
        return f"<font color='red'>Error: Missing expected column in data for New Members: {e}</font>"
    except Exception as e:
        return f"<font color='red'>An error occurred in New Members report: {str(e)}</font>"


def technogym_reporting(df: pd.DataFrame) -> str:
    """
    Calculates number of PT and health consult sessions.
    Returns an HTML formatted table string.
    """
    try:
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

        # Ensure 'Activity' column exists
        if "Activity" not in df.columns:
            return "<font color='red'>Error: DataFrame missing 'Activity' column for Technogym report.</font>"

        df_consults = df.loc[df["Activity"].isin(consults_list)]
        df_pts = df.loc[df["Activity"].isin(pt_list)]

        consults_no = len(df_consults)
        pts_no = len(df_pts)

        # HTML Table Output
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
    except KeyError as e:  # Should be caught by the check above, but as a fallback
        return f"<font color='red'>Error: Missing expected column 'Activity' in data: {e}</font>"
    except Exception as e:
        return (
            f"<font color='red'>An error occurred in Technogym report: {str(e)}</font>"
        )


def groupFitness(df: pd.DataFrame) -> str:
    """
    Generates an HTML table string for Group Fitness Summary.
    (This function was already updated in the previous response and is good)
    """
    try:
        if not {"Club", "UserActive"}.issubset(df.columns):
            return "<font color='red'>Error: DataFrame missing 'Club' or 'UserActive' column.</font>"

        df_filtered = df[["Club", "UserActive"]].copy()

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
            "<td align='right' style='padding-right: 20px; padding: 5px 0;'><b>Classes Run</b></td>",  # Added padding to header
            "<td align='right' style='padding: 5px 0;'><b>Total Attendees</b></td>",  # Added padding to header
            "</tr>",
            "<tr><td colspan='3' style='line-height: 0.5em;'><hr></td></tr>",
        ]

        for club_name in clubs_to_report:
            club_df = df_filtered[df_filtered["Club"] == club_name]
            num_classes = (
                len(club_df) / 2
            )  # realised that there are two tables in the sheet, resulting in dupicate rows
            total_attendees = (
                pd.to_numeric(club_df["UserActive"], errors="coerce").sum()
                / 2  # Handling dupllicate rows
                if not club_df.empty
                else 0
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
    except Exception as e:
        return f"<font color='red'>An error occurred in Group Fitness report: {str(e)}</font>"


def booking_zones(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processes booking data, applies weights, and returns a summary DataFrame.
    (This function returns a DataFrame for CSV export, formatting is handled by main.py for display if needed, but current goal is CSV)
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
                f"Booking Zones report: Missing required columns: {', '.join(missing_cols)}"
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
                # Consider logging a warning here about unparseable "Length of Booking" values

        df_sum = (
            filtered_df.groupby(["Club", "Club Zone Type Name"])["Length of Booking"]
            .sum()
            .reset_index()
        )

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
        raise KeyError(
            f"Missing expected column for Booking Zones report: {e}. Please check your Excel file."
        )
    except ValueError as ve:  # Catch potential ValueError from the column check
        raise ve
    except Exception as e:
        raise Exception(f"An error occurred in Booking Zones processing: {str(e)}")

import pandas as pd


# Function to count current members
def current_members(df: pd.DataFrame, target_club: str, end_date: str) -> int:
    target_payment_plan = ["Fortnightly Fixed", "Upfront"]
    df_filtered = df.loc[
        (df["Club"] == target_club)
        & (df["Payment plan type"] == target_payment_plan[0])
        & (
            (df["End date"].isnull())
            | (
                pd.to_datetime(df["End date"], format="%Y-%m-%d")
                > pd.to_datetime(end_date, format="%Y-%m-%d")
            )
        )
    ]

    fortnightly_fixed_members = len(df_filtered)

    df_filtered = df.loc[
        (df["Club"] == target_club)
        & (df["Payment plan type"].isin(target_payment_plan))
        & (
            (df["End date"].isnull())
            | (
                pd.to_datetime(df["End date"], format="%Y-%m-%d")
                > pd.to_datetime(end_date, format="%Y-%m-%d")
            )
        )
    ]

    total_members = len(df_filtered)

    return [fortnightly_fixed_members, total_members]


# Function to count new members in a month
def new_members(df: pd.DataFrame, target_club: str, end_date: str) -> int:
    target_payment_plan = ["Fortnightly Fixed", "Upfront"]
    end_date = pd.to_datetime(end_date, format="%Y-%m-%d")
    month = end_date.month
    year = end_date.year
    start_date = pd.to_datetime(f"{year}-{month}-01", format="%Y-%m-%d")

    df_filtered = df.loc[
        (df["Club"] == target_club)
        & (df["Payment plan type"] == target_payment_plan[0])
        & (
            (df["End date"].isnull())
            | (pd.to_datetime(df["End date"], format="%Y-%m-%d") > end_date)
        )
        & (pd.to_datetime(df["Join date"], format="%Y-%m-%d") >= start_date)
    ]

    new_fortnightly_fixed_members = len(df_filtered)

    df_filtered = df.loc[
        (df["Club"] == target_club)
        & (df["Payment plan type"].isin(target_payment_plan))
        & (
            (df["End date"].isnull())
            | (pd.to_datetime(df["End date"], format="%Y-%m-%d") > end_date)
        )
        & (pd.to_datetime(df["Join date"], format="%Y-%m-%d") >= start_date)
    ]

    new_total_members = len(df_filtered)

    return [new_fortnightly_fixed_members, new_total_members]


# Function to calculate number of PT and health consult sessions in a month
def technogym_reporting(df: pd.DataFrame) -> list:
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

    return [consults_no, pts_no]


def groupFitness(df: pd.DataFrame) -> dict:
    df = df[["Club", "UserActive"]]

    df_burwood = df.loc[df["Club"] == "DeakinACTIVE Burwood"]
    df_waterfront = df.loc[df["Club"] == "DeakinACTIVE Waterfront"]
    df_wp = df.loc[df["Club"] == "DeakinACTIVE Waurn Ponds"]
    df_warrnambool = df.loc[df["Club"] == "DeakinACTIVE Warrnambool"]

    no_burwood = len(df_burwood)
    no_waterfront = len(df_waterfront)
    no_wp = len(df_wp)
    no_warrnambool = len(df_warrnambool)

    burwood_attendees = df_burwood["UserActive"].sum()
    waterfront_attendees = df_waterfront["UserActive"].sum()
    wp_attendees = df_wp["UserActive"].sum()
    warrnambool_attendees = df_warrnambool["UserActive"].sum()

    results_dict = {
        "Burwood": [no_burwood, burwood_attendees],
        "Waterfront": [no_waterfront, waterfront_attendees],
        "Waurn Ponds": [no_wp, wp_attendees],
        "Warrnambool": [no_warrnambool, warrnambool_attendees],
    }

    return results_dict


def booking_zones(df: pd.DataFrame) -> pd.DataFrame:
    weights_dict = {
        "BUR - Badminton Court": 0.16,
        "WP - Badminton Court": 0.16,
        "BUR - Court": 0.5,
        "WP - Court": 0.5,
        "WP - Athletic Track Lane": 0.25,
    }

    # --- Corrected part to avoid SettingWithCopyWarning ---
    # Perform filtering and column selection in one step using .loc
    # This ensures the subsequent operations are on a known DataFrame slice/copy
    filtered_df = df.loc[
        (df["Facility Booking Definition"].str[0:11] != "Unavailable")
        & (df["Facility Booking Definition"] != "University Class"),
        [
            "Club",
            "Club Zone Type Name",
            "Length of Booking",
            "Facility Booking Definition",
        ],
    ].copy()  # Use .copy() to explicitly work on a copy and prevent future warnings

    # Now perform the timedelta conversion on this filtered and selected DataFrame
    # The warning should not appear here
    filtered_df["Length of Booking"] = pd.to_timedelta(filtered_df["Length of Booking"])
    # --- End of corrected part ---

    # Group and sum as before, using the filtered_df
    df_sum = (
        filtered_df.groupby(["Club", "Club Zone Type Name"])["Length of Booking"]
        .sum()
        .reset_index()
    )

    # Apply the weighting logic
    df_sum["Adjusted Time"] = df_sum.apply(
        lambda row: row["Length of Booking"] * weights_dict[row["Club Zone Type Name"]]
        if row["Club Zone Type Name"] in weights_dict
        else row["Length of Booking"],
        axis=1,
    )

    # Convert Timedelta columns to hours as floats
    df_sum["Length of Booking"] = df_sum["Length of Booking"].dt.total_seconds() / 3600
    df_sum["Adjusted Time"] = df_sum["Adjusted Time"].dt.total_seconds() / 3600

    return df_sum

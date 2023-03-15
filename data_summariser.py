import pandas as pd
import re
import csv
import os
import numpy as np

# --- User input ---

date_format = "%d/%m/%Y"

# ------------------


def obtaining_person_identity(manifest_location):
    """Going through the manifest and then obtain the relevant information
    Args:
        manifest_location (str): The absolute filepath for the manifest.
    Returns:
        Dict: The subject code of the person together with the corresponding student's identity.
    """
    df = pd.read_excel(manifest_location)
    df = df[
        ["Name", "ACT Subject Code", "AY", "Trimester (1/2/3)", "Arm (LTLB/ Control)"]
    ]
    df = df.loc[(df["ACT Subject Code"] != "N") & (df["ACT Subject Code"].notnull())]

    return df


def obtaining_dataframe(filenames):
    """Obtain the summary statistics for the actigraphy
    input: CSV
    output: Pandas dataframe containing all stats for REST interval
    """
    # --- this section tries to isolate the "Statistics" section of the CSV file ---
    searching_statistics_header = re.compile(r"-+\s?Statistics\s?-+")
    searching_marker_header = re.compile(r"-+\s?Marker/Score List\s?-+")
    row_lower_limit = 1
    row_upper_limit = 1
    if len(filenames) == 0:
        print("ERROR: You need to make sure that you provide at least one file!")
        return
    else:
        data_frames = []
        for filename in filenames:
            with open(filename) as csv_file:

                csv_reader = csv.reader(csv_file)
                for item in csv_reader:
                    if (
                        not re.search(searching_statistics_header, str(item))
                        and row_upper_limit <= row_lower_limit
                    ):
                        row_lower_limit += 1
                        row_upper_limit += 1
                    elif not re.search(searching_marker_header, str(item)):
                        row_upper_limit += 1
                    elif re.search(searching_marker_header, str(item)):
                        break
            df = pd.read_csv(
                filename,
                skiprows=row_lower_limit + 1,
            )
            pd.set_option("display.max_columns", None)
            df.drop(list(df)[18:], axis=1, inplace=True)
            data_frames.append(df)
        final_df = pd.concat(data_frames, ignore_index=True)

        return final_df


def obtaining_rest_dataframe(filenames):
    """Obtaining the data for Rest Interval
    input: CSV file
    Output: pandas dataframe"""

    rest_df = obtaining_dataframe(filenames)
    rest_df = rest_df.loc[rest_df["Interval Type"] == "REST"]
    rest_df.drop(list(rest_df)[7:], axis=1, inplace=True)
    rest_df.drop(columns=[r"Interval#", "Interval Type"], axis=1, inplace=True)

    rest_df.rename(
        columns={
            "Start Time": "Bed Time",
            "End Time": "Get Up Time",
            "Duration": "Total Time in Bed (hours)",
        },
        inplace=True,
    )
    rest_df = rest_df[
        [
            "Start Date",
            "Bed Time",
            "Get Up Time",
            "Total Time in Bed (hours)",
            "End Date",
        ]
    ]

    return rest_df


def obtaining_sleep_dataframe(filenames):
    """Obtaining the data for the Sleep Interval
    input: CSV file
    Return: pandas dataframe containing full stats for sleep interval"""

    sleep_df = obtaining_dataframe(filenames)
    sleep_df = sleep_df.loc[sleep_df["Interval Type"] == "SLEEP"]

    sleep_df = sleep_df.loc[:, "Onset Latency":"%Sleep"]
    sleep_df.drop(columns=["Wake Time", "%Wake", "%Sleep"], axis=1, inplace=True)

    sleep_df.rename(
        columns={
            r"#Wake Bouts": r"#Awake",
            "Efficiency": "Sleep Efficiency (percent)",
            "Onset Latency": "Onset Latency (minutes)",
            "Sleep Time": "Total Sleep Time (hours)",
            "WASO": "WASO (minutes)",
        },
        inplace=True,
    )

    # --- Rearranging the columns in this dataframe ---
    sleep_df = sleep_df[
        [
            "Total Sleep Time (hours)",
            "Onset Latency (minutes)",
            "Sleep Efficiency (percent)",
            "WASO (minutes)",
            r"#Awake",
        ]
    ]

    return sleep_df


def combined_stats(filenames):
    """Combine the dataframe from the preceding functions and try to get the summarised statistics
    Input: CSV file names in list format
    Return: pandas dataframe containing the summarised data"""

    resting_df = obtaining_rest_dataframe(filenames)
    sleeping_df = obtaining_sleep_dataframe(filenames)
    resting_df.reset_index(drop=True, inplace=True)
    sleeping_df.reset_index(drop=True, inplace=True)
    # Combining the resting_df and sleeping_df into a single dataframe

    new_df = pd.concat([resting_df, sleeping_df], axis=1)
    # The list contains the variables which should be changed to numeric format rather than string format.
    column_to_change = [
        "Onset Latency (minutes)",
        "Sleep Efficiency (percent)",
        "WASO (minutes)",
    ]
    # Converting the relevant column to datetime objects.

    new_df["Get Up Time"] = pd.to_datetime(new_df["Get Up Time"])
    # For the Total Sleep Time column, if there is an NA cell, refer to the corresponding cell in the Bed Time column. If that cell in the Bed Time column is not null, then fill the NA cell with 0, else leave it as it is.
    new_df["Total Sleep Time (hours)"] = np.where(
        (new_df["Bed Time"].notnull()) & (new_df["Total Sleep Time (hours)"].isna()),
        0,
        new_df["Total Sleep Time (hours)"],
    )
    new_df["Total Time in Bed (hours)"] = pd.to_datetime(
        new_df["Total Time in Bed (hours)"], unit="m"
    )
    new_df["Total Sleep Time (hours)"] = pd.to_datetime(
        new_df["Total Sleep Time (hours)"], unit="m"
    )
    # Changing the pm to am and vice versa in the Bed Time column so that Python will be able to order the Bed Time correctly
    new_df["Bed Time"] = new_df["Bed Time"].str.upper()
    new_df["Bed Time"] = new_df["Bed Time"].str.replace("PM", "am")
    new_df["Bed Time"] = new_df["Bed Time"].str.replace("AM", "pm")

    new_df["Bed Time"] = pd.to_datetime(new_df["Bed Time"])
    # Change type of data in the column_to_change from string to numeric. Numeric includes float.
    new_df[column_to_change] = new_df[column_to_change].apply(pd.to_numeric)
    # If there is an NA cell in Onset Latency, check the corresponding cell in the Bed Time column. If that cell is not NA, then replace NA with 0 in the Onset Latency cell.
    new_df["Onset Latency (minutes)"] = np.where(
        (new_df["Bed Time"].notna()) & (new_df["Onset Latency (minutes)"].isna()),
        0,
        new_df["Onset Latency (minutes)"],
    )

    new_df["Onset Latency (minutes)"] = np.where(
        (new_df["Bed Time"].notna()) & (new_df["Onset Latency (minutes)"].isna()),
        0,
        new_df["Onset Latency (minutes)"],
    )
    new_df["Sleep Efficiency (percent)"] = np.where(
        (new_df["Bed Time"].notna()) & (new_df["Sleep Efficiency (percent)"].isna()),
        0.0,
        new_df["Sleep Efficiency (percent)"],
    )
    new_df["WASO (minutes)"] = np.where(
        (new_df["Bed Time"].notna()) & (new_df["WASO (minutes)"].isna()),
        0,
        new_df["WASO (minutes)"],
    )
    new_df["#Awake"] = np.where(
        (new_df["Bed Time"].notnull()) & (new_df["#Awake"].isna()),
        0,
        new_df["#Awake"],
    )
    # The describe function gathers the summarised stats of the dataframe. This includes the mean, median, and their percentile. Datetime_is_numeric enables us to obtain these stats for datetime objects.
    summary_stats = new_df.describe(datetime_is_numeric=True)
    # Changing the data format to string.
    summary_stats["Bed Time"] = pd.to_datetime(
        summary_stats["Bed Time"], errors="coerce"
    ).dt.strftime("%I:%M:%S %p")
    # Replacing the AM back to PM and vice versa so that the summarised stats make more sense
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("AM", "pm")
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("PM", "am")
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.capitalize()

    summary_stats["Get Up Time"] = pd.to_datetime(
        summary_stats["Get Up Time"], errors="coerce"
    ).dt.strftime("%I:%M:%S %p")

    summary_stats["Total Time in Bed (hours)"] = pd.to_datetime(
        summary_stats["Total Time in Bed (hours)"], errors="coerce"
    ).dt.strftime("%H:%M:%S")
    summary_stats["Total Sleep Time (hours)"] = pd.to_datetime(
        summary_stats["Total Sleep Time (hours)"], errors="coerce"
    ).dt.strftime("%H:%M:%S")

    summary_stats.iloc[:, 3:] = summary_stats.iloc[:, 3:].round(2)

    return summary_stats

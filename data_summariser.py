import pandas as pd
import re
import csv
import os

# --- User input ---
# input_directory = (
#     "Actiware output/CSV output"  # Which folder contains the raw csv input?
# )
# output_directory = (
#     "formatted_csv"  # Which folder do you want to export the formatted csv to?
# )
date_format = "%d/%m/%Y"

# ------------------
# os.chdir(input_directory)


def obtaining_person_identity():
    """Obtain the Subject Code and the corresponding students name
    Return: dictionary containing Subject Code and corresponding Name.
    """
    subject_name_and_code = {}

    for i in os.listdir():
        if i.endswith("pdf"):
            relevant_info = i.split(" Actigraphy")
            search_term = r"SIT\d{3}"
            subject_name = relevant_info[0]
            subject_code = re.findall(search_term, relevant_info[-1])[0]
            subject_name_and_code[subject_code] = subject_name

    return subject_name_and_code


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
    sleep_df.drop(list(sleep_df)[7:10], axis=1, inplace=True)

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

    Input: CSV file

    Return: pandas dataframe containing the summarised data"""
    print(filenames)

    resting_df = obtaining_rest_dataframe(filenames)
    sleeping_df = obtaining_sleep_dataframe(filenames)
    resting_df.reset_index(drop=True, inplace=True)
    sleeping_df.reset_index(drop=True, inplace=True)

    new_df = pd.concat([resting_df, sleeping_df], axis=1)

    column_to_change = [
        "Onset Latency (minutes)",
        "Sleep Efficiency (percent)",
        "WASO (minutes)",
    ]

    new_df["Get Up Time"] = pd.to_datetime(new_df["Get Up Time"])
    new_df["Total Time in Bed (hours)"] = pd.to_datetime(
        new_df["Total Time in Bed (hours)"], unit="m"
    )
    new_df["Total Sleep Time (hours)"] = pd.to_datetime(
        new_df["Total Sleep Time (hours)"], unit="m"
    )
    new_df["Bed Time"] = new_df["Bed Time"].str.replace("pm", "AM")
    new_df["Bed Time"] = new_df["Bed Time"].str.replace("am", "PM")

    new_df["Bed Time"] = pd.to_datetime(new_df["Bed Time"])

    new_df[column_to_change] = new_df[column_to_change].apply(pd.to_numeric)

    summary_stats = new_df.describe(datetime_is_numeric=True)

    summary_stats["Bed Time"] = pd.to_datetime(
        summary_stats["Bed Time"], errors="coerce"
    ).dt.strftime("%I:%M:%S %p")

    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("AM", "pm")
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("PM", "am")
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("am", "AM")
    summary_stats["Bed Time"] = summary_stats["Bed Time"].str.replace("pm", "PM")

    summary_stats["Get Up Time"] = pd.to_datetime(
        summary_stats["Get Up Time"], errors="coerce"
    ).dt.strftime("%I:%M:%S %p")

    summary_stats["Total Time in Bed (hours)"] = pd.to_datetime(
        summary_stats["Total Time in Bed (hours)"], errors="coerce"
    ).dt.strftime("%I:%M:%S")
    summary_stats["Total Sleep Time (hours)"] = pd.to_datetime(
        summary_stats["Total Sleep Time (hours)"], errors="coerce"
    ).dt.strftime("%I:%M:%S")

    summary_stats.iloc[:, 3:] = summary_stats.iloc[:, 3:].round(2)

    return summary_stats

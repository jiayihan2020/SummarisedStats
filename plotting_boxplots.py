import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import os


working_directory = Path("Combining the median")
os.chdir(working_directory)

df_control_group = pd.read_excel("List of Median Values for Control group.xlsx")
df_ltlb_group = pd.read_excel("List of Median Values for LTLB group.xlsx")


# --- Dealing with the difficult case of handling datetime objects ---


df_control_timing = df_control_group.iloc[:, 0:4]
df_ltlb_timing = df_ltlb_group.iloc[:, 0:4]


cols = df_control_timing.columns[2:]
df_control_timing[cols] = df_control_timing[cols].apply(pd.to_timedelta)

df_control_timing = df_control_timing.assign(
    Total_Time_in_Bed_hours=[
        y.seconds / (60.0 * 60.0)
        for y in df_control_timing["Total Time in Bed (hours)"]
    ]
)

df_control_timing = df_control_timing.assign(
    Total_Sleep_Time_hours=[
        (x.seconds) / (60.0 * 60.0)
        for x in df_control_timing["Total Sleep Time (hours)"]
    ]
)


df_control_timing["Bed Time"] = df_control_timing["Bed Time"].str.replace("AM", "pm")

df_control_timing["Bed Time"] = df_control_timing["Bed Time"].str.replace("PM", "am")

cols_to_change_to_dt = df_control_timing.columns[:2]

df_control_timing[cols_to_change_to_dt] = df_control_timing[cols_to_change_to_dt].apply(
    pd.to_datetime
)

df_control_timing["Bed Time (unix)"] = (
    df_control_timing["Bed Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")


df_control_timing["Get Up Time (unix)"] = (
    df_control_timing["Get Up Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")


df_control_timing = df_control_timing.drop(
    columns=[
        "Bed Time",
        "Get Up Time",
        "Total Time in Bed (hours)",
        "Total Sleep Time (hours)",
    ]
)


df_ltlb_timing[cols] = df_ltlb_timing[cols].apply(pd.to_timedelta)

df_ltlb_timing["Bed Time"] = df_ltlb_timing["Bed Time"].str.replace("AM", "pm")

df_ltlb_timing["Bed Time"] = df_ltlb_timing["Bed Time"].str.replace("PM", "am")

df_ltlb_timing = df_ltlb_timing.assign(
    Total_Time_in_Bed_hours=[
        a.seconds / (60.0 * 60.0) for a in df_ltlb_timing["Total Time in Bed (hours)"]
    ]
)

df_ltlb_timing = df_ltlb_timing.assign(
    Total_Sleep_Time_hours=[
        (b.seconds) / (60.0 * 60.0) for b in df_ltlb_timing["Total Sleep Time (hours)"]
    ]
)
df_ltlb_timing[cols_to_change_to_dt] = df_ltlb_timing[cols_to_change_to_dt].apply(
    pd.to_datetime
)

df_ltlb_timing["Bed Time (unix)"] = (
    df_ltlb_timing["Bed Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")

df_ltlb_timing["Get Up Time (unix)"] = (
    df_ltlb_timing["Get Up Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")

df_ltlb_timing = df_ltlb_timing.drop(
    columns=[
        "Bed Time",
        "Get Up Time",
        "Total Time in Bed (hours)",
        "Total Sleep Time (hours)",
    ]
)


for timing in df_control_timing.columns:
    figure = plt.figure()
    concat_df = pd.DataFrame(
        {
            "Control Group": df_control_timing[timing],
            "LTLB Group": df_ltlb_timing[timing],
        }
    )
    medians = concat_df.median()
    box_plot = sns.boxplot(data=concat_df, palette="flare")

    sns.set(rc={"figure.figsize": (11.37, 8.27)})
    ax = plt.gca()
    y_axis = ax.get_yticks()
    if timing == "Bed Time (unix)" or timing == "Get Up Time (unix)":

        ax.set_yticklabels(
            [pd.to_datetime(tm, unit="s").strftime("%I:%M:%S %p") for tm in y_axis]
        )
        if timing == "Bed Time (unix)":
            plt.ylabel("Bed Time")
        else:
            plt.ylabel("Get Up Time")
    else:
        ax.set_yticklabels([pd.to_timedelta(a, unit="h") for a in y_axis])
        plt.ylabel(f"{timing}")
    plt.xlabel("Research groups")

    print(f"Generating boxplot for {timing} using seaborn...")
    figure.savefig(f"{timing}.png", format="png")
    plotly_timing = px.box(concat_df)
    print(f"Generating boxplot for {timing} using plotly...")
    plotly_timing.write_html(f"{timing}.html")

df_control_summarised_timing = df_control_timing.describe()
df_control_summarised_timing["Bed Time"] = pd.to_datetime(
    df_control_summarised_timing["Bed Time (unix)"], unit="s"
)

df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].dt.strftime("%I:%M:%S %p")

df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].str.replace("AM", "pm")
df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].str.replace("PM", "am")

df_control_summarised_timing["Get Up Time"] = pd.to_datetime(
    df_control_summarised_timing["Get Up Time (unix)"], unit="s"
)

df_control_summarised_timing["Get Up Time"] = df_control_summarised_timing[
    "Get Up Time"
].dt.strftime("%I:%M:%S %p")

df_ltlb_summarised_timing = df_ltlb_timing.describe()
df_ltlb_summarised_timing["Bed Time"] = pd.to_datetime(
    df_ltlb_summarised_timing["Bed Time (unix)"], unit="s"
)
df_ltlb_summarised_timing["Get Up Time"] = pd.to_datetime(
    df_ltlb_summarised_timing["Get Up Time (unix)"], unit="s"
)
df_ltlb_summarised_timing["Bed Time"] = df_ltlb_summarised_timing[
    "Bed Time"
].dt.strftime("%I:%M:%S %p")

df_ltlb_summarised_timing["Get Up Time"] = df_ltlb_summarised_timing[
    "Get Up Time"
].dt.strftime("%I:%M:%S %p")

df_ltlb_summarised_timing["Bed Time"] = df_ltlb_summarised_timing[
    "Bed Time"
].str.replace("AM", "pm")
df_ltlb_summarised_timing["Bed Time"] = df_ltlb_summarised_timing[
    "Bed Time"
].str.replace("PM", "am")

df_ltlb_summarised_timing["Bed Time"] = df_ltlb_summarised_timing[
    "Bed Time"
].str.replace("am", "AM")
df_ltlb_summarised_timing["Bed Time"] = df_ltlb_summarised_timing[
    "Bed Time"
].str.replace("pm", "PM")

print(
    "Generating the summarised stats for the timing portion (Bed Time, Get Up Time, Total time in Bed, and Total Sleep Time)..."
)
data_timing = []

data_timing.extend([df_control_summarised_timing, df_ltlb_summarised_timing])


# --- Obtaining and plotting the stats for Onset Latency, Sleep Efficiency, WASO, and Number of Awake. ---


df_just_numbers_control = df_control_group.iloc[:, 4:]
df_just_numbers_ltlb = df_ltlb_group.iloc[:, 4:]

for column in df_just_numbers_control.columns:
    combined_dfs = pd.DataFrame(
        {
            "Control Group": df_just_numbers_control[column],
            "LTLB Group": df_just_numbers_ltlb[column],
        }
    )
    fig = plt.figure()
    sns.boxplot(data=combined_dfs, palette="flare")
    print(f"Generating box plot for {column} using matplotlib....")
    fig.savefig(f"{column}.png", format="png")
    figure_plotly = px.box(combined_dfs)
    print(f"Generating boxplot for {column} using plotly...")
    figure_plotly.write_html(f"{column}.html")
print("All plots have been generated successfully!")
summarised_stats_control = df_just_numbers_control.describe()
summarised_stats_ltlb = df_just_numbers_ltlb.describe()
print("Generating the summarised stats as an Excel file...")

data_sets = []

data_sets.extend([summarised_stats_control, summarised_stats_ltlb])

row_number = 0

writer = pd.ExcelWriter("Summarised Stats.xlsx", engine="xlsxwriter")

for time_summarised in data_timing:
    time_summarised.to_excel(
        writer, sheet_name="Datetime variables", startrow=row_number
    )
    row_number += 10

row_number = 0

for summarise in data_sets:
    summarise.to_excel(
        writer, sheet_name="Float and Integer variables", startrow=row_number
    )
    row_number += 10
writer.save()


print("The excel file has been generated successfully!")

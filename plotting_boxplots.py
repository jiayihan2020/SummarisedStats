import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import os


working_directory = Path("For plotting graph")
os.chdir(working_directory)

df_control_group = pd.read_excel(
    "Median values for Control AY 2021 to AY2022 Tri 2 and 3.xlsx"
)
df_ltlb_group = pd.read_excel(
    "Median values for LTLB AY2021 to AY2022 Tri 2 and 3.xlsx"
)


# --- Dealing with the difficult case of handling datetime objects ---

# Isolating the dataframe from Bed Time to Total Sleep Time (hours)
df_control_timing = df_control_group.iloc[:, 0:4]
df_ltlb_timing = df_ltlb_group.iloc[:, 0:4]

# Converting the Total Bed Time and Total Sleep Time to timedelta
cols = df_control_timing.columns[2:]
df_control_timing[cols] = df_control_timing[cols].apply(pd.to_timedelta)
#  Convert the resultant timedeltas to hours represented as float
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

# Converting AM to pm and vice versa for the Bed Time so that python sorts bedtime correctly
df_control_timing["Bed Time"] = df_control_timing["Bed Time"].str.replace("AM", "pm")

df_control_timing["Bed Time"] = df_control_timing["Bed Time"].str.replace("PM", "am")
#  Converting Bed Time and Get up Time to datetime objects
cols_to_change_to_dt = df_control_timing.columns[:2]

df_control_timing[cols_to_change_to_dt] = df_control_timing[cols_to_change_to_dt].apply(
    pd.to_datetime
)

df_control_timing["Bed Time (unix)"] = (
    df_control_timing["Bed Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")

# Converting the resultant datetime objects to unix timing to simplify the boxplotting process
df_control_timing["Get Up Time (unix)"] = (
    df_control_timing["Get Up Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")

#  Remove unneeded columns so that only relevant variables are plotted
df_control_timing = df_control_timing.drop(
    columns=[
        "Bed Time",
        "Get Up Time",
        "Total Time in Bed (hours)",
        "Total Sleep Time (hours)",
    ]
)

#  Converting the Total Time in Bed, and Total Sleep Time to timedeltas, this time for LTLB group
df_ltlb_timing[cols] = df_ltlb_timing[cols].apply(pd.to_timedelta)
# Replacing AM to pm and vice versa for Bed Time and Get Up Time, this time for LTLB group
df_ltlb_timing["Bed Time"] = df_ltlb_timing["Bed Time"].str.replace("AM", "pm")

df_ltlb_timing["Bed Time"] = df_ltlb_timing["Bed Time"].str.replace("PM", "am")
#  Convert resultant timedeltas to float in hours.
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
# Converting the Bed Time and Get up Time to datetime objects
df_ltlb_timing[cols_to_change_to_dt] = df_ltlb_timing[cols_to_change_to_dt].apply(
    pd.to_datetime
)
# Converting resultant datetime objects to unix timing to simplify process of plotting boxplots.
df_ltlb_timing["Bed Time (unix)"] = (
    df_ltlb_timing["Bed Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")

df_ltlb_timing["Get Up Time (unix)"] = (
    df_ltlb_timing["Get Up Time"] - pd.Timestamp("1970-01-01")
) // pd.Timedelta("1s")
# Remove unneeded columns so that only relevant variables will be plotted.
df_ltlb_timing = df_ltlb_timing.drop(
    columns=[
        "Bed Time",
        "Get Up Time",
        "Total Time in Bed (hours)",
        "Total Sleep Time (hours)",
    ]
)

# --- Start of plotting of boxplots ---

# Iterate through the columns of one of the dataframes (since column heading are the same for control and LTLB groups)
for timing in df_control_timing.columns:

    figure = plt.figure()
    # Concatenate two series from Control and LTLB groups together as one Dataframe so that we can plot two boxplots in one chart.
    concat_df = pd.DataFrame(
        {
            "Control Group": df_control_timing[timing],
            "LTLB Group": df_ltlb_timing[timing],
        }
    )

    # sns.set(rc={"figure.figsize": (11.7, 8.3)})

    box_plot = sns.boxplot(data=concat_df, palette="flare")

    # This gets the current axes of the graphs
    ax = plt.gca()
    y_axis = ax.get_yticks()
    if timing == "Bed Time (unix)" or timing == "Get Up Time (unix)":
        # Converts the y-axis back to human readable datetime format, depending on y-axis variable.

        ax.set_yticklabels(
            [pd.to_datetime(tm, unit="s").strftime("%I:%M:%S %p") for tm in y_axis]
        )
        box_plot.set_xlabel("Research Group")
        if timing == "Bed Time (unix)":
            box_plot.set_ylabel("Bed Time")
        else:
            box_plot.set_ylabel("Get Up Time")

    else:
        ax.set_yticklabels([pd.to_timedelta(a, unit="h") for a in y_axis])
        box_plot.set_xlabel("Research Group")
        if timing == "#Awake":
            box_plot.set_ylabel("Number of Awakes")
        else:
            box_plot.set_ylabel(f"{timing}")

    print(f"Generating boxplot for {timing} using seaborn...")

    figure.savefig(f"{timing}.png", format="png", bbox_inches="tight")
    # This part plots the boxplot using plotly so that there is something to cross check and ensure the validity of the boxplot.
    plotly_timing = px.box(concat_df)
    print(f"Generating boxplot for {timing} using plotly...")
    plotly_timing.write_html(f"{timing}.html")
# --- This section aims to provide the summarise stats that are found in the boxplots in table format. This will then be exported later in the script ---
df_control_summarised_timing = df_control_timing.describe()
df_control_summarised_timing["Bed Time"] = pd.to_datetime(
    df_control_summarised_timing["Bed Time (unix)"], unit="s"
)

df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].dt.strftime("%I:%M:%S %p")
# Chaning AM to PM and vice versa to provide a more accurate interpretation of the data for Bed Time.
df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].str.replace("AM", "pm")
df_control_summarised_timing["Bed Time"] = df_control_summarised_timing[
    "Bed Time"
].str.replace("PM", "am")
# Convert Unix timing back to human readable timing
df_control_summarised_timing["Get Up Time"] = pd.to_datetime(
    df_control_summarised_timing["Get Up Time (unix)"], unit="s"
)

df_control_summarised_timing["Get Up Time"] = df_control_summarised_timing[
    "Get Up Time"
].dt.strftime("%I:%M:%S %p")
#
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
# For Bed Time variable, switch PM to AM and vice versa to produce a more correct interpretation of the data
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
# Storing the dataframe containing the summarised stats in the list so that it can be exported later in the script.
data_timing.extend([df_control_summarised_timing, df_ltlb_summarised_timing])


# --- Obtaining and plotting the stats for Onset Latency, Sleep Efficiency, WASO, and Number of Awake. ---

# Isolating the variables of interests
df_just_numbers_control = df_control_group.iloc[:, 4:]
df_just_numbers_ltlb = df_ltlb_group.iloc[:, 4:]
# Iterate through the columns of one of the dataframes (since column heading are the same for control and LTLB groups)
for column in df_just_numbers_control.columns:
    # Concat the two series of column as one Dataframe so that we can plot two boxplots in one chart.
    combined_dfs = pd.DataFrame(
        {
            "Control Group": df_just_numbers_control[column],
            "LTLB Group": df_just_numbers_ltlb[column],
        }
    )
    fig = plt.figure()
    boxy = sns.boxplot(data=combined_dfs, palette="flare")
    boxy.set_xlabel("Research Group")
    boxy.set_ylabel(f"{column}")
    print(f"Generating box plot for {column} using matplotlib....")
    fig.savefig(f"{column}.png", format="png", bbox_inches="tight")
    figure_plotly = px.box(combined_dfs)
    print(f"Generating boxplot for {column} using plotly...")
    figure_plotly.write_html(f"{column}.html")
print("All plots have been generated successfully!")
summarised_stats_control = df_just_numbers_control.describe()
summarised_stats_ltlb = df_just_numbers_ltlb.describe()
print("Generating the summarised stats as an Excel file...")

#  This section will create a table containing key stats that are found in the boxplots so that we can check the validity of the boxplot stats.
data_sets = []

data_sets.extend([summarised_stats_control, summarised_stats_ltlb])

row_number = 0

writer = pd.ExcelWriter("Summarised Stats.xlsx", engine="xlsxwriter")
# Exporting the keystats as an excel file.
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

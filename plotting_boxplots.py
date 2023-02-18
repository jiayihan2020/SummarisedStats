import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import os
import send2trash


working_directory = Path("Combining the median")
os.chdir(working_directory)

df_control_group = pd.read_excel("List of Median Values for Control group.xlsx")
df_ltlb_group = pd.read_excel("List of Median Values for LTLB group.xlsx")


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
print(
    "Generating the summarised stats for all the median collected as an excel file..."
)
# --- Exporting the summarised data to excel form for checking ---
data_sets = []

data_sets.extend([summarised_stats_control, summarised_stats_ltlb])

row_number = 0

writer = pd.ExcelWriter(
    "Consolidated median stats for control and LTLB group.xlsx", engine="xlsxwriter"
)
for summarise in data_sets:
    summarise.to_excel(writer, sheet_name="Results", startrow=row_number)
    row_number += 10
writer.save()


print("The excel file has been generated successfully!")

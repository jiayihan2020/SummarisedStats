import pandas as pd
import plotly.express as px
import plotly
import matplotlib.pyplot as plt
import data_summariser
from pathlib import Path
import os


working_directory = Path("LTLB data/Control/formatted data")
os.chdir(working_directory)

df = pd.read_excel("Median Value.xlsx")


# --- Obtaining and plotting the stats for Onset Latency, Sleep Efficiency, WASO, and Number of Awake. ---


df_just_numbers = df.iloc[:, 4:]

for column in df_just_numbers.columns:
    fig = plt.figure()
    outbox = df_just_numbers.boxplot([column])
    print(f"Generating box plot for {column} using matplotlib....")
    fig.savefig(f"{column}.png", format="png")
    figure_plotly = px.box(df_just_numbers[column])
    print(f"Generating boxplot for {column} using plotly...")
    figure_plotly.write_html(f"{column}.html")
print("All plots have been generated successfully!")
summarised_stats = df_just_numbers.describe()
print(
    "Generating the summarised stats for all the median collected as an excel file..."
)
summarised_stats.to_excel("Summarised stats for all the median for students.xlsx")
print("The excel file has been generated successfully!")

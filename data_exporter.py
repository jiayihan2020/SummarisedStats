import data_summariser
import os
import pandas as pd
import openpyxl

# --- User Input ---
working_directory = "./"

output_directory = "./formatted summarised data excel"

# --------------------

subject_code_and_identity = data_summariser.obtaining_person_identity()

if not os.path.isdir(os.path.join(working_directory, output_directory)):
    print("The output directory does not exist. Creating the output directory... ")
    os.mkdir(os.path.join(os.getcwd(), output_directory))
    print("Output directory created!")

for raw_files in os.listdir():

    if raw_files.endswith(".csv"):
        print(f"Opening and reading {raw_files}...")
        subject_code = raw_files.split("_")[0]
        summarised_data = data_summariser.combined_stats(raw_files)
        summarised_data.reset_index(inplace=True)
        summarised_data = summarised_data.rename(
            columns={
                "index": f"Summarised stats for {subject_code} - {subject_code_and_identity[subject_code]}"
            }
        )

        print(f"Exporting summarised data for {subject_code}")

        summarised_data.to_excel(
            os.path.join(
                os.getcwd(), output_directory, f"{subject_code} summarised data.xlsx"
            ),
            index=False,
            header=True,
        )

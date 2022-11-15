import data_summariser
import condensing_excel_output
import os
import pandas as pd
import openpyxl
from collections import Counter

# --- User Input ---
working_directory = "./LTLB data"

output_directory = "formatted data/"

# --------------------

os.chdir(working_directory)
if not os.path.isdir(os.path.join(os.getcwd(), output_directory)):
    print("The folder could not be found. Creating the folder...")
    os.mkdir(os.path.join(os.getcwd(), "formatted data"))
    print("Folder successfully created!")

subject_code_and_identity = data_summariser.obtaining_person_identity()
changed_watch_person = [
    same_person
    for same_person, v in Counter(list(subject_code_and_identity.values())).items()
    if v > 1
]

person_subject_code = {}

if changed_watch_person:
    list_of_file_for_person_who_changed_watch = []
    for k, v in subject_code_and_identity.items():
        if v in changed_watch_person:
            if v not in person_subject_code:
                person_subject_code[v] = [k]
            else:
                person_subject_code[v].append(k)
    excluded_file = []
    for person in changed_watch_person:
        file_of_interest = []
        for subject_code in person_subject_code[person]:
            for csv_file in os.listdir():
                if csv_file.endswith(".csv") and subject_code in csv_file:
                    file_of_interest.append(csv_file)
                    excluded_file.append(csv_file)
            summarised_data = data_summariser.combined_stats(file_of_interest)
            summarised_data.reset_index(inplace=True)
            summarised_data = summarised_data.rename(
                columns={
                    "index": f"Summarised data for {person_subject_code[person]} - {person}"
                }
            )
            summarised_data.to_excel(
                os.path.join(
                    os.getcwd(),
                    output_directory,
                    f"{person_subject_code[person]} summarised data.xlsx",
                ),
                index=False,
                header=True,
            )
    for raw_data in os.listdir():
        if raw_data.endswith(".csv") and raw_data not in excluded_file:
            print(raw_data)
            person_code = raw_data.split("_")[0]
            data_summarised = data_summariser.combined_stats([raw_data])
            data_summarised.reset_index(inplace=True)
            data_summarised = data_summarised.rename(
                columns={
                    "index": f"Summarised data for {person_code} - {subject_code_and_identity[person_code]}"
                }
            )
            data_summarised.to_excel(
                os.path.join(
                    os.getcwd(),
                    output_directory,
                    f"{person_code} summarised data.xlsx",
                ),
                index=False,
                header=True,
            )
else:
    for raw_data in os.listdir():
        if raw_data.endswith(".csv"):
            person_code = raw_data.split("_")[0]
            data_summarised = data_summariser.combined_stats(raw_data)
            data_summarised.reset_index(inplace=True)
            data_summarised = data_summarised.rename(
                columns={
                    "index": f"Summarised data for {person_code} - {subject_code_and_identity[person_code]}"
                }
            )
            data_summarised.to_excel(
                os.path.join(
                    os.getcwd(), output_directory, f"{person_code} summarised data.xlsx"
                ),
                index=False,
                header=True,
            )
condensing_excel_output.consolidating_excel_files()

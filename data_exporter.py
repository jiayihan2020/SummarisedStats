import data_summariser
import os
import pandas as pd
import openpyxl
import re
import glob

# --- User Input ---
working_directory = "./Actiware csv outputs post sleep diary/Control Group"

output_directory = "formatted data/"

location_of_manifest = "D:/OneDrive - Singapore Institute Of Technology/LTLB/Research/3. Data/Participant Manifest.xlsx"

# --------------------

os.chdir(working_directory)
if not os.path.isdir(os.path.join(os.getcwd(), output_directory)):
    print("The folder could not be found. Creating the folder...")
    os.mkdir(os.path.join(os.getcwd(), "formatted data"))
    print("Folder successfully created!")

nom_roll = data_summariser.obtaining_person_identity(location_of_manifest)
options_for_research = {"1": "Control", "2": "LTLB"}
while True:
    academic_year = input(
        "Which academic year (e.g. 21/22) do you want to filter? Note that you can only select ONE academic year:"
    )
    if not re.match(r"\d{2}/\d{2}", academic_year):
        print("ERROR: Sorry, input must be in the format XX/XX, where X is an integer")
    else:
        break
while True:
    trimester = input(
        "Which Trimester (1,2, or 3) do you want? Note that you can only input ONE Trimester:"
    )
    if not re.match(r"\b[1-3]\b", trimester):
        print("ERROR: Please input an integer between 1-3.")
    else:
        break
while True:
    which_arm = input(
        """Which research arm are you interested in?
    1) Control
    2) LTLB\n
    Select only the corresponding (e.g. 1)"""
    )
    if not re.match(r"\b[1-2]\b", which_arm):
        print(
            "ERROR: Please key in the correct number corresponding to the research arm (e.g. 1, or 2):"
        )
    else:
        break


nom_roll = nom_roll.loc[
    (nom_roll["AY"] == academic_year)
    & (nom_roll["Trimester (1/2/3)"] == float(trimester))
    & (nom_roll["Arm (LTLB/ Control)"] == options_for_research[which_arm])
]
subject_code_and_identity = nom_roll[["ACT Subject Code", "Name"]]
subject_code_and_identity.reset_index(drop=True, inplace=True)

subject_code_identity = subject_code_and_identity.set_index("ACT Subject Code")[
    "Name"
].to_dict()

people_who_changed_watch = {k: v for k, v in subject_code_identity.items() if "&" in k}
if people_who_changed_watch:
    omitted_file = []
    file_of_interest = []
    for person_code in people_who_changed_watch:
        individual_codes = person_code.split("&")
        individual_codes_no_space = [s.strip() for s in individual_codes]
        for individual_code in individual_codes_no_space:
            file_name = glob.glob(f"{individual_code}*.*")[0]
            file_of_interest.append(file_name)
            omitted_file.append(file_name)
        data_summarised = data_summariser.combined_stats(file_of_interest)
        data_summarised.reset_index(inplace=True)
        data_summarised = data_summarised.rename(
            columns={
                "index": f"{person_code} - {people_who_changed_watch[person_code]}"
            }
        )
        data_summarised.to_excel(
            os.path.join(
                os.getcwd(), output_directory, f"{person_code} summarised_data.xlsx"
            ),
            index=False,
            header=True,
        )
        file_of_interest = []
    for raw_data in os.listdir():
        if raw_data.endswith(".csv") and raw_data not in omitted_file:
            person_code = raw_data.split("_")[0]
            data_summarised = data_summariser.combined_stats([raw_data])
            data_summarised.reset_index(inplace=True)
            data_summarised = data_summarised.rename(
                columns={
                    "index": f"Summarised data for {person_code} - {subject_code_identity[person_code]}"
                }
            )
            data_summarised.to_excel(
                os.path.join(
                    os.getcwd(), output_directory, f"{person_code} summarised data.xlsx"
                ),
                index=False,
                header=True,
            )


else:
    for raw_data in os.listdir():
        if raw_data.endswith(".csv"):
            person_code = raw_data.split("_")[0]
            data_summarised = data_summariser.combined_stats([raw_data])
            data_summarised.reset_index(inplace=True)
            data_summarised = data_summarised.rename(
                columns={
                    "index": f"Summarised data for {person_code} - {subject_code_identity[person_code]}"
                }
            )
            data_summarised.to_excel(
                os.path.join(
                    os.getcwd(), output_directory, f"{person_code} summarised data.xlsx"
                ),
                index=False,
                header=True,
            )

import data_summariser
import os
import re
import glob
import subprocess
import sys

try:
    import pandas as pd
    import openpyxl
except ModuleNotFoundError:
    print("The required packages are not installed. Installing required packages...")
    required_packages = ["pandas", "openpyxl"]
    for packaging in required_packages:
        subprocess.call([sys.executable, "-m", "pip", "install", packaging])
    print(
        "Download completed! data_exporter.py will attempt to import the required packages again. If error occurs, please run the script again."
    )
    import pandas as pd
    import openpyxl


# --- User Input ---
working_directory = "LTLB Group"

output_directory = "formatted data/"

location_of_manifest = "Z:/Participant Manifest updated.xlsx"

# --------------------

os.chdir(working_directory)
if not os.path.isdir(os.path.join(os.getcwd(), output_directory)):
    print(
        "The formatted files folder to store the formatted xlsx files could not be found. Creating the folder..."
    )
    os.mkdir(os.path.join(os.getcwd(), "formatted data"))
    print("Folder successfully created!")
#  Obtaining the nominal roll of students from the Participants's Manifest. This will then isolate data based on the academic year, trimester, and research groups based on user's input.
nom_roll = data_summariser.obtaining_person_identity(location_of_manifest)
options_for_research = {"1": "Control", "2": "LTLB", "3": "None"}
while True:
    academic_year = input(
        "Which academic year (e.g. 21/22) are you interested in? Note that you can only select ONE academic year. If you do not wish to indicate an academic year, answer 'none' without quotations:"
    )
    if academic_year.casefold() == "none":
        print("No Academic Year is set!")
        break
    elif not re.match(r"\d{2}/\d{2}", academic_year):
        print("ERROR: Sorry, input must be in the format XX/XX, where X is an integer")
    else:
        break
while True:
    trimester = input(
        "Which Trimester (1,2, or 3) are you interested in? Note that you can only input ONE Trimester. If you do not wish to indicate a Trimester, answer 'none' with quotations:"
    )
    if trimester.casefold() == "none":
        print("No trimester is set!")
        break
    elif not re.match(r"\b[1-3]\b", trimester):
        print("ERROR: Please input an integer between 1-3.")
    else:
        break

if (academic_year == "22/23" and trimester == 2) or (
    academic_year == "none" and trimester.casefold() == "none"
):
    while True:
        user_warning = input(
            "WARNING: Your data may include those from AY22/23 Tri 2. For this batch of data, there is a subtle difference in how the 'AM/PM' is written compared to the other data. This will affect how the bed time statistics is calculated. Please ensure you do the necessary adjustments before carrying on with the analysis. Do you still want to continue? (Y/N): "
        )
        if user_warning.casefold() == "y" or user_warning.casefold() == "yes":
            print(
                "The script will continue to run as per normal. Please do check the output of the xlsx files with those from the Clinician Report to ensure the accuracy of the statistics."
            )
            break
        elif user_warning.casefold() == "n" or user_warning.casefold() == "no":
            sys.exit(
                "This python script has been terminated. Please run the script again after necessary adjustments have been made."
            )
        else:
            print(
                "ERROR: Invalid input! Please key in the correct input as shown in the previous prompt!"
            )
while True:
    which_arm = input(
        """Which research arm are you interested in?
    1) Control
    2) LTLB
    3) None\n
    Select only the corresponding options(e.g. 1):"""
    )
    if which_arm == "3":
        print("No research arm is selected!")
    elif not re.match(r"\b[1-2]\b", which_arm):
        print(
            "ERROR: Please key in the correct integer corresponding to the research arm (e.g. 1, or 2):"
        )
    else:
        break

# isolate the data based on the user's inputs.
if which_arm != "3":
    nom_roll = nom_roll.loc[
        (nom_roll["Arm (LTLB/ Control)"] == options_for_research[which_arm])
    ]

if academic_year.casefold() != "none" and trimester.casefold() != "none":
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
# --- If there is someone who is changing their watch ---

# Check to see if there are students who have changed their watches. If so, identify and isolate their subject IDs
people_who_changed_watch = {k: v for k, v in subject_code_identity.items() if "&" in k}
if people_who_changed_watch:
    omitted_file = []
    file_of_interest = []
    for person_code in people_who_changed_watch:
        individual_codes = person_code.split("&")
        individual_codes_no_space = [s.strip() for s in individual_codes]
        # Iterate through the list and identify the csv files of the student who has changed their watches.
        for individual_code in individual_codes_no_space:
            file_name = glob.glob(f"{individual_code}*.*")[0]
            file_of_interest.append(file_name)
            # Add the filename into the ommited file list so that we do not process their data when we are handling those who did not change their watches.
            omitted_file.append(file_name)
        # Call upon the combined stats function to combine the dataframes of the student who has changed their watches and obtained their summarised stats so that their data will be more representative.
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
        # Proceed with coming up with the summarised stats for those who did not change their watches.
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
#  --- If there is no one changing their watch ---

# If there is no one changing the watch, then proceed to generate the summarised stats.
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

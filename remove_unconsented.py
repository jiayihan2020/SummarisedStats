import pandas as pd
from pathlib import Path
import re
import send2trash


pd.set_option("display.max_rows", 200)

working_directory = Path("LTLB data/Control")
participant_manifest_location = Path(
    "D:/OneDrive - Singapore Institute Of Technology/LTLB/Research/3. Data/Participant Manifest updated.xlsx"
)


def remove_non_consented_people():
    """Remove the files of those who did not consent/withdraw from the research"""

    df = pd.read_excel(participant_manifest_location)
    df = df.loc[(df["Consent (Y/N)"] == "Y") | (df["Consent (Y/N)"] == "2Y")]
    df = df.loc[(df["ACT Subject Code"] != "N")]

    df = df["ACT Subject Code"].dropna()

    list_of_consented_students = df.to_list()

    non_consented_students = []

    for i in working_directory.iterdir():
        subject_id = i.name.split("_")[0]
        matcher = re.compile(f".*{subject_id}.*")
        if i.suffix == ".csv" and not filter(matcher.match, list_of_consented_students):
            non_consented_students.append(subject_id)
            send2trash.send2trash(i)

    if non_consented_students:

        print(
            "The following subject code did not indicate that they consented to the research, their files have been sent to the recyling bin:"
        )

        for non_consent in non_consented_students:
            print(non_consent)

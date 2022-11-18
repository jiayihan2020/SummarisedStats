# Actigraphy Summary Data Extractor and Consolidator

---

## Introduction

The scripts in this repository aim to obtain the summary data obtained from the Philips Actiwatch data for eachof the sleep participants, and consolidate them into a single Excel file.

---

## Functions of each scripts

There are three scripts that work in conjuction to produced the desired output. They are:

1) ```data_summariser.py```
2) ```data_exporter.py```
3) ```condensing_excel_output.py```

```data_summariser.py``` will open the csv and then iterate through the data that was obtained from the Philips Actiware software for each subject, and extract out the relevant data. The script will then calculate the minimum, maximum, mean, 25th, 50th (A.K.A median), and 75th percentile, as well as the standard deviation for each of the variables that were extracted.

```data_exporter.py``` will then export the resultant summarised data as a table into their respective individual excel workbook based on the subject's code.

If there are multiple subjects in the study, then ```condensing_excel_output.py``` will combine all of the excel files that were produced by the ```data_exporter.py``` script into a single Excel file.

---

## Pre-requisites

please ensure that you have the following software installed on your computer:

- Python 3.9+
- ```Pandas``` library installed
  - You may install pandas using pip via the following command in Command Prompt:
  ```pip install -U pandas```
- `xlsxwriter` module installed
  - You may enter the following command in the Command Prompt to install the module:
  ```pip install -U xlsxwriter```

---

## How to use

You will only need to change the setting in the ```condensing_excel_output.py```. Simply open the python file in an IDE (e.g. Python's built in editor, VSCode, or Notepad++ etc.) and edit the `working_directory` variable.

Ensure that the csv file(s) is/are located in the directory as specified in the `working_directory` variable, and then run the `condensing_excel_output.py` using python.

To run the `condensing_excel_output.py` script, simply press `F5` or `CTRL + F5` (depending on which IDE you use, refer to the IDE help menu for more information) and the excel file should be generated.

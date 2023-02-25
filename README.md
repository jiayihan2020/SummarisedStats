# SummarisedStats

## Introduction

The scripts in this repository aim to obtain the summary data obtained from the Philips Actiwatch data for each of the sleep participants, and consolidate them into a single Excel file.

## Functions of each scripts

There are three scripts that work hand in hand to produced the desired output. They are:

1) `data_summariser.py`
2) `data_exporter.py`
3) `condensing_excel_output.py`
4) `plotting_boxplots.py`

`data_summariser.py` will open the csv and then iterate through the data that was obtained from the Philips Actiware software for each subject, and extract out the relevant data. The script will then calculate the minimum, maximum, mean, 25th, 50th (A.K.A median), and 75th percentile, as well as the standard deviation for each of the variables that were extracted.

`data_exporter.py` will first create a folder called `formatted_data` within the location that houses the xlsx files generated by the Philips Actiware, then export the resultant summarised data as a table into their respective individual excel workbook based on the subject's code into the `formatted_data` folder.

If there are multiple subjects in the study, then `condensing_excel_output.py` will combine all of the excel files that were produced by the `data_exporter.py` script into a single Excel file called `Consolidated stats.xlsx`. `condensing_excel_output.py` will then send the individual xlsx files that were previously  generated by the `data_exporter.py` into the recycling bin. The use of this script is entirely optional and does not affect how the data is interpreted.

`plotting_boxplot.py` generates boxplot using seaborn and plotly. It will also generate a summarised data as an excel file to enable one to check the validity of the data.

## Pre-requisites

Please ensure that you have the following software installed on your computer:

- Python 3.9+
- ```Pandas``` library installed
  - You may install pandas using pip via the following command in Command Prompt/Powershell/terminal:
  ```pip install -U pandas```
   <br/>
- `xlsxwriter` module installed
  - You may enter the following command in the Command Prompt/Powershell/terminal to install the module:
  ```pip install -U xlsxwriter```
  <br/>
- `send2trash` module installed
  - You may enter the following command in the Command Prompt/Powershell/terminal:
  ```pip install -U send2trash```
  <br/>
- `seaborn` and `plotly` modules installed
  - You may enter the following command in the Command Prompt/Powershell/terminal:
 ```pip install -U seaborn plotly```

## How to use

You will only need to change the settings in the ```data_exporter.py``` and ```condensing_excel_output.py```.

### data_exporter.py

In the `User Input` section of the script, edit the `working_directory` variable with the filepath containing the CSV outputs generated from Philips Actiware software. Replace all `\` symbols with `/` symbol (without the quotation marks) in the filepath if you are using Windows. An example is shown below:

Do the same for the `location_of_manifest` variable with the filepath that points to the manifest location. Replace all `\` symbols with `/` (without the quotation marks) in the filepath if you are using Windows.

The result should look something like this:
<img src='img/data exporter user input.png'>

Run the script. If you are using the default Python IDE, press 'F5' to run the script. If you are using a different IDE (e.g. Sublime Text or VSCode, etc.), use the respective IDE's keyboard shortcut to run the script.

The summarised xlsx files generated by this script is now located in the `formatted_data` folder.

### condensing_excel_output.py

After using the `data_exporter.py`, open `condensing_excel_output.py` using an IDE and edit the `working_directory` variable in the `user input` section with the filepath containing the summarised xlsx files. If you are using Windows, replace all '\\' symbols with '/' (without the quotation marks).

Run the script. If you are using the IDE that came with Python, the keyboard shortcut is 'F5'. For other IDEs (e.g. Sublime Text, VSCode, etc), refer to the IDE's manual on the shortcut to run the file.

The consolidated excel file containing all the summarised stats will be generated.This consolidated excel file has the default filename of `Consolidated stats.xlsx`

### plotting_boxplot.py

Once the `Consolidated stats.xlsx` has been generated, isolate the median for each research participants into another excel workbook. Make sure you have a header for each column of the data. For best results, you can copy the headers from `Consolidated stats.xlsx`. You should format the excel data as shown below:

<img src= 'img/Median consolidation examples.png'>

Simply run the script using your IDE and the boxplot for each of the variables will be generated. Note that by default, the script will generate both png file and html file. The former is created using the seaborn package while the latter is generated using the plotly package. For this script, the html file should only be used as a reference to check the validity of the boxplot that was generated as a png format.

This script will also generate another summarised stats based on the data that you just created. This summarised stats is similar to the ones that were generated by `data_exporter.py` and `data_summariser.py` scripts, and can be used to check on the validity of the boxplot. The filename for this summarised stat is `Summarised Stats.xlsx`.

## Caveats

There are some caveats you should take note of while using the script.

1) Currently, boxplots that are plotted using the seaborn package (i.e. file whose extension is with `.png`) are not labelled with their median. To circumvent this, the script will generate the `Summarised Stats.xlsx` and the plotly graphs so that you can check the validity of the graphs plotted.
<br/>
2) The `data_exporter.py` and `data_summariser.py` require a participant manifest to work. A participants manifest is a file (usually an excel file) that contains the participants' details. Different research groups will have different way in formatting their participants' manifest. These scripts are tailored to the structure of our research group's participants manifest. Thus, you will need to edit both `data_exporter.py` and `data_summariser.py` to suit the format of your participants manifest.

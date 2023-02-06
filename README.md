# UCI Downhill Data Cleaning with VBA

This script was used for cleaning and preparing data for my "UCI Downsill analysis" project (https://rpubs.com/vit_kolarik/uci_dh_analysis), which tested the impact of rain on Scottish riders' results.

## Goal
The data from the oficial UCI were unorganised and unfit to further analysis (as can be seen on sheet "Source Data"). Data about each rider were spread on multiple rows and the purpose of this script was to get them nicely on one row. 

## Data
The source data I've used are from the official UCI webpage for the "Men elite - downhill". 

https://www.uci.org/competition-details/2022/MTB/65467

## Viewing the code
You can either view just the plain VBA code in text form here in GitHub or download the Excel file with macro and run it on your computer. For that, you will need to download the "uci-dh-data-cleaning.xlsm". After that, you need to right-click on the file, go to "Properties" and in "Security" click "Unblock". This might be blocking the macro from running. For running the macro you also need the "Developer" tab in your Excel. In case you don't have it, go to Files - Options - Customize Ribbon - Main Tabs - select the Developer check box.

## Running the macro
Once in "uci-dh-data-cleaning.xlsm", go to the Developers tab and Macros. You will see two options - "execute" and "reset_to_original". The "execute" macro will run the code and do all the changes. If you would like to run the code again, the "reset_to_original" macro will delete all changes and you can run the code again.

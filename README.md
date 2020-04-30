# Redline Map Generator


Generates a map in a html file for all the trails in the White Mountain Guide 30th edition. The trails are colored based on the hiked status taken from spreadsheet used to monitor red-lining progress. Read more about <a href="http://www.48x12.com/white-mountains-red-lining-objective.shtml" target="_blank">red-lining here</a>


## Notes
* The map has geometries for 643 trails
* If you find any trails missing or geometries to be wrong, please let me know at pvvktheja@gmail.com

## Instructions
1. CLick on <a href="https://mybinder.org/v2/gh/theja/redline_map/master" target="_blank"><img src="http://mybinder.org/badge_logo.svg"></a>, or follow the URL: <a href="https://mybinder.org/v2/gh/theja/redline_map/master" target="_blank">https://mybinder.org/v2/gh/theja/redline_map/master</a>
2. Open the `Input` folder and upload the excel spreadsheet you use for monitoring redline progress
    * The spreadsheet should follow the format of the `AMC_30th_Edition_Redlining_v30_9_template.xls` file in the `Input` folder
    * The code accounts for some common deviations from the format but it might raise an error sometimes
    * If an error is raised due to format, I tried to include helpful information in the error message. Adjust the format accordingly and upload the sheet again
    * If you don't have the spreadsheet, feel free to download the template and use it to monitor your hiking progress
3. Run the `redline_to_map.ipynb` jupyter notebook and when prompted, enter the name of spreadsheet (e.g. `AMC_30th_Edition_Redlining_v30_9.xls`) with the correct extension (.xls or .xlsx)
4. Output will be saved as a HTML file in the `Output` folder. You can download the file and use it as you please.
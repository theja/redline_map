# Redline Map Generator

[![Binder](http://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/theja/redline_map/master)

Generates a map in a html file for all the trails in the White Mountain Guide 30th edition. The trails are colored based on the hiked status taken from spreadsheet used to monitor progress

Access this Binder by clicking the blue badge above or at the following URL:

https://mybinder.org/v2/gh/theja/redline_map/master

## Notes
* The map has geometries for 643 trails
* If you find any trails missing or geometries to be wrong, please let me know at pvvktheja@gmail.com

## Instructions
1. Upload your excel spreadsheet for monitoring redline progress to the project
    * The spreadsheet should follow the format of the 'AMC_30th_Edition_Redlining_v30_9_template.xls' file
    * The code accounts for some common deviations from the format but it might raise an error sometimes
    * If an error is raised due to format, I tried to include helpful information in the error message. Adjust the format accordingly and upload the sheet again
    * If you don't have the spreadsheet, feel free to download the template and use it to monitor teh progress
2. Run the jupyter notebook and when prompted, enter the name of the file with the correct extension (.xls or .xlsx)

## Common Issues
*   **Issue:** The spreadsheet does not follow the format of the template\
    **Solution:** Modify the spreadsheet format, upload it and run the code again
*   **Issue:** Spelling of the trail name, tab, or section does not match what is in the geojson file with the geometries\
    **Solution:** Check the spellings on teh sheet to ensure. The error message might provide some indication on what spelling is not correct




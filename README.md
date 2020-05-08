# Redline Map Generator


Generates a map in a html file for all the trails in the White Mountain Guide 30th edition. The trails are colored based on the hiked status taken from spreadsheet used to monitor red-lining progress. Read more about red-lining here `http://www.48x12.com/white-mountains-red-lining-objective.shtml`


## Notes
* This works with 29th and 30th edition wrokbooks.
* The map has geometries for 643 trails. While this is not a complete list of trails, it has most of the ones in 30th edition. Some trails in the older 29th edition are missing
* If you find any trails missing or geometries to be wrong, please let me know at pvvktheja@gmail.com
* I won't have access to your spreadsheet or the map that is generated
* If your spreadsheet has any sensitive information, I would recommend modifying it before uploading it for security reasons

## Instructions
1. Click on <a href="https://mybinder.org/v2/gh/theja/redline_map/master"><img src="http://mybinder.org/badge_logo.svg"></a>, or copy an paste the following URL in your browser:
`https://mybinder.org/v2/gh/theja/redline_map/master`. I would recommend opening the binder in a new tab if you want to keep this page open.
2. Open the `Input` folder and upload the excel spreadsheet you use for monitoring red-lining progress.
    * The spreadsheet should follow the format of the `Redlining_template.xls` file in the `Input` folder.
    * The code accounts for some common deviations from the format but it might raise an error sometimes.
    * If an error is raised due to formatting, I tried to include helpful information in the error message. Adjust the format accordingly and upload the spreadsheet again.
    * If you don't have the spreadsheet, feel free to download the template and use it to monitor your hiking progress.
    * You are also welcome to just use the spreadsheet template to build a sample map.
3. Run the `excel_to_redline_map.ipynb` jupyter notebook and when prompted, enter the name of spreadsheet (e.g. `AMC_30th_Edition_Redlining_v30_9.xls`) with the correct extension (.xls or .xlsx).
4. Output will be saved as a HTML file in the `Output` folder. You can download the file and use it as you please.
5. The map generated will have all the trails that are completely hiked in red and everything else in blue.

## Acknowledgments
I would like to thank Beth Zimmer, Andreas Frese and others in the red-lining community for providing gps tracks of trails used here. Thank you to all those involved in making the spreadsheet template that most of use to keep track of our hiking progress. My gratitude goes to all those who maintain the trails we enjoy and those involved in making the White mountain Guide possible.

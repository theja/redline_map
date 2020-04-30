<html>
    <h1>Redline Map Generator</h1>
      <p>Generates a map in a html file for all the trails in the White Mountain Guide 30th edition. The trails are colored based on the hiked status taken from spreadsheet used to monitor progress</p>

  <h3>Notes:</h3>
    <ol>
      <li>The map has geometries for 643 trails</li>
      <li>If you find any trails missing or geometries to be wrong, please let me know at pvvktheja@gmail.com</li>
    </ol>

  <h3>Instructions:</h3>
    <ol>
      <li>Upload your excel spreadsheet for monitoring redline progress to the project
        <ul>
          <li>The spreadsheet should follow the format of the 'AMC_30th_Edition_Redlining_v30_9_template.xls' file</li>
          <li>The code accounts for some common deviations from the format but it might raise an error sometimes</li>
          <li>If an error is raised due to format, I tried to include helpful information in the error message. Adjust the format accordingly and upload the sheet again</li>
          <li>If you don't have the spreadsheet, feel free to download the template and use it to monitor teh progress</li>
       </ul>
      </li>
      <li>Run the jupyter notebook and when prompted, enter the name of the file with the correct extension (.xls or .xlsx)</li>
    </ol>

  <h23>Common Issues</h3>
    <ul>
      <li>
        <b>Issue:</b> The spreadsheet does not follow the format of the template<br>
        <b>Solution:</b> Reformat the spreadsheet, upload it and run the code again
      </li>
      <li>
        <b>Issue:</b> Spelling of the trail name, tab, or section does not match what is in my json<br>
        <b>Solution:</b> Reformat the spreadsheet, upload it and run the code again
      </li>  
    </ul>
</html>

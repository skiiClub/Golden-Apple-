# Golden-Apple-
One of the scripts I wrote for the Golden Apple Foundation during my internship with them last year. 

This script parses an excel workbook that contains a list of schools in the CPS school system and returns a new workbook that contains their respective performance rating. 

After getting all the schools in the input workbook I launch Selenium to open the CPS webpage and search the school in its search box. After the new page loads, I parse the html document and extract the performance ranking of the school using beautiful soup. After the page is scraped I write the data into a new excel sheet that contains each school and its respective performance rating. 


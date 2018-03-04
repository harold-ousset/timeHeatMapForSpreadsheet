# timeHeatMapForSpreadsheet
create a time heat map in Google Spreadsheet
  
Copy the gs file in your Google apps script and it's ready to use  
eg:  
```  
var hm = new HeatMap(); // create a new heat map  
hm.getData('sample data!A:C').render('Sheet4!H1'); // get data from sheet "sample data!A:C" and render it in "SheetA!H1"  
```  
  
Several options can be applied: select the start / end date of the rendered period, sum, average, for days that have several values, locale (en, fr) to set week starting on sunday or monday, there is also differents way to address the working ranges.  
Next planned evolutions:  
- change the gradiant color
- change the display: disable display of month, weekday , sparkline graphs
- add labels with a third column for data with comment
- remove weekends
- allow differents date imput format (sync with the locale option)
- prevent data crush on the rendering zone

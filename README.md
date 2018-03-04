# timeHeatMapForSpreadsheet
create a time heat map in Google Spreadsheet
  
Copy the gs file in your Google apps script and it's ready to use  
eg: 
  var hm = new HeatMap(); // create a new heat map
  hm.getData('sample data!A:C').render('Sheet4!H1'); // get data from sheet "sample data!A:C" and render it in "SheetA!H1"

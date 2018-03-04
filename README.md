# timeHeatMapForSpreadsheet
create a time heat map in Google Spreadsheet
  
Copy the gs file in your Google apps script and it's ready to use  
eg:  
```  
var hm = new HeatMap(); // create a new heat map  
hm.getData('sample data!A:C').render('Sheet4!H1'); // get data from sheet "sample data!A:C" and render it in "SheetA!H1"  

```  
  
  
Several options can be applied: select the start / end date of the rendered period, sum, average, for days that have several values, locale (en, fr) to set week starting on sunday or monday, there is also differents way to address the working ranges.   
More complex and unachieved sample:  
```  
function displayMap() {
  var hm = new HeatMap();
  hm.locale = 'en';
  hm.getData('sample data!A:C');
  var startDate = new Date('2016-01-01T00:00:00');
  var endDate = new Date('2016-04-30T23:00:00');
  hm.render('goalSheet', 2, 10);
  
  var startDate = new Date('2016-05-01T00:00:00');
  endDate = new Date('2016-08-31T23:00:00');
  hm.consolidate({endDate:endDate, startDate:startDate});
  hm.buildRenderingDataForBlock()
  hm.render('goalSheet', 2, 22);
  
  var startDate = new Date('2016-09-01T00:00:00');
  endDate = new Date('2016-12-31T23:00:00');
  hm.consolidate({endDate:endDate, startDate:startDate});
  hm.buildRenderingDataForBlock();
  hm.render('goalSheet', 2, 34);
}
```  
  
  
Next planned evolutions:  
- change the gradiant color  
- change the display: disable display of month, weekday , sparkline graphs  
- add labels with a third column for data with comment  
- remove weekends  
- allow differents date imput format (sync with the locale option)  
- prevent data crush on the rendering zone  
- allow splits (eg: 4 month by 4) and keep the minMax global to the whole period  

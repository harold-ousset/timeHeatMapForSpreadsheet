# timeHeatMapForSpreadsheet  
## create a time heat map in Google Spreadsheet  
  
### Install  
  
Include the library id: **1rGvNrQ6wrtmsyE6Wf0L8Kfu38DQQx6Fp4KZKbbHUqSDkxq_R6G3C_mLu** or alternatively copy the gs file "timeHeatMap.gs" in your Google apps script  
  
### Syntax  
  
> **simple demo script**
```  
// create a new heat map from the library
var hm = new timeHeatMap.HeatMap(); 
// Aleternative option to create a new heat map from the gs file
var hm = new HeatMap(); 
// get data from sheet "sample data!A:C" and render it in "SheetA!H1"  
hm.getData('sample data!A:C').render('Sheet4!H1'); 

```  
  
### Parameters  
  
**getData**  
Define where to take the source of information  
It can be an array with two or three columns  
- first column [Date] to position theh event in the map
- second column [Value] what is the score for this date (if there is more than one score for a given date, they will be aggregated)
- third column *optional* [Label] label to display on the event (if there is more than one entry for a given date only the last entry will be taken in account)  

Arguments can be from one up to five:  
- 1 argument **SheetName!A1:B1**  
- 2 arguments **startRowNumber**, **startColumnNumber**   it will be transformed as: **ActiveSheet!A2:B**  
- 3 arguments **SheetName**, **A**, **C** it will be transformed as: **SheetName!A:C**  
- 4 arguments **startRowNumber**, **startColumnNumber**, **numberOfRows**, **numberOfColumns** it will be transformed as: **ActiveSheet!A1:B2**
- 5 arguments **SheetName**, **startRowNumber**, **startColumnNumber**, **numberOfRows**, **numberOfColumns** it will be transformed as: **SheetName!A1:B2**  
> example  
```  
hm.getData('sample data!A:C');
hm.getData(2,3); // will get the data from the active sheet starting second line column C and D
hm.getData('sample data', 'A', 'C'); // will get data from "sample data!A:C"  
hm.getData('sample data', 'A2', 'B10'); // will get data from "sample data!A2:B10"  
hm.getData(1,2,3,4); // will get data from active sheet starting first line column B up to three lines for 4 columns ==> remark only column B, C D will be used column D will be ignored  
hm.getData('sample data', 2, 1, 10, 2); // will get data from "sample data!A2:B10"

```  
  
**render**  
Define where the output should happen  
It is a coordinate reference where to drop the data  
- 1 argument **SheetName!A1**  
- 2 arguments **startRowNumber**, **startColumnNumber** it will be transformed as **ActiveSheet!A1** where A <=> startRowNumber and 1 <=> startColumnNumber  
- 3 arguments **SheetName**, **startRowNumber**, **startColumnNumber**  
> example  
```
hm.render('timeHeatMap!A1');
hm.render(4, 8); // will render the heat map on the active Sheet at H4
hm.render('timeHeatMap', 2, 10); // will render the heat map on timeHeatMap!J2
```  
  
**consolidate**  
*optional parameter*  
Allow you to calculate what period to take in account and how to aggregate data from same date  
It take two arguments period and aggregateMethod  
- period, is an object with two optional arguments startDate and endDate (to be given as JS date)  
- aggregateMethod, is a String that can take the values: [sum, average, median]. sum is the default value  
>example  
```  
var startDate = new Date('2016-01-01T00:00:00');
var endDate = new Date('2016-12-31T23:00:00');
hm.consolidate({endDate:endDate, startDate:startDate}); // will take whole year 2016
var halfDate = new Date('2016-06-30T00:00:00');
hm.consolidate({startDate:halfDate}); // will start the 30 of June until the end of the year
hm.consolidate({}, "average"); // will take the default period but will make an average with the values from the same days 
```  

>Note: if no period is selected by defect the period will start the 1 January of the year where there is data until the end of this first year. (data on the following years will be ignored)  
  
**locale**  
*optional parameter*
default option "fr" set week start on Monday
alternative option "en-us" set week start on Sunday
  
> example  
```
hm.locale = "en-us";
``` 


**gradient**  
*optional parameter*
- default  
![default](https://i.imgur.com/NmcIIYs.png)
- weddingDayBlues  
![weddingDayBlues](https://i.imgur.com/j9758OU.png)  
- kingYna  
![kingYna](https://i.imgur.com/oEwJv7z.png)  
- visionOfGrandeur  
![visionOfGrandeur](https://i.imgur.com/Zr8dNjq.png)  
- timber  
![timber](https://i.imgur.com/TZXjXPD.png)  

> example    
```
hm.gradient = "weddingDayBlues";
```


Several options can be applied: select the start / end date of the rendered period, sum, average, for days that have several values, locale (en-us, fr) to set week starting on sunday or monday, there is also differents way to address the working ranges.   
More complex and unachieved sample:  

```  
function displayMap() {
  var hm = new timeHeatMap.HeatMap();
  hm.locale = 'en';
  hm.getData('sample data!A:C');
  var startDate = new Date('2016-01-01T00:00:00');
  var endDate = new Date('2016-04-30T23:00:00');
  hm.consolidate({endDate:endDate, startDate:startDate});
  hm.render('timeHeatMap', 2, 10);
  
  var startDate = new Date('2016-05-01T00:00:00');
  endDate = new Date('2016-08-31T23:00:00');
  hm.consolidate({endDate:endDate, startDate:startDate});
  hm.render('timeHeatMap', 2, 22);
  
  var startDate = new Date('2016-09-01T00:00:00');
  endDate = new Date('2016-12-31T23:00:00');
  hm.consolidate({endDate:endDate, startDate:startDate});
  hm.render('timeHeatMap', 2, 34);
}
```  
![4months](https://i.imgur.com/grR9L7F.png)  
>example of a heat map  
```
  var hm = new HeatMap();
  hm.getData('sample data!A:C');
  hm.gradient = "kingYna";
  var startDate = new Date('2017-01-01T00:00:00');
  hm.consolidate({startDate:startDate});
  hm.render('goalSheet', 2, 2);
```
![sampleDataPlusGradientHeatMap](https://i.imgur.com/E12RHM6.png)
  
  
*** 
  
### Next planned evolutions:  
- Behaviour: always start at the begining of the month even if the given start date is not the first day of the month, or allow random starts?
- add methods to change rendering optins
- change the display: disable display of month, weekday , sparkline graphs  
- remove weekends  
- allow differents date imput format (sync with the locale option)  
- prevent data crush on the rendering zone  
- allow splits (eg: 4 month by 4) and keep the minMax global to the whole period  

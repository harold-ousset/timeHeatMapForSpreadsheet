# timeHeatMapForSpreadsheet  
## create a time heat map in Google Spreadsheet  
  
  
### Install  
  
Include the library id: **1rGvNrQ6wrtmsyE6Wf0L8Kfu38DQQx6Fp4KZKbbHUqSDkxq_R6G3C_mLu** or alternatively copy the gs file "timeHeatMap.gs" in your Google apps script  
  
### Syntax  
Time heat map can be invoked by calling `new HeatMap()`if called as a library you'll first need to call the library name eg: `var hm = new timeHeatMap.HeatMap()`where "timeHeatMap" is the name of the library.  
A HeatMap can take two arguments: the source of the data and the destination where to render the map. They can be given directly when creating the heat map eg: `new HeatMap(source, output);` where *source* is a string representing a range coordinates: SheetName!A1:B2 and *output* is a string representing a cell coordinate where to render the graph: SheetName!A1.  
To draw the heat map call the method `render()`
  
> **Demo script**
```javascript
function minimalDemo() {
  // create a new heat map from the library, get data from sheet "sampleDataSet!A:C" and set destination as "renderSheet!A1"
  var hm = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'renderSheet!A1');
  // render the heat map
  hm.render();
}
```  
  
**source** and **output** can also be given later with the parameters `setSource` and `setOutputCoordinates`  
```javascript
function demoWithInAndOutParam(){
  // initialize a heat map objeect
  var hm = new timeHeatMap.HeatMap();
  // define the input coordinates
  hm.setSource('sampleDataSet!A:C');
  // define the output coordinates
  hm.setOutputCoordinates('renderSheet!A1');
  // render teh heat map
  hm.render();
}
```
    

### Parameters  
#### setSource  
Define where to take the source of information  
It can be an array with two or three columns  
- first column [Date] to position the event in the map
- second column [Value] what is the score for this date (if there is more than one score for a given date, they will be aggregated)
- third column *optional* [Label] label to display on the event (if there is more than one entry for a given date only the last entry will be taken in account)  

Arguments can be from one up to five:  
- 1 argument **SheetName!A1:B1**  
- 2 arguments **startRowNumber**, **startColumnNumber**   it will be transformed as: **ActiveSheet!A2:B**  
- 3 arguments **SheetName**, **A**, **C** it will be transformed as: **SheetName!A:C**  
- 4 arguments **startRowNumber**, **startColumnNumber**, **numberOfRows**, **numberOfColumns** it will be transformed as: **ActiveSheet!A1:B2**
- 5 arguments **SheetName**, **startRowNumber**, **startColumnNumber**, **numberOfRows**, **numberOfColumns** it will be transformed as: **SheetName!A1:B2**  
> example  
```javascript  
hm.setSource('sample data!A:C');
hm.setSource(2,3); // will get the data from the active sheet starting second line column C and D
hm.setSource('sample data', 'A', 'C'); // will get data from "sample data!A:C"  
hm.setSource('sample data', 'A2', 'B10'); // will get data from "sample data!A2:B10"  
hm.setSource(1,2,3,4); // will get data from active sheet starting first line column B up to three lines for 4 columns ==> remark only column B, C D will be used column D will be ignored  
hm.setSource('sample data', 2, 1, 10, 2); // will get data from "sample data!A2:B10"

```  
  
#### setOutputCoordinates  
Define where the output should happen  
It is a coordinate reference where to drop the data  
- 1 argument **SheetName!A1**  
- 2 arguments **startRowNumber**, **startColumnNumber** it will be transformed as **ActiveSheet!A1** where A <=> startRowNumber and 1 <=> startColumnNumber  
- 3 arguments **SheetName**, **startRowNumber**, **startColumnNumber**  
> example  
```javascript
hm.setOutputCoordinates('timeHeatMap!A1');
hm.setOutputCoordinates(4, 8); // will render the heat map on the active Sheet at H4
hm.setOutputCoordinates('timeHeatMap', 2, 10); // will render the heat map on timeHeatMap!J2
```  
  
#### setPeriod  
Allow you to define what period to take in account. By defect all the values contained in the source range will be parsed and only the first year will be rendered. You can here specify a year from when to start.
- period, is an object with two optional arguments **startDate** and **endDate** (to be given as JS date)  
>example  
```javascript  
var startDate = new Date('2016-01-01T00:00:00');
var endDate = new Date('2016-12-31T23:00:00');
hm.setPeriod({endDate:endDate, startDate:startDate}); // will take whole year 2016

var halfDate = new Date('2016-06-30T00:00:00');
hm.setPeriod({startDate:halfDate}); // will start the 30 of June until the end of the year
```  

>Note: if no period is selected by defect the period will start the 1 January of the year where there is data until the end of this first year. (data on the following years will be ignored)  

#### setAggregationMethod 
 When you have more than one value for a given day, the data need to be aggregated. There is three options available: sum, average and median.
By default (if nothing is specified) the used aggregation method is a sum.
eg: 
```javascript
hm.setAggregationMethod('average');
```  

demo:
```javascript  
function changeAggregationMethod(){
  var period = {
  startDate: new Date('2016-01-01T00:00:00'),
  endDate: new Date('2016-01-31T00:00:00')
  }; // period cover January 2016
  
  // create a normal HM
  var standardHM = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'aggregate demo!A4').setPeriod(period).render();
  
  // create a HM where data for the same day are smothed by their average
  var averageAggregationHM = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'aggregate demo!M4').setPeriod(period);
  averageAggregationHM.setAggregationMethod('average').render();
}
``` 
![aggregate method](https://i.imgur.com/BY95LKM.png)  

#### setLocale  
default option "fr" set week start on Monday and language to french
alternative option "en-us" set week start on Sunday and language is in english
  (language for displayed month and week day)
> example  
```javascript
hm.setLocale("en-us");
``` 
*Note: more options are to come*

#### setGradient  
Let you change the color gradient used to show the different map temperatures
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
```javascript
hm.setGradient("weddingDayBlues");
```

setGradient can take one or two arguments:
If only one argument is given it's one of the gradient color listed above.
If two argument are given, the first one is a name for a new color gradient, the second is an array with the hexColor representing the gradient. eg:
```javascript
hm.setGradient("Atlas", ["#FEAC5E", "#C779D0", "#4BC0C8"]); // idea: https://uigradients.com/#Atlas
```  

Demo:  
```javascript
function changeGradient(){
  var period = { startDate: new Date('2017-01-01T00:00:00'), endDate: new Date('2017-03-31T23:00:00') }; // period cover January , February, March 2017
  var standardHM = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'gradient demo!A4').setPeriod(period).render(); // create a normal HM
  // create a HM with Atlas gradient
  var atlasHM = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'gradient demo!M4').setPeriod(period).setGradient("Atlas", ["#FEAC5E", "#C779D0", "#4BC0C8"]).render();
}
```

![gradient modification](https://i.imgur.com/GWb5TME.png)


#### setGraphsize  
let you change the global size of the graph with a multiplication factor. eg:
```javascript
hm.setGraphsize(2); // graph will be twice the standard size
```  

#### setSplit  
Allow you to display the graph in several columns. The split is made every X months
```javascript
hm.setSplit(4); // for a whole year graph it will be splited in 3 columns Jan-Aprl | May-Aug | Sept-Dec
``` 

#### displayCommentLabel  
If you define 3 column as the input source the third column will let you display the text as comment for the given date. 
```javascript  
hm.displayCommentLabel(TRUE);
```  
The default value is FALSE  
Note: if you have several entries for one date, only the last comment will be taken in account

#### displayMonthLabel  
For each month display a three letter code  is also displayed at it's left. You can set this setting off
```javascript
hm.displayMonthLabel(FALSE);
``` 

#### displayGradientLegend  
At the end of the graphic a legend is displayed showing the min / max and color code, you can swith it off.  
```javascript
hm.displayGradientLegend(FALSE);
```  

#### drawBackground  
By default no background is rendered (you can remove the grid display in the spreadsheet settings). With this option you can force the background to be rendered (in white)  
```javascript  
hm.drawBackground(TRUE);
```  

#### displayValues  
IT will write the value for each cell in the same color than the background of the cell, this will let you know the actual value by selecting the cell.  
```javascript
hm.displayValues(TRUE);
```  

#### equalize  
If your data are not homogenous you may observe a color shift. Using this option will let you maximise the color range available (the option do the same as an histogram equalization for a photograph)
```javascript
hm.equalize(TRUE);
```  
demo:  
```javascript  
function equalizeModification(){
  var period = { startDate: new Date('2018-01-01T00:00:00'), endDate: new Date('2018-05-31T23:00:00')}; // start in 2018
  var hm = new timeHeatMap.HeatMap('sampleDataSet!A:C', 'equalize demo!A4');
  hm.setPeriod(period).render(); // create a normal HM
  // apply equalization
  hm.equalize(true);
  hm.setOutputCoordinates('equalize demo!M4').render();
}
```
![equalization demo](https://i.imgur.com/QXKpPIb.png)
  
    
     
#### params  
let you change several parameters on one single step.
eg:
```javascript
hm.params( { equalize: True, displayValues: True } );
```  

### Other sample
Several options can be applied: select the start / end date of the rendered period, sum, average, for days that have several values, locale (en-us, fr) to set week starting on sunday or monday, there is also differents way to address the working ranges.   
More complex and unachieved sample:  

```javascript  
function displayMap() {
  var hm = new timeHeatMap.HeatMap();
  hm.setLocale('en-us');
  hm.setSource('sample data!A:C');
  var startDate = new Date('2016-01-01T00:00:00');
  var endDate = new Date('2016-04-30T23:00:00');
  hm.setPeriod({endDate:endDate, startDate:startDate});
  hm.setOutputCoordinates('timeHeatMap', 2, 10);
  hm.render();
  
  var startDate = new Date('2016-05-01T00:00:00');
  endDate = new Date('2016-08-31T23:00:00');
  hm.setPeriod({endDate:endDate, startDate:startDate});
  hm.setOutputCoordinates('timeHeatMap', 2, 22);
  hm.render();
  
  var startDate = new Date('2016-09-01T00:00:00');
  endDate = new Date('2016-12-31T23:00:00');
  hm.setPeriod({endDate:endDate, startDate:startDate});
  hm.setOutputCoordinates('timeHeatMap', 2, 34);
  hm.render();
}
```  
![4months](https://i.imgur.com/grR9L7F.png)  
>example of a heat map  
```javascript
  var hm = new HeatMap();
  hm.setSource('sample data!A:C');
  hm.setGradient("kingYna");
  var startDate = new Date('2017-01-01T00:00:00');
  hm.setPeriod({startDate:startDate});
  hm.setOutputCoordinates('goalSheet', 2, 2);
  hm.render();
```
![sampleDataPlusGradientHeatMap](https://i.imgur.com/E12RHM6.png)
  
  
*** 
  
### Next planned evolutions:  
- Option to remove weekends  
- Allow differents date input format (sync with the locale option)  
- add more locales options (english starting monday...)
- Prevent data crush on the rendering zone (if there is some data it will throw an error instead of rendering the map)  





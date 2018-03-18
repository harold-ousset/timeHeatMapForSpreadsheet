/*
Copyright 2018 Harold Ousset

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

var HeatMap = function (){
  this.locales = {
    "en-us": {
      days:['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
      months:['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      weekStart:0
    },
    "fr": {
      days:['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'],
      months:['Jan', 'Feb', 'Mar', 'Avr', 'Mai', 'Jui', 'Jul', 'Aou', 'Sep', 'Oct', 'Nov', 'Dec'],
      weekStart:1
    }
  };
  
  this.gradients = {
  // inspiration: https://uigradients.com/
  "weedingDayBlues": ["#40E0D0", "#FF8C00", "#FF0080"],
  "kingYna": ["#1a2a6c", "#b21f1f", "#fdbb2d"],
  "visionOfGrandeur": ["#000046", "#1CB5E0"],
  "timber": ["#fc00ff", "#00dbde"],
  "default": ["#00FF00", "#FFFF00","#FF0000"],
  };
  this.gradient = "default";
  
  this.width = 16;
  this.height = 16;
  this.locale = 'fr';
  
  this.minMax = {minDate:undefined, minVal:undefined, maxDate:undefined, maxVal:undefined};
  
  this.labelOn = false; // is there a column with labels
  this.data = {}; // {YYYYMMJJ:{values:[], label: String}}
  this.errors = []; // {err, record}
  this.timelapse = []; // currated content that will be displayed day by day value: [day, displayDay, value, label]
  this.timelapseBounds = {} // {minDate:undefined, minVal:undefined, maxDate:undefined, maxVal:undefined, heavisetWeekVolume:undefined};
  this.constructionMaterial = {weekDayMatrix:[], weekSumMatrix:[], borderedRanges:[], labelledCells:[] , colorMatrix:[], monthLabelMatrix:[]};
  this.isSet = {timelapse: false, constructionMaterial: false};
};

/**
* Retrieve the data froom the spreadsheet
* @param{String} eref,reference a range with notation SheetName!A1:B2 || The Column to start with (A, B , C...)
* @param{String=} erow, last Column or first Row if there is a following argument, or first Column if there is 4 args
* @parma{String=} ecol, 
* ca me route a finir un jour...
**/
HeatMap.prototype.getData = function(eref, erow, ecol, enumRow, enumCol){
  // https://stackoverflow.com/questions/9905533/convert-excel-column-alphabet-e-g-aa-to-number-e-g-25
  var lettersToNumber = function(val) {
    var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;
    
    for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
      result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
    }
    
    return result;
  };
  
  // https://stackoverflow.com/questions/45787459/convert-number-to-alphabet-string-javascript
  var numToSSColumn = function (num){
    var s = '', t;
    
    while (num > 0) {
      t = (num - 1) % 26;
      s = String.fromCharCode(65 + t) + s;
      num = (num - t)/26 | 0;
    }
    return s || undefined;
  }
  
  var r1c1Exp = /^(?:['"]{0,1}(.+?)[['"]{0,1}!){0,1}R(\d+)C(\d+):R(\d+)C(\d+)$/;  // recognise the sheet ref as R1C1
  var a1Exp = /^(?:['"]{0,1}(.+?)[['"]{0,1}!){0,1}([a-zA-Z]+)(\d*):([a-zA-Z])+(\d*)$/; // recongnise the sheet ref as A1
  var range; // dataSource to be retrieved
  
  /**
  * retrieve the min and max val of the selected interval
  * @this{Object} HeatMap
  * @param{Array} record, [date, value, label] (label is optional)
  **/
  function reduceData_(record, idx){
    if(record[0] == '' && record[1] == ''){
      return;
    }
    try {
      if(isNaN(record[1])) {
        throw 'not a number';
      }
      // TODO handle other date format
      var date = record[0].getFullYear()+("0"+(record[0].getMonth()+1)).slice(-2)+('0'+record[0].getDate()).slice(-2); 
      this.data[date] = this.data[date] || {values:[],label:''};
      this.data[date].values.push(record[1]);
      if(this.labelOn){
        this.data[date].label = record[2];
      }
      if (this.minMax.minDate == undefined || this.minMax.minDate > record[0]) {
        this.minMax.minDate = record[0];
      }
      else if (this.minMax.maxDate == undefined || this.minMax.maxDate < record[0]) {
        this.minMax.maxDate = record[0];
      }
      if (this.minMax.minVal == undefined || this.minMax.minVal > record[1]) {
        this.minMax.minVal = record[1];
      }
      else if (this.minMax.maxVal == undefined || this.minMax.maxVal < record[1]) {
        this.minMax.maxVal = record[1];
      }
    }
    catch (err) {
      this.errors.push({err:err, record:record});
    }
  }
  
  if(erow == undefined){ // 1 arg
    range = SpreadsheetApp.getActive().getRange(eref); // SheetName!A1:B2
  }
  else if(ecol == undefined){ // 2 args
    var startCol = numToSSColumn(erow);
    var endCol = numToSSColumn(Number(erow)+1);
    range = SpreadsheetApp.getActiveSheet().getRange(startCol+eref+":"+endCol); // notation A1:B
  }
  else if(enumRow == undefined){ // 3 args
    var sheet = SpreadsheetApp.getActive().getSheetByName(eref); // eref, erow, ecol
    range = sheet.getRange(erow+':'+ecol);
  }
  else if(enumCol == undefined){ // 4 args
    var startCol = numToSSColumn(erow);
    var endCol = numToSSColumn(Number(erow)+enumRow);
    range = SpreadsheetApp.getActiveSheet().getRange(startCol+eref+":"+endCol+ecol); // notation A1:B2
  }
  else{ // 5 args
    var startCol = numToSSColumn(ecol);
    var endCol = numToSSColumn(Number(ecol)+enumCol);
    range = SpreadsheetApp.getActiveSheet().getRange(eref+"!"+startCol+erow+":"+endCol+enumRow);
  }
  
  if(range.getNumColumns()>2){
    this.labelOn = true;
  }
  var rawData = range.getValues();
  var that = this;
  rawData.forEach(reduceData_, that);
  return this; // for chaining;
}


/**
* consolidate the data arround a defined period

* @param{[Object]} period - {startDate, endDate}, the period that will be used to cover the calculation
* @param{[String]} aggregateMethod, how does the values on a same date are aggregated: [sum, average, median] default: sum
* @create{Object} heatMapDescriptor, {monthIdx:[3,6,9,...], out:[week[{day, value, color, top, bottom, right, left},],], minMax:{minVal, maxVal, minDate, maxDate}}
* @return{Object} this, for chaining purpose
**/
HeatMap.prototype.consolidate = function(period, aggregateMethod) {
  this.timelapse = [];
  this.timelapseBounds = {} 
  
  // https://gist.github.com/caseyjustus/1166258
  function median_(values) { // return the median of the results
    values.sort( function(a,b) {return a - b;} );
    var half = Math.floor(values.length/2);
    if(values.length % 2){
      return values[half];
    }
    return (values[half-1] + values[half]) / 2.0;
  }

  period = period || {};
  if(period.startDate == undefined){ // set as the start of the year
    this.timelapseBounds.minDate = new Date(this.minMax.minDate.getFullYear()+'-01-01T00:00:00');
  } 
  else{
    this.timelapseBounds.minDate = period.startDate;
  }
  if(period.endDate == undefined){ // set as the end of the year
    this.timelapseBounds.maxDate = new Date(this.timelapseBounds.minDate.getFullYear()+'-12-31T23:59:59');
  }
  else{
    this.timelapseBounds.maxDate = period.endDate;
  }
  
  this.timelapseBounds.minVal = undefined;
  this.timelapseBounds.maxVal = undefined;
  
  aggregateMethod = aggregateMethod || 'sum';
  
  var dd;  // displayed date
  var pd = new Date(this.timelapseBounds.minDate); // parsed Date
  
  do{ // parse every day during the timelapse
    dd = pd.getFullYear()+('0'+(pd.getMonth()+1)).slice(-2)+('0'+pd.getDate()).slice(-2);
    var raw = this.data[dd];
    if(raw == undefined || raw.values.length == 0){
      this.timelapse.push([new Date(pd),dd,null,null]);
    }
    else{
      var value;
      switch (aggregateMethod) {
        case 'sum':
          value = raw.values.reduce(function(a, b) { return a + b; });
          break;
        case 'average':
          value = raw.values.reduce(function(a, b) { return a + b; }) / raw.values.length;
          break;
        case 'median':
          value = median_(raw.values);
          break;
        default:
          value = raw.values.reduce(function(a, b) { return a + b; });
      }
      
      this.timelapse.push([new Date(pd),dd,value,raw.label]);  
      if (this.timelapseBounds.minVal == undefined || this.timelapseBounds.minVal > value) {
        this.timelapseBounds.minVal = value;
      }
      if (this.timelapseBounds.maxVal == undefined || this.timelapseBounds.maxVal < value) {
        this.timelapseBounds.maxVal = value;
      }
      
    }
    pd.setDate(pd.getDate()+1);
  }
  while(pd < this.timelapseBounds.maxDate);

  this.isSet.timelapse = true;
  this.isSet.constructionMaterial = false;
  return this; // return self for chaining
}


/**
*  Build rendering data for a given block
*  fill constructionMaterial object
* @param{[Array]} dateArray, the date block that will be parsed if not provided will use timelapse [day, displayDay, value, label]
* @return{Object} this, for chaining purpose
**/
HeatMap.prototype.buildRenderingDataForBlock = function (dateArray){
  
  /**
  * addBorderTo create a super object containing all the border 
  * it need to be improved (copy the borders on adjacent cells
  **/
  function addBorderTo_(row, column, side){
    this[row] = this[row] || {};
    this[row][column] = this[row][column] || {top:false, right:false, bottom:false, left:false};
    this[row][column][side] = true;
  }
  
  this.constructionMaterial = {weekDayMatrix:[], weekSumMatrix:[], borderedRanges:[], labelledCells:[] , colorMatrix:[], monthLabelMatrix:[]};
  this.isSet.constructionMaterial = true;
  
  dateArray = dateArray || this.timelapse;
  
  // check the integrity of the data
  if(dateArray.length != 1+Math.floor((dateArray[dateArray.length-1][0].getTime() - dateArray[0][0].getTime())/(3600*24*1000))){
    Logger.log('WARNING INCONSISTENT DATA');
    Logger.log(1+Math.floor((dateArray[dateArray.length-1][0].getTime() - dateArray[0][0].getTime())/(3600*24*1000))+' VS '+dateArray.length);
  }
  
  var weekDayMatrix = [[],[],[],[],[],[],[]]; // help build the the top char per day of the week
  var weekSumMatrix = []; // help build the graph for each week
  var borderedRanges = {}; // tell where to create the border (row:{column:{top, right, bottom, left}})
  var labelledCells = []; // where to add comments
  var colorMatrix = []; // how to color the shit
  var monthLabelMatrix = []; // label for the month (txt, row, column)
  
  var weekColorMatrix = []; // TEMP 7 colors line data for each parsed week
  var weekSum = null; // TEMP data to be used with weekSumMatrix
  var monthLabelWeekVal = ''; // can be '' or month to display
  var idx = 0;
  var wdCorrection = this.locales[this.locale].weekStart; // change the start of the day from sunday to monday (F..... murican)
  
  var start = dateArray[0]; // day where we start working
  var plus7 = new Date(start[0]); // week +1 from start (used for border to next month
  plus7.setDate(start[0].getDate()+7);
  var wd = start[0].getDay(); // day of the week M, T, W, ....
  
  if(wd-wdCorrection == -1){
    wd = 7 // weed to find something better than that when I'm going to implement weekstart on saturday...
  }
  while(weekColorMatrix.length < wd-wdCorrection){ // init the array with blank and border ==> what if start is different than 1 day of month?
    weekColorMatrix.push(''); // add blank cell at start
    addBorderTo_.call(borderedRanges, 0, weekColorMatrix.length - 1, 'bottom'); // add the header cap for WEEK 2 (row | column)
  }
  
  addBorderTo_.call(borderedRanges, 0, weekColorMatrix.length, 'left'); // first day of the month border ==> what if this is not the case?
  
  do{ // parse each day in the array and do the job with it! // while(dateArray.length > 0)
    var dayVals = dateArray.shift();
    plus7 = new Date(dayVals[0]);
    plus7.setDate(dayVals[0].getDate()+7); // week +1 from start
    
    wd = dayVals[0].getDay();
    if(dayVals[2] != null){ // if there is data handle them
      weekDayMatrix[wd].push(dayVals[2]);
      weekSum += dayVals[2];
      weekColorMatrix.push(this.getGradientColor_(dayVals[2], this.timelapseBounds.minVal, this.timelapseBounds.maxVal, this.gradients[this.gradient])); // TODO include change palette color
      if(dayVals[3] != null && dayVals[3] != undefined && dayVals[3] != ""){
      labelledCells.push({label:dayVals[3], column:wd-wdCorrection, row: colorMatrix.length});
      }
    }
    else{ // there is no data just add a GREY cell
      weekColorMatrix.push('#E6E6E6');
    }
    
    // capping
    if(colorMatrix.length == 0){
      addBorderTo_.call(borderedRanges, 0, weekColorMatrix.length -1, 'top'); // cap the head for WEEK 1  ==> what if this is not the first week of month???
    }
    else if(dayVals[0].getMonth() !== plus7.getMonth()){ // cap the bottom if next week is on an other month
      addBorderTo_.call(borderedRanges, colorMatrix.length, weekColorMatrix.length - 1, 'bottom'); 
      var plus1 = new Date(dayVals[0]);
      plus1.setDate(plus1.getDate()+1);
      if(dayVals[0].getMonth() !== plus1.getMonth()){ // cap the last day of the month
        addBorderTo_.call(borderedRanges, colorMatrix.length, weekColorMatrix.length -1, 'right');
      }
    }
    
    // label the month
    if(dayVals[0].getDate() == 8){
      monthLabelMatrix.push([this.locales[this.locale].months[dayVals[0].getMonth()], colorMatrix.length + 1, 0]); // txt, row, column (index 0)
    }
    
    // it's a new week store the color, weeksum data
    if((wdCorrection == 0 && wd == 6) || (wdCorrection == 1 && wd == 0)){
      if(this.timelapseBounds.heavisetWeekVolume == undefined || this.timelapseBounds.heavisetWeekVolume < weekSum){
        this.timelapseBounds.heavisetWeekVolume = weekSum;
      }
      weekSumMatrix.push(weekSum);
      colorMatrix.push(weekColorMatrix);
      weekSum = null;
      weekColorMatrix = [];
    }
  }
  while(dateArray.length > 0);
  
  // properly finish the array
  while(weekColorMatrix.length > 0 ){ // add the final blanks if color Matrix is not full
    weekColorMatrix.push(''); // add blank cell at start
    if(weekColorMatrix.length == 7){
      colorMatrix.push(weekColorMatrix);
      weekSumMatrix.push(weekSum);
      weekColorMatrix = [];
    }
  }
  
  this.constructionMaterial.weekDayMatrix = weekDayMatrix;
  this.constructionMaterial.weekSumMatrix = weekSumMatrix;
  this.constructionMaterial.borderedRanges = borderedRanges;
  this.constructionMaterial.labelledCells = labelledCells;
  this.constructionMaterial.colorMatrix = colorMatrix;
  this.constructionMaterial.monthLabelMatrix = monthLabelMatrix;
  return this;
}


/**
* render tha map in a spreadsheet
* @param{[String]} eref, the full reference where to render the map (SheetName!A1) OR the SheetName OR the starting Column
* @param{[String]} erow, if ecol == undefined ==> the Column where to render the map ELSE the Row
* #param{[String]} ecol, the Column where to render the map
**/
HeatMap.prototype.render = function (eref, erow, ecol){
  /** 
  * set the border of a cell 
  * @param{Object} sheet, the sheet where to apply the data
  * @param{String} row, cell row
  * @param{String} column, cell column
  * @param{Object} borderObj, where to paint a border {top, left, bottom, right}
  **/
  function setBorders_(sheet, row, column, borderObj){
    sheet.getRange(row, column).setBorder(
      borderObj.top, 
      borderObj.left, 
      borderObj.bottom, 
      borderObj.right, 
      null, 
      null, 
      "black", 
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
   
  /**
  * create the formula that display the sparkline graph for each row
  * @param{Number} max for the week
  * @return{String} the sparkline formula
  **/
  function setWeekSumSparks_(val){
    if(val && this.timelapseBounds.heavisetWeekVolume != null){
      return ['=SPARKLINE('+val+', {"charttype","bar";"max",'+this.timelapseBounds.heavisetWeekVolume+';"color1","#2E9AFE"})'];
    }
    return [''];
  } 
  
  /**
  * create the formula that display the sparkline graph for each row
  * @this{Object} vals, {minMax: {minVal, maxVal, minDate, maxDate}, errors:[]}
  * @param{Array} raw values
  * @return{String} the sparkline formula
  **/
  function setDaySparks_(arr){
    if(arr !== null && arr.length > 1) { // >1  because if there is only one value this is not enough to build a graph
      var sum = arr.reduce(function(a, b) {
        return a + b;
      }, 0);
      var avg = (sum/arr.length)-this.timelapseBounds.minVal;
      return '=SPARKLINE('+avg+', {"charttype","column";"ymin",0;"ymax",'+(this.timelapseBounds.maxVal-this.timelapseBounds.minVal)+';"firstcolor","#2E9AFE"})';
    }
    return '';
  }

  if(!this.isSet.timelapse){
    this.consolidate();
  }
  if(!this.isSet.constructionMaterial){
    this.buildRenderingDataForBlock();
  }
  this.isSet = {timelapse: false, constructionMaterial: false};
  
  var that = this;
  var sheet = SpreadsheetApp.getActiveSheet();
  var column = ecol || 1;
  var row = erow || 1;
  
  if(eref != undefined && erow == undefined){ // 1 arg
    var range = SpreadsheetApp.getActive().getRange(eref); // SheetName!A1
    sheet = range.getSheet();
    row = range.getRow();
    column = range.getColumn(); 
  }
  else if(ecol == undefined && erow != undefined){
    row = eref;
    column = erow;
  }
  else if(eref != undefined && erow != undefined && ecol != undefined){
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eref);
  }
  
  
  // clean area before doing anything
  sheet.getRange(row, column, this.constructionMaterial.colorMatrix.length+7, 11).clear().clearNote();
  
  // set comment labels, start at position 4 | 2
  var labels = this.constructionMaterial.labelledCells;
  labels.forEach(function (lblObj){
    sheet.getRange(lblObj.row+ row+ 4, lblObj.column + 2 + column).setNote(lblObj.label);
  }); 
  
  // set weeks sparks, start at position 4 | 10
  var sparks = this.constructionMaterial.weekSumMatrix.map(setWeekSumSparks_, that);
  sheet.getRange(row+4, column+10, sparks.length, 1).setFormulas(sparks);
  
  // set colors, start at position 4 | 2
  sheet.getRange(row+4, column+2, this.constructionMaterial.colorMatrix.length, 7).setBackgrounds(this.constructionMaterial.colorMatrix);;
  
  // set months labels, start at position 4 | 0
  this.constructionMaterial.monthLabelMatrix.forEach(setMonthLabel_, sheet);
  function setMonthLabel_(cellRef){
    this.getRange(cellRef[1]+row + 4, cellRef[2]+column).setValue(cellRef[0]);
  }
  
  // set upper days label, start at position 2 | 2
  var weekDays = this.locales[this.locale].days.map(
    function(day){
      return day.slice(0,1);
    }
  );
  sheet.getRange(row+2, column+2, 1, 7).setValues([weekDays]);
  
  // set color gradient legend, start at position x+6 | 2
  var grad = this.gradients[this.gradient];
  var colors = [];
  colors.push(this.getGradientColor_(0, 0, 1, grad));
  for(var i = 1; i < 6; i++){
    var valmin = i/7;
    colors.push(this.getGradientColor_(valmin, 0, 1, grad));
  }
  colors.push(this.getGradientColor_(1, 0, 1, grad));
  sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 6, column+2, 1,7).setBackgrounds([colors]);
  
  // set sparks by day type, start at position 0 | 2
  var daySparks = this.constructionMaterial.weekDayMatrix.map(setDaySparks_, that);
  var lastDays = daySparks.splice(0,this.locales[this.locale].weekStart);
  daySparks = daySparks.concat(lastDays);
  sheet.getRange(row, column+2, 1, 7).setFormulas([daySparks]);
  
  // set borders
  for(var brow in this.constructionMaterial.borderedRanges){
    for(var bcol in this.constructionMaterial.borderedRanges[brow]){
      setBorders_(sheet, Number(brow)+4+row, Number(bcol)+2+column, this.constructionMaterial.borderedRanges[brow][bcol]);   
    }
  }
  
  // set columns width of map
  for(var i = 0 ;i<7; i++) {
    sheet.setColumnWidth(i+2+column, this.width);
  }
  
  // set rows height of map
  for( var j = 0; j < this.constructionMaterial.colorMatrix.length; j++){ 
    sheet.setRowHeight(j+4, this.height);
  }
  
  // set row height for gradient
  sheet.setRowHeight(this.constructionMaterial.colorMatrix.length + row + 6, this.height*0.4);
  
  sheet.setColumnWidth(column, this.width*2.2); // set width of month label column
  sheet.setColumnWidth(column+1, this.width*0.6); // separation month label | map
  sheet.setColumnWidth(column+9, this.width*0.6); // separation map | spark week
  sheet.setColumnWidth(column+10, this.width*2.2); // set width Spark week
}

/**
* get gradient color, a sub function to generate a color code from an interval
* @param{Number} rawVal, the value to transform in color code
* @param{Number} min, minimal value from the interval
* @param{Number} max, the maximal value from the interval
* @param{Array} gradient, the colors codes used in the gradient
* @return{String} colorCode, the color code in hex code
**/
HeatMap.prototype.getGradientColor_ =   function (rawVal, min, max, gradient){
  var val = (rawVal - min) / (max - min);
  if(val == 0 || val == undefined || isNaN(val)){
    return gradient[0];
  }
  var interval = gradient.length - 1;
  var spaces = 1/interval;
  var chevron = 0;
  while(chevron*spaces < val && chevron*spaces<=1){
    chevron++;
  }
  val = (val - (chevron-1)*spaces)/ spaces;
  var myMax = gradient[chevron];
  var myMin = gradient[chevron-1];
  
  var r1,g1,b1, r2, g2, b2;
  r1 = myMin.substr(1,2);
  g1 = myMin.substr(3,2);
  b1 = myMin.substr(5,2);
  r2 = myMax.substr(1,2);
  g2 = myMax.substr(3,2);
  b2 = myMax.substr(5,2);
  var out = '#'
  +('0'+Math.round(parseInt(r1,16)+((parseInt(r2,16)-parseInt(r1,16))*val)).toString(16)).slice(-2)
  +('0'+Math.round(parseInt(g1,16)+((parseInt(g2,16)-parseInt(g1,16))*val)).toString(16)).slice(-2)
  +('0'+Math.round(parseInt(b1,16)+((parseInt(b2,16)-parseInt(b1,16))*val)).toString(16)).slice(-2);
  return out;
}

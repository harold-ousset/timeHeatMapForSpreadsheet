
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

function HeatMap (){
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
  this.locale = 'fr';
  
  this.gradients = {
    // inspiration: https://uigradients.com/
    "weddingDayBlues": ["#40E0D0", "#FF8C00", "#FF0080"],
    "kingYna": ["#1a2a6c", "#b21f1f", "#fdbb2d"],
    "visionOfGrandeur": ["#000046", "#1CB5E0"],
    "timber": ["#fc00ff", "#00dbde"],
    "default": ["#00FF00", "#FFFF00","#FF0000"],
  };
  this.gradient = "default";
  
  this.width = 16;
  this.height = 16;
  
  this.noValueColorCode = '#E6E6E6';
  
  this.draw = {monthLabel:true, gradient:true, weekDaySum:true, weekDay:true, labels:false, weekSum:true,background:false,values:false};
  
  this.minMax = {minDate:undefined, minVal:undefined, maxDate:undefined, maxVal:undefined}; // will be filled upon getSource invocation
  this.labelOn = false; // is there a column with labels that CAN be displayed as comment on the selected cells // initiated upon getSource invocation
  this.data = {}; // {YYYYMMJJ:{values:[], label: String}} will be filled upon getSource invocation
  this.errors = []; // {err, record}
  
  this.splits = 0; // divide period into several columns, defect option is no splits
  
  this.equalizeHistogram = true;
  
  this.aggregatedMethod = undefined; // how to aggregate differents data for the same day? median, average, sum? defect: sum.
  this.timelapse = []; // currated content that will be displayed day by day value: [day, displayDay, value, label]
  this.timelapseBounds = {} // {minDate:undefined, minVal:undefined, maxDate:undefined, maxVal:undefined, heavisetWeekVolume:undefined}; 
  this.constructionMaterial = {weekDayMatrix:[], weekSumMatrix:[], borderedRanges:[], labelledCells:[] , colorMatrix:[], monthLabelMatrix:[]};
  this.isSet = {timelapse: false, constructionMaterial: false, parseData:false};
};

/**
* parseData from the range ref will parse every row into a usable object this.data
@ return {Object} this, for chaining purpose.
**/
HeatMap.prototype.parseData = function(){
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
  
  
  if(this.range.getNumColumns()>2){
    this.labelOn = true;
  }
  var rawData = this.range.getValues();
  var that = this;
  rawData.forEach(reduceData_, that);
  this.isSet.parseData = true;
  return this; // for chaining;
}


/**
* consolidate the data arround a defined period

* @param{[Object]} period - {startDate, endDate}, the period that will be used to cover the calculation
* @param{[String]} aggregateMethod, how does the values on a same date are aggregated: [sum, average, median] default: sum
* @param{[Boolean]} slice, are we working on a slice of the whole period or is it the global timerange?
* @create{Object} heatMapDescriptor, {monthIdx:[3,6,9,...], out:[week[{day, value, color, top, bottom, right, left},],], minMax:{minVal, maxVal, minDate, maxDate}}
* @return{Object} this, for chaining purpose
**/
HeatMap.prototype.consolidate = function(period, aggregateMethod, slice) {
  if(!this.isSet.parseData){
    this.parseData();
  }
  this.timelapse = [];
  this.timelapseBounds = slice == true ? this.timelapseBounds : {}; 
  
  // https://gist.github.com/caseyjustus/1166258
  function median_(values) { // return the median of the results
    values.sort( function(a,b) {return a - b;} );
    var half = Math.floor(values.length/2);
    if(values.length % 2){
      return values[half];
    }
    return (values[half-1] + values[half]) / 2.0;
  }
  
  period = period || this.period || {};
  
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
    if(slice){
      var tempEndMonth = this.timelapseBounds.maxDate.getMonth();
      var tempEnd = new Date(this.timelapseBounds.maxDate);
      tempEnd.setMonth(tempEndMonth+1);
      tempEnd.setDate(1);
      tempEnd.setHours(0);
      tempEnd.setMinutes(0);
      tempEnd.setMilliseconds(0);
      this.timelapseBounds.maxDate = tempEnd; 
    }
  }
  
  this.timelapseBounds.minDate.setDate(1);
  
  if(slice != true){
    this.timelapseBounds.minVal = undefined;
    this.timelapseBounds.maxVal = undefined;
  }
  
  aggregateMethod = this.aggregatedMethod || aggregateMethod || 'sum';
  this.aggregatedMethod = aggregateMethod;
  
  var dd;  // displayed date
  var pd = new Date(this.timelapseBounds.minDate); // parsed Date
  
  var ehSource = [] // equalization matrix (raw indexed values to equalize) [idx, value]
  
  do{ // parse every day during the timelapse
    dd = pd.getFullYear()+('0'+(pd.getMonth()+1)).slice(-2)+('0'+pd.getDate()).slice(-2); // displayed day YYYYMMDD
    var raw = this.data[dd];
    if(raw == undefined || raw.values.length == 0){
      this.timelapse.push([new Date(pd),dd,null,null]); // TODO do not set pd as Date it's already the case
    }
    else{
      if(this.period != undefined){ // if data are out of the perimeter do not render them
        if(this.period.startDate != undefined && pd < this.period.startDate ||
           this.period.endDate != undefined && pd > this.period.endDate){
          this.timelapse.push([new Date(pd),dd,null,null]);
          pd.setDate(pd.getDate()+1);
          continue;
        }
      }
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
      
      if(slice != true){ // why? I mean if something lower/upper is found it should be applyed no? ==> NO slice is a local arg this is a control statment (useless unless some bullshit is done)
        if (this.timelapseBounds.minVal == undefined || this.timelapseBounds.minVal > value) {
          this.timelapseBounds.minVal = value;
        }
        if (this.timelapseBounds.maxVal == undefined || this.timelapseBounds.maxVal < value) {
          this.timelapseBounds.maxVal = value;
        }
        ehSource.push([this.timelapse.length-1, value]); 
      }
    }
    pd.setDate(pd.getDate()+1);
  }
  while(pd < this.timelapseBounds.maxDate);
  
  if(slice != true && this.equalizeHistogram == true){
    this.cdfe = buildCdfe(ehSource.map(function(ehs){return ehs[1]}), this.timelapseBounds.minVal, this.timelapseBounds.maxVal);
  }
  if(this.equalizeHistogram == true){
    var that = this;
    this.timelapse = this.timelapse.map(function(val){
      if (val[2] != null){
        val.push(this.cdfe.get(val[2]));
      }
      return val;
    }, that);
  }

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
  * it need to be improved (copy the borders on adjacent cells)
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
  var dataMatrix = []; // the values to display
  var monthLabelMatrix = []; // label for the month (txt, row, column)
  
  var weekColorMatrix = []; // TEMP 7 colors line data for each parsed week
  var weekDataMatrix = []; // TEMP 7 data line for each parsed week
  
  var weekSum = null; // TEMP data to be used with weekSumMatrix
  var monthLabelWeekVal = ''; // can be '' or month to display
  var idx = 0;
  var wdCorrection = this.locales[this.locale].weekStart; // change the start of the day from sunday to monday (F..... murican)
  
  var start = dateArray[0]; // day where we start working
  
  var plus7 = new Date(start[0]); // week +1 from start (used for border to next month
  plus7.setDate(start[0].getDate()+7);
  var wd = start[0].getDay(); // day of the week M, T, W, ....
  
  if(wd-wdCorrection == -1){
    wd = 7 // we need to find something better than that when I'm going to implement weekstart on saturday...
  }
  while(weekColorMatrix.length < wd-wdCorrection){ // init the array with blank and border ==> what if start is different than 1 day of month?
    weekColorMatrix.push(''); // add blank cell at start
    weekDataMatrix.push('');
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
      if(this.equalizeHistogram != true){
        weekColorMatrix.push(this.getGradientColor_(dayVals[2], this.timelapseBounds.minVal, this.timelapseBounds.maxVal, this.gradients[this.gradient])); // TODO include change palette color
      }
      else{
        weekColorMatrix.push(this.getGradientColor_(dayVals[4], 0, 255, this.gradients[this.gradient]));
      }
      weekDataMatrix.push(dayVals[2]);
      if(dayVals[3] != null && dayVals[3] != undefined && dayVals[3] != ""){
        labelledCells.push({label:dayVals[3], column:wd-wdCorrection, row: colorMatrix.length});
      }
    }
    else{ // there is no data just add a GREY cell
      weekColorMatrix.push(this.noValueColorCode);
      weekDataMatrix.push('no data');
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
      dataMatrix.push(weekDataMatrix);
      weekSum = null;
      weekColorMatrix = [];
      weekDataMatrix = [];
    }
  }
  while(dateArray.length > 0);
  
  // properly finish the array
  while(weekColorMatrix.length > 0 ){ // add the final blanks if color Matrix is not full
    weekColorMatrix.push(''); // add blank cell at start
    weekDataMatrix.push('');
    if(weekColorMatrix.length == 7){
      colorMatrix.push(weekColorMatrix);
      dataMatrix.push(weekDataMatrix);
      weekSumMatrix.push(weekSum);
      weekColorMatrix = [];
      weekDataMatrix = [];
    }
  }
  
  this.constructionMaterial = {
    weekDayMatrix: weekDayMatrix, 
    weekSumMatrix:weekSumMatrix, 
    borderedRanges:borderedRanges, 
    labelledCells:labelledCells , 
    colorMatrix:colorMatrix, 
    dataMatrix:dataMatrix,
    monthLabelMatrix:monthLabelMatrix};
  return this;
}


/**
* render tha map in a spreadsheet
* @param{[String]} eref, the full reference where to render the map (SheetName!A1) OR the SheetName OR the starting Column
* @param{[String]} erow, if ecol == undefined ==> the Column where to render the map ELSE the Row
* @param{[String]} ecol, the Column where to render the map
**/
HeatMap.prototype.render = function (eref, erow, ecol){
  var that = this;
  if(eref != undefined){
    this.setOutputCoordinates(eref, erow, ecol);
  }
  else if(this.outCorr == undefined){
    this.setOutputCoordinates();
  } 
  // retrieve data from setOutputCoordinates
  var row = this.outCorr.row;
  var column = this.outCorr.column;
  var sheet = this.outCorr.sheet;
  
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
  
  /**
  * display the month to a given cell
  * @param{Array} cellRef, [label, row, column]
  * @this{GoogleObject} Sheet  
  **/
  function setMonthLabel_(cellRef){
    this.getRange(cellRef[1]+row, cellRef[2]+column - 2).setValue(cellRef[0]);
  }
  
  // check that the data are ready for the render operation
  if(!this.isSet.timelapse){
    this.consolidate(); 
  }
  if(!this.isSet.constructionMaterial){
    this.buildRenderingDataForBlock();
  }
  
  if(this.splits  == 0){
    this.isSet = {timelapse: false, constructionMaterial: false};
  }
  
  
  // material needed to determine if the graph need to be splitted in differents truncs
  var startMonth = this.timelapseBounds.minDate.getMonth();
  var endMonth = this.timelapseBounds.maxDate.getMonth();
  var endPeriodMonth;
  var startDate = new Date(this.timelapseBounds.minDate);
  var endDate = new Date(this.timelapseBounds.maxDate);
  var shift = 0;
  var pusher = 7;
  pusher += this.draw.monthLabel == true ? 2 : 0;
  pusher += this.draw.weekSum == true ? 2 : 0;
  pusher++;
  
  var initRow = row;
  var initColumn = column;
  
  do{ // loop to draw the shit (as many as truncs)
    
    column = initColumn + shift * (pusher);
    row = initRow;
    
    if(this.splits > 0){      // yauritilpasunprobleme? averifier!
      endPeriodMonth = this.splits + (startMonth - startMonth%this.splits) - 1; //startMonth + (12 / this.splits) -1 - (startMonth %( 12 / this.splits));
      startDate.setMonth(startMonth);
      endDate = new Date(startDate);
      endDate.setMonth(endPeriodMonth);
      var period = {startDate: startDate, endDate: endDate};
      this.consolidate(period, undefined, true); 
      this.buildRenderingDataForBlock();
    }
    else{
      endPeriodMonth = endMonth;
    }
    
    // #####################################################
    
    // clean area before doing anything
    sheet.getRange(row, column, this.constructionMaterial.colorMatrix.length+7, 11).clear().clearNote();
    
    if(this.draw.background == true){
      var numRows = this.draw.weekDaySum == true ? 2 : 0;
      numRows += this.draw.weekDay == true ? 2 : 0;
      var numColumns = 7;
      numColumns += this.draw.monthLabel == true? 2 : 0;
      numColumns += this.draw.weekSum == true ? 2 : 0;
      
      if(numRows > 0){
        //Logger.log('head: '+row+' | '+column+' | '+numRows+' | '+numColumns);
        sheet.getRange(row, column, numRows, numColumns).setBorder(true, true, false, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
      }
      
      var cml = this.constructionMaterial.colorMatrix.length;
      cml += this.draw.gradient == true ? 3 : 0;
      var colPusher = 0;
      if(this.draw.monthLabel == true){
        //Logger.log('monthLab: '+row+' | '+column+' | '+numRows+' | '+numColumns);
        sheet.getRange(row + numRows, column, cml, 2).setBorder(true, true, true, false, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
        colPusher += 2;
      }
      colPusher += 7;
      if(this.draw.weekSum == true){
        //Logger.log('weekSum: '+row+' | '+column+' | '+numRows+' | '+numColumns);
        sheet.getRange(row + numRows, column+colPusher, cml, 2).setBorder(true, false, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
      }  
      
      colPusher += -7;
      numRows += this.constructionMaterial.colorMatrix.length;
      if(this.draw.gradient == true){
        //Logger.log('grad: '+row+' | '+column+' | '+numRows+' | '+numColumns);
        sheet.getRange(row + numRows, column+ colPusher, 3, 7).setBorder(false, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
        //numRows += 3;
      }
    }
    
    if(this.draw.monthLabel == true){
      column +=2;
    }
    
    
    // set sparks by day type, start at position 0 | 2
    if(this.draw.weekDaySum == true){
      var daySparks = this.constructionMaterial.weekDayMatrix.map(setDaySparks_, that);
      var lastDays = daySparks.splice(0,this.locales[this.locale].weekStart);
      daySparks = daySparks.concat(lastDays);
      sheet.getRange(row, column, 1, 7).setFormulas([daySparks]);
      sheet.setRowHeight(row, this.height*1.2);
      sheet.setRowHeight(row+1, this.height*0.2);
      row +=2;
    }
    
    
    // set upper days label, start at position 2 | 2
    if(this.draw.weekDay == true){
      var weekDays = this.locales[this.locale].days.map(
        function(day){
          return day.slice(0,1);
        }
      );
      sheet.getRange(row, column, 1, 7).setValues([weekDays]).setFontSize(this.height*0.65);
      sheet.setRowHeight(row, this.height);
      sheet.setRowHeight(row + 1, this.height*0.2);
      row +=2;
    }
    
    
    // set months labels, start at position 4 | 0
    if(this.draw.monthLabel == true){
      this.constructionMaterial.monthLabelMatrix.forEach(setMonthLabel_, sheet);
      sheet.getRange(row, column - 2, this.constructionMaterial.colorMatrix.length, 1).setFontSize(this.height*0.55);
      sheet.setColumnWidth(column - 2, this.width*1.6); // set width of month label column
      sheet.setColumnWidth(column - 1, this.width*0.4); // separation month label | map
    }
    
    // set comment labels, start at position 4 | 2
    if(this.draw.labels == true && this.labelOn == true){
      var labels = this.constructionMaterial.labelledCells;
      labels.forEach(function (lblObj){
        sheet.getRange(lblObj.row+ row, lblObj.column + column).setNote(lblObj.label);
      });
    }
    // set colors, start at position 4 | 2
    sheet.getRange(row, column, this.constructionMaterial.colorMatrix.length, 7).setBackgrounds(this.constructionMaterial.colorMatrix);
    if(this.draw.values){
      sheet.getRange(row, column, this.constructionMaterial.colorMatrix.length, 7).setValues(this.constructionMaterial.dataMatrix).
      setFontColors(this.constructionMaterial.colorMatrix).
      setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).
      setFontSize(6);
    }
    
    // set borders 4 | 2  
    for(var brow in this.constructionMaterial.borderedRanges){
      for(var bcol in this.constructionMaterial.borderedRanges[brow]){
        setBorders_(sheet, Number(brow)+row, Number(bcol)+column, this.constructionMaterial.borderedRanges[brow][bcol]);   
      }
    }
    // set columns width of map
    for(var i = 0 ;i<7; i++) {
      sheet.setColumnWidth(i+column, this.width);
    }  
    // set rows height of map
    for( var j = 0; j < this.constructionMaterial.colorMatrix.length; j++){ 
      sheet.setRowHeight(j+row, this.height);
    }
    
    // set color gradient legend, start at position x+6 | 2
    if(this.draw.gradient == true){
      var grad = this.gradients[this.gradient];
      var colors = [];
      colors.push(this.getGradientColor_(0, 0, 1, grad));
      for(var i = 1; i < 6; i++){
        var valmin = i/7;
        colors.push(this.getGradientColor_(valmin, 0, 1, grad));
      }
      colors.push(this.getGradientColor_(1, 0, 1, grad));
      sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 1, column, 1,7).setBackgrounds([colors]);
      // set row height for gradient
      sheet.setRowHeight(this.constructionMaterial.colorMatrix.length + row + 0, this.height*0.4);
      sheet.setRowHeight(this.constructionMaterial.colorMatrix.length + row + 1, this.height*0.4);
      // set legend value
      sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 2, column, 1,3).merge().setHorizontalAlignment("left");
      sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 2, column).setValue(this.timelapseBounds.minVal).setFontSize(this.height*0.55);
      sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 2, column + 4, 1,3).merge().setHorizontalAlignment("normal");
      sheet.getRange(this.constructionMaterial.colorMatrix.length + row + 2, column + 4).setValue(this.timelapseBounds.maxVal).setFontSize(this.height*0.55);
      sheet.setRowHeight(this.constructionMaterial.colorMatrix.length + row + 2, this.height*0.6);
    }
    
    // set weeks sparks, start at position 4 | 10
    if(this.draw.weekSum == true){
      var sparks = this.constructionMaterial.weekSumMatrix.map(setWeekSumSparks_, that);
      sheet.getRange(row, column+8, sparks.length, 1).setFormulas(sparks);
      sheet.setColumnWidth(column+8, this.width*2.2); // set width Spark week
      sheet.setColumnWidth(column+7, this.width*0.4); // separation map | spark week
    }    
    // #####################################################
    
    startMonth = endPeriodMonth + 1;
    shift++;
    this.draw.gradient = false; // do not draw additionnal legend ==> this break the output
  }
  while(startMonth <= endMonth);
  
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
  if(val == 0 || val == undefined || isNaN(val)){ // Zero OR not recognized will return the lowes value
    return gradient[0];
  }
  
  var interval = gradient.length - 1; // numbers of working space beetwen two color stops
  var spaces = 1/interval; // weigh of an interval
  var chevron = 0; // Index of the interval we have to work on
  while(chevron * spaces < val && chevron * spaces <= 1){ // count on what interval we have to work
    chevron++;
  }
  val = (val - (chevron-1)*spaces)/ spaces; // relative value for a given space (reported to it's interval index)
  var myMax = gradient[chevron]; // up color code
  var myMin = gradient[chevron-1]; // down color code
  
  var r1, g1, b1, r2, g2, b2;
  
  r1 = myMin.substr(1,2);
  g1 = myMin.substr(3,2);
  b1 = myMin.substr(5,2);
  r2 = myMax.substr(1,2);
  g2 = myMax.substr(3,2);
  b2 = myMax.substr(5,2);
  var out = '#'
  +('0'+Math.round(parseInt(r1,16)+((parseInt(r2,16)-parseInt(r1,16))*val)).toString(16)).slice(-2) // red
  +('0'+Math.round(parseInt(g1,16)+((parseInt(g2,16)-parseInt(g1,16))*val)).toString(16)).slice(-2) // green
  +('0'+Math.round(parseInt(b1,16)+((parseInt(b2,16)-parseInt(b1,16))*val)).toString(16)).slice(-2); // blue
  return out;
}



// https://en.wikipedia.org/wiki/Histogram_equalization
// https://stackoverflow.com/questions/12114643/how-can-i-perform-in-browser-contrast-stretching-normalization#  ==> Not working
/**
* Equalizes the histogram of an unsigned 1-channel image with values
*
* @param {Array} src 1-channel source image
* @param {Number} min, minimum value of the range
* @param {Number} max, maximum value of the range
* @return {Array} Destination image
*/
equalizeHistogram = function(src, min, max) {
  
  min = min || 0;
  max = max || 255;
  
  var cdfe = buildCdfe(src, min, max);
  
  // Equalize image:
  var out = []; // new equalized values
  src.forEach(function(val){
    val = ~~(255 * ((val - min) / (max - min)));
    this.push(cdfe[val].hv);
  }, out);
  
  return out;
}

/**
* build CDFE, cumulative distribution function equalized
*
**/
buildCdfe = function  (src, min, max){
  min = min || 0;
  max = max || 255;
  
  // Compute histogram  
  var v = {}; // object with rounded vals
  src.forEach(function(val){
    val = ~~(255 * ((val - min) / (max - min)));
    this[val] = this[val] || 0;
    this[val]++;
  }, v);
  
  var d = []; // distribution or founded vals
  for(var i in v){
    d.push([i, v[i]]);
  }
  d.sort(function(a,b){return a[0]-b[0];});
  
  // Compute integral histogram:
  var cdf = []; // cumulative distribution of rounded vals
  d.reduce(function(acc, val, idx){acc+=val[1]; cdf.push([val[0], acc]);return acc;},0);
  
  var cdfe = {min:min, max:max, get:function(val){val = ~~(255 * ((val - this.min) / (this.max - this.min))); return this[val].hv;}}; // cumulative distribution function equalized {roundedValue: {cumulativeDistributionValue(cdfv), equalization(hv)}}
  
  var cdfmin = d[0][1]; // plus petite valeur de la distribution
  
  var h = function(cdfv){ // used to get the cdfe for each value
    return ~~(((cdfv - cdfmin) / (src.length - cdfmin)) * (256-1));
  };
  
  cdf.forEach(function(cudi){this[cudi[0]]={cdfv:cudi[1], hv:h(cudi[1]) }}, cdfe); // build cdfe
  
  return cdfe;
}

/* ######## METHODS ########## */

/**
* set locale: change the languages and options according to the selected language
* @param{String} locale, fr | en-us
**/
HeatMap.prototype.setLocale = function (locale){
  if(this.locales[locale] == undefined){
    throw 'you tried to set "'+locale+'" as the new locale, but it\' not yet implemented in the system. Please select one of: '+Object.keys(this.locales).join(', ');
  }
  this.locale = locale;
  return this; // for parsing
};

/**
* set Gradient: change gradient color for an alternative option NOTE: hexCode with percent val not yet iplemented
* @param{String} gradientName, the gradient name you want to set
* @param{[Array]} gradientCodes, OPTIONAL param: The hexacolor codes representing the gradient from the lowest value to the highest {[#Code, ]} || {[{hexcode:#code, percent:%Val},]}
**/
HeatMap.prototype.setGradient = function (gradientName, gradientCodes){
  if(gradientCodes == undefined || typeof gradientCodes != 'Array' || gradientCodes.length < 2){ // no valid gradientCodes
    if(Object.keys(this.gradients).indexOf(gradientName) < 0){
      throw 'The gradient name '+ gradientName +' is unknow, please select in the list: '+Object.keys(this.gradients).join(', ');
    }
  } else{
    this.gradients[gradientName] = gradientCodes;
  }
  this.gradient = gradientName;
  return this; // for parsing
}

/**
* set graph size allow you to create graph larger or smaller
* @param{Number} factor, multiplication factor to change the size
**/
HeatMap.prototype.setGraphsize = function (factor){
  
  this.width = this.width*factor;
  this.height = this.height*factor;
  return this; // for parsing
}


/**
* set data source
* @param{String} eref,reference a range with notation SheetName!A1:B2 || The Column to start with (A, B , C...)
* @param{String=} erow, last Column or first Row if there is a following argument, or first Column if there is 4 args
* @param{String=} 
IT'S A FUCKING MESS
**/
HeatMap.prototype.setSource = function (eref, erow, ecol, enumRow, enumCol){
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
    range = SpreadsheetApp.getActive().getRange(eref+"!"+startCol+erow+":"+endCol+enumRow);
  }
  
  if(range.getNumColumns()>2){
    this.labelOn = true;
  }
  this.range = range;
  return this;
}


/**
* set output coordinates (destination where to write teh data
* @param{[String]} eref, the full reference where to render the map (SheetName!A1) OR the SheetName OR the starting Column
* @param{[String]} erow, if ecol == undefined ==> the Column where to render the map ELSE the Row
* @param{[String]} ecol, the Column where to render the map
**/
HeatMap.prototype.setOutputCoordinates = function(eref, erow, ecol){
  // check the parameter to determine where to drop the data
  
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
  this.outCorr = {column:column, row: row, sheet:sheet};
  return this;
}

/**
* set splits, split into new column every X months
* @param{Number} months to group by
**/
HeatMap.prototype.setSplit = function (splt){
  if (splt != 0 && splt != 1 && splt != 2 && splt != 3 && splt != 4 && splt != 6){
    throw "we do not know how to handle " + splt + " splits yet, please select an other value for the \"split\" parameter.";
  }
  this.splits = splt;
  return this;
}

/**
* set aggregation method, how does values from the same day
* @param{String} aggregationMethod, can by sum, average, median. by defect sum method is applyed
**/
HeatMap.prototype.setAggregationMethod = function (aggregationMethod){
  aggregationMethod = aggregationMethod.toLocaleLowerCase(); 
  if(aggregationMethod != "sum" && aggregationMethod != "average" && aggregationMethod != "median"){
    throw 'the aggregation method "' + aggregationMethod + '" is not unknown, please select an other method or ignore the parameter';
  }
  this.aggregatedMethod = aggregationMethod;
  return this;
}

/**
* display comment labels, as a third column with text that can be displayed as comments
* @param{Boolean} commentLabelDisplay, do we have to display the label yes or no
**/
HeatMap.prototype.displayCommentLabel = function (lblOn){
  this.draw.labels = lblOn;
  return this;
}

/**
* display month label, display/hide a column with the month name (3 letters)
* @param{Boolean} monthLabelDisplay, default value TRUE
**/
HeatMap.prototype.displayMonthLabel = function (monthLabelDisplay){
  this.draw.monthLabel = monthLabelDisplay;
  return this;
}

/**
* display gradient legend at the bottom of the graph
* @param{Boolean} gradientLegendDisplay, default value TRUE
**/
HeatMap.prototype.displayGradientLegend = function (gradientLegendDisplay){
  this.draw.gradient = gradientLegendDisplay;
  return this;
}

/**
* draw background, draw a white background instead of letting everithing transparent
* @param{Boolean} backgroundDisplay, default values: FALSE
**/
HeatMap.prototype.drawBackground = function(backgroundDisplay){
  this.draw.background = backgroundDisplay;
  return this;
}

/**
* set period: define min and max dates that will be rendered
* @param{Object} period, {startDate: [Date], endDate:[Date]}
**/
HeatMap.prototype.setPeriod = function(period){
  this.period = period;
}


/**
* display value: display the values for each cell the same color as the background
* @param{Boolean} displayVal, default FALSE
**/
HeatMap.prototype.displayValues = function(displayVal){
  this.draw.values = displayVal;
}

/**
* histogram equalization, equalize the value to improve teh repartition
* @param{Boolean} histogramEqualization, default FALSE
**/
HeatMap.prototype.equalize = function(histogramEqualization){
  this.equalizeHistogram = histogramEqualization;
}

/**
* set output display options (default is all options on)
* @param{Object} options, 
**/
HeatMap.prototype.Params = function (options){
  for(var i in options){
    switch(i){
      case "setSource":
        if(options[i].eref == undefined){
          this.setSource(options[i]);
        } else {
          this.setSource(options[i].eref, options[i].erow, options[i].ecol, options[i].enumRow, options[i].enumCol);
        }
        break;
      case "setOutputCoordinates":
        if(options[i].eref == undefined){
          this.setOutputCoordinates(options[i]);
        } else{
          this.setOutputCoordinates(options[i].eref, options[i].erow, options[i].ecol);
        }
        break;
      case "displayValues":
        this.displayValues(options[i]);
        break;
      case "displayGradientLegend":
        this.displayGradientLegend(options[i]);
        break;
      case "equalize":
        this.equalize(options[i]);
        break;
      case "drawBackground":
        this.drawBackground(options[i]);
        break;
      case "displayMonthLabel":
        this.displayMonthLabel(options[i]);
        break;
      case "displayCommentLabel":
        this.displayCommentLabel(options[i]);
        break;
      case "setAggregationMethod":
        this.setAggregationMethod(options[i]);
        break;
      case "setSplit":
        this.setSplit(options[i]);
        break;
      case "setLocale":
        this.setLocale(options[i]);
        break;
      case "setGradient":
        this.setGradient(options[i]);
        break;
      case "setGraphsize":
        this.setGraphsize(options[i]);
        break;
      case "setPeriod":
        this.setPeriod(options[i]);
        break;
    }
  }
}

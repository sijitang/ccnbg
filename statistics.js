/*
 *  Copyright 2014.9.18  Jing Tang  tangjing725@ccnbg.org
 *  此程序是利用google app script api 和 google drive service 实现的营会分房程序。
 *  在通知原作者的前提下，你可以使用传播以及修改本程序，但禁止用于任何商业用途。
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation; either version 2 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *  General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>. 
 *
 *  For more information on using the Spreadsheet API, see
 *  https://developers.google.com/apps-script/service_spreadsheet
 *
 */

//test
function showStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number")
     .addRow(["基督徒", 231])
     .addRow(["慕道友", 151])
     .addRow(["小孩", 70])
     .build();


   var chartBuilder = Charts.newPieChart()
       .setTitle('2014南德福音营')
       .setDimensions(600, 500)
       .set3D()
       .setDataTable(data)
   
   var chart = chartBuilder.build();
   var app = UiApp.createApplication()
   var scroll = app.createScrollPanel().setPixelSize(500, 500);
   scroll.add(chart);
   app.add(scroll);
   app.setHeight(1000);
   app.setWidth(1000);
   app.setTitle("2014南德福音营");
   ss.show(app);
  
}

// use statistics as web services
function doGet() {
   // Get data from a spreadsheet.

  
   var app = UiApp.createApplication()

   var data1 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number")
     .addRow(["基督徒", 231])
     .addRow(["慕道友", 151])
     .addRow(["小孩", 70])
     .build();


   var chartBuilder1 = Charts.newPieChart()
       .setTitle('2014南德福音营报名统计')
       .setDimensions(550, 500)
       .set3D()
       .setDataTable(data1)
   
   var data2 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number")
     .addRow(["住宿", 452 - 94])
     .addRow(["日营", 94])
     .build();


   var chartBuilder2 = Charts.newPieChart()
       .setTitle('住宿/日营 比例')
       .setDimensions(550, 500)
       .set3D()
       .setDataTable(data2)
   
   var data3 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number")
     .addRow(["报道的人", 452 - 23])
     .addRow(["临时取消的人", 23])
     .build();
  
  var chartBuilder3 = Charts.newPieChart()
       .setTitle('临时取消报名与实际报道人数的比例')
       .setDimensions(550, 500)
       .set3D()
       .setDataTable(data3)
  
   var data4 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number")
     .addRow(["取消报名的基督徒", 8])
     .addRow(["取消报名的慕道友", 12])
     .addRow(["取消报名的孩子", 3])
     .build();
  
  var chartBuilder4 = Charts.newPieChart()
       .setTitle('临时取消人员类型')
       .setDimensions(550, 500)
       .set3D()
       .setDataTable(data4)
  
  var data5 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number");
  
  var data6 = Charts.newDataTable()
     .addColumn(Charts.ColumnType.STRING, "type")
     .addColumn(Charts.ColumnType.NUMBER, "number");
  var data7 = Charts.newDataTable()
       .addColumn(Charts.ColumnType.STRING, "教会")
       .addColumn(Charts.ColumnType.NUMBER, "基督徒")
       .addColumn(Charts.ColumnType.NUMBER, "慕道友")
       .addColumn(Charts.ColumnType.NUMBER, "孩子")


  
  var registerSheet = SpreadsheetApp.openById('1TDqx4qx3S_QmK41XG1htZiCBvaPq69UoQAo12gUzyCs').getActiveSheet();
  var regData = registerSheet.getDataRange().getValues();
  for (var i = 1; i <= regData.length - 2; i++) {
    var row = regData[i];
    var fs = row[0];
    var chrs = row[1];
    var normal = row[2];
    var kids = row[3];
    var sumfs = row[4];
    var cancelfs = row[8];
    data5.addRow([fs, sumfs]);
    if(cancelfs!=''){
      data6.addRow([fs, sumfs]);
    }
    data7.addRow([fs, chrs, normal, kids]);
    Logger.log(row);
  }
  var chartBuilder5 = Charts.newPieChart()
       .setTitle('各教会报名人数')
       .setDimensions(550, 820)
       .set3D()
       .setDataTable(data5.build())
  
 var chartBuilder6 = Charts.newPieChart()
       .setTitle('各教会取消报名人数')
       .setDimensions(550, 820)
       .set3D()
       .setDataTable(data6.build())
 
var chartBuilder7 = Charts.newColumnChart()
       .setTitle('各教会/团契 报名概要')
       .setXAxisTitle('教会')
       .setYAxisTitle('报名人数')
       .setDimensions(1200, 600)
       .setDataTable(data7.build())

  
   var chart1 = chartBuilder1.build();
   var chart2 = chartBuilder2.build();
  
   var chart3 = chartBuilder3.build();
   var chart4 = chartBuilder4.build();
   
   var chart5 = chartBuilder5.build();
   var chart6 = chartBuilder6.build();
  
   var chart7 = chartBuilder7.build();

  
   var panel1 = app.createHorizontalPanel();
   panel1.add(chart1)
   panel1.add(chart2);
   
   var panel2 = app.createHorizontalPanel();
   panel2.add(chart3);
   panel2.add(chart4);
  
   var panel3 = app.createHorizontalPanel();
   panel3.add(chart5);
   panel3.add(chart6);

   var panel4 = app.createHorizontalPanel();
   panel4.add(chart7);
  
  app.add(panel1);
  app.add(panel2);
  app.add(panel3);
  app.add(panel4);

  return app;
 }


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "show Statitics",
    functionName : "showStatistics"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
};

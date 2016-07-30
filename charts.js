// The initialize function must be run each time a new page is loaded

var chartsData = [];
var chartNames = [];
var chartRanges = [];
var seriesNames = [];

function playChart(){
    var myChartIndex = parseInt(document.getElementById("selectItem").value);
    var mySeriesIndex = parseInt(document.getElementById("selectSeries").value);
    console.log(chartsData[myChartIndex][mySeriesIndex]);
    var a = dtm.array(chartsData[myChartIndex][mySeriesIndex]);
    a.range(60, 96, chartRanges[myChartIndex][0], chartRanges[myChartIndex][1]); 
    dtm.synth().play().nn([60,96]).dur(1);
    dtm.synth().play().nn(a).dur(3).offset(2);
}

function selectChart(event){
    var chartIndex = parseInt(event.target.value);
    var seriesPicker = document.getElementById("selectSeries");
    document.getElementById("sonify").disabled = true;
    seriesPicker.innerHTML = "<option selected='selected'>Select data series to listen</option>";
    for(var seriesIndex = 0;seriesIndex < seriesNames[chartIndex].length;seriesIndex++){
        seriesPicker.innerHTML += "<option value='" + String(seriesIndex) + "'>" + seriesNames[chartIndex][seriesIndex] + "</option>";
    }
    seriesPicker.disabled = false;
    
}

function selectSeries(event){
    var seriesIndex = parseInt(event.target.value);
    console.log(seriesIndex);
    document.getElementById("sonify").disabled = false;
}

function clickFunction(event){
    
}
    // retrieve charts
function refreshCharts(){    
    console.log("refreshing chartS");
    Excel.run(function (ctx) {  
        var myChartCollection = ctx.workbook.worksheets.getActiveWorksheet().charts;
        var mySeriesCollections = [];
        var myPointCollections = [];

        //myChartCollection.load("series/items/points/value");
        //myChartCollection.load("axes/valueAxis/maximum");
        myChartCollection.load("name","axes");


        return ctx.sync().then(function(){

            for(var chartIndex = 0;chartIndex < myChartCollection.items.length;chartIndex++){

                chartNames.push(myChartCollection.items[chartIndex].name);


                //chartRanges.push([myChartCollection.items[chartIndex].axes.valueAxis.minimum,myChartCollection.items[chartIndex].axes.valueAxis.maximum]);
                mySeriesCollections.push(myChartCollection.items[chartIndex].series);
                mySeriesCollections[chartIndex].load('name');  

                myChartCollection.items[chartIndex].axes.load('valueAxis/maximum');    
                myChartCollection.items[chartIndex].axes.load('valueAxis/minimum')

            }
        }).then(ctx.sync).then(function(){
            for(var chartIndex = 0;chartIndex < mySeriesCollections.length;chartIndex++){
                seriesNames.push([]);
                myPointCollections.push([]);

                //console.log("number of series is " + mySeriesCollections[chartIndex].items.length);
                chartRanges.push([myChartCollection.items[chartIndex].axes.valueAxis.minimum, myChartCollection.items[chartIndex].axes.valueAxis.maximum]);

                
                for(var seriesIndex = 0;seriesIndex < mySeriesCollections[chartIndex].items.length;seriesIndex++){
                    seriesNames[chartIndex].push(String(seriesIndex + 1) + ". " + mySeriesCollections[chartIndex].items[seriesIndex].name);
                    myPointCollections[chartIndex].push(mySeriesCollections[chartIndex].items[seriesIndex].points);
                    myPointCollections[chartIndex][seriesIndex].load('value');
                }
            }
            console.log("ready to issue");

        }).then(ctx.sync).then(function(){
            console.log("afterreception");
            document.getElementById("selectItem").innerHTML = "<option selected='selected' disabled='disabled'>Select a chart</option>";
            for(var chartIndex = 0;chartIndex < mySeriesCollections.length;chartIndex++){
               // console.log("max is " + myChartCollection.items[chartIndex].axes.valueAxis.maximum);
                chartsData.push([]);
                for(var seriesIndex = 0;seriesIndex < myPointCollections[chartIndex].length;seriesIndex++){
                    chartsData[chartIndex].push([]);
                    if(myPointCollections[chartIndex][seriesIndex]){
                        // then the series is not undefined
                        for(var pointIndex = 0;pointIndex < myPointCollections[chartIndex][seriesIndex].items.length;pointIndex++){
                            chartsData[chartIndex][seriesIndex].push(myPointCollections[chartIndex][seriesIndex].items[pointIndex].value);
                        }
                    }
                }
                console.log("adding a chart");
                document.getElementById("selectItem").innerHTML += "<option value='" + String(chartIndex) + "'>" + chartNames[chartIndex] + "</option>";                
            }

            console.log(chartsData);
        });       
    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

(function () {
    Office.initialize = function (reason) {

        if(getParameterByName('action') == 'charts'){
            console.log("it's a charts!");
            console.log("about to assign onchange");
            document.getElementById("selectItem").onchange = selectChart;
            document.getElementById("selectSeries").onchange = selectSeries;
            document.getElementById("sonify").onclick = playChart;
            refreshCharts();     

        }
        else{
            console.log("it's not a task pane");

        }
        
    };
})();


//Notice function needs to be in global namespace
/*
function insertAStock(event){

}

function insertTable(event) {
    // user clicks "Insert Stocks Table"
    
    Office.context.document.bindings.addFromPromptAsync(
        Office.BindingType.Table,
        {
            id: "stocksTable",
            promptText: "Select a cell to insert the Stocks table"
        },
        function(asyncResult){
            
        }
    );
    */

    //writing
/*
    Office.context.document.settings.set('a', '42');
    Office.context.document.settings.saveAsync(function (asyncResult) {});

    //Office.context.document.setSelectedDataAsync("wrote 42");
    
    Office.context.document.setSelectedDataAsync(i);
    i++;
}


function insertTable2(event) {

    Excel.run(function (ctx) {
        var selectedRange = ctx.workbook.getSelectedRange();
        selectedRange.load('address');
        selectedRange.values = "Hello!";
        return ctx.sync().then(function () {
            Office.context.document.setSelectedDataAsync(selectedRange.address);
        });
    }).catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
*/






function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}




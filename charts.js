// The initialize function must be run each time a new page is loaded

var chartsData = [];
var chartNames = [];
var seriesNames = [];

function playChart(myChartIndex){
    console.log(chartsData[parseInt(myChartIndex)][0]);
    var a = dtm.array(chartsData[parseInt(myChartIndex)][0]);
    a.range([60, 97]);
    dtm.synth().play().nn(a).dur(3);
}

function clickFunction(event){
    // retrieve charts
    
    Excel.run(function (ctx) {  
        var myChartCollection = ctx.workbook.worksheets.getActiveWorksheet().charts;
        var mySeriesCollections = [];
        var myPointCollections = [];

        myChartCollection.load('name');

        return ctx.sync().then(function(){
            for(var chartIndex = 0;chartIndex < myChartCollection.items.length;chartIndex++){
                chartNames.push(myChartCollection.items[chartIndex].name);
                mySeriesCollections.push(myChartCollection.items[chartIndex].series);
                mySeriesCollections[chartIndex].load('name');                
            }
        }).then(ctx.sync).then(function(){
            for(var chartIndex = 0;chartIndex < mySeriesCollections.length;chartIndex++){
                seriesNames.push([]);
                myPointCollections.push([]);
                for(var seriesIndex = 0;seriesIndex < mySeriesCollections[chartIndex].items.length;seriesIndex++){
                    seriesNames[chartIndex].push(mySeriesCollections[chartIndex].items[seriesIndex].name);
                    myPointCollections[chartIndex].push(mySeriesCollections[chartIndex].items[seriesIndex].points);
                    myPointCollections[chartIndex][seriesIndex].load('value');
                }
            }
        }).then(ctx.sync).then(function(){
            for(var chartIndex = 0;chartIndex < mySeriesCollections.length;chartIndex++){
                chartsData.push([]);
                for(var seriesIndex = 0;seriesIndex < myPointCollections.length;seriesIndex++){
                    chartsData[chartIndex].push([]);
                    if(myPointCollections[chartIndex][seriesIndex]){
                        // then the series is not undefined
                        for(var pointIndex = 0;pointIndex < myPointCollections[chartIndex][seriesIndex].items.length;pointIndex++){
                            chartsData[chartIndex][seriesIndex].push(myPointCollections[chartIndex][seriesIndex].items[pointIndex].value);
                        }
                    }
                }
                document.getElementById("players").innerHTML += "<button id='playchart" + String(chartIndex) + "' type=button onclick='playChart(\x22" + String(chartIndex) + "\x22);'>Play " + chartNames[chartIndex] + "</button>";                
            }

            console.log(chartsData);
        });       
    }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

    event.completed();
}

(function () {
    Office.initialize = function (reason) {
        console.log("ran in the chartsjs file");

        document.getElementById("button1").onclick = clickFunction;
        console.log("added the click handler");

        Office.context.document.setSelectedDataAsync("ran in the chartsjs file");
        /* if(getParameterByName('action') == 'taskpane'){
            console.log("it's a task pane");      

        }
        else{
            console.log("it's not a task pane");

        }
        */
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




/*

function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

*/



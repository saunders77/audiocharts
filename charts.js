// The initialize function must be run each time a new page is loaded

var chartsData = [];
var chartNames = [];
var chartRanges = [];
var seriesNames = [];

function playChart(){
    var myChartIndex = parseInt(document.getElementById("selectItem").value);
    console.log("series box value is: " + document.getElementById("selectSeries").value);
    var seriesOptions = document.getElementById("selectSeries").options;
	var mySpeed = parseFloat(document.getElementById("speed").value);
	dtm.synth().play().nn([56, 104]).dur(1);

	for(var i = 0;i < seriesOptions.length;i++){
		if(seriesOptions[i].selected){
			var mySeriesIndex = parseInt(seriesOptions[i].value);
			var b = dtm.array(chartsData[myChartIndex][mySeriesIndex]);

			b.range(56, 104, chartRanges[myChartIndex][0], chartRanges[myChartIndex][1]); 
			if(i % 2 == 0){
				dtm.synth().play().nn(b).dur(3/mySpeed).offset(2);
			}
			else{
				dtm.synth().play().nn(b).wt([-1, -0.75, -0.5, -0.25, 0, 0.25, 0.5, 0.75, 1]).dur(3/mySpeed).offset(2);
			}
			
		}
	}

}

function selectChart(event){
    var chartIndex = parseInt(event.target.value);
    var seriesPicker = document.getElementById("selectSeries");
    document.getElementById("sonify").disabled = true;
    seriesPicker.innerHTML = "";
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

function playSelection(event){
	Office.context.document.getSelectedDataAsync("matrix",function(asyncResult){
		dtm.synth().play().nn([56, 104]).dur(1);
		var minval = asyncResult.value[0][0];
		var maxval = asyncResult.value[0][0];
		for(var m = 0;m < asyncResult.value.length;m++){
			for(var n = 0;n < asyncResult.value[m].length;n++){
				if(asyncResult.value[m][n] > maxval){
					maxval = asyncResult.value[m][n];
				}
				if(asyncResult.value[m][n] < minval){
					minval = asyncResult.value[m][n];
				}
			}
		}
		
		for(var i = 0;i < asyncResult.value.length;i++){
		

			var b = dtm.array(asyncResult.value[i]);

			b.range(56, 104, minval, maxval); 
			if(i % 2 == 0){
				dtm.synth().play().nn(b).dur(3).offset(2);
			}
			else{
				dtm.synth().play().nn(b).wt([-1, -0.75, -0.5, -0.25, 0, 0.25, 0.5, 0.75, 1]).dur(3).offset(2);
			}
			
		
		}
		event.completed();
	});
}

function clickFunction(event){
    
}
    // retrieve charts
function refreshCharts(){
	chartsData = [];
	chartNames = [];
	chartRanges = [];
	seriesNames = [];
	document.getElementById("selectSeries").disabled = true;    
	document.getElementById("selectSeries").innerHTML = "";
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


        }).then(ctx.sync).then(function(){

            document.getElementById("selectItem").innerHTML = "<option selected='selected' disabled='disabled'>Select a chart</option>";
            for(var chartIndex = 0;chartIndex < mySeriesCollections.length;chartIndex++){

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
            document.getElementById("selectItem").onchange = selectChart;
            document.getElementById("selectSeries").onchange = selectSeries;
            document.getElementById("sonify").onclick = playChart;
			document.getElementById("refresh").onclick = refreshCharts;
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




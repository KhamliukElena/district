$(document).ready( function() {
    $.support.cors = true; 
    var workbook = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    var excelIO = new GC.Spread.Excel.IO();  
    function ImportFile() {
        var excelUrl = $("#importUrl").val();  
        var oReq = new XMLHttpRequest();  
        oReq.open('get', excelUrl, true);  
        oReq.responseType = 'blob';  
        oReq.onload = function () {  
            var blob = oReq.response;  
            excelIO.open(blob, LoadSpread, function (message) {  
                console.log(message);  
            });  
        };  
        oReq.send(null);  
    }
    function LoadSpread(json) {  
        jsonData = json;  
        workbook.fromJSON(json);  
        workbook.setActiveSheet("Лист1");  
    } 
    changeData = function() {
        var sheet = workbook.getActiveSheet();
        sheet.setValue(1, 1, "blalb");
        console.log("change");
    }

    $("#importUrl").focusout( function () {
        ImportFile();
        LoadSpread();
        setTimeout(changeData, 3000);

    })
 
})
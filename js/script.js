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
    function ExportFile(fileName) {
        var json = JSON.stringify(workbook.toJSON());  
        excelIO.save(json, function (blob) {  
            saveAs(blob, fileName);  
        }, function (e) {  
            if (e.errorCode === 1) {  
                alert(e.errorMessage);  
            }  
        });  
    }
    changeData = function() {
        var peopleSheet = workbook.getSheet(1);
        peopleSheet.setValue(1, 1, "blalb");
        var districtSheet = workbook.getSheet(2);
        var j = districtSheet.getValue(0,0);
        console.log(j);
    }

    $("#importUrl").focusout( function () {
        ImportFile();
        LoadSpread();
        setTimeout(changeData, 1000);
        $('#ready').prop('disabled', false);
    })

    $("#ready").click( function() {
        fileName = $("#importUrl").val();
        fileName = fileName.replace('./', '');
        ExportFile(fileName);
    })
 
})
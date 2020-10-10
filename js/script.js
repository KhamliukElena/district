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
    createBuildingList = function () {
        var districtSheet = workbook.getSheet(2);
        let buildingList = {};
        return buildingList;
    }
    changeData = function() {
        let buildings = createBuildingList();
        var peopleSheet = workbook.getSheet(1);
        var i = 1;
        var data;
        const toponims = ["вул.", "пров.", "пр.", "м.", "провулок"];
        const build = "буд.";
        while ((data = peopleSheet.getValue(i, 0)) != null) {
            data = data.toLowerCase().replace(/\s/g, '').replace(",,", ",").split(',');
            console.log(data);
            if (data.length == 1) { //if there is only city, w/o street and flat, do not make any marks
                i++;
                continue; 
            } 
            else if (data.length > 1) { //street name parser
                for (let i=0; i< toponims.length; i++) {
                    data[1] = data[1].replace(toponims[i], '');
                }
                if (data.length > 2) { //building number parser
                    console.log(data[2]);
                    data[2] = data[2].replace(build, '').replace('.', '');
                    if (data[2].includes('/')) { //if apt number is specified with /
                        let tmp = data[2].split('/'); 
                        data[2] = tmp[0];
                        data.push(tmp[1]);
                    }
                }
            }
            peopleSheet.setValue(i, 1, data);
            i++;
        }
        $('#ready').prop('disabled', false);
    }

    $("#importUrl").focusout( function () {
        ImportFile();
        LoadSpread();
        setTimeout(changeData, 1000);
    })

    $("#ready").click( function() {
        fileName = $("#importUrl").val();
        fileName = fileName.replace('./', '');
        ExportFile(fileName);
    })
 
})
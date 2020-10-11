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
    isSubBuilding = function(str) {
        let number = [];
        let el = "";
        for (let k=0; k<=str.length; k++) {
            let l = parseInt(str[k]);
            if (isNaN(l)) {
                number.push(el);
                l = str[k];
                el = "";
            }
            el+=l;
        }
        return number;
    }
    buildingNumber = function (str) {
        listNum = [];
        if (str == null) {
            listNum = null;
        }
        else {
            str = str.replace(/\s/g, '').split(',');
            for (let i = 0; i<str.length; i++) {
                if (str[i].includes('-')) {
                    let tmp = str[i].split('-');
                    listNum.push(tmp[0]);
                    number1 = isSubBuilding(tmp[0]);
                    number2 = isSubBuilding(tmp[1]);
                    for (let j=parseInt(number1[0])+1; j<=parseInt(number2[0]); j++) {
                        listNum.push(j.toString());
                    }
                    if (number2.length == 2) {
                        listNum.push(tmp[1]);
                        while (number2[1].toString().charCodeAt(0) > 1040) {
                            let code = number2[1].toString().charCodeAt(0)-1;
                            number2[1] = String.fromCharCode(code);
                            listNum.push(number2[0].concat(number2[1]));
                        }
                    }
                }
                else {
                    listNum.push(str[i]);
                }
            }
        }
        return listNum;
    }
    createBuildingList = function () {
        var districtSheet = workbook.getSheet(2);
        let buildingList = [];
        let street;
        let i = 1;
        while ((street = districtSheet.getValue(i, 0)) != null) {
            var element = {};
            element.street = street.toLowerCase().replace(/\s/g, '');
            element.buildings = buildingNumber(districtSheet.getValue(i,1));
            element.district = districtSheet.getValue(i,2);
            buildingList.push(element);
            i++;
        }
        return buildingList;
    }
    changeData = function() {
        let buildings = createBuildingList(); //create a structure with addresses belonging to a district
        var peopleSheet = workbook.getSheet(1);
        let i = 1;
        let data;
        const toponims = ["вул.", "пров.", "пр.", "м.", "провулок"];
        const build = "буд.";
        while ((data = peopleSheet.getValue(i, 0)) != null) {
            data = data.toLowerCase().replace(/\s/g, '').replace(",,", ",").split(',');
            if (data.length == 1) { //if there is only city, w/o street and flat, do not make any marks
                i++;
                continue; 
            } 
            else if (data.length > 1) { //street name parser
                for (let i=0; i< toponims.length; i++) {
                    data[1] = data[1].replace(toponims[i], '').replace(/\./g, '');
                }
                if (data.length > 2) { //building number parser
                    data[2] = data[2].replace(build, '').replace(/\./g, '').replace(/\s/g, '').toUpperCase();
                    if (data[2].includes('/')) { //if apt number is specified with /
                        data[2] = data[2].split('/')[0]; 
                    }
                }
            }
            let district = [];
            for (let j = 0; j<buildings.length; j++) {
                if (data[1] == buildings[j].street || data[1].endsWith(buildings[j].street)) {
                    if (buildings[j].buildings == null || buildings[j].buildings.indexOf(data[2]) != -1) {
                        district.push(buildings[j].district);
                    }
                    else if (data[2] != undefined && data[2] != null) {
                        let sub = isSubBuilding(data[2]);
                        if (sub.length > 1 && buildings[j].buildings.indexOf(sub[0]) != -1 &&
                        buildings[j].buildings.indexOf((parseInt(sub[0])+1).toString()) != -1) {
                                district.push(buildings[j].district);
                        }
                    }
                }
            }
            if (district.length == 1) {
                peopleSheet.setValue(i, 1, district[0]);
            }
            else if (district.length > 1) {
                let msg = "Несколько доступных округов:"
                for (let j=0; j<district.length; j++) {
                    msg+=' ' + district[i];
                }
                peopleSheet.setValue(i, 1, msg);
            }
            else {
                peopleSheet.setValue(i, 1, "Округ не найден");
            }
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
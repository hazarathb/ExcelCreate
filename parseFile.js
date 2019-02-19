"use strict";
var fs = require('fs');
const xl = require("excel4node");
var xml2js = require('xml2js');
var pattern = /^={7}\s*[A-Za-z0-9._\-]+\s*={7}$/;
var  spattern = /Task:SENSORS/
var logDataArray = {};
var nextArrFlag = true;
var currArrName = "";
var wb = new xl.Workbook({ dateFormat: "m/d/yyyy hh:mm:ss" });
fs.readFile("C:\\Samples\\capture.txt", (err, data) => {

    var array = data.toString().split("\n")
   
    var line;
   
    var fileLen = array.length;
    for (var i = 0; i < fileLen; ++i) {
        line = array[i].trim();
        if (line && pattern.exec(line)) {
            if (nextArrFlag) {
           // console.log(line);
                currArrName = line;
                logDataArray[currArrName] = new Array();
                nextArrFlag = false;
            }else {
                nextArrFlag = true;
            }
        }else if (!nextArrFlag) {
            if(currArrName==="=======SystemLog======="){
                if (spattern.exec(line)) {
                    logDataArray[currArrName].push(line);
                }
            }else{
                
                logDataArray[currArrName].push(line);
            }
        }
   
    }
       
    var sensorData = getSensorInfo(logDataArray);
	
	createtSensorInfo(sensorData,wb);
	//getDaily_QC(logDataArray,wb);
    //getWeekly_QC(logDataArray,wb);
	//getIterReport(logDataArray,wb);
	//getPHAStaticAcquisitionReport(logDataArray,wb);
	//getUniformitytestReport(logDataArray,wb);
	//getUniformitymapcreationReport(logDataArray,wb);
	//getEnergyMapCreationReport(logDataArray,wb);
    createOutPutFile(wb);   
})

var getSensorInfo = function (logDataArray) {

    let systemIDArr = logDataArray["=======SystemLog======="];
    var sensordataArray =[];
	 var finaldataArray =[];
	
	var keyWiseArrayData = {};
    var sensorStr="";
    let line;
    var pattern = /Task:SENSORS Parameters.*/;
	
    for (let i = 0, len = systemIDArr !== undefined ? systemIDArr.length : 0; i < len; ++i) {
	
        line = systemIDArr[i].trim();
        if (!line) continue;
        var linearr = line.toString().split(",");
		var sensordataObj ={};
        var date = linearr[0];
        var time = linearr[1];
		
		sensordataObj["Date"] = date;
		sensordataObj["Time"] = time;
        var sensordata = pattern.exec(line);
        var paramArray = sensordata.toString().split(",")[0].split(":").pop().split("|");
        var length = paramArray.length;
		
		for (var j = 0; j < length; ++j) {
		
            if(paramArray[j]!==""){
                var arr = paramArray[j].toString().split("=");
                if(arr[1]!== undefined){
									
					if(keyWiseArrayData.hasOwnProperty(arr[0])){
						keyWiseArrayData[arr[0]].push(parseFloat(arr[1]));
					}else{
					 keyWiseArrayData[arr[0]] = new Array();
					 keyWiseArrayData[arr[0]].push(parseFloat(arr[1]));
					}
				}
            }
					
        }
        sensordataArray.push(sensordataObj); 
       
    }
	finaldataArray.push(sensordataArray)
	finaldataArray.push(keyWiseArrayData)
	
	return finaldataArray;	
 }
 
 var createtSensorInfo = function (sensorData, wb) {
    
	var ws = wb.addWorksheet('Sensor Data');
	var numv = wb.createStyle({ numberFormat: '0.000',border: {top: {style: 'thick', color: '000000'}} });
	var hedStyle = wb.createStyle({font: {bold: true,color: 'FFFAF0'},fill: {type: 'pattern',patternType: 'solid', fgColor: '000080',bgColor: '000080'}});
	
	var minStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: '800080',bgColor: '800080'},numberFormat: '0.000'});
	
	var maxStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: '0000FF',bgColor: '0000FF'},numberFormat: '0.000'});
	
	var stdStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: 'FFA500',bgColor: '000080'},numberFormat: '0.00'});
	
	var minbStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: '800080',bgColor: '800080'},border: {top: {style: 'thick', color: '000000'}},numberFormat: '0.000'});
	
	var maxbStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: '0000FF',bgColor: '0000FF'},border: {top: {style: 'thick', color: '000000'}},numberFormat: '0.000'});
	
	var stdbStyle = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: 'FFA500',bgColor: '000080'},border: {top: {style: 'thick', color: '000000'}},numberFormat: '0.00'});
	
	var std1Style = wb.createStyle({fill: {type: 'pattern',patternType: 'solid', fgColor: 'FF0000',bgColor: '000080'},numberFormat: '0.000'});
	
	var data =sensorData[1];
	var datedata =sensorData[0];
	
	var j=0;
	var  prevDate ="";
	
	ws.cell(1, 1).string("Legend");
	ws.cell(2, 1).string("Min/Max");
	ws.cell(3, 1).string("1 StedDev");
	ws.cell(4, 1).string("2 StedDev");
	
	ws.cell(1, 2).string("Average");
	ws.cell(2, 2).string("Min");
	ws.cell(3, 2).string("Max");
	ws.cell(4, 2).string("Deviation");
	
	ws.cell(5, 1).string("Date").style(hedStyle);
    ws.cell(5, 2).string("Time").style(hedStyle);
	ws.row(5).filter();
	
	//Math.min.apply(Math,_array)
	for (var key in data)
	{
		
		var arr = data[key];
		 ws.cell(5, 3+j).string(key).style(hedStyle);
		 var minval = Math.min.apply(Math,arr);
		 var maxnval = Math.max.apply(Math,arr);
		 var averagval = average(arr);
		 var stdval = standardDeviation(arr);
		 
		 ws.cell(1, 3+j).number(averagval);
		 ws.cell(2, 3+j).number(minval).style(minStyle);
		 ws.cell(3, 3+j).number(maxnval).style(maxStyle);
		 ws.cell(4, 3+j).number(stdval).style(stdStyle);
		 
		//console.log(minval);
		for(var i = 0; i < arr.length; i++) {
			var date = datedata[i]["Date"]
			if(prevDate==""){
				prevDate = date;
			}
			if(prevDate==date){
				ws.cell(6+i, 1).string(date)
				ws.cell(6+i,2).string(datedata[i]["Time"])
				if(arr[i] ==minval ){
					ws.cell(6+i, 3+j).number(parseFloat(arr[i])).style(minStyle);
				}else if(arr[i] ==maxnval){
					ws.cell(6+i, 3+j).number(parseFloat(arr[i])).style(maxStyle);
				}else{
					ws.cell(6+i, 3+j).number(parseFloat(arr[i]));
				}
			}else{
				ws.cell(6+i, 1).string(date).style(numv);
				ws.cell(6+i,2).string(datedata[i]["Time"]).style(numv);
				if(arr[i] ==minval ){
					ws.cell(6+i, 3+j).number(parseFloat(arr[i])).style(minbStyle);
				}else if(arr[i] ==maxnval){
					ws.cell(6+i, 3+j).number(parseFloat(arr[i])).style(maxbStyle);
				}else{
					ws.cell(6+i, 3+j).number(parseFloat(arr[i])).style(numv);
				}
				
				prevDate = date;
			
			}
		 
		}
		
		j = j+1;

	}
	
 }
 
 function standardDeviation(values){
  var avg = average(values);
  
  var squareDiffs = values.map(function(value){
    var diff = value - avg;
    var sqrDiff = diff * diff;
    return sqrDiff;
  });
  
  var avgSquareDiff = average(squareDiffs);

  var stdDev = Math.sqrt(avgSquareDiff);
  return stdDev;
}

function average(data){
  var sum = data.reduce(function(sum, value){
    return sum + value;
  }, 0);

  var avg = sum / data.length;
  return avg;
}
 

var getDaily_QC = function (logDataArray,wb) {
    var ws = wb.addWorksheet('Daily_QC');
	
    ws.column(1).setWidth(20);
    var  s = 0;
    var  k = 0;
	
	
    let Daily_QC =logDataArray["=======daily_qc_Co57======="];
  
    var arr = Daily_QC.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
	
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
			if (err) {
				console.error(err);                         
				return;
			}
			extracteddata = result['Report']['detResults'];
      
			if(j==1){
			   k =2 
			}else{
			   k = s;
			}
			
			var dateTime = result['Report'].$.date;
			var  date = dateTime.toString().split(" ")[0];
			var  time = dateTime.toString().split(" ")[1];
			
			Object.keys(extracteddata).forEach(function(skey) {
						
				if(skey.startsWith("Detector")){
				let index =skey.substring(8);
				
				var c = index*4
				if(c==0){
				   c =4;
				   }else{
				   c = c+4;
				   }
				
				 ws.cell(1, 1).string("date");
				 ws.cell(1, 2).string("time");
				 
				ws.cell(1, 3).string("isotope");
				ws.column(3).setWidth(20);
				ws.cell(1, c).string(skey+"FWHM");
				ws.column(c).setWidth(20);
    
				ws.cell(1, c+1).string(skey+"UFOV_Integral_Uniformity");
				ws.column(c+1).setWidth(30);
    
				ws.cell(1, c+2).string(skey+"Peak");
				ws.column(c+2).setWidth(20);
    
				ws.cell(1, c+3).string(skey+"CFOV_Integral_Uniformity");
				ws.column(c+3).setWidth(30);
				
					ws.cell(k, 1).string(date)
					ws.cell(k, 2).string(time)
					
					Object.keys(extracteddata[skey]["$"]).forEach(function(key) {
					
						if( key === "isotope"){
							ws.cell(k, 3).string(extracteddata[skey]["$"].isotope)
						}
					
						if( key === "FWHM"){
							ws.cell(k, c).string(extracteddata[skey]["$"].FWHM)
						}
						if( key === "UFOV_Integral_Uniformity"){
							ws.cell(k, c+1).string(extracteddata[skey]["$"].UFOV_Integral_Uniformity)
						}
						
						if( key === "Peak"){
							ws.cell(k, c+2).string(extracteddata[skey]["$"].Peak)
						}
						if( key === "CFOV_Integral_Uniformity"){
							ws.cell(k, c+3).string(extracteddata[skey]["$"].CFOV_Integral_Uniformity)
						}
					});
									
                    s = k+1;
				}
		 
			});
          
		})   
	}

}
var getWeekly_QC = function (logDataArray,wb) {

    var ws = wb.addWorksheet('Weekly_QC');
	
    ws.column(1).setWidth(20);
    var  s = 0;
    var  k = 0;
	
	
    let Weekly_QC =logDataArray["=======weekly_qc_tc99m======="];
    var arr = Weekly_QC.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
	
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
		
			if (err) {
				console.error(err);                         
				return;
			}
			
			extracteddata = result['Report']['detResults'];
			
			if(j==1){
			   k =2 
			}else{
			   k = s;
			}
			
			var dateTime = result['Report'].$.date;
			var  date = dateTime.toString().split(" ")[0];
			var  time = dateTime.toString().split(" ")[1];
		 
			Object.keys(extracteddata).forEach(function(skey) {
						
				if(skey.startsWith("Detector")){
					let index =skey.substring(8);
					var c = index*4
					if(c==0){
					   c =4;
					 }else{
					   c = c+4;
					 }
				
					ws.cell(1, 1).string("date");
					ws.cell(1, 2).string("time");
					ws.cell(1, 3).string("isotope");
					ws.column(3).setWidth(20);
					ws.cell(1, c).string(skey+"UFOV_Integral_Uniformity");
					 ws.column(c).setWidth(20);
   
					ws.cell(1, c+1).string(skey+"FWHM");
					 ws.column(c+1).setWidth(30);
   
					ws.cell(1, c+2).string(skey+"Peak");
					ws.column(c+2).setWidth(20);
					ws.cell(1, c+3).string(skey+"CFOV_Integral_Uniformity");
					ws.column(c+3).setWidth(30);

				
					ws.cell(k, 1).string(date)
					ws.cell(k, 2).string(time)
					Object.keys(extracteddata[skey]["$"]).forEach(function(key) {
					
						if( key === "isotope"){
							ws.cell(k, 3).string(extracteddata[skey]["$"].isotope)
						}
					
						if( key === "UFOV_Integral_Uniformity"){
							ws.cell(k, c).string(extracteddata[skey]["$"].UFOV_Integral_Uniformity)
						}
						
						if( key === "FWHM"){
							ws.cell(k, c+1).string(extracteddata[skey]["$"].FWHM)
						}
						
						if( key === "Peak"){
							ws.cell(k, c+2).string(extracteddata[skey]["$"].Peak)
						}
						
						if( key === "CFOV_Integral_Uniformity"){
							ws.cell(k, c+3).string(extracteddata[skey]["$"].CFOV_Integral_Uniformity)
						}
					});
					
					
                    s = k+1;
				}
		 	});
     	})   
	}
}

var getIterReport = function (logDataArray,wb) {

    var ws = wb.addWorksheet('IterReport');
    
    ws.column(1).setWidth(20);
     
    ws.column(2).setWidth(20);
	ws.cell(1, 1).string("date");
	ws.cell(1, 2).string("time");
   
	
    var  s = 0;
    var  k = 0;
	var  c = 3
	
    let IterReport =logDataArray["=======iter======="];
    var arr = IterReport.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
        if (err) {
                console.error(err);                         
                return;
            }
            
            extracteddata = result['Report']['pmInfo']['default'];
            if(j==1){
               k =2 
            }else{
               k = k+1;
            }
			c = 2;
			var p = 2;
           
			var dateTime = result['Report'].$.date;
			var  date = dateTime.toString().split(" ")[0];
			var  time = dateTime.toString().split(" ")[1];
            Object.keys(extracteddata).forEach(function(dtkey) {
					var temp = 0
					if(dtkey.startsWith("det")){
					var detNo = extracteddata[dtkey]["$"].detectorNumber;
						detNo =parseInt(detNo)+1
						temp = p;
						Object.keys(extracteddata[dtkey]).forEach(function(stkey) {
							if(stkey.startsWith("pm")){
								var index =stkey.substring(2);
								
								if(p==2){
								//console.log(detNo);
									if(detNo==1){
										c = parseInt(index)+2;
									}else{
										//console.log("PM---->"+index);
										c = parseInt(index)+61;
										//console.log("Position---->"+c);
									}
								}else{
								
								if(detNo==1){
									c = parseInt(index)+temp;
									
								}else{
								
									c = parseInt(index)+temp+59;
								}
									
								}
								
								//console.log(c);
								ws.cell(1, c).string("det"+detNo+"."+stkey);
								ws.column(c).setWidth(30);
								
								ws.cell(k, 1).string(date)
								ws.cell(k, 2).string(time)
								
								Object.keys(extracteddata[dtkey][stkey]["$"]).forEach(function(key) {
								
									if( key === "gain"){
										ws.cell(k, c).string(extracteddata[dtkey][stkey]["$"].gain)
									} 
									
								});
								p = p+1;
							
							}
						});
							
						
					}
            });
        })   
    }
}

var getPHAStaticAcquisitionReport = function (logDataArray,wb) {
    var ws = wb.addWorksheet('PHAStaticReport');
    ws.column(1).setWidth(20);
    ws.column(2).setWidth(30);
    ws.column(3).setWidth(30);
    ws.cell(1, 1).string("date");
    ws.cell(1, 2).string("FWHM");
     ws.cell(1, 3).string("Peak"); 
     ws.cell(1, 4).string("TotalCount");
     ws.cell(1, 5).string("Rate");
    var  s = 0;
    var  k = 0;
    let phastatictc99 =logDataArray["=======phastatic_tc99m======="];
	let phastatictc99_30 =logDataArray["=======phastatic_tc99m.30======="];
	let phastatic_co57=logDataArray["=======phastatic_Co57======="];
	let PHAStaticAcquisitionReport = phastatictc99.concat(phastatictc99_30, phastatic_co57);
    var arr = PHAStaticAcquisitionReport.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
		if (err) {
				console.error(err);                         
				return;
			}
			
			extracteddata = result['Report']['detResults'];
			if(j==1){
			   k =2 
			}else{
			   k = s;
			}
			
			var date = result['Report'].$.date;
		 
			Object.keys(extracteddata).forEach(function(st4key) {
			
				if(st4key.startsWith("Detector")){
					ws.cell(k, 1).string(date)
					Object.keys(extracteddata[st4key]["$"]).forEach(function(key) {
					
						if( key === "FWHM"){
                            ws.cell(k, 2).string(extracteddata[st4key]["$"].FWHM)
                           
						} 
                        if( key === "Peak"){
                            ws.cell(k, 3).string(extracteddata[st4key]["$"].Peak)
                           
                        } 
                        if( key === "TotalCount"){
                            ws.cell(k, 4).string(extracteddata[st4key]["$"].TotalCount)
                           
						} if( key === "Rate"){
                            ws.cell(k, 5).string(extracteddata[st4key]["$"].Rate)
                           
						} 
                        
					});
					k = k+1;
					s = k;
				}
		 	});
     	})   
	}
}

var getUniformitytestReport = function (logDataArray,wb) {

    var ws = wb.addWorksheet('UniformitytestReport');
    
    ws.column(1).setWidth(10);
	ws.column(2).setWidth(10);
   
    var  s = 0;
    var  k = 0;
    let UniformitytestReport =logDataArray["=======Utest======="];
    var arr = UniformitytestReport.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
        if (err) {
                console.error(err);                         
                return;
            }
            
            extracteddata = result['Report']['detResults'];
            if(j==1){
               k =2 
            }else{
               k = s;
            }
            
            var dateTime = result['Report'].$.date;
            var date =dateTime.toString().split(" ")[0];
            var time =dateTime.toString().split(" ")[1];
         
            Object.keys(extracteddata).forEach(function(strkey) {
            
                if(strkey.startsWith("Detector")){
                    let index =strkey.substring(8);
                    var c = index*2
                    if(c==0){
                        c=3;
                    }else{
                        c=c+3;
                    }
                    
                    ws.cell(1, 1).string("date");
                    ws.cell(1, 2).string("time");
                    //ws.column(1,3).setWidth(50);
                    ws.cell(1,c).string(strkey+"uFOVResult_intgUniformity");
                    ws.column(c).setWidth(40);
                    ws.cell(1,c+1).string(strkey+"cFOVResult_intgUniformity");
                    ws.column(c+1).setWidth(40);
                    ws.cell(k, 1).string(date)
                    ws.cell(k, 2).string(time)
                    
                    Object.keys(extracteddata[strkey]["$"]).forEach(function(key) {
                    
                        if( key === "uFOVResult_intgUniformity"){
                            ws.cell(k, c).string(extracteddata[strkey]["$"].uFOVResult_intgUniformity)
                           
                        } 
                        if( key === "cFOVResult_intgUniformity"){
                            ws.cell(k, c+1).string(extracteddata[strkey]["$"].cFOVResult_intgUniformity)
                        }
                        
                    });
                    s = k+1;
            
                }
            });
        })   
    }
}
var getUniformitymapcreationReport = function (logDataArray,wb) {

    var ws = wb.addWorksheet('UniformitymapcreationReport');
    ws.column(1).setWidth(20);
    
    var  s = 0;
    var  k = 0;
    let UniformitymapcreationReport =logDataArray["=======Uniformitymapcreation======="];
    var arr = UniformitymapcreationReport.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
			if (err) {
					console.error(err);                         
					return;
			 }
            
            extracteddata = result['Report']['detResults'];
            if(j==1){
               k =2 
            }else{
               k = s;
            }
            
            var dateTime = result['Report'].$.date;
			var title = result['Report'].$.title;
			var isotope = result['Report']['inputParam'].$.Isotope;
			var MapName = result['Report']['inputParam'].$.MapName;
            var date =dateTime.toString().split(" ")[0];
            var time =dateTime.toString().split(" ")[1];
         
              
            Object.keys(extracteddata).forEach(function(strkey) {
            
                if(strkey.startsWith("Detector")){
				
				let index =strkey.substring(8);
                    var c = index*1
                    if(c==0){
                        c=6;
                    }else{
                        c=c+6;
                    }
				
					ws.cell(1, 1).string("date");
					ws.cell(1, 2).string("time");
					ws.cell(1, 3).string("title");
					ws.column(3).setWidth(40);
					ws.cell(1, 4).string("Isotope");
					ws.column(4).setWidth(20);
					ws.cell(1, 5).string("MapName");
					ws.column(5).setWidth(40);
					 ws.cell(1,c).string(strkey+"MapName");
                    ws.column(c).setWidth(40);
				
                    ws.cell(k, 1).string(date)
                    ws.cell(k, 2).string(time)
					ws.cell(k, 3).string(title)
					ws.cell(k, 4).string(isotope)
					ws.cell(k, 5).string(MapName)
                    Object.keys(extracteddata[strkey]["$"]).forEach(function(key) {
                    
                        if( key === "MapName"){
                            ws.cell(k, c).string(extracteddata[strkey]["$"].MapName)
                           
                        } 
                        
                        
                    });
                   s = k+1;
                }
            });
        })   
    }
}


var getEnergyMapCreationReport = function (logDataArray,wb) {

    var ws = wb.addWorksheet('EnergyMapCreationReport');
    ws.column(1).setWidth(20);
    
    var  s = 0;
    var  k = 0;
    let UniformitymapcreationReport =logDataArray["=======EnergyMapCreation======="];
    var arr = UniformitymapcreationReport.toString().split("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
    var parser = new xml2js.Parser({explicitArray : false});
    var extracteddata = "";
    for(var j=1;j<arr.length;j++){
        var cleanedString = arr[j].trim().replace(",", "");
        parser.parseString(cleanedString, function(err,result){
			if (err) {
					console.error(err);                         
					return;
			 }
            
            extracteddata = result['Report']['detResults'];
            if(j==1){
               k =2 
            }else{
               k = s;
            }
            
            var dateTime = result['Report'].$.date;
			var title = result['Report'].$.title;
			var isotope = result['Report']['inputParam'].$.Isotope;
			var MapName = result['Report']['inputParam'].$.MapName;
            var date =dateTime.toString().split(" ")[0];
            var time =dateTime.toString().split(" ")[1];
         
		 
			ws.cell(1, 1).string("date");
			ws.cell(1, 2).string("time");
			ws.cell(1, 3).string("title");
			ws.column(3).setWidth(40);
			ws.cell(1, 4).string("Isotope");
			ws.column(4).setWidth(20);
			ws.cell(1, 5).string("MapName");
			ws.column(5).setWidth(40);
			
		
			ws.cell(k, 1).string(date)
			ws.cell(k, 2).string(time)
			ws.cell(k, 3).string(title)
			ws.cell(k, 4).string(isotope)
			ws.cell(k, 5).string(MapName)
            if(extracteddata)  {
            Object.keys(extracteddata).forEach(function(strkey) {
            
                if(strkey.startsWith("Detector")){
				
				let index =strkey.substring(8);
                    var c = index*1
                    if(c==0){
                        c=6;
                    }else{
                        c=c+6;
                    }
				
					
                    Object.keys(extracteddata[strkey]["$"]).forEach(function(key) {
                    
                        if( key === "MapName"){
							ws.cell(1,c).string(strkey+"MapName");
							ws.column(c).setWidth(40);
                            ws.cell(k, c).string(extracteddata[strkey]["$"].MapName)
                           
                        } 
                        
                        
                    });
                  
                }
            });
			}
			 s = k+1;
        })   
    }
}

var createOutPutFile = function (wb) {
   
    wb.write('ExcelFile.xlsx', function(err, stats) {
      if (err) {
        console.error(err);
      } else {
        console.log(stats);
      }
    });
}

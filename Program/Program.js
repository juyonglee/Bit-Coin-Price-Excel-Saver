const request = require('sync-request');
const Binance = require('node-binance-api');
var Excel = require('exceljs');
var BinanceList = [], GOPAXList = [], HuobiList = [];
var binanceResult = [], gopaxResult = [], huobiResult = [];
var binanceEndpoint = 'https://api.binance.com', gopaxEndpoint = 'https://api.gopax.co.kr/trading-pairs', huobiEndPoint = "http://api.huobi.pro";
var workbook = new Excel.Workbook();
var maxLength;
var worksheet;
var testResult;
var ProgressBar = require("cli-progress-bar")
var bar = new ProgressBar();

workbook.xlsx.readFile("List.xlsx")
  .then(function() {
      console.log("[Step01] Read List");
      readListFunction();
      return;
  }).then(function() {
      console.log("[Step02] Price Data Parsing");
      binanceParsingFunction();
      gopaxParsingFunction();
      huobiParsingFunction();
    //   huobiList();
      return;
  }).then(function() {
      console.log("[Step03] Excel File Creation");
      writeExcelFile();
      return;
  });


function readListFunction() {
    BinanceList = workbook.getWorksheet("List").getColumn(1).values;
    BinanceList.splice(BinanceList.indexOf("Binance"), 1);
    BinanceList.splice(BinanceList.indexOf("종목명"), 1);
    BinanceList.sort();

    GOPAXList = workbook.getWorksheet("List").getColumn(2).values;
    GOPAXList.splice(GOPAXList.indexOf("GOPAX"), 1);
    GOPAXList.splice(GOPAXList.indexOf("종목명"), 1);
    GOPAXList.sort();

    HuobiList = workbook.getWorksheet("List").getColumn(3).values;
    HuobiList.splice(HuobiList.indexOf("Huobi"), 1);
    HuobiList.splice(HuobiList.indexOf("종목명"), 1);
    HuobiList.sort();;
    maxLength = Math.max(BinanceList.length, GOPAXList.length, HuobiList.length);
}

function binanceParsingFunction() {
    for(var i=0; i<BinanceList.length; i++) {
        if(BinanceList[i] != undefined) {
            bar.show("", ((i+1)/BinanceList.length)); 
            bar.pulse("Binance: " + BinanceList[i] + " 정보 다운로드 중.....");
            var res = request('GET', binanceEndpoint +'/api/v3/ticker/price?symbol=' + BinanceList[i]);
            var parsingData = JSON.parse(res.getBody());
            binanceResult.push({name: BinanceList[i], price: parseFloat(parsingData['price'])});
        } else {
            binanceResult.push({name: " ", price: " "});
        }
        
    }  
    bar.hide();  
    console.log("Binance에서 정보 수신 완료!");     
}

function gopaxParsingFunction() {
      //  GOPAX Pasing Function
              for(var i=0; i<GOPAXList.length; i++) {
                bar.show("", ((i+1)/GOPAXList.length)); 
                bar.pulse("GOPAX: " + GOPAXList[i] + " 정보 다운로드 중.....");
                //   console.log(parseInt(((i+1)/GOPAXList.length)*100) + "%");
                  if(GOPAXList[i] != undefined) {
                      var res = request('GET', gopaxEndpoint+'/'+GOPAXList[i]+'/ticker');
                      var parsingData = JSON.parse(res.getBody());
                      gopaxResult.push({name: GOPAXList[i], price: parsingData['price']});
                  } 
                  else {
                      gopaxResult.push({name: " ", price: " "});
                  }
              }
              bar.hide();
              console.log("GOPAX에서 정보 수신 완료!");      
}


function huobiParsingFunction() {
      //  Huobi Parsing Function
      console.log("Huobi에서 정보 수신 중....");
      for(var i=0; i<HuobiList.length; i++) {
          bar.show("", ((i+1)/HuobiList.length)); 
          bar.pulse("Huobi: "+ HuobiList[i] + " 정보 다운로드 중.....");
        //   console.log(parseInt(((i+1)/HuobiList.length)*100) + "%");
          if(HuobiList[i] != undefined) {
              var res =  request('GET', huobiEndPoint+ "/market/history/kline?period=1day&size=1&symbol="+HuobiList[i]);
              var price = JSON.parse(res.getBody());
              huobiResult.push({name: HuobiList[i], price: price['data'][0]['close']});
          } else {
              huobiResult.push({name: " ", price: " "});
          }
      }
      bar.hide();
      console.log("Huobi에서 정보 수신 완료!");
}

function huobiList() {
            var res =  request('GET', huobiEndPoint+ "/market/tickers");
            var price = JSON.parse(res.getBody());
            console.log(price['data'].length);
            var nameList = [];
            for(var i=0; i<price['data'].length; i++) {
                if(price['data'][i]['symbol'].match('btc')) {
                    huobiResult.push({name: price['data'][i]['symbol'], price: price['data'][i]['close']});
                    nameList.push(price['data'][i]['symbol']);
                }
            }
            nameList.sort();
            // for(var i=0; i<nameList.length; i++) {
            //     huobiResult.push({name: nameList[i], price: price['data'][i]['close']});
            // }
    bar.hide();
    console.log("Huobi에서 정보 수신 완료!");
}

function writeExcelFile() {
    var currentdate = new Date(); 
    var datetime =""+
                currentdate.getFullYear() + "." + (currentdate.getMonth()+1)  + "." 
                + currentdate.getDay() + "(" +currentdate.getHours()+"시"+currentdate.getMinutes()+"분"+currentdate.getSeconds()+"초)";
                // var worksheet = workbook.addWorksheet(datetime);
                var workbook = new Excel.Workbook();
                var worksheet = workbook.getWorksheet('aaqqww');
    
                if(!worksheet) {
                    worksheet = workbook.addWorksheet('aaqqww');

                }

                worksheet.getColumn(1).values = printFormCreation('바이낸스', '종목명', binanceResult, 'name');
                worksheet.getColumn(1).width = 12;
                worksheet.getColumn(2).values = printFormCreation('', '가격', binanceResult, 'price');
                worksheet.getColumn(2).width = 15;
                worksheet.getColumn(3).values = printFormCreation('고팍스', '종목명', gopaxResult, 'name');
                worksheet.getColumn(3).width = 12;
                worksheet.getColumn(4).values = printFormCreation('', '가격(원화)', gopaxResult, 'price');
                worksheet.getColumn(4).width = 15;
                worksheet.getColumn(5).values = printFormCreation('후오비', '종목명', huobiResult, 'name');
                worksheet.getColumn(5).width = 12;
                worksheet.getColumn(6).values = printFormCreation('', '가격', huobiResult, 'price');
                worksheet.getColumn(6).width = 15;
                workbook.xlsx.writeFile("aaqqww.xlsx");
}

function printFormCreation(name, value, DataSet, type) {
    var data = [];
    data.push(name);
    data.push(value);
    for(var i=0; i<DataSet.length; i++) {
        data.push(DataSet[i][type]);    
    }
    return data;
}
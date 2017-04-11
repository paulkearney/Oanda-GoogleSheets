function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Oanda')
      .addItem('Update', 'update')
      .addToUi();
}

function update() {
  var accountNum = SpreadsheetApp.getActive().getSheetByName('Account').getRange('B100').getValue();
  updateAccountData(accountNum);
  updateOpenTrades(accountNum);
  updateHistory(accountNum);
  SpreadsheetApp.getActive().getSheetByName('Account').getRange('B9').setValue(new Date());
}

function updateHistory(accountId) {
  var formulas = [['=R[0]C[-8]-R[0]C[-9]', '=YEAR(R[0]C[-10])', '=MONTH(R[0]C[-11])', '=DAY(R[0]C[-12])', '=WEEKDAY(R[0]C[-13])', '=WEEKNUM(R[0]C[-14])', '=HOUR(R[0]C[-15])' ]];
  var sheet = SpreadsheetApp.getActive().getSheetByName('History');
  writeHistoryHeaders(sheet);
  
  var done = false;
  var lastHistoryId = 0;
  while (done == false) {
    var url = 'https://api-fxtrade.oanda.com/v3/accounts/'+accountId+'/trades?state=CLOSED&count=500';
    if (lastHistoryId > 0) {
      url = url + '&beforeID=' + lastHistoryId;
    }
    var data = fetchJson(url);
    var rowUpdates = [];
    var existingIds = sheet.getRange('A:A').getValues().map(function(row) {return parseInt(row[0]);});
    
    for (var t = 0; t < data.trades.length; t++) {
      var trade = data.trades[t];
      // see if this trade is already in the list
      if (existingIds.indexOf(parseInt(trade.id)) > -1) {
        Logger.log('Already have trade id ' + trade.id); 
        continue; // we already have this trade
      }
      rowUpdates.push([trade.id, trade.clientExtensions.id, trade.initialUnits < 0 ? 'SELL':'BUY', getDateFromIso(trade.openTime), getDateFromIso(trade.closeTime),
                       formatSymbol(trade.instrument), trade.clientExtensions.tag, Math.abs(trade.initialUnits/100000), trade.price, trade.averageClosePrice, trade.realizedPL, getComment(trade),
                      ]);
      lastHistoryId = trade.id;
    }
        
    if (rowUpdates.length == 0) {
      Logger.log('Last fetch resulted in 0 new rows. Exiting loop.');
      break;
    }
    
    var numLastRow = sheet.getLastRow();
    //add row(s) at end if necessary
    var totalRowsNeeded = rowUpdates.length + sheet.getLastRow();
    if (totalRowsNeeded > sheet.getMaxRows()) {
      for (var r = sheet.getMaxRows() - totalRowsNeeded; r > 0; r--) sheet.insertRowAfter(sheet.getMaxRows());
    }

    numLastRow++;
    for (var r = 0; r < rowUpdates.length; r++) {
      var thisRow = numLastRow + r;
      var valueRange = sheet.getRange('A'+thisRow+':L'+thisRow);
      var formulaRange = sheet.getRange('M'+thisRow+':S'+thisRow);      
      valueRange.setValues([rowUpdates[r]]);
      formulaRange.setFormulasR1C1(formulas)
    }
  }

  var maxId = Math.max.apply(null, data.trades.map(function(t) {return t.id;}));
  var minId = Math.min.apply(null, data.trades.map(function(t) {return t.id;}));
  SpreadsheetApp.getActive().getSheetByName('Account').getRange('B8').setValue(maxId);

  sheet.sort(1, false);
}

function updateOpenTrades(accountId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('OpenTrades');
  var row = 2;
  sheet.getRange('A2:L999').clearContent();
  writeOpenTradeHeaders(sheet);
  //get open trades for acct
  var openTrades = getOpenTrades(accountId);
  for (var ot = 0; ot < openTrades.trades.length; ot++) {
    var t = openTrades.trades[ot];
    var range = sheet.getRange('A'+row+':L'+row);
    var d = getDateFromIso(t.openTime);
    range.setValues([[t.id, t.clientExtensions.id, t.initialUnits < 0 ? 'SELL' : 'BUY', d, formatSymbol(t.instrument), t.clientExtensions.tag, Math.abs(t.initialUnits/100000), t.price, t.unrealizedPL, 
                      getComment(t) ,
                      '=now()-d'+row, '=i'+row+'/Account!$b$1']]);
    row++;
  }
}

function updateAccountData(accountId) {
  var data = getAccountData(accountId);
  var sheet = SpreadsheetApp.getActive().getSheetByName('Account');
  var range = sheet.getRange('B1:B7');
  range.clearContent();
  range.setValues([[data.balance],
                  [data.marginCloseoutNAV],
                  [data.marginUsed],
                  ['=b2/b3*100'],
                  [data.marginAvailable],
                  [data.unrealizedPL],
                  [data.openTradeCount]])
}

function getAccountData(accountId) {
  var data = fetchJson('https://api-fxtrade.oanda.com/v3/accounts/' + accountId);
  return data.account;
}

function getAccounts() {
  var data = fetchJson('https://api-fxtrade.oanda.com/v3/accounts');
  return data.accounts.map(function(a) {return a.id;});
}

function getOpenTrades(accountId) {
  var data = fetchJson('https://api-fxtrade.oanda.com/v3/accounts/' + accountId + '/openTrades');
  return data;
}

function writeOpenTradeHeaders(sheet) {
  var range = sheet.getRange('A1:L1');
  range.setValues([['Id', 'Order Number', 'Type', 'Open Time', 'Symbol', 'Magic Number', 'Lots', 'Open Price', 'OpenPL', 'Comment', 'Duration', 'PL %']]);
}

function writeHistoryHeaders(sheet) {
  var range = sheet.getRange('A1:S1');
  range.setValues([['Id', 'Order Number', 'Type', 'Open Time', 'Close Time',  'Symbol', 'Magic Number', 'Lots', 'Open Price', 'Close Price', 'Profit', 'Comment', 
                    'Duration', 'Year', 'Month', 'Day', 'DOW', 'Week Number', 'Hour']]);
}

function fetchJson(url) {
  Logger.log(url);
  var apiKey = SpreadsheetApp.getActive().getSheetByName('Account').getRange('B99').getValue();
  var headers = {
    'Authorization': 'Bearer ' + apiKey
  };
  var options = {
    'method': 'get',
    'headers': headers
  };
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  return data;
}


// http://delete.me.uk/2005/03/iso8601.html
function getDateFromIso(string) {
  try{
    var aDate = new Date();
    var regexp = "([0-9]{4})(-([0-9]{2})(-([0-9]{2})" +
        "(T([0-9]{2}):([0-9]{2})(:([0-9]{2})(\\.([0-9]+))?)?" +
        "(Z|(([-+])([0-9]{2}):([0-9]{2})))?)?)?)?";
    var d = string.match(new RegExp(regexp));

    var offset = 0;
    var date = new Date(d[1], 0, 1);

    if (d[3]) { date.setMonth(d[3] - 1); }
    if (d[5]) { date.setDate(d[5]); }
    if (d[7]) { date.setHours(d[7]); }
    if (d[8]) { date.setMinutes(d[8]); }
    if (d[10]) { date.setSeconds(d[10]); }
    if (d[12]) { date.setMilliseconds(Number("0." + d[12]) * 1000); }
    if (d[14]) {
      offset = (Number(d[16]) * 60) + Number(d[17]);
      offset *= ((d[15] == '-') ? 1 : -1);
    }
    offset -= date.getTimezoneOffset();
    time = (Number(date) + (offset * 60 * 1000));
    return new Date(aDate.setTime(Number(time)));
  } catch(e){
    return;
  }
}

function getComment(trade) {
  return typeof(trade.clientExtensions.comment) === 'undefined' ? '' : trade.clientExtensions.comment;
}

function formatSymbol(symbol) {
  return symbol.replace(/_/g, '');
}

/**
 * Fetches discount code of an order.
 *
 * @param {id} id of an order.
 * @return {code} Discount code of a given order.
 * @customfunction
 */
function ImportEcwidOrder(id=0) {
  if (id == 0) {
    return null;
  }
  let url = ''+id;
  let options = {
    'method' : 'get',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'headers' : {
      "Authorization" : ""
    }
  };
  let response = null;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    code = "Can't connect!";
  }
  try {
    code = JSON.parse(response)["discountCoupon"]["code"];
  } catch (e) {
    code = 'No Code';
  }
  return code;
}

/**
 * Fetches discount code usage in Ecwid.
 *
 * @param {code} Discount code.
 * @return {code} Uses count in Ecwid.
 * @customfunction
 */
function ImportEcwidDiscountUses(code) {
  let url = ''+code;
  let options = {
    'method' : 'get',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'headers' : {
      "Authorization" : ""
    }
  };
  let response = null;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    total = "Can't connect!";
  }
  try {
    total = JSON.parse(response)["total"];
  } catch (e) {
    total = 'No Code';
  }
  return total;
}

function sendNotification(data) {
  let code = '';
  if (data.length > 10) {
    code = data[176].split(' ').pop();
    Logger.log(code);
  } else {
    code = data[data.length - 1];
    Logger.log(code);
  }

  address = getExpertMailByCode(code);

  const mailtemp = HtmlService.createTemplateFromFile('mailtemp');
  mailtemp.code = code;
  const message = mailtemp.evaluate().getContent();
  GmailApp.sendEmail(address, 'Использован промокод эксперта', message, {
    name: 'Программа лояльности LitteLifeLab',
    from: 'experts@mylifelab.ru',
    htmlBody: message
  });
}

function mailTest() {
  const mailtemp = HtmlService.createTemplateFromFile('mailtemp');
  mailtemp.code = '8U4KSELP';
  const message = mailtemp.evaluate().getContent();
  GmailApp.sendEmail('supreme.opp@gmail.com', 'Использован промокод эксперта', message, {
    name: 'LitteLifeLab — Программа лояльности ',
    from: 'experts@mylifelab.ru',
    htmlBody: message
  });
}

function rewardExperts(){
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var expertsDataSheet = activeSheet.getSheetByName('experts');

  // 0 name; 1 id; 2 mail; 3 code; 6 cache
  expertData = expertsDataSheet.getRange("A2:G").getValues();
  usesData = expertsDataSheet.getRange("H2:H").getValues();
  for (i = 0; i < usesData.length; ++i) {
    if (usesData[i][0] - expertData[i][6] >= 11) {
      expertData[i][6] += 11;
      addPoints(expertData[i][1]);
      const rewardtemp = HtmlService.createTemplateFromFile('rewardtemp');
      rewardtemp.code = expertData[i][3];
      const message = rewardtemp.evaluate().getContent();
      GmailApp.sendEmail(expertData[i][2], 'Награда за 11 заказов по промокоду', message, {
        name: 'LitteLifeLab — Программа лояльности ',
        from: 'experts@mylifelab.ru',
        htmlBody: message
      });
    };
  }
  expertsDataSheet.getRange("A2:G").setValues(expertData);
};

function addPoints(expertId) {
  var raw = JSON.stringify({
    "bonus_system_transaction": {
      "bonus_points": 5250,
      "description": "bonus"
    }
  });

  var headers = {
    "Content-Type": "application/json",
    "Authorization": "Basic " + Utilities.base64Encode("")
  };

  var options = {
    "method": "POST",
    "contentType": "application/json",
    "headers": headers,
    "payload": raw
  };

  UrlFetchApp.fetch("" + expertId + "/bonus_system_transactions.json", options)
}

function aliasTest() {
  var aliases = GmailApp.getAliases();
  Logger.log(aliases[0]);
}

function getExpertMailByCode(code) {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var expertsDataSheet = activeSheet.getSheetByName('experts');

  // Find text within sheet.
  var textSearch = expertsDataSheet.createTextFinder(code).findAll();

  if (textSearch.length > 0) {
    // Get single row from search result.
    var row = textSearch[0].getRow();    
    // Get the last column so we can use for the row range.
    var rowLastColumn = expertsDataSheet.getLastColumn();
    // Get all values for the row.
    var rowValues = expertsDataSheet.getRange(row, 1, 1, rowLastColumn).getValues();

    return rowValues[0][2]; // email address
  }
  else {
    return "";
  }
}

function insertRowAtTop_v1(data, sheetName, targetRow) {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
 
  // Insert a row
  // NOTE show what happens if we use insertRowAfter and how it carries over formatting from the top.
  sheet.insertRowBefore(targetRow);
  sheet
    .getRange(targetRow, 1, 1, data[0].length)
    .setValues(data);
 
  SpreadsheetApp.flush();
}
 
 
/**
* Runs example1.
* This function simulates how the insertRowAtTop funciton can be used.
*/
function runsies_example1(){
  const targetRow = 2;
  const sheetName = "ecwid_stream"
 
  // Dummy Data
  const myDate = new Date();
  const myTime = myDate.getTime() // The id
  const data = [
    [
      myTime,
      myDate,
      `${myTime}@example.com`
    ]
  ]
 
  insertRowAtTop_v1(data, sheetName, targetRow)
}


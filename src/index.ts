
const gas: any = global;

type ResponseData = {
  spinach?: {
    date: Date,
    chart: string,
    quantity: number,
    price: {
      al: number, am: number, as: number
    }
  },
  watermelon?: {
    date: Date,
    chart: string,
    quantity: number,
    price: {
      average: number,
      s4: number, s5: number, sl: number, sm: number,
      y4: number, y5: number, yl: number, ym: number
    }
  }
}

gas.doGet = (e: GoogleAppsScript.Events.DoGet): GoogleAppsScript.Content.TextOutput => {
  const SPINACH_SPREAD_SHEET_ID = getProperty('SPINACH_SPREAD_SHEET_ID');
  const WATERMELON_SPREAD_SHEET_ID = getProperty('WATERMELON_SPREAD_SHEET_ID');

  const sheetName = String(new Date().getFullYear());
  const responseData: ResponseData = {};

  const spinachSpreadSheet = SpreadsheetApp.openById(SPINACH_SPREAD_SHEET_ID);
  const spinachSheet = spinachSpreadSheet.getSheetByName(sheetName);
  if (!spinachSheet) {
    return responseJson({ errorMessage: 'spinachSheet is null.' });
  }
  // 日時, AL, AM, AS, 出荷数量
  const spinachData = spinachSheet.getRange(
    spinachSheet.getLastRow(), 1, 1, 5).getValues()[0];
  responseData.spinach = {
    date: spinachData[0],
    chart: getChart(spinachSheet),
    quantity: toInt(spinachData[4]),
    price: {
      al: toInt(spinachData[1]),
      am: toInt(spinachData[2]),
      as: toInt(spinachData[3]),
    }
  };

  const watermelonSpreadSheet = SpreadsheetApp.openById(WATERMELON_SPREAD_SHEET_ID);
  const watermelonSheet = watermelonSpreadSheet.getSheetByName(sheetName);
  if (!watermelonSheet) {
    return responseJson({ errorMessage: 'watermelonSheet is null.' });
  }
  // 日時, 秀4, 秀5, 秀L, 秀M, 優4, 優5, 優L, 優M, 平均単価, 出荷箱数
  const watermelonData = watermelonSheet.getRange(
    watermelonSheet.getLastRow(), 1, 1, 11).getValues()[0];

  responseData.watermelon = {
    date: watermelonData[0],
    chart: getChart(watermelonSheet),
    quantity: toInt(watermelonData[10]),
    price: {
      average: toInt(watermelonData[9]),
      s4: toInt(watermelonData[1]),
      s5: toInt(watermelonData[2]),
      sl: toInt(watermelonData[3]),
      sm: toInt(watermelonData[4]),
      y4: toInt(watermelonData[5]),
      y5: toInt(watermelonData[6]),
      yl: toInt(watermelonData[7]),
      ym: toInt(watermelonData[8]),
    }
  };

  return responseJson(responseData);
};

function toInt(value: any) {
  if (typeof value === 'number') {
    return value;
  } else if (typeof value === 'string' && value.match(/^\d+$/)) {
    return Number(value);
  } else {
    return 0;
  }
}

function getChart(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  return Utilities.base64Encode(sheet.getCharts()[0].getBlob().getBytes());
}

function responseJson(data: any): GoogleAppsScript.Content.TextOutput {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function getProperty(key: string, defaultValue?: any): string {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value) return value;
  if (defaultValue) return defaultValue;
  throw new Error(`Undefined property: ${key}`);
}

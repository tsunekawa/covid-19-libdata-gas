
// 情報取得系関数

function getSpreadsheet() {
  if (!this.spreadSheet) {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  }
  
  return this.spreadSheet
}

function getMasterSheet() {
  const MASTER_SHEET_NAME = '全国調査（4月9日分）'
  
  if (!this.sheet) {
    this.sheet = getSpreadsheet().getSheetByName(MASTER_SHEET_NAME)
  }  

  return this.sheet;
}

function getHeaders() {
  let sheet = getMasterSheet()
  let headers = sheet.getRange('1:1').getValues().flat().filter( (value) => value.length > 0 )
  
  return headers
}

function groupRowsByPrefecture(sheet) {
  const PREFECTURE_LABEL = '都道府県'
  
  let data = sheet.getDataRange().getValues()
  let headers = data.shift()
  let prefectureColumnIndex = headers.indexOf(PREFECTURE_LABEL)
  
  let groups = data.map((rowValues) => {
    return [rowValues[prefectureColumnIndex], rowValues]
  }).reduce( (obj, entry) =>{
    if (!Array.isArray(obj[entry[0]])) {
      obj[entry[0]] = []
    }
    
    obj[entry[0]].push(entry[1])
    
    return obj
  }, {})
  
  return groups
}

// 更新系の関数

function createPrefectureSheet(name, rows) {
  const CODE_COLUMN_NAME = '市町村コード'
  let sheetName = "分割_" + name
  let spreadSheet = getSpreadsheet()
  let headers = getHeaders()
  
  let sheet = spreadSheet.insertSheet(sheetName)
  
  // 市町村コードの列を文字列属性にする
  let codeColumnRange = sheet.getRange(1, headers.indexOf(CODE_COLUMN_NAME)+1, rows.length, 1)
  codeColumnRange.setNumberFormat('@')
  
  // ヘッダーとそのスタイルを設定
  let headerRange = sheet.getRange(1, 1, 1, headers.length)
  headerRange.setValues([headers])
  headerRange.setBackground('#fff2cc')
  headerRange.setBorder(true, true, true, true, true, true)
  
  // データと各行のスタイルを設定
  let dataRange = sheet.getRange(2, 1, rows.length, headers.length)
  dataRange.setValues(rows)
  dataRange.setBorder(true, true, true, true, true, true)
  
  sheet.getDataRange().createFilter()
  
  return sheet
}

function splitMasterSheetByPrefecture() {
  let masterSheet = getMasterSheet()
  let groups = groupRowsByPrefecture(masterSheet)
  
  let createdSheets = Object.entries(groups).map( (entry) => {
    let name = entry[0]
    let rows = entry[1]
    return createPrefectureSheet(name, rows)
  })
  
  return createdSheets
}

function getPartSheets() {
  let spreadsheet = getSpreadsheet()
  let partSheets  = spreadsheet.getSheets().filter( (sheet) => sheet.getName().match(/^分割_.+/) )
  
  return partSheets
}

function deleteAllPartSheet() {
  let spreadsheet = getSpreadsheet()
  let partSheets  = getPartSheets()
  partSheets.forEach( (sheet) => spreadsheet.deleteSheet(sheet) )
  
  return true
}

function integrateAllPartSheets() {
  let masterSheet = getMasterSheet()
  let integratedSheet = getSpreadsheet().insertSheet('【統合】')
  let headersRange = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn())
  headersRange.copyTo(integratedSheet.getRange(1, 1))
  
  let rangeData = getPartSheets().reduce( (obj, sheet) => {  
    const range = sheet.getDataRange()
    const dataRange = range.offset(1, 0, range.getNumRows()-1)
    
    obj.totalNumRows += dataRange.getNumRows()
    obj.ranges.push(dataRange)
    
    return obj
  }, { totalNumRows: 0, ranges: []})  

  integratedSheet.insertRowsAfter(1, rangeData.totalNumRows - 1000 + 1)
  
  rangeData.ranges.forEach((range) => {
    let dataRange = integratedSheet.getDataRange()
    let lastRowIndex = (dataRange != undefined) ? dataRange.getNumRows() : 0
    let destination = integratedSheet.getRange(lastRowIndex+1, 1)
    range.copyTo(destination)
  })
    
  return integratedSheet
}

function test() {
  let result = splitMasterSheetByPrefecture()
  
  console.log(groups)
}
/*

Copyright 2020 Tsunekawa Mao

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/

type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet

// 情報取得系関数

function getSpreadsheet() {
  if (!this.spreadSheet) {
    this.spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  }
  
  return this.spreadSheet
}

function getMasterSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const MASTER_SHEET_NAME: string = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_NAME')
  
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

// 更新系の関数

function createPrefectureSheet(name, rows) {
  const CODE_COLUMN_NAME = PropertiesService.getScriptProperties().getProperty('CODE_COLUMN_NAME') || '市町村コード'
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

function copyPrefectureSheet(name: string, sourceRange: GoogleAppsScript.Spreadsheet.Range, prefix='分割_') {
  let sheetName = prefix + name
  let spreadSheet = getSpreadsheet()
  let sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.insertSheet(sheetName)

  sourceRange.copyTo(sheet.getRange(1, 1))
  sheet.getDataRange().createFilter()
  
  return sheet
}

function filterSheetByPrefecture(dataRange: GoogleAppsScript.Spreadsheet.Range, prefectureName: string): GoogleAppsScript.Spreadsheet.Range {
  const PREFECTURE_LABEL = PropertiesService.getScriptProperties().getProperty('PREFECTURE_LABEL') || '都道府県'
  
  let prefectureColumnIndex: number = dataRange.getValues()[0].indexOf(PREFECTURE_LABEL) + 1
  
  let filter = dataRange.getFilter() || dataRange.createFilter()
  let criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(prefectureName).build()
 
  filter.removeColumnFilterCriteria(prefectureColumnIndex)
  filter.setColumnFilterCriteria(prefectureColumnIndex, criteria)

  return filter.getRange()
}

function splitMasterSheetByPrefecture(prefix='分割_'): GoogleAppsScript.Spreadsheet.Sheet[] {
  const PREFECTURE_LABEL = PropertiesService.getScriptProperties().getProperty('PREFECTURE_LABEL') || '都道府県'

  let masterSheet = getMasterSheet()
  let masterHeader = masterSheet.getRange('1:1')
  let masterTable = masterSheet.getDataRange()
  let prefectureColumnIndex: number = masterHeader.getValues()[0].indexOf(PREFECTURE_LABEL) + 1
  let masterData = masterSheet.getRange(2, prefectureColumnIndex, masterSheet.getLastRow() - 1)

  let prefectureNames: string[] = Array.from(new Set(masterData.getValues().flat()))

  let prefectureSheets = prefectureNames.map((prefectureName: string): GoogleAppsScript.Spreadsheet.Sheet => {
    let filteredRange = filterSheetByPrefecture(masterTable, prefectureName)
    return copyPrefectureSheet(prefectureName, filteredRange, prefix=prefix)
  })

  return prefectureSheets
}

function getSheetURL(sheet: Sheet): string {
  let spreadsheetURL: string = sheet.getParent().getUrl()
  let sheetId: string = sheet.getSheetId().toString()

  return `${spreadsheetURL}#gid=${sheetId}`
}

function setupWorksheets(): Sheet {
  const prefectureSheetPrefix = '分割_'
  const spreadsheet: Spreadsheet = getSpreadsheet()
  const prefectureSheets: Sheet[] = splitMasterSheetByPrefecture(prefectureSheetPrefix)
  const headers = ['都道府県名', '件数', '調査済', '完了率', '調査完了日', '担当', 'メモ']
  const headerColor = '#fff2cc'
  const sheetName = '【分割シート一覧】'

  let statusSheet = spreadsheet.insertSheet(sheetName, 0)

  const data: string[][] = prefectureSheets.map((sheet, idx) => {
    let sheetName = sheet.getSheetName().replace(prefectureSheetPrefix, '')
    let sheetURL = getSheetURL(sheet)
    let rowIndex = idx + 2

    return [
      `=HYPERLINK("${sheetURL}", "${sheetName}")`,
      `=IFERROR(COUNTIF(INDIRECT(JOIN("", "${prefectureSheetPrefix}", $A${rowIndex}, "!E2:E")), "○"), "-")`,
      `=IFERROR(COUNTIFS(INDIRECT(JOIN("", "${prefectureSheetPrefix}", $A${rowIndex}, "!E2:E")), "○", INDIRECT(JOIN("", "${prefectureSheetPrefix}", $A${rowIndex}, "!H2:H")), ">0"), "-")`,
      `=IFERROR($C${rowIndex}/$B${rowIndex}, "-")`
    ]
  })

  statusSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])

  statusSheet.getRange(2, 1, data.length, data[0].length)
    .setFormulas(data)
  
  return  statusSheet
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
  
  let rangeData = getPartSheets().reduce( (obj, sheet: GoogleAppsScript.Spreadsheet.Sheet) => {  
    const range = sheet.getDataRange()

    let filter = range.getFilter()
    if (filter != null) {
      filter.remove()
    }

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
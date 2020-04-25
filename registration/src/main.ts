type IEMail = string
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type RegistrationStatus = '承認' | '却下' | ''

interface IRegistrant {
    name: string
    email: IEMail
    affiliation: string
    registeredAt: Date
    status: RegistrationStatus
}

class Registrant implements IRegistrant {
    name: string
    email: IEMail
    affiliation: string
    status: RegistrationStatus
    registeredAt: Date

    constructor(obj: IRegistrant) {
        this.name = obj.name
        this.email = obj.email
        this.affiliation = obj.affiliation
        this.status = obj.status
        this.registeredAt = obj.registeredAt
    }
}

const NAME_COLUMN_LABEL = '名前（ニックネーム可）'
const EMAIL_COLUMN_LABEL = 'メールアドレス'
const AFFILIATION_COLUMN_LABEL = '所属（任意）'
const TIMESTAMP_COLUMN_LABEL = 'タイムスタンプ'
const STATUS_COLUMN_LABEL = '登録処理'

function getSheet(): Sheet {
    if (typeof this.sheet == 'undefined') {
        this.sheet = SpreadsheetApp.getActiveSheet()
    }

    return this.sheet
}

function getWorkSheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
    if (typeof this.worksheet == 'undefined') {
        const worksheetId = PropertiesService.getScriptProperties().getProperty('WORKSHEET_ID')
        this.worksheet = SpreadsheetApp.openById(worksheetId)
    }

    return this.worksheet
}

function notifyRegistration(registrant: IRegistrant): boolean {
    const { email } = registrant
    const worksheetUrl = getWorkSheet().getUrl()
    let templateFilePath = ''
    let subject = ''

    templateFilePath = 'mail_templates/approved.html'
    subject = '【saveMLAK】COVID-19全国図書館調査へのご参加を承認いたしました'

    // ToDo: リジェクト時のメッセージ送信

    const template = HtmlService.createTemplateFromFile(templateFilePath)
    template.registrant = registrant
    template.worksheetUrl = worksheetUrl
    const body = template.evaluate().getContent()
    
    GmailApp.sendEmail(email, subject, body)

    return true
}

function notifyRegistratonToAdmin(registrants: IRegistrant[]): boolean {
    const adminEmail: IEMail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
    const template = HtmlService.createTemplateFromFile('mail_templates/admin')
    const subject = '【covid-19-libdata】新たに%名を作業者として承認しました。'.replace('%', registrants.length.toString())

    template.registrants = registrants
    const body = template.evaluate().getContent()

    GmailApp.sendEmail(adminEmail, subject, body)

    return true
}

function addCollabolator(registrant: IRegistrant): boolean {
    let spreadsheet = getWorkSheet()
    spreadsheet.addEditor(registrant.email)

    notifyRegistration(registrant)
    notifyRegistratonToAdmin([ registrant ])
    
    return true
}

function addCollabolators(registrants: IRegistrant[]): boolean {
    let spreadsheet = getWorkSheet()
    const emailList: IEMail[] = registrants.map( (registrant) => registrant.email )
    spreadsheet.addEditors(emailList)

    registrants.forEach((registrant) => {
        notifyRegistration(registrant)
    })
    
    notifyRegistratonToAdmin(registrants)

    return true
}

function getHeader(sheet: Sheet): string[] {
    if (typeof this.headers == 'undefined') {
        let dataRange = sheet.getDataRange()
        this.headers = sheet.getDataRange().getValues().shift()
    }

    return this.headers
}

function getRegistrantFromRow(rowValues: string[]) {
    let sheet   = getSheet()
    let headers = getHeader(sheet)

    let nameColumnIndex = headers.indexOf(NAME_COLUMN_LABEL)
    let emailColumnIndex = headers.indexOf(EMAIL_COLUMN_LABEL)
    let affiliationColumnIndex = headers.indexOf(AFFILIATION_COLUMN_LABEL)
    let timestampColumnIndex = headers.indexOf(TIMESTAMP_COLUMN_LABEL)
    let statusColumnIndex = headers.indexOf(STATUS_COLUMN_LABEL)

    let status: RegistrationStatus = <RegistrationStatus> rowValues[statusColumnIndex]

    return new Registrant({
        name: rowValues[nameColumnIndex],
        email: rowValues[emailColumnIndex],
        affiliation: rowValues[affiliationColumnIndex],
        registeredAt: new Date(rowValues[timestampColumnIndex]),
        status: status
    })
}

function updateRegistration(range: GoogleAppsScript.Spreadsheet.Range, status: string): GoogleAppsScript.Spreadsheet.Range {
    let sheet   = getSheet()
    let headers = getHeader(sheet)
    let approvedColumnIndex = headers.indexOf(STATUS_COLUMN_LABEL)
    let values = range.getValues()
    values[0][approvedColumnIndex] = status

    range.setValues(values)

    return range
}

function addCollabolatorFromRow(range: GoogleAppsScript.Spreadsheet.Range): boolean {
    if (range.getNumRows() != 1) {
        throw range.toString() + "is not row"
    }

    let rowValues = range.getValues().flat()
    let registrant = getRegistrantFromRow(rowValues)

    if (addCollabolator(registrant)) {
        updateRegistration(range, '承認')
    } else {
        updateRegistration(range, '却下')
    }

    return true
}

function addCollabolatorFromForm(event: GoogleAppsScript.Events.SheetsOnFormSubmit) {
    return addCollabolatorFromRow(event.range)
}
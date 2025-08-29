/** สำหรับสร้าง web app */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/1170/1170170.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('ระบบบันทึกการกระทำความผิด กองบิน 21')
}

/** สำหรับแยกหน้า html */
function include(f) {
  return HtmlService.createHtmlOutputFromFile(f).getContent()
}

/** สร้างตัวแปรใช้ร่วมกัน */
const ss = SpreadsheetApp.getActive()
const studentDataSheet = ss.getSheetByName('Data')
const logSheet = ss.getSheetByName('log')

/** ฟังก์ชันสำหรับตั้งค่า ID โฟลเดอร์ */
function setFolderId() {
  const folderId = '1Ly-M0hNBJKRC-58RJmrYXsHqN4rdOpVd' // 👈 แก้ไขตรงนี้
  PropertiesService.getScriptProperties().setProperty('idFolder', folderId)
  Logger.log('ตั้งค่า idFolder เรียบร้อยแล้ว: ' + folderId)
}

/** รับข้อมูลหัวตาราง */
function getSchemaFromSheet() {
  const schema = studentDataSheet.getDataRange().getValues()[0]
  console.log(schema)
  return schema
}

/** รับ object รายการ */
function getStudentList() {
  const payload = {}
  const values = studentDataSheet.getDataRange().getDisplayValues()
  const header = values.shift()
  values.forEach(row => {
    const obj = {}
    header.forEach((h, i) => {
      obj[h] = row[i]
    })
    payload[row[0]] = obj
  })
  console.log(payload)
  return payload
}

/** บันทึกข้อมูล */
function saveStudentData(obj) {
  const idFolder = PropertiesService.getScriptProperties().getProperty('idFolder')
  if (!idFolder) throw new Error('ยังไม่ได้ตั้งค่า idFolder ใน Script Properties กรุณารันฟังก์ชัน setFolderId() ก่อน')

  const folder = DriveApp.getFolderById(idFolder)

  const studentId = obj['เลขทะเบียนรถ']
  const values = studentDataSheet.getDataRange().getValues()
  const header = values[0]
  const idColumnIndex = header.indexOf('เลขทะเบียนรถ')
  if (idColumnIndex === -1) throw new Error('ไม่พบเลขทะเบียนรถที่หัวตาราง')

  const rowIndex = values.map(r => r[idColumnIndex]).findIndex(x => x == studentId)
  if (rowIndex === -1) throw new Error('ไม่พบเลขทะเบียนรถในฐานข้อมูล')

  let saveArr = []
  let LogMsg = []

  for (let i = 0; i < header.length; i++) {
    const name = header[i]
    const numberString = ['เลขทะเบียนรถ', 'หมายเลขโทรศัพท์', 'หมายเลขโทรศัพท์ฉุกเฉิน']

    if (name === 'รูปประจำตัว') {
      const newBlob = obj[name]
      if (newBlob && newBlob.getBytes().length > 0) {
        try {
          const files = folder.searchFiles(`title = '${studentId}'`)
          while (files.hasNext()) {
            const fileToDelete = files.next()
            fileToDelete.setTrashed(true)
            LogMsg.push(`ลบไฟล์เก่า: ${fileToDelete.getName()} (${fileToDelete.getId()})`)
          }

          const file = folder.createFile(newBlob).setName(studentId)
          const url = file.getUrl()
          saveArr.push(url)
          LogMsg.push(`อัปโหลดไฟล์ใหม่: ${file.getName()} (${file.getId()})`)
        } catch (e) {
          LogMsg.push(`เกิดข้อผิดพลาดในการจัดการรูปภาพ: ${e.message}`)
          const existingUrl = values[rowIndex][i]
          saveArr.push(existingUrl || '')
        }
      } else {
        const existingUrl = values[rowIndex][i]
        saveArr.push(existingUrl || '')
        LogMsg.push(`ไม่มีการเลือกไฟล์รูปภาพใหม่ ใช้ URL เดิม: ${existingUrl}`)
      }

    } else if (numberString.includes(name)) {
      const value = obj[name] ? "'" + obj[name] : ''
      saveArr.push(value)
    } else {
      const value = obj[name] || ''
      saveArr.push(value)
    }
  }

  studentDataSheet.getRange(rowIndex + 1, 1, 1, saveArr.length).setValues([saveArr])
  return LogMsg.join(', ')
}

/** บันทึก log */
function saveLog(logData) {
  const header = logSheet.getDataRange().getValues()[0]
  const rowData = header.map(h => logData[h] || '')
  logSheet.appendRow(rowData)
  return 'success'
}

function onOpen () {
    var ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Create Docs', 'createNewGoogleDocs');
    menu.addToUi();
}

function createNewGoogleDocs() {
    const googleDocTemplate = DriveApp.getFileById('')
    const sheet = SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('')

    const rows = sheet.getDataRange().getValues();

    rows.shift()
    rows.forEach((row, index) => {
        if (index === 0) return;
        if (row[5]) return;

        const urlIsBlank = sheet.getRange(index + 1, 6).isBlank

        if (urlIsBlank) {
          const copy = googleDocTemplate.makeCopy(`สัญญาจ้างงาน - ${row[0]}`)
          const doc = DocumentApp.openById(copy.getId())
          const body = doc.getBody()

          const date_today_th = Utilities.formatDate(new Date(), "GMT+7", "dd MMMM yyyy")
          const amount_text = ArabicNumberToText(row[3])

          body.replaceText('{{date_today}}', date_today_th)
          body.replaceText('{{customer_name}}', row[0])
          body.replaceText('{{tax_id}}', row[1])
          body.replaceText('{{job_name}}', row[2])
          body.replaceText('{{total_amount}}', row[3].toLocaleString('th'))
          body.replaceText('{{terms}}', row[4])
          body.replaceText('{{amount_text}}', amount_text)

          doc.saveAndClose()
          const url = doc.getUrl()
          sheet.getRange(index + 2, 6).setValue(url)
        }
    })
}

function ArabicNumberToText(Number) {
    var Number = CheckNumber(Number);
    var NumberArray = new Array("ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า", "สิบ");
    var DigitArray = new Array("", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน");
    var DecimalLen = 2;
    var BahtText = "";
    if (isNaN(Number)) {
        return "ข้อมูลนำเข้าไม่ถูกต้อง";
    } else {
        if ((Number - 0) > 9999999.9999) {
            return "ข้อมูลนำเข้าเกินขอบเขตที่ตั้งไว้";
        } else {
            Number = Number.split(".");
            if (Number[1].length > 0) {
                Number[1] = Number[1].substring(0, 2);
            }
            var NumberLen = Number[0].length - 0;
            for (var i = 0; i < NumberLen; i++) {
                var tmp = Number[0].substring(i, i + 1) - 0;
                if (tmp != 0) {
                    if ((i == (NumberLen - 1)) && (tmp == 1)) {
                        BahtText += "เอ็ด";
                    } else
                        if ((i == (NumberLen - 2)) && (tmp == 2)) {
                            BahtText += "ยี่";
                        } else
                            if ((i == (NumberLen - 2)) && (tmp == 1)) {
                                BahtText += "";
                            } else {
                                BahtText += NumberArray[tmp];
                            }
                    BahtText += DigitArray[NumberLen - i - 1];
                }
            }
            BahtText += "บาท";
            if ((Number[1] == "0") || (Number[1] == "00")) {
                BahtText += "ถ้วน";
            } else {
                DecimalLen = Number[1].length - 0;
                for (var i = 0; i < DecimalLen; i++) {
                    var tmp = Number[1].substring(i, i + 1) - 0;
                    if (tmp != 0) {
                        if ((i == (DecimalLen - 1)) && (tmp == 1)) {
                            BahtText += "เอ็ด";
                        } else
                            if ((i == (DecimalLen - 2)) && (tmp == 2)) {
                                BahtText += "ยี่";
                            } else
                                if ((i == (DecimalLen - 2)) && (tmp == 1)) {
                                    BahtText += "";
                                } else {
                                    BahtText += NumberArray[tmp];
                                }
                        BahtText += DigitArray[DecimalLen - i - 1];
                    }
                }
                BahtText += "สตางค์";
            }
            return BahtText;
        }
    }
}

function CheckNumber(Number) {
    var decimal = false;
    Number = Number.toString();
    Number = Number.replace(/ |,|บาท|฿/gi, '');
    for (var i = 0; i < Number.length; i++) {
        if (Number[i] == '.') {
            decimal = true;
        }
    }
    if (decimal == false) {
        Number = Number + '.00';
    }
    return Number
}
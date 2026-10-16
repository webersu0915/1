var SHEET_NAME = "Equipment_List";
var LOG_SHEET_NAME = "Log";

function doGet(e) {
  var equipmentId = e.parameter.id;
  var mode = e.parameter.mode || "borrow"; // borrow 或 return

  if (!equipmentId) {
    return json_(false, "請提供設備 ID，例如：?id=1001&mode=borrow");
  }
  if (mode !== "borrow" && mode !== "return") {
    return json_(false, "不支援的模式：" + mode);
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_(false, "系統忙碌中，請稍後再試");
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return json_(false, "找不到工作表：" + SHEET_NAME);
    }
    var data = sheet.getDataRange().getValues();
    var targetId = String(equipmentId).trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== targetId) continue;

      var status = String(data[i][2]).trim();
      var eqName = data[i][1];
      var eqBrand = data[i][4];
      var title = (eqBrand ? eqBrand + " " : "") + eqName;

      if (mode === "borrow") {
        if (status === "可借用") {
          sheet.getRange(i + 1, 3).setValue("借出中");
          SpreadsheetApp.flush();
          log_(ss, targetId, title, "借出");
          return json_(true, "借用成功！\n" + title + " 狀態已更新為「借出中」。");
        }
        return json_(false, "借用失敗！\n" + title + " 目前狀態為：" + status);
      }

      // mode === "return"
      if (status === "借出中") {
        sheet.getRange(i + 1, 3).setValue("可借用");
        SpreadsheetApp.flush();
        log_(ss, targetId, title, "歸還");
        return json_(true, "歸還成功！\n" + title + " 狀態已更新為「可借用」。");
      }
      if (status === "可借用") {
        return json_(false, "歸還失敗！\n" + title + " 目前並未借出。");
      }
      return json_(false, "歸還失敗！\n" + title + " 目前狀態為：" + status);
    }

    return json_(false, "找不到設備編號：" + equipmentId);
  } catch (err) {
    return json_(false, "系統錯誤：" + err.message);
  } finally {
    lock.releaseLock();
  }
}

// 寫入借閱紀錄（Log 工作表不存在會自動建立）
function log_(ss, id, name, action) {
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["時間", "設備編號", "名稱", "動作"]);
  }
  logSheet.appendRow([new Date(), id, name, action]);
}

function json_(ok, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

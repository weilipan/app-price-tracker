function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📱 追蹤 App 價格")
    .addItem("📄 產生空白表格", "createTrackingSheet")
    .addItem("🔍 查詢 App ID", "searchAppID")
    .addItem("🔄 更新 App 價格並發信", "checkAppPrices")
    .addItem("✉️ 設定通知 Email", "setNotificationEmail")
    .addSeparator()
    .addItem("ℹ️ 關於這個工具", "showAboutInfo")
    .addToUi();
}

// 📄 產生空白表格（先建立必要的表格，再刪除其他不相關的表格）
function createTrackingSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var trackingSheet = ss.getSheetByName("App Tracking");
  var settingsSheet = ss.getSheetByName("Settings");

  if (trackingSheet) {
    var response = ui.alert(
      "⚠️ 已存在「App Tracking」工作表！",
      "是否要清除所有不相關的工作表，並重新建立？\n\n⚠️ 這將 **清除所有數據**，無法還原！",
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.NO) {
      ui.alert("✅ 已取消操作，保留現有工作表。");
      return;
    }
  }

  // **建立或清空 "App Tracking" 工作表**
  if (trackingSheet) {
    trackingSheet.clear();
  } else {
    trackingSheet = ss.insertSheet("App Tracking");
  }

  var headers = ["App ID", "App Name", "目前價格", "前一天價格", "國家"];
  trackingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 設定標題格式
  trackingSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f4b400");

  // **建立或清空 "Settings" 工作表**
  if (settingsSheet) {
    settingsSheet.clear();
  } else {
    settingsSheet = ss.insertSheet("Settings");
  }

  settingsSheet.getRange("A1").setValue("通知 Email").setFontWeight("bold").setBackground("#34a853");

  // **詢問使用者 Email**
  var emailResponse = ui.prompt("📧 請輸入預設的通知 Email：");
  var email = emailResponse.getResponseText().trim();
  if (!email || !email.includes("@")) {
    email = "your-email@gmail.com"; // 預設值
  }
  settingsSheet.getRange("A2").setValue(email);

  // **刪除所有不相關的工作表**
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName !== "App Tracking" && sheetName !== "Settings") {
      ss.deleteSheet(sheet);
    }
  });

  ui.alert("📄 已成功建立「App Tracking」和「Settings」表格，並移除其他不相關的工作表。\n預設 Email: " + email);
}

// 🔄 更新 App 價格並發信
function checkAppPrices() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("App Tracking");
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("⚠️ 找不到 'App Tracking' 表格，請先執行「產生空白表格」。");
    return;
  }

  var emailRecipient = settingsSheet.getRange("A2").getValue(); // 讀取 Email
  var data = sheet.getDataRange().getValues(); 
  var messages = [];  

  for (var i = 1; i < data.length; i++) {  
    var appId = data[i][0];  
    var appName = data[i][1];  
    var prevPrice = data[i][2];  
    var country = data[i][4];  
    
    if (!appId || !country) continue;  

    var apiUrl = "https://itunes.apple.com/lookup?id=" + appId + "&country=" + country;
    try {
      var response = UrlFetchApp.fetch(apiUrl);
      var json = JSON.parse(response.getContentText());
      
      if (json.resultCount > 0) {
        var newPrice = json.results[0].price;  

        sheet.getRange(i + 1, 4).setValue(prevPrice);  
        sheet.getRange(i + 1, 3).setValue(newPrice);   

        if (newPrice < prevPrice) {
          messages.push(appName + " 降價了！\n原價: $" + prevPrice + " -> 現價: $" + newPrice);
        }
      }
    } catch (e) {
      Logger.log("⚠️ API 請求失敗: " + e.toString());
    }
  }
  
  if (messages.length > 0) {
    var subject = "📢 App Store 降價通知";
    var body = messages.join("\n\n");
    MailApp.sendEmail(emailRecipient, subject, body);
    SpreadsheetApp.getUi().alert("📧 已發送降價通知至 " + emailRecipient);
  } else {
    SpreadsheetApp.getUi().alert("✅ 今日沒有 App 降價。");
  }
}

// ✉️ 設定通知 Email
function setNotificationEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("📧 請輸入新的通知 Email：");

  if (response.getSelectedButton() == ui.Button.OK) {
    var newEmail = response.getResponseText();
    var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    
    if (!settingsSheet) {
      settingsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings");
      settingsSheet.getRange("A1").setValue("通知 Email").setFontWeight("bold").setBackground("#34a853");
    }

    settingsSheet.getRange("A2").setValue(newEmail);
    ui.alert("✅ 通知 Email 已更新為：" + newEmail);
  }
}

// 🔍 查詢 App ID（允許直接輸入 App ID，避免 API 搜尋不到）
function searchAppID() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("🔍 請輸入 App 名稱 或 App ID 和 國家代碼（例如：XCOM 2, tw 或 1458655678, tw）：");

  if (response.getSelectedButton() == ui.Button.OK) {
    var input = response.getResponseText().split(",");
    if (input.length < 2) {
      ui.alert("⚠️ 格式錯誤，請輸入「App 名稱 或 App ID」和「國家代碼」，如「XCOM 2, tw」或「1458655678, tw」。");
      return;
    }

    var query = input[0].trim();
    var country = input[1].trim();
    var apiUrl;

    if (!isNaN(query)) {  // **如果輸入的是數字，則視為 App ID，直接使用 lookup API**
      apiUrl = "https://itunes.apple.com/lookup?id=" + query + "&country=" + country;
    } else {  // **如果輸入的是名稱，則使用 search API**
      apiUrl = "https://itunes.apple.com/search?term=" + encodeURIComponent(query) + "&country=" + country + "&entity=software&limit=5";
    }

    try {
      var response = UrlFetchApp.fetch(apiUrl);
      var json = JSON.parse(response.getContentText());

      if (json.resultCount > 0) {
        var chosenApp;
        if (json.resultCount == 1) {
          chosenApp = json.results[0];  // 只有一個結果時，直接選擇
        } else {
          // **如果有多個符合條件的 App，讓使用者選擇**
          var appNames = json.results.map(function(app, index) {
            return (index + 1) + ". " + app.trackName;
          }).join("\n");

          var selection = ui.prompt(
            "🔍 找到多個符合條件的 App，請輸入選擇的編號：\n" + appNames
          );

          var selectedIndex = parseInt(selection.getResponseText()) - 1;
          if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= json.resultCount) {
            ui.alert("⚠️ 選擇無效，請重新查詢。");
            return;
          }

          chosenApp = json.results[selectedIndex];
        }

        var appId = chosenApp.trackId;
        var appStoreName = chosenApp.trackName;
        var appPrice = chosenApp.price || "N/A";

        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("App Tracking");
        if (!sheet) {
          ui.alert("⚠️ 找不到 'App Tracking' 表格，請先執行「產生空白表格」。");
          return;
        }

        var lastRow = sheet.getLastRow() + 1;
        sheet.getRange(lastRow, 1).setValue(appId);
        sheet.getRange(lastRow, 2).setValue(appStoreName);
        sheet.getRange(lastRow, 3).setValue(appPrice);
        sheet.getRange(lastRow, 4).setValue("N/A");
        sheet.getRange(lastRow, 5).setValue(country);

        ui.alert("✅ 查詢成功！已將「" + appStoreName + "」加入追蹤列表。\nApp ID：" + appId + "\n價格：" + appPrice);
      } else {
        ui.alert("⚠️ 找不到該 App，請確認名稱或 App ID 是否正確。");
      }
    } catch (e) {
      Logger.log("⚠️ API 請求失敗: " + e.toString());
      ui.alert("❌ API 請求失敗，請稍後再試。");
    }
  }
}

// ℹ️ 顯示關於工具
// ℹ️ 更新關於工具的說明
function showAboutInfo() {
  SpreadsheetApp.getUi().alert(
    "📱 追蹤 App 價格工具\n\n" +
    "📄 **產生空白表格**\n" +
    "  - 建立「App Tracking」和「Settings」表格\n" +
    "  - 詢問並設定預設通知 Email\n" +
    "  - 自動刪除其他不相關的工作表\n\n" +
    
    "🔍 **查詢 App ID**\n" +
    "  - 允許輸入 **App 名稱 或 App ID** 來查詢\n" +
    "  - 如果找到多個 App，讓使用者選擇正確的應用\n" +
    "  - 自動將選擇的 App ID、新增到「App Tracking」表格\n\n" +

    "🔄 **更新 App 價格並發信**\n" +
    "  - 每天自動從 App Store API 取得最新價格\n" +
    "  - 若價格下降，發送 Email 通知\n\n" +

    "✉️ **設定通知 Email**\n" +
    "  - 允許手動修改通知 Email，存入「Settings」表格\n\n" +

    "⚙️ **自動化功能**\n" +
    "  - 可設定每日自動更新價格，並在降價時發送 Email\n\n" +

    "🔹 **由 Google Apps Script 自動運行**"
  );
}



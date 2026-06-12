/**
 * 當試算表打開時，自動在上方工具列新增自訂按鈕選單
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚾️ MLB工具')
    .addItem('📊 更新球員賽季數據 (P/B)', 'updatePlayerStatsAction')
    .addSeparator()
    .addItem('🔄 一鍵更新【基本資訊】', 'updateAll')
    .addSeparator()
    .addItem('⚡️ 局部更新：最新異動', 'updateTransactions')
    .addItem('⚡️ 局部更新：狀態與球隊', 'updateStatusAndTeam')
    .addSeparator()
    .addItem('⚙️ 建立/重設【郵件寄送管理中心】', 'setupEmailCenter')
    .addItem('📧 寄送測試信 (對準有勾選【立即寄送】的用戶)', 'sendTestEmailAction')
    .addSeparator()
    .addItem('⏰ 啟用：每日早上 7 點自動更新與發信', 'enableDailyTrigger')
    .addItem('🛑 關閉：自動更新排程', 'disableDailyTrigger')
    .addToUi();
}

/**
 * 總排程核心：每日早上 7 點（自動）或手動測試信（手動）皆執行此函式
 * @param {boolean} isDaily 是否為每日定時自動更新（若為物件或未傳入，則預設視為每日自動）
 */
function performDailyAutomation(isDaily) {
  var targetDaily = (isDaily === false) ? false : true;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  initLogSheet(ss);
  initHistorySheet(ss);
  initEmailCenterSheet(ss); 
  
  writeLog("INFO", "=== 開始執行 [MLB自動更新與發信總排程] (模式: " + (targetDaily ? "每日自動" : "手動測試") + ") ===");
  
  // 1. 更新 台灣球員清單 (基本資訊與異動)
  if (ss.getSheetByName("台灣球員清單")) {
    executeCoreUpdate("台灣球員清單", "ALL", "台灣球員清單-基本資訊");
  } else {
    writeLog("WARN", "找不到「台灣球員清單」工作表，跳過基本資訊更新。");
  }
  
  // 2. 更新 投手賽季數據
  var pitcherSheet = ss.getSheetByName("Pitcher");
  if (pitcherSheet) {
    processStatSheet(pitcherSheet, "Pitcher");
  }
  
  // 3. 更新 打者賽季數據
  var batterSheet = ss.getSheetByName("Batter");
  if (batterSheet) {
    processStatSheet(batterSheet, "Batter");
  }
  
  // 4. 依據模式撈取對應名單並發送電子郵件
  sendSummaryEmail(ss, targetDaily);
  
  writeLog("INFO", "=== [MLB自動更新與發信總排程] 執行完畢 ===");
}

/**
 * 手動按鈕：寄送測試信（僅針對勾選了「⚡ 立即寄送」的收件人）
 */
function sendTestEmailAction() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("確認執行", "這將會花費1~3分鐘更新所有球員資料，並在完成後立即發送摘要郵件給有勾選【⚡ 立即寄送】的用戶。\n發送成功後，系統會自動取消勾選其「立即寄送」欄位。是否繼續？", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    performDailyAutomation(false);
    ui.alert("⚾️ 測試信更新與發送程序已執行完畢！詳細進程請見 「Log」 工作表與您的管理員信箱。");
  }
}

/**
 * 使用者編輯試算表時自動觸發 (即時檢查 Email 格式、除錯並標記)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.source.getActiveSheet();
  
  if (sheet.getName() !== "郵件寄送管理中心") return;
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  if (col === 4 && row >= 4) {
    var value = range.getValue();
    var statusCell = sheet.getRange(row, 6); 
    
    if (value) {
      var originalVal = String(value);
      var cleaned = originalVal.trim().replace(/\s+/g, "");
      
      cleaned = cleaned.replace(/@gamail\.com/i, "@gmail.com")
                       .replace(/@gamil\.com/i, "@gmail.com")
                       .replace(/@gmai\.com/i, "@gmail.com")
                       .replace(/@gmail\.con/i, "@gmail.com");
                       
      if (cleaned !== originalVal) {
        range.setValue(cleaned);
      }
      
      var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(cleaned)) {
        statusCell.setValue("❌ 格式錯誤，請檢查！").setFontColor("red");
        range.setBackground("#FFD2D2");
      } else {
        statusCell.setValue("✅ 格式正確").setFontColor("green");
        range.setBackground(null);
      }
    } else {
      statusCell.setValue("").setFontColor("black");
      range.setBackground(null);
    }
  }
}

// ==========================================================
// 📧 郵件發送與摘要資料擷取核心模組
// ==========================================================

function sendSummaryEmail(ss, isDaily) {
  var recipientSheet = ss.getSheetByName("郵件寄送管理中心");
  if (!recipientSheet) {
    writeLog("ERROR", "未找到「郵件寄送管理中心」工作表，取消發信程序。");
    return;
  }
  
  var adminEmail = String(recipientSheet.getRange("B1").getValue()).trim();
  var recipientData = recipientSheet.getDataRange().getValues();
  
  var activeSendEmails = [];       
  var rowsToUpdate = [];           
  var skippedWithErrorRecipients = []; 
  
  if (recipientData.length > 3) {
    for (var i = 3; i < recipientData.length; i++) {
      var isDailyChecked = recipientData[i][0];     
      var isImmediateChecked = recipientData[i][1]; 
      var name = String(recipientData[i][2]).trim(); 
      var email = String(recipientData[i][3]).trim(); 
      var currentStatus = String(recipientData[i][5]).trim(); 
      
      var isTargeted = isDaily ? isDailyChecked : isImmediateChecked;
      
      if (isTargeted === true) {
        if (!email || currentStatus.indexOf("❌") !== -1) {
          skippedWithErrorRecipients.push("👤 " + (name || "未填姓名") + " [" + (email || "未填信箱") + "] → 原因：" + (currentStatus || "欄位缺失"));
          continue; 
        }
        
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
          recipientSheet.getRange(i + 1, 6).setValue("❌ 格式異常").setFontColor("red");
          skippedWithErrorRecipients.push("👤 " + name + " [" + email + "] → 原因：格式異常");
          continue;
        }
        
        activeSendEmails.push(email);
        rowsToUpdate.push(i + 1);
      }
    }
  }
  
  var htmlBody = generateHtmlReport(ss);
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if (activeSendEmails.length === 0 && skippedWithErrorRecipients.length === 0) {
    if (adminEmail) {
      writeLog("INFO", "未偵測到收件人，系統自動切換為 [管理者個人追蹤模式] 發送。");
      
      var adminHtmlWrapper = '<div style="font-family: Arial, sans-serif; background-color: #e8f0fe; padding: 15px; border-radius: 6px; border: 1px solid #1a73e8; margin-bottom: 20px;">';
      adminHtmlWrapper += '<h3 style="color: #174ea6; margin-top: 0;">👑 MLB 台灣球員每日動態（管理者個人追蹤模式）</h3>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>🔹 運行模式：</b>' + (isDaily ? '⏰ 每日早上自動排程更新' : '⚡ 管理者手動即時測試信') + '</p>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>🔹 統計時間：</b>' + nowStr + '</p>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 12px; color: #5f6368;">💡 提示：目前下方收件者清單為空，系統自動將今日戰報直接寄給您供個人追蹤使用。</p>';
      adminHtmlWrapper += '</div>';
      
      MailApp.sendEmail({
        to: adminEmail,
        subject: "⚾️ [個人追蹤] MLB 台灣球員每日動態與近五日出賽摘要報告",
        htmlBody: adminHtmlWrapper + htmlBody
      });
      return;
    } else {
      writeLog("WARN", "未偵測到任何收件人且無管理者信箱，取消發信。");
      return;
    }
  }
  
  if (activeSendEmails.length === 0 && skippedWithErrorRecipients.length > 0) {
    writeLog("WARN", "有用戶勾選發送，但因皆包含錯誤，發信程序被系統全數攔截。");
    if (adminEmail) {
      var warnBody = "⚠️ 【MLB系統發信中斷警告】\n\n系統發現有用戶被勾選發送（模式: " + (isDaily ? "每日自動" : "手動測試") + "），但因其資料或信箱格式原本就存在錯誤，系統為確保安全已全數自動攔截，今日報告並未發送出去。\n\n請管理員撥冗回到試算表進行修正：\n\n【❌ 被系統攔截的錯誤名單】\n" + skippedWithErrorRecipients.join("\n");
      MailApp.sendEmail(adminEmail, "⚠️ [警告] MLB 報告發送中斷（名單內含格式錯誤）", warnBody);
    }
    return;
  }

  try {
    MailApp.sendEmail({
      bcc: activeSendEmails.join(","),
      subject: "⚾️ MLB 台灣球員每日動態與近五日出賽摘要報告",
      htmlBody: htmlBody
    });
    
    rowsToUpdate.forEach(function(row) {
      recipientSheet.getRange(row, 5).setValue(nowStr); 
      recipientSheet.getRange(row, 6).setValue("✅ 發送成功").setFontColor("green"); 
      if (!isDaily) {
        recipientSheet.getRange(row, 2).setValue(false); 
      }
    });
    writeLog("INFO", "摘要郵件已成功群發至 " + activeSendEmails.length + " 位收件人。");
    
    if (adminEmail) {
      var adminHtmlWrapper = '<div style="font-family: Arial, sans-serif; background-color: #f1f3f4; padding: 15px; border-radius: 6px; border: 1px solid #dadce0; margin-bottom: 20px;">';
      adminHtmlWrapper += '<h3 style="color: #202124; margin-top: 0;">👑 MLB 系統發送成功報告（管理者審查頁面）</h3>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>🔹 發送類型：</b>' + (isDaily ? '⏰ 每日早上自動排程更新' : '⚡ 管理者手動即時測試信') + '</p>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>🔹 發送時間：</b>' + nowStr + '</p>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>🔹 成功發送數量：</b>' + activeSendEmails.length + ' 人</p>';
      adminHtmlWrapper += '<p style="margin: 5px 0; font-size: 13px;"><b>📬 收件人清單：</b><span style="color: #5f6368; font-size: 12px;">' + activeSendEmails.join(', ') + '</span></p>';
      
      if (skippedWithErrorRecipients.length > 0) {
        adminHtmlWrapper += '<div style="margin-top: 10px; padding: 10px; background-color: #feefe3; border-left: 4px solid #e67e22; border-radius: 4px; font-size: 13px; color: #b06000;">';
        adminHtmlWrapper += '<b>⚠️ 警告：以下用戶因資料/信箱格式原先存在錯誤，已被系統安全攔截（今日未收到信）：</b><br>';
        adminHtmlWrapper += skippedWithErrorRecipients.join('<br>');
        adminHtmlWrapper += '</div>';
      }
      adminHtmlWrapper += '</div>';
      adminHtmlWrapper += '<hr style="border: 0; border-top: 2px dashed #ccc; margin: 30px 0;">';
      adminHtmlWrapper += '<p style="color: #5f6368; font-size: 12px; margin-bottom: 5px;">👇 以下為「發送給收件人的實際郵件內容」供您檢視品質：</p>';
      
      MailApp.sendEmail({
        to: adminEmail,
        subject: "✅ [成功] MLB 報告發送通知 " + (skippedWithErrorRecipients.length > 0 ? "(含欄位錯誤警告)" : "(內容審查)"),
        htmlBody: adminHtmlWrapper + htmlBody
      });
    }
    
  } catch(e) {
    writeLog("ERROR", "郵件發送失敗: " + e.toString());
    rowsToUpdate.forEach(function(row) {
      recipientSheet.getRange(row, 6).setValue("❌ 發送失敗").setFontColor("red");
    });
    
    if (adminEmail) {
      var errorBody = "⚠️ 【MLB 系統發送嚴重失敗報告】\n\n郵件伺服器在群發過程中回報了錯誤...\n\n【🛑 錯誤原因与代碼】\n" + e.toString() + "\n\n【📬 本次嘗試發送的名單】\n" + activeSendEmails.join("\n");
      MailApp.sendEmail(adminEmail, "❌ [嚴重錯誤] MLB 報告發送失敗，請管理者進場檢查", errorBody);
    }
  }
}

/**
 * ⚙️ 功能選單：初始化或重設管理中心
 */
function setupEmailCenter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "郵件寄送管理中心";
  var oldSheet = ss.getSheetByName(sheetName);
  var ui = SpreadsheetApp.getUi();
  
  if (oldSheet) {
    var response = ui.alert("重設確認", "偵測到已存在「" + sheetName + "」工作表。\n若執行重設，系統會刪除此分頁並重新建構全新的「雙核取方塊」架構（舊的名單將會消失）。是否確認重設？", ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
    ss.deleteSheet(oldSheet);
  }
  
  initEmailCenterSheet(ss);
  ui.alert("✅ 「郵件寄送管理中心」已全新建置完畢！請切換至該分頁填寫設定。");
}

function initEmailCenterSheet(ss) {
  var sheetName = "郵件寄送管理中心";
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    sheet.getRange("A1").setValue("👑 管理者信箱：").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange("B1").setValue(Session.getActiveUser().getEmail()).setFontWeight("bold").setFontColor("#1a73e8");
    sheet.getRange("C1").setValue("⬅️ 所有系統通知、成功通報（含內文審查）、錯誤攔截警報都會寄到這個信箱");
    
    sheet.getRange("A2:F2").setBackground("#f3f3f3"); 
    
    var headers = ["⏰ 每日自動", "⚡ 立即寄送", "👤 收件人姓名", "📧 收件人信箱", "📅 最近發送時間", "📝 系統狀態與備註"];
    sheet.getRange("A3:F3").setValues([headers]).setFontWeight("bold").setBackground("#d2e3fc").setHorizontalAlignment("center");
    
    sheet.getRange("A4").insertCheckboxes().check();
    sheet.getRange("B4").insertCheckboxes().check();
    sheet.getRange("C4").setValue("管理員(自試)");
    sheet.getRange("D4").setValue(Session.getActiveUser().getEmail());
    
    sheet.setColumnWidth(1, 95);  
    sheet.setColumnWidth(2, 95);  
    sheet.setColumnWidth(3, 120); 
    sheet.setColumnWidth(4, 260); 
    sheet.setColumnWidth(5, 160); 
    sheet.setColumnWidth(6, 280); 
    
    var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange("A4:A100").setDataValidation(checkboxRule).setHorizontalAlignment("center");
    sheet.getRange("B4:B100").setDataValidation(checkboxRule).setHorizontalAlignment("center");
    
    sheet.getRange("D4:D100").setNumberFormat('@');
  }
}

// ==========================================================
// HTML 生成與網頁爬蟲核心模組
// ==========================================================

function generateHtmlReport(ss) {
  var html = '<!DOCTYPE html><html><head>';
  html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">';
  html += '<style>';
  html += 'body { font-family: "Microsoft JhengHei", Arial, sans-serif; color: #333; line-height: 1.6; margin: 0; padding: 10px; background-color: #f9f9f9;}';
  html += '.email-container { max-width: 900px; margin: 0 auto; background-color: #ffffff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }';
  html += 'h2 { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 8px; margin-top: 0; }';
  html += 'h3 { padding: 8px 12px; margin-top: 25px; border-radius: 4px; }';
  html += '.section-tx { color: #d93025; background-color: #fce8e6; border-left: 5px solid #d93025; }';
  html += '.section-stats { color: #1e8e3e; background-color: #e6f4ea; border-left: 5px solid #1e8e3e; }';
  html += '.table-wrapper { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; margin-bottom: 25px; border: 1px solid #dadce0; border-radius: 4px; box-shadow: inset 0 0 5px rgba(0,0,0,0.05); }';
  html += 'table { width: 100%; border-collapse: collapse; min-width: 800px; font-size: 13px; white-space: nowrap; }';
  html += 'th, td { padding: 8px 10px; border: 1px solid #dadce0; text-align: center; }';
  html += 'th { background-color: #f1f3f4; font-weight: bold; color: #444; }';
  html += 'tr:nth-child(even) { background-color: #fcfcfc; }';
  html += '.sticky-col { position: -webkit-sticky; position: sticky; left: 0; background-color: #ffffff; z-index: 1; border-right: 2px solid #9aa0a6; font-weight: bold; box-shadow: 2px 0 4px -2px rgba(0,0,0,0.15); }';
  html += 'th.sticky-col { background-color: #f1f3f4; }';
  html += '</style></head><body>';
  html += '<div class="email-container">';
  html += '<h2>⚾️ MLB 台灣球員每日動態摘要</h2>';
  
  html += '<h3 class="section-tx">🚨 近期異動球員通知</h3>';
  var rosterSheet = ss.getSheetByName("台灣球員清單");
  var txCount = 0;
  
  if (rosterSheet) {
    var rHeaders = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    var colName = rHeaders.indexOf("姓名") + 1;
    var colDate = rHeaders.indexOf("異動日期") + 1;
    var colDesc = rHeaders.indexOf("異動內容") + 1;
    var rLastRow = rosterSheet.getLastRow();
    
    if (rLastRow > 1 && colDate > 0 && colDesc > 0) {
      var rValues = rosterSheet.getRange(2, 1, rLastRow - 1, rosterSheet.getLastColumn()).getValues();
      var rColors = rosterSheet.getRange(2, colDate, rLastRow - 1, 1).getBackgrounds();
      
      html += '<ul style="padding-left: 20px;">';
      for (var i = 0; i < rValues.length; i++) {
        if (rColors[i][0].toUpperCase() === "#FFD2D2") {
          txCount++;
          var name = colName > 0 ? rValues[i][colName - 1] : "未知";
          var date = "";
          if (rValues[i][colDate - 1] instanceof Date) {
            date = Utilities.formatDate(rValues[i][colDate - 1], Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            date = String(rValues[i][colDate - 1]);
          }
          var desc = rValues[i][colDesc - 1];
          html += '<li style="margin-bottom: 10px; font-size: 14px;"><strong>' + name + '</strong> <span style="color:#5f6368; font-size:12px;">(' + date + ')</span><br><span style="color: #333;">↳ ' + desc + '</span></li>';
        }
      }
      html += '</ul>';
    }
  }
  if (txCount === 0) html += '<p style="color: #666; padding-left: 10px;">本日尚無近期球員異動紀錄。</p>';

  html += '<h3 class="section-stats">🔥 近五天有出賽球員統計</h3>';
  html += '<p style="font-size:12px; color:#666; margin-top:-5px; margin-bottom:15px;">👉 手機用戶請於表格區域<strong style="color:#1a73e8;">左右滑動</strong>查看完整數據。</p>';

  var sheetsConfig = [
    { sheetName: "Pitcher", title: "⚾ 投手 (Pitchers)", fields: ["W", "L", "ERA", "G", "GS", "SV", "IP", "SO", "WHIP"] },
    { sheetName: "Batter", title: "🏏 打者 (Batters)", fields: ["AB", "R", "H", "HR", "RBI", "SB", "AVG", "OBP", "OPS"] }
  ];

  var totalPlayCount = 0;

  sheetsConfig.forEach(function(config) {
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colName = headers.indexOf("姓名") + 1;
    var colEng = headers.indexOf("英文名字") + 1;
    var colPos = headers.indexOf("守備位置") + 1;
    var colTeam = headers.indexOf("所屬球隊") + 1;
    var colLvl = headers.indexOf("球隊層級") + 1;

    var baseCols = {}, deltaCols = {};
    config.fields.forEach(function(f) {
      var count = 0;
      headers.forEach(function(h, idx) {
        if (h.trim() === f) {
          count++;
          if (count === 1) baseCols[f] = idx + 1;
          if (count === 2) deltaCols[f] = idx + 1;
        }
      });
    });

    var checkField = (config.sheetName === "Pitcher") ? "G" : "AB";
    var colDeltaCheck = deltaCols[checkField] || -1;
    
    if (colDeltaCheck !== -1 && sheet.getLastRow() > 1) {
      var dataRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      var activePlayersHtml = "";

      for (var i = 0; i < dataRows.length; i++) {
        var rowData = dataRows[i];
        var deltaCheckVal = String(rowData[colDeltaCheck - 1]).trim();
        
        if (deltaCheckVal !== "No Game" && deltaCheckVal !== "累積中..." && deltaCheckVal !== "—" && deltaCheckVal !== "") {
          totalPlayCount++;
          var pName = colName > 0 ? rowData[colName - 1] : "—";
          var pEng = colEng > 0 ? rowData[colEng - 1] : "—";
          var pPos = colPos > 0 ? rowData[colPos - 1] : "—";
          var pTeam = colTeam > 0 ? rowData[colTeam - 1] : "—";
          var pLvl = colLvl > 0 ? rowData[colLvl - 1] : "—";

          activePlayersHtml += '<tr>';
          activePlayersHtml += '<td class="sticky-col">' + pName + '</td>';
          activePlayersHtml += '<td style="color:#5f6368; font-size:12px;">' + pEng + '</td>';
          activePlayersHtml += '<td>' + pPos + '</td>';
          activePlayersHtml += '<td style="text-align:left;">' + pTeam + ' <span style="font-size:11px; color:#fff; background:#80868b; padding:2px 5px; border-radius:3px; margin-left:4px;">' + pLvl + '</span></td>';

          config.fields.forEach(function(f) {
            var val = baseCols[f] ? String(rowData[baseCols[f] - 1]).trim() : "—";
            activePlayersHtml += '<td style="color:#1a73e8;">' + val + '</td>';
          });
          
          config.fields.forEach(function(f) {
            var val = deltaCols[f] ? String(rowData[deltaCols[f] - 1]).trim() : "—";
            activePlayersHtml += getEmailDeltaCellHtml(f, val);
          });

          activePlayersHtml += '</tr>';
        }
      }

      if (activePlayersHtml !== "") {
        html += '<h4 style="color: #333; margin-top: 20px; margin-bottom: 8px; font-size: 16px;">' + config.title + '</h4>';
        html += '<div class="table-wrapper">';
        html += '<table>';
        html += '<thead>';
        html += '<tr>';
        html += '<th rowspan="2" class="sticky-col">姓名</th>';
        html += '<th rowspan="2">英文名字</th>';
        html += '<th rowspan="2">守備</th>';
        html += '<th rowspan="2">球隊 (層級)</th>';
        html += '<th colspan="' + config.fields.length + '" style="background-color:#d2e3fc; color:#174ea6;">當季最新數據</th>';
        html += '<th colspan="' + config.fields.length + '" style="background-color:#ceead6; color:#0d652d;">近五天數據變化</th>';
        html += '</tr><tr>';
        config.fields.forEach(function(f) { html += '<th style="background-color:#e8f0fe; font-size:12px;">' + f + '</th>'; });
        config.fields.forEach(function(f) { html += '<th style="background-color:#e6f4ea; font-size:12px;">' + f + '</th>'; });
        html += '</tr></thead><tbody>';
        html += activePlayersHtml;
        html += '</tbody></table></div>';
      }
    }
  });

  if (totalPlayCount === 0) {
    html += '<p style="padding: 15px; text-align: center; color: #666; border: 1px dashed #dadce0; background-color: #f8f9fa; border-radius: 4px;">近五天內無球員出賽或資料尚在累積中。</p>';
  }

  html += '<p style="font-size: 11px; color: #9aa0a6; margin-top: 35px; border-top: 1px solid #eee; padding-top: 10px; text-align: center;">此郵件為 MLB 試算表系統自動統整發送，請勿直接回覆。</p>';
  html += '</div></body></html>';
  return html;
}

/**
 * 💡 修正點：信件生成不再疊加符號，直接讀取已美化完成的字串
 */
function getEmailDeltaCellHtml(field, deltaValStr) {
  if (!deltaValStr || deltaValStr === "No Game" || deltaValStr === "累積中..." || deltaValStr === "—") {
    return '<td style="color: #aaa; font-size: 12px;">' + (deltaValStr || "—") + '</td>';
  }
  
  // 濾除符號抽取純數字，用以辨別表格背景高亮顏色
  var cleanStr = deltaValStr.replace(/[^0-9.-]/g, "");
  var delta = parseFloat(cleanStr);
  
  if (isNaN(delta) || delta === 0) return '<td style="color: #5f6368;">' + deltaValStr + '</td>';
  
  var isReverseGood = ["L", "ERA", "WHIP"].includes(field); 
  var isGood = isReverseGood ? (delta < 0) : (delta > 0);
  var bgColor = isGood ? "#e6f4ea" : "#fef7e0"; 
  var textColor = isGood ? "#0d652d" : "#b06000"; 
  return '<td style="background-color: ' + bgColor + '; color: ' + textColor + '; font-weight: bold;">' + deltaValStr + '</td>';
}

/**
 * 按鈕觸發函式：更新數據統計分頁 (Pitcher & Batter)
 */
function updatePlayerStatsAction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToUpdate = ["Pitcher", "Batter"];
  
  initLogSheet(ss);
  initHistorySheet(ss); 
  writeLog("INFO", "=== 開始執行 [球員賽季數據] 更新程序 ===");

  sheetsToUpdate.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      processStatSheet(sheet, sheetName);
    } else {
      writeLog("WARN", "找不到分頁: " + sheetName + "，跳過。");
    }
  });

  SpreadsheetApp.getUi().alert("⚾️ 球員賽季數據更新完成！詳細執行進程請見 「Log」 工作表。");
}

function processStatSheet(sheet, type) {
  var ss = sheet.getParent();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statFields = (type === "Pitcher") ? 
    ["W", "L", "ERA", "G", "GS", "SV", "IP", "SO", "WHIP"] : 
    ["AB", "R", "H", "HR", "RBI", "SB", "AVG", "OBP", "OPS"];

  var colMap = {}; var baseCols = {}; var deltaCols = {};
  headers.forEach((h, i) => {
    var headerName = h.trim();
    if (statFields.includes(headerName)) {
      if (!baseCols[headerName]) { baseCols[headerName] = i + 1; } else { deltaCols[headerName] = i + 1; }
    } else { colMap[headerName] = i + 1; }
  });

  var timeCol = colMap["最新數據異動時間"] || colMap["資料更新時間"];
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var currentYear = new Date().getFullYear();

  for (var row = 2; row <= lastRow; row++) {
    var engName = String(sheet.getRange(row, colMap["英文名字"]).getValue()).trim();
    if (!engName) continue;

    writeLog("INFO", "正在處理 [" + type + "] 球員: " + engName);

    var playerId = colMap["Player ID"] ? String(sheet.getRange(row, colMap["Player ID"]).getValue()).trim() : "";
    var milbUrl = colMap["MiLB頁面"] ? String(sheet.getRange(row, colMap["MiLB頁面"]).getValue()).trim() : "";
    var mlbUrl = colMap["MLB頁面"] ? String(sheet.getRange(row, colMap["MLB頁面"]).getValue()).trim() : "";

    if (!playerId || !milbUrl || !mlbUrl) {
      var apiResult = fetchPlayerIdFromAPI(engName);
      if (apiResult) {
        if (colMap["Player ID"]) sheet.getRange(row, colMap["Player ID"]).setValue(apiResult.id);
        if (colMap["MiLB頁面"]) sheet.getRange(row, colMap["MiLB頁面"]).setValue(apiResult.milbUrl);
        if (colMap["MLB頁面"]) sheet.getRange(row, colMap["MLB頁面"]).setValue(apiResult.mlbUrl);
        milbUrl = apiResult.milbUrl; mlbUrl = apiResult.mlbUrl;
      }
    }

    var teamLevel = colMap["球隊層級"] ? String(sheet.getRange(row, colMap["球隊層級"]).getValue()).trim() : "";
    var targetUrl = milbUrl;
    if (teamLevel.toUpperCase().indexOf("MLB") !== -1 || teamLevel.toUpperCase().indexOf("MAJOR") !== -1) {
      if (mlbUrl && mlbUrl.indexOf("http") === 0) targetUrl = mlbUrl;
    }
    if (!targetUrl) continue;

    var stats = scrapeDetailedStats(targetUrl, type, teamLevel, currentYear);
    if (stats) {
      for (var field in baseCols) {
        if (stats[field] !== undefined) updateCellOnly(sheet, row, baseCols[field], stats[field]);
      }
      if (timeCol) sheet.getRange(row, timeCol).setValue(todayStr);

      saveToHistorySheet(ss, todayStr, engName, type, stats);
      var oldStats = getStatsFromHistory(ss, engName, type, todayStr);

      var hasPlayed = true;
      if (oldStats) {
        if (type === "Pitcher") {
          if (stats["G"] !== "今年賽季尚無出賽紀錄" && stats["G"] !== "—") {
            var currentG = parseInt(stats["G"]) || 0;
            var oldG = parseInt(oldStats["G"]) || 0;
            if (currentG - oldG === 0) hasPlayed = false;
          }
        } else if (type === "Batter") {
          if (stats["AB"] !== "今年賽季尚無出賽紀錄" && stats["AB"] !== "—") {
            var currentAB = parseInt(stats["AB"]) || 0;
            var oldAB = parseInt(oldStats["AB"]) || 0;
            if (currentAB - oldAB === 0) hasPlayed = false;
          }
        }
      } else { 
        hasPlayed = null; 
      }

      for (var field in deltaCols) {
        var deltaColIdx = deltaCols[field];
        if (!deltaColIdx) continue;

        if (hasPlayed === false || stats[field] === "今年賽季尚無出賽紀錄" || stats[field] === "—") {
          applyDeltaColor(sheet, row, deltaColIdx, field, "No Game");
        } else if (hasPlayed === null) {
          applyDeltaColor(sheet, row, deltaColIdx, field, "累積中...");
        } else {
          var deltaValue = calculateDelta(field, stats[field], oldStats[field]);
          applyDeltaColor(sheet, row, deltaColIdx, field, deltaValue);
        }
      }
    }
    Utilities.sleep(400); 
  }
}

function isWithinFiveDays(dateStr) {
  if (!dateStr || dateStr === "未找到" || dateStr === "—") return false;
  try {
    var parsedDate = new Date(dateStr);
    if (isNaN(parsedDate.getTime())) return false;
    var today = new Date();
    today.setHours(0, 0, 0, 0); parsedDate.setHours(0, 0, 0, 0);
    var dayDiff = (today.getTime() - parsedDate.getTime()) / (1000 * 3600 * 24);
    return (dayDiff >= -1 && dayDiff <= 5);
  } catch (e) { return false; }
}

/**
 * 💡 修正點：集中格式化邏輯。絕對數字加單一正號，百分比數值（WHIP, AVG, OBP, OPS）計算增減百分比。
 */
function calculateDelta(field, currVal, oldVal) {
  if (currVal === "今年賽季尚無出賽紀錄" || currVal === "—" || oldVal === "今年賽季尚無出賽紀錄" || oldVal === "—") return "—";
  if (field === "IP") return calcIPDiff(currVal, oldVal);
  
  var c = parseFloat(currVal); var o = parseFloat(oldVal);
  if (isNaN(c) || isNaN(o)) return "—";
  
  // 📈 百分比增減率分支：針對率類指標（WHIP, AVG, OBP, OPS）
  if (["WHIP", "AVG", "OBP", "OPS"].includes(field)) {
    if (o === 0) return "0%";
    var pctChange = ((c - o) / o) * 100;
    var roundedPct = Math.round(pctChange);
    if (roundedPct > 0) return "+" + roundedPct + "%";
    if (roundedPct < 0) return roundedPct + "%"; // 負數自帶減號
    return "0%";
  }
  
  // 🔢 絕對數字增減分支：ERA 獨立取兩位，其餘取整數
  var diff = c - o;
  if (field === "ERA") {
    if (diff > 0) return "+" + diff.toFixed(2);
    if (diff < 0) return diff.toFixed(2);
    return "0.00";
  }
  
  var intDiff = Math.round(diff);
  if (intDiff > 0) return "+" + intDiff;
  if (intDiff < 0) return String(intDiff);
  return "0";
}

function calcIPDiff(curr, old) {
  function toOuts(ip) {
    let parts = String(ip).split('.');
    let full = parseInt(parts[0]) || 0; let partial = parseInt(parts[1]) || 0;
    return full * 3 + partial;
  }
  let diffOuts = toOuts(curr) - toOuts(old);
  let sign = diffOuts < 0 ? -1 : 1; let absOuts = Math.abs(diffOuts);
  let fullDiff = Math.floor(absOuts / 3); let partDiff = absOuts % 3;
  var val = (parseFloat(fullDiff + "." + partDiff) * sign).toFixed(1);
  
  var num = parseFloat(val);
  if (num > 0) return "+" + val;
  if (num < 0) return val;
  return "0.0";
}

/**
 * 💡 修正點：直接套用字串，不進行重複字串疊加
 */
function applyDeltaColor(sheet, row, col, field, deltaValStr) {
  var cell = sheet.getRange(row, col);
  if (deltaValStr === "No Game" || deltaValStr === "累積中..." || deltaValStr === "—") {
    cell.setValue(deltaValStr).setBackground(null); return;
  }
  
  cell.setValue(deltaValStr);
  
  var cleanStr = deltaValStr.replace(/[^0-9.-]/g, "");
  var delta = parseFloat(cleanStr);
  if (isNaN(delta) || delta === 0) { cell.setBackground(null); return; }
  
  var isReverseGood = ["L", "ERA", "WHIP"].includes(field); 
  var isGood = isReverseGood ? (delta < 0) : (delta > 0);
  if (isGood) { cell.setBackground("#b7e1cd"); } else { cell.setBackground("#fce8b2"); }
}

function initHistorySheet(ss) {
  var sheet = ss.getSheetByName("Stats_History");
  if (!sheet) {
    sheet = ss.insertSheet("Stats_History");
    sheet.appendRow(["日期", "姓名", "類型", "數據JSON"]);
    sheet.hideSheet();
  }
}

function saveToHistorySheet(ss, todayStr, engName, type, statsObj) {
  var sheet = ss.getSheetByName("Stats_History"); if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = data[i][0];
    var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowDate).trim();
    
    if (rowDateStr === todayStr && data[i][1] === engName && data[i][2] === type) {
      sheet.getRange(i + 1, 4).setValue(JSON.stringify(statsObj)); return;
    }
  }
  sheet.appendRow([todayStr, engName, type, JSON.stringify(statsObj)]);
}

function getStatsFromHistory(ss, engName, type, todayStr) {
  var sheet = ss.getSheetByName("Stats_History"); if (!sheet) return null;
  var data = sheet.getDataRange().getValues(); 
  
  var parts = todayStr.split('-');
  var todayMidnight = new Date(parts[0], parts[1] - 1, parts[2]); 
  
  var bestRecord = null; 
  var maxDiffDays = -1; 

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === engName && data[i][2] === type) {
      var rowDate = data[i][0];
      var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowDate).trim();
      
      if (rowDateStr === todayStr) continue;
      
      var rParts = rowDateStr.split('-');
      var recordMidnight = new Date(rParts[0], rParts[1] - 1, rParts[2]);
      
      var diffTime = todayMidnight.getTime() - recordMidnight.getTime();
      var diffDays = Math.round(diffTime / (1000 * 3600 * 24)); 
      
      if (diffDays >= 1 && diffDays <= 5) {
        if (diffDays > maxDiffDays) { 
          maxDiffDays = diffDays; 
          bestRecord = data[i][3]; 
        }
      }
    }
  }
  if (bestRecord) { try { return JSON.parse(bestRecord); } catch(e) {} }
  return null;
}

function updateCellOnly(sheet, row, col, newValue) {
  if (!col) return;
  var cell = sheet.getRange(row, col);
  if (String(newValue).trim() !== String(cell.getValue()).trim()) cell.setValue(newValue);
}

function updateCellAndHighlight(sheet, row, col, newValue, todayStr) {
  if (!col) return;
  var cell = sheet.getRange(row, col);
  if (String(newValue).trim() !== String(cell.getValue()).trim()) { cell.setValue(newValue); cell.setBackground("#FFFF00"); }
}

function updateAll() { executeCoreUpdate(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(), "ALL", "所有資訊"); }
function updateTransactions() { executeCoreUpdate(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(), "TRANSACTIONS", "最新異動"); }
function updateStatusAndTeam() { executeCoreUpdate(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(), "STATUS", "狀態與球隊"); }

function executeCoreUpdate(sheetName, mode, processName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap = {}; headers.forEach((h, i) => colMap[h.trim()] = i + 1);
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (var rowIdx = 2; rowIdx <= sheet.getLastRow(); rowIdx++) {
    var englishName = String(sheet.getRange(rowIdx, colMap["英文名字"]).getValue()).trim();
    if (!englishName) continue;
    var milbUrl = colMap["MiLB頁面"] ? String(sheet.getRange(rowIdx, colMap["MiLB頁面"]).getValue()).trim() : "";
    
    var currentId = String(sheet.getRange(rowIdx, colMap["Player ID"]).getValue()).trim();
    if ((!milbUrl || !currentId) && colMap["Player ID"]) {
      var apiResult = fetchPlayerIdFromAPI(englishName);
      if (apiResult) {
        sheet.getRange(rowIdx, colMap["Player ID"]).setValue(apiResult.id);
        if (colMap["MiLB頁面"]) sheet.getRange(rowIdx, colMap["MiLB頁面"]).setValue(apiResult.milbUrl);
        if (colMap["MLB頁面"]) sheet.getRange(rowIdx, colMap["MLB頁面"]).setValue(apiResult.mlbUrl);
        milbUrl = apiResult.milbUrl;
      }
    }

    if (milbUrl) {
      var webData = scrapeMiLBPage(milbUrl);
      if (webData) {
        var dataFields = {};
        if (mode === "ALL") {
          dataFields = { "守備位置": webData.position, "目前年齡": webData.age, "目前狀態": webData.status, "所屬球隊": webData.currentTeam, "球隊層級": webData.teamLevel, "下一場比賽": webData.nextGame, "異動日期": webData.transactionDate, "異動內容": webData.transactionDesc };
        } else if (mode === "TRANSACTIONS") {
          dataFields = { "異動日期": webData.transactionDate, "異動內容": webData.transactionDesc };
        } else if (mode === "STATUS") {
          dataFields = { "目前狀態": webData.status, "所屬球隊": webData.currentTeam, "球隊層級": webData.teamLevel, "下一場比賽": webData.nextGame };
        }

        for (var fieldName in dataFields) { if (colMap[fieldName]) sheet.getRange(rowIdx, colMap[fieldName]).setBackground(null); }
        if (dataFields["異動內容"] !== undefined && colMap["異動日期"]) sheet.getRange(rowIdx, colMap["異動日期"]).setBackground(null);

        var levelKey = colMap["球隊層級"] ? "球隊層級" : (colMap["目前層級"] ? "目前層級" : null);
        var hasStatusChanged = false, hasTeamChanged = false, hasLevelChanged = false;
        if (colMap["目前狀態"]) { if (webData.status && String(webData.status).trim() !== String(sheet.getRange(rowIdx, colMap["目前狀態"]).getValue()).trim()) hasStatusChanged = true; }
        if (colMap["所屬球隊"]) { if (webData.currentTeam && String(webData.currentTeam).trim() !== String(sheet.getRange(rowIdx, colMap["所屬球隊"]).getValue()).trim()) hasTeamChanged = true; }
        if (levelKey) { if (webData.teamLevel && String(webData.teamLevel).trim() !== String(sheet.getRange(rowIdx, colMap[levelKey]).getValue()).trim()) hasLevelChanged = true; }

        for (var fieldName in dataFields) { if (colMap[fieldName]) updateCellAndHighlight(sheet, rowIdx, colMap[fieldName], dataFields[fieldName], todayStr); }
        if (colMap["資料更新時間"]) sheet.getRange(rowIdx, colMap["資料更新時間"]).setValue(todayStr);

        if (dataFields["異動內容"] !== undefined && isWithinFiveDays(webData.transactionDate)) {
          if (colMap["異動內容"]) sheet.getRange(rowIdx, colMap["異動內容"]).setBackground("#FFD2D2");
          if (colMap["異動日期"]) sheet.getRange(rowIdx, colMap["異動日期"]).setBackground("#FFD2D2");
        }
        if (dataFields["目標狀態"] !== undefined || dataFields["目前狀態"] !== undefined) {
          if (hasStatusChanged && (colMap["目前狀態"] || colMap["現狀/狀態"])) { var sKey = colMap["目前狀態"] ? "目前狀態" : "現狀/狀態"; sheet.getRange(rowIdx, colMap[sKey]).setBackground("#FFEA00"); }
          if (hasTeamChanged && colMap["所屬球隊"]) sheet.getRange(rowIdx, colMap["所屬球隊"]).setBackground("#FFEA00");
          if (hasLevelChanged && levelKey) sheet.getRange(rowIdx, colMap[levelKey]).setBackground("#FFEA00");
        }
      }
    }
    Utilities.sleep(300);
  }
}

function scrapeDetailedStats(url, type, teamLevel, currentYear) {
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true, 'headers': { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36' } });
    if (response.getResponseCode() !== 200) return null;
    const html = response.getContentText(); const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let trMatch; let dataObj = null; let allRows = [];
    while ((trMatch = trRegex.exec(html)) !== null) {
      const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi; let tds = []; let tdMatch;
      while ((tdMatch = tdRegex.exec(trMatch[1])) !== null) tds.push(tdMatch[1].replace(/<[^>]+>/g, '').trim());
      if (tds.length >= 10) allRows.push(tds);
    }
    for (let i = 0; i < allRows.length; i++) {
      let tds = allRows[i]; let rowHeader = tds[0];
      if (teamLevel.toUpperCase().indexOf("MLB") !== -1 || teamLevel.toUpperCase().indexOf("MAJOR") !== -1) {
        if (rowHeader.indexOf(currentYear + " Regular Season") !== -1 || rowHeader.indexOf(currentYear + " MLB Stats") !== -1) { dataObj = buildStatsObject(type, tds); break; }
      } else {
        if (rowHeader.indexOf(currentYear + " MiLB Stats") !== -1 || rowHeader.indexOf(currentYear + " Regular Season") !== -1) { dataObj = buildStatsObject(type, tds); break; }
      }
    }
    if (!dataObj) {
      for (let i = 0; i < allRows.length; i++) { if (allRows[i][0].indexOf(String(currentYear)) !== -1) { dataObj = buildStatsObject(type, allRows[i]); break; } }
    }
    if (!dataObj) {
      dataObj = {}; let fields = (type === "Pitcher") ? ["W", "L", "ERA", "G", "GS", "SV", "IP", "SO", "WHIP"] : ["AB", "R", "H", "HR", "RBI", "SB", "AVG", "OBP", "OPS"];
      fields.forEach((field, idx) => { dataObj[field] = (idx === 0) ? "今年賽季尚無出賽紀錄" : "—"; });
    }
    return dataObj;
  } catch (e) { return null; }
}

function buildStatsObject(type, tds) {
  if (type === "Pitcher") { return { "W": tds[1], "L": tds[2], "ERA": tds[3], "G": tds[4], "GS": tds[5], "SV": tds[6], "IP": tds[7], "SO": tds[8], "WHIP": tds[9] }; } 
  else { return { "AB": tds[1], "R": tds[2], "H": tds[3], "HR": tds[4], "RBI": tds[5], "SB": tds[6], "AVG": tds[7], "OBP": tds[8], "OPS": tds[9] }; }
}

function fetchPlayerIdFromAPI(englishName) {
  var url = "https://statsapi.mlb.com/api/v1/people/search?names=" + encodeURIComponent(englishName) + "&sportIds=1,11,12,13,14,16,21";
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); var json = JSON.parse(response.getContentText());
    if (response.getResponseCode() === 200 && json.people && json.people.length > 0) {
      var player = json.people[0]; var formattedName = englishName.toLowerCase().replace(/[^a-z0-9\s-]/g, '').replace(/[\s_]+/g, '-');
      return { id: player.id, milbUrl: "https://www.milb.com/player/" + formattedName + "-" + player.id, mlbUrl: "https://www.mlb.com/player/" + formattedName + "-" + player.id };
    }
  } catch (e) {} return null;
}

function scrapeMiLBPage(url) {
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true, 'headers': { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36' } }); if (response.getResponseCode() !== 200) return null;
    const html = response.getContentText();
    var data = { currentTeam: "未找到", teamLevel: "未找到", transactionDate: "未找到", transactionDesc: "未找到", position: "P", status: "未找到", nextGame: "未找到", age: "未找到" };
    const statusMatch = html.match(/<span class="label">Status:<\/span>\s*([^<]+)/i); if (statusMatch) data.status = statusMatch[1].trim();
    const nextGameMatch = html.match(/<span class="label">Next [^<]+ Game:\s*<\/span>\s*<a href="[^"]+">([^<]+)<\/a>/i); if (nextGameMatch) data.nextGame = nextGameMatch[1].trim();
    const posMatch = html.match(/<li>([A-Za-z0-9]{1,2})<\/li>/i); if (posMatch) data.position = posMatch[1].trim();
    const ageMatch = html.match(/<li class="player-header--vitals-age">Age:\s*(\d+)<\/li>/i); if (ageMatch) data.age = ageMatch[1].trim();
    const teamMatch = html.match(/class="player-header--vitals-currentTeam-name"[^>]*>[\s\S]*?<span class="player-header--vitals-name">([^<]+)<\/span>/i); if (teamMatch) data.currentTeam = teamMatch[1].trim();
    const levelMatch = html.match(/<div class="header__info-bar">([\s\S]*?)<\/div>/i); if (levelMatch) data.teamLevel = levelMatch[1].replace(/<[^>]+>/g, '').replace(/\s*Affiliate\s*/i, '').trim();
    if ((data.teamLevel === "" || data.teamLevel === "未找到") && data.currentTeam !== "未找到") data.teamLevel = "MLB";
    const tableMatch = html.match(/<table class="transactions-table[^>]*>[\s\S]*?<tbody>\s*<tr>([\s\S]*?)<\/tr>/i);
    if (tableMatch) {
      const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi; let tds = []; let tdMatch;
      while ((tdMatch = tdRegex.exec(tableMatch[1])) !== null) tds.push(tdMatch[1].replace(/<[^>]+>/g, '').trim());
      if (tds.length >= 3) { data.transactionDate = tds[1]; data.transactionDesc = tds[2]; }
    }
    return data;
  } catch (e) { return null; }
}

function initLogSheet(ss) {
  var logSheet = ss.getSheetByName("Log");
  if (!logSheet) { logSheet = ss.insertSheet("Log"); logSheet.appendRow(["系統時間", "日誌等級", "詳細訊息"]); }
}

function writeLog(level, message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var logSheet = ss.getSheetByName("Log");
  if (logSheet) { var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); logSheet.appendRow([timeStamp, level, message]); }
}

function enableDailyTrigger() {
  disableDailyTrigger();
  ScriptApp.newTrigger('performDailyAutomation').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert("每日早上 7 點自動更新與信件摘要已啟用。");
}

function disableDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'performDailyAutomation' || triggers[i].getHandlerFunction() === 'updateAll') ScriptApp.deleteTrigger(triggers[i]);
  }
}
/**
 * 試算表編輯監聽器 v4（精簡審核版）
 * 
 * 欄位架構對照：
 * A–J 欄：原 Google 表單資料
 * K 欄 (11)：申請人 Email
 * L 欄 (12)：☑ 執行換班（更動日曆，撞班自動退件）
 * M 欄 (13)：執行狀態（更新成功 / 撞班退件原因）
 * N 欄 (14)：☑ 審核通知（寄出核准信）
 * O 欄 (15)：核准通知記錄（防重複寄信）
 * P 欄 (16)：退件通知記錄（防重複寄信）
 * 
 * ⚠️ 安裝式觸發器只需 1 個：
 * 執行函式：handleSheetEdit ｜ 事件來源：來自試算表 ｜ 事件類型：編輯時
 */

// 單次觸發最多處理的列數
const MAX_ROWS_PER_RUN = 30;
// Apps Script 執行上限煞車點（毫秒）
const TIME_BUDGET_MS = 5 * 60 * 1000;

// ============================================================
// CONFIG 設定值
// ============================================================
const CONFIG = {
  // --- 換班與狀態欄位 ---
  EXECUTE_CHECK_COL: 12, // L欄：執行換班用的勾選框
  STATUS_COL: 13,        // M欄：執行狀態

  // --- 審核通知欄位 ---
  APPROVE_CHECK_COL: 14, // N欄：審核通知勾選框
  APPROVE_LOG_COL: 15,   // O欄：核准通知寄送記錄
  REJECT_LOG_COL: 16,    // P欄：退件通知寄送記錄

  // --- Email 相關設定 ---
  APPLICANT_EMAIL_COL: 11, // K欄：表單收集之申請人 Email

  // --- 員工姓名對應 Email 名冊（雙向換班時通知互換對象） ---
  STAFF_DIRECTORY_SHEET_NAME: "員工名冊",
  STAFF_NAME_COL: 1,  // 名冊分頁：姓名欄（A=1）
  STAFF_EMAIL_COL: 2, // 名冊分頁：Email欄（B=2）
};

/**
 * 試算表編輯事件監聽器
 */
function handleSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  const startCol = range.getColumn();
  const endCol = range.getLastColumn();

  const touchesExecuteCol = !(startCol > CONFIG.EXECUTE_CHECK_COL || endCol < CONFIG.EXECUTE_CHECK_COL);
  const touchesApproveCol = !(startCol > CONFIG.APPROVE_CHECK_COL || endCol < CONFIG.APPROVE_CHECK_COL);
  if (!touchesExecuteCol && !touchesApproveCol) return;

  // N欄（審核通知）被編輯：走獨立的寄信通知流程，不碰日曆
  if (touchesApproveCol && !touchesExecuteCol) {
    handleApprovalEditRange(sheet, range);
    return;
  }

  const startRow = Math.max(2, range.getRow());
  let endRow = range.getLastRow();
  if (startRow > endRow) return;

  let truncated = false;
  if (endRow - startRow + 1 > MAX_ROWS_PER_RUN) {
    const realEndRow = endRow;
    endRow = startRow + MAX_ROWS_PER_RUN - 1;
    truncated = true;
    sheet.getRange(endRow + 1, CONFIG.STATUS_COL, realEndRow - endRow, 1)
      .setValue(`尚未處理：單次批次上限為 ${MAX_ROWS_PER_RUN} 列，請稍後重新勾選此列（或分批操作）`);
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    sheet.getRange(startRow, CONFIG.STATUS_COL, endRow - startRow + 1, 1).setValue("系統忙碌中，請稍候重試");
    return;
  }

  const runStart = Date.now();

  try {
    const calendars = CalendarApp.getCalendarsByName("達賢館創新組助理值班");
    if (calendars.length === 0) {
      sheet.getRange(startRow, CONFIG.STATUS_COL, endRow - startRow + 1, 1).setValue("錯誤：找不到指定日曆");
      return;
    }
    const calendar = calendars[0];

    for (let r = startRow; r <= endRow; r++) {
      if (Date.now() - runStart > TIME_BUDGET_MS) {
        sheet.getRange(r, CONFIG.STATUS_COL, endRow - r + 1, 1)
          .setValue("尚未處理：本次執行時間已達上限，請重新勾選此列");
        truncated = true;
        break;
      }

      try {
        processSingleRow(sheet, calendar, r);
      } catch (rowErr) {
        console.error(`第 ${r} 列處理發生非預期錯誤: ${rowErr.message}`);
        sheet.getRange(r, CONFIG.STATUS_COL).setValue("執行失敗（非預期錯誤）: " + rowErr.message);
        const checkCell = sheet.getRange(r, CONFIG.EXECUTE_CHECK_COL);
        if (checkCell.getValue() === true) checkCell.setValue(false);
      }
    }
  } catch (err) {
    console.error(err);
    sheet.getRange(startRow, CONFIG.STATUS_COL, endRow - startRow + 1, 1)
      .setValue("系統發生錯誤，請聯絡管理員: " + err.message);
  } finally {
    lock.releaseLock();
  }

  if (truncated) {
    console.log(`本次觸發因批次列數或時間上限而部分列未處理（範圍 ${startRow}-${range.getLastRow()}）`);
  }
}

/**
 * 核心處理單一列日曆變更
 */
function processSingleRow(sheet, calendar, row) {
  const checkCell = sheet.getRange(row, CONFIG.EXECUTE_CHECK_COL);
  const statusCell = sheet.getRange(row, CONFIG.STATUS_COL);
  const isChecked = checkCell.getValue() === true;
  const statusVal = statusCell.getValue().toString();

  // 讀取欄位資料
  const origPerson = sheet.getRange(row, 2).getValue().toString().trim();
  const origDate = sheet.getRange(row, 3).getValue();
  const origStartStr = sheet.getRange(row, 4).getDisplayValue().toString().trim();
  const origEndStr = sheet.getRange(row, 5).getDisplayValue().toString().trim();
  const targetPerson = sheet.getRange(row, 6).getValue().toString().trim();

  const swapDate = sheet.getRange(row, 7).getValue();
  const swapStartStr = sheet.getRange(row, 8).getDisplayValue().toString().trim();
  const swapEndStr = sheet.getRange(row, 9).getDisplayValue().toString().trim();

  if (!origPerson || !origDate) {
    if (isChecked) {
      denyRow(sheet, row, checkCell, statusCell, "錯誤：請填寫原值班人員與原值班日期");
    }
    return;
  }
  if (isChecked && (!origStartStr || !origEndStr)) {
    denyRow(sheet, row, checkCell, statusCell, "錯誤：請填寫原值班起訖時間");
    return;
  }

  const isSwap = Boolean(
    swapDate && swapStartStr && swapEndStr && targetPerson &&
    targetPerson !== "請假" && targetPerson !== "無"
  );

  const tokenA = `${row}-A`;
  const tokenB = `${row}-B`;

  try {
    // 動作 1：取消勾選 (還原回原本班表)
    if (!isChecked) {
      if (!statusVal.includes("已更新日曆")) return;

      const origStartTime = combineDateTimeByStr(origDate, origStartStr);
      const origEndTime = combineDateTimeByStr(origDate, origEndStr);

      revertEvents(calendar, origStartTime, origEndTime, origPerson, tokenA);

      if (isSwap) {
        const swapStartTime = combineDateTimeByStr(swapDate, swapStartStr);
        const swapEndTime = combineDateTimeByStr(swapDate, swapEndStr);
        revertEvents(calendar, swapStartTime, swapEndTime, targetPerson, tokenB);
      }

      statusCell.setValue("");
      return;
    }

    // 動作 2：勾選 TRUE (執行更動日曆)
    if (isChecked) {
      if (statusVal.includes("已更新日曆")) return;

      const origStartTime = combineDateTimeByStr(origDate, origStartStr);
      const origEndTime = combineDateTimeByStr(origDate, origEndStr);

      const eventsO = getEventsForWindow(calendar, origStartTime, origEndTime);

      const events1 = matchPersonEvents(eventsO, origStartTime, origEndTime, origPerson);
      if (events1.length === 0) {
        denyRow(sheet, row, checkCell, statusCell, `錯誤：找不到 ${origPerson} 的原值班行程`);
        return;
      }
      if (events1.length > 1) {
        denyRow(sheet, row, checkCell, statusCell, `錯誤：${origPerson} 同時段有多筆行程，請手動確認日曆`);
        return;
      }

      const check1 = validateTimeRange(events1[0], origStartTime, origEndTime);
      if (!check1.valid) {
        denyRow(sheet, row, checkCell, statusCell, `錯誤：${origPerson} 的填寫時段超出其原班表 (${check1.actualRange})`);
        return;
      }

      // === 情況 A：雙向換班 ===
      if (isSwap) {
        const swapStartTime = combineDateTimeByStr(swapDate, swapStartStr);
        const swapEndTime = combineDateTimeByStr(swapDate, swapEndStr);
        
        const eventsS = getEventsForWindow(calendar, swapStartTime, swapEndTime);
        const events2 = matchPersonEvents(eventsS, swapStartTime, swapEndTime, targetPerson);

        if (events2.length === 0) {
          denyRow(sheet, row, checkCell, statusCell, `錯誤：找不到 ${targetPerson} 的互換時段行程`);
          return;
        }
        if (events2.length > 1) {
          denyRow(sheet, row, checkCell, statusCell, `錯誤：${targetPerson} 互換時段有多筆行程，請手動確認日曆`);
          return;
        }

        const check2 = validateTimeRange(events2[0], swapStartTime, swapEndTime);
        if (!check2.valid) {
          denyRow(sheet, row, checkCell, statusCell, `錯誤：${targetPerson} 的填寫時段超出其原班表 (${check2.actualRange})`);
          return;
        }

        // 撞班檢測 1
        const conflictForTarget = findRealConflicts(eventsO, origStartTime, origEndTime, targetPerson, events2[0], swapStartTime, swapEndTime);
        if (conflictForTarget.length > 0) {
          const timeRangeStr = formatEventTime(conflictForTarget[0]);
          denyRow(sheet, row, checkCell, statusCell, `錯誤退件：${targetPerson} 在原值班時段已有班 (${conflictForTarget[0].getTitle()} ${timeRangeStr})`);
          return;
        }

        // 撞班檢測 2
        const conflictForOrig = findRealConflicts(eventsS, swapStartTime, swapEndTime, origPerson, events1[0], origStartTime, origEndTime);
        if (conflictForOrig.length > 0) {
          const timeRangeStr = formatEventTime(conflictForOrig[0]);
          denyRow(sheet, row, checkCell, statusCell, `錯誤退件：${origPerson} 在互換時段已有值班 (${conflictForOrig[0].getTitle()} ${timeRangeStr})`);
          return;
        }

        const origDateStr = origDate instanceof Date ? Utilities.formatDate(origDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : origDate;
        const swapDateStr = swapDate instanceof Date ? Utilities.formatDate(swapDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : swapDate;

        let firstSideApplied = false;
        try {
          applyShiftChange(
            calendar, events1[0], origStartTime, origEndTime, targetPerson, origPerson,
            `【換班紀錄】\n- 實際到勤：${targetPerson}\n- 原定值班：${origPerson}\n- 互換對象時段：${swapDateStr}`,
            "[換班]", tokenA
          );
          firstSideApplied = true;

          applyShiftChange(
            calendar, events2[0], swapStartTime, swapEndTime, origPerson, targetPerson,
            `【換班紀錄】\n- 實際到勤：${origPerson}\n- 原定值班：${targetPerson}\n- 互換對象時段：${origDateStr}`,
            "[換班]", tokenB
          );
        } catch (swapErr) {
          if (firstSideApplied) {
            try {
              revertEvents(calendar, origStartTime, origEndTime, origPerson, tokenA);
            } catch (rollbackErr) {
              console.error(`換班回滾失敗（第 ${row} 列）: ${rollbackErr.message}`);
              throw new Error(`換班失敗且自動回滾也失敗，請手動檢查日曆: ${swapErr.message}`);
            }
          }
          throw swapErr;
        }

        statusCell.setValue("已更新日曆（雙向換班完成）");

      } else if (targetPerson && targetPerson !== "請假" && targetPerson !== "無") {
        // === 情況 B：單向代班 ===
        const conflictForTarget = findRealConflicts(eventsO, origStartTime, origEndTime, targetPerson, null, null, null);
        if (conflictForTarget.length > 0) {
          const timeRangeStr = formatEventTime(conflictForTarget[0]);
          denyRow(sheet, row, checkCell, statusCell, `錯誤退件：${targetPerson} 在代班時段已有值班 (${conflictForTarget[0].getTitle()} ${timeRangeStr})`);
          return;
        }

        const desc = `【代班紀錄】\n- 實際到勤：${targetPerson}\n- 原定值班：${origPerson}（請假由他人代班）`;
        applyShiftChange(calendar, events1[0], origStartTime, origEndTime, targetPerson, origPerson, desc, "[代班]", tokenA);
        statusCell.setValue("已更新日曆（代班完成）");

      } else {
        // === 情況 C：純請假 ===
        const desc = `【請假紀錄】\n- 原定值班：${origPerson}（請假無人代理）`;
        applyShiftChange(calendar, events1[0], origStartTime, origEndTime, origPerson, origPerson, desc, "【請假】", tokenA);
        statusCell.setValue("已更新日曆（請假完成）");
      }
    }

  } catch (err) {
    const failMsg = "執行失敗: " + err.message;
    statusCell.setValue(failMsg);
    if (isChecked) {
      checkCell.setValue(false);
      notifyRejectionIfNeeded(sheet, row, failMsg);
    }
  }
}

/**
 * 抓取某時間窗前後 12 小時內的所有行程
 */
function getEventsForWindow(calendar, reqStart, reqEnd) {
  const searchStart = new Date(reqStart.getTime() - 12 * 60 * 60 * 1000);
  const searchEnd = new Date(reqEnd.getTime() + 12 * 60 * 60 * 1000);
  return calendar.getEvents(searchStart, searchEnd);
}

/**
 * 依姓名篩選原班表行程
 */
function matchPersonEvents(events, reqStart, reqEnd, personName) {
  const matched = [];
  for (let i = 0; i < events.length; i++) {
    const evt = events[i];
    if (!evt.getTitle().includes(personName)) continue;

    const evStart = evt.getStartTime();
    const evEnd = evt.getEndTime();
    if (reqStart < evEnd && evStart < reqEnd) {
      matched.push(evt);
    }
  }
  return matched;
}

/**
 * 精確衝突計算（考量讓出時段後的殘餘行程）
 */
function findRealConflicts(events, targetStart, targetEnd, personName, cedingEvent, cededStart, cededEnd) {
  const cedingId = cedingEvent ? cedingEvent.getId() : null;
  const conflicts = [];

  for (let i = 0; i < events.length; i++) {
    const evt = events[i];
    if (!evt.getTitle().includes(personName)) continue;

    const evStart = evt.getStartTime();
    const evEnd = evt.getEndTime();

    if (cedingId && evt.getId() === cedingId) {
      if (evStart.getTime() < cededStart.getTime()) {
        if (targetStart < cededStart && evStart < targetEnd) {
          conflicts.push(evt);
          continue;
        }
      }
      if (cededEnd.getTime() < evEnd.getTime()) {
        if (targetStart < evEnd && cededEnd < targetEnd) {
          conflicts.push(evt);
          continue;
        }
      }
    } else {
      if (targetStart < evEnd && evStart < targetEnd) {
        conflicts.push(evt);
      }
    }
  }
  return conflicts;
}

/**
 * 格式化行程時間為 HH:mm-HH:mm
 */
function formatEventTime(evt) {
  const timeZone = Session.getScriptTimeZone();
  const s = Utilities.formatDate(evt.getStartTime(), timeZone, "HH:mm");
  const e = Utilities.formatDate(evt.getEndTime(), timeZone, "HH:mm");
  return `${s}-${e}`;
}

/**
 * 驗證填寫時間是否落在原班表內
 */
function validateTimeRange(event, reqStart, reqEnd) {
  const evStart = event.getStartTime().getTime();
  const evEnd = event.getEndTime().getTime();
  const s = reqStart.getTime();
  const e = reqEnd.getTime();
  const timeZone = Session.getScriptTimeZone();

  const actualStartStr = Utilities.formatDate(event.getStartTime(), timeZone, "HH:mm");
  const actualEndStr = Utilities.formatDate(event.getEndTime(), timeZone, "HH:mm");
  const actualRange = `${actualStartStr}-${actualEndStr}`;

  if (s < evStart - 60000 || e > evEnd + 60000) {
    return { valid: false, actualRange: actualRange };
  }
  return { valid: true, actualRange: actualRange };
}

/**
 * 時段切割並套用更動
 */
function applyShiftChange(calendar, mainEvent, subStart, subEnd, newWorker, origWorker, descText, tag, rowToken) {
  const evStart = mainEvent.getStartTime();
  const evEnd = mainEvent.getEndTime();
  const floor = extractFloorSuffix(mainEvent.getTitle());
  const rowMeta = `[ROW_ID:${rowToken}]`;

  if (Math.abs(evStart.getTime() - subStart.getTime()) < 60000 && Math.abs(evEnd.getTime() - subEnd.getTime()) < 60000) {
    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    mainEvent.setTitle(newTitle);
    mainEvent.setDescription(`${rowMeta}\n${descText}\n--------------------\n` + cleanDescription(mainEvent.getDescription(), rowToken));
    return;
  }

  const backupMeta = `[SPLIT_ORIG_TIME:${evStart.getTime()}-${evEnd.getTime()}]`;

  if (subStart.getTime() > evStart.getTime() && subEnd.getTime() < evEnd.getTime()) {
    mainEvent.setTime(evStart, subStart);
    mainEvent.setDescription(`${rowMeta}\n${backupMeta}\n` + (mainEvent.getDescription() || ""));

    const midTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    calendar.createEvent(midTitle, subStart, subEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n${descText}\n--------------------\n`
    });

    calendar.createEvent(`${origWorker}${floor}`, subEnd, evEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n【原班後半段】\n--------------------\n`
    });
    return;
  }

  if (subStart.getTime() > evStart.getTime() && Math.abs(subEnd.getTime() - evEnd.getTime()) < 60000) {
    mainEvent.setTime(evStart, subStart);
    mainEvent.setDescription(`${rowMeta}\n${backupMeta}\n` + (mainEvent.getDescription() || ""));

    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    calendar.createEvent(newTitle, subStart, subEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n${descText}\n--------------------\n`
    });
    return;
  }

  if (Math.abs(subStart.getTime() - evStart.getTime()) < 60000 && subEnd.getTime() < evEnd.getTime()) {
    mainEvent.setTime(subEnd, evEnd);
    mainEvent.setDescription(`${rowMeta}\n${backupMeta}\n` + (mainEvent.getDescription() || ""));

    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    calendar.createEvent(newTitle, subStart, subEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n${descText}\n--------------------\n`
    });
    return;
  }

  throw new Error("無法計算時段切割，請確認起訖時間是否正確");
}

/**
 * 還原行程
 */
function revertEvents(calendar, rangeStart, rangeEnd, origWorker, rowToken) {
  const searchStart = new Date(rangeStart.getTime() - 24 * 60 * 60 * 1000);
  const searchEnd = new Date(rangeEnd.getTime() + 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(searchStart, searchEnd);
  const rowTokenTag = `[ROW_ID:${rowToken}]`;

  events.forEach(evt => {
    const desc = evt.getDescription() || "";

    if (!desc.includes(rowTokenTag)) return;

    if (desc.includes("[AUTO_SPLIT_CREATED]")) {
      evt.deleteEvent();
      return;
    }

    const splitMatch = desc.match(/\[SPLIT_ORIG_TIME:(\d+)-(\d+)\]/);
    if (splitMatch) {
      const origS = new Date(parseInt(splitMatch[1]));
      const origE = new Date(parseInt(splitMatch[2]));
      evt.setTime(origS, origE);
      evt.setDescription(desc.replace(/\[SPLIT_ORIG_TIME:\d+-\d+\]\n?/g, ""));
    }

    const title = evt.getTitle();
    if (title.includes("[代班]") || title.includes("[換班]") || title.includes("【請假】")) {
      const floor = extractFloorSuffix(title);
      evt.setTitle(`${origWorker}${floor}`);
      evt.setDescription(cleanDescription(desc, rowToken));
    }
  });
}

/**
 * 擷取括號後綴（樓層資訊）
 */
function extractFloorSuffix(title) {
  const pureTitle = title.replace(/\s*\[(換班|代班)\]/g, "").replace(/【請假】/g, "");
  const match = pureTitle.match(/(\([^\)]+\)|（[^）]+）)$/);
  return match ? match[0] : "";
}

/**
 * 清除行程說明中的中繼紀錄標籤
 */
function cleanDescription(desc, rowToken) {
  if (!desc) return "";
  let res = desc.replaceAll(`[ROW_ID:${rowToken}]\n`, "").replaceAll(`[ROW_ID:${rowToken}]`, "");
  res = res.replace(/【(換班|代班|請假)紀錄】[\s\S]*?--------------------\n?/g, "").trim();
  return res;
}

/**
 * 純文字時間字串精準解析為 Date 物件
 */
function combineDateTimeByStr(dateVal, timeStr) {
  const d = new Date(dateVal);
  let hours = 0;
  let minutes = 0;

  const str = (timeStr || "").toString().trim();
  const match = str.match(/(\d{1,2}):(\d{2})/);

  if (match) {
    hours = parseInt(match[1], 10);
    minutes = parseInt(match[2], 10);

    const isPM = str.includes("下午") || str.toUpperCase().includes("PM");
    const isAM = str.includes("上午") || str.toUpperCase().includes("AM");

    if (isPM) {
      if (hours < 12) hours += 12;
    } else if (isAM) {
      if (hours === 12) hours = 0;
    }
  }

  d.setHours(hours, minutes, 0, 0);
  return d;
}

/**
 * 讀取並驗證申請人 Email（含格式驗證防呆）
 */
function getApplicantEmail(sheet, row) {
  if (!CONFIG.APPLICANT_EMAIL_COL) {
    console.error("尚未設定 CONFIG.APPLICANT_EMAIL_COL");
    return "";
  }

  const email = sheet.getRange(row, CONFIG.APPLICANT_EMAIL_COL).getValue().toString().trim();
  if (!email) return "";

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(email)) {
    console.error(`第 ${row} 列 K欄內容並非有效 Email：${email}`);
    return "";
  }

  return email;
}

/**
 * 依姓名查詢「員工名冊」分頁中的 Email（雙向換班通知互換對象）
 */
function getStaffEmailByName(name) {
  if (!name) return "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dirSheet = ss.getSheetByName(CONFIG.STAFF_DIRECTORY_SHEET_NAME);
  if (!dirSheet) {
    console.error(`找不到員工名冊分頁「${CONFIG.STAFF_DIRECTORY_SHEET_NAME}」，請確認分頁名稱是否一致`);
    return "";
  }
  const lastRow = dirSheet.getLastRow();
  if (lastRow < 2) return "";

  const numCols = Math.max(CONFIG.STAFF_NAME_COL, CONFIG.STAFF_EMAIL_COL);
  const data = dirSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const target = name.toString().trim();

  for (let i = 0; i < data.length; i++) {
    const rowName = (data[i][CONFIG.STAFF_NAME_COL - 1] || "").toString().trim();
    if (rowName === target) {
      return (data[i][CONFIG.STAFF_EMAIL_COL - 1] || "").toString().trim();
    }
  }
  console.error(`員工名冊中找不到姓名「${target}」對應的 Email`);
  return "";
}

/**
 * 統一退件處理：寫入狀態、還原 L 欄勾選框、觸發退件通知信
 */
function denyRow(sheet, row, checkCell, statusCell, message) {
  statusCell.setValue(message);
  checkCell.setValue(false);
  notifyRejectionIfNeeded(sheet, row, message);
}

/**
 * 退件通知信（具備防重複發送機制）
 */
function notifyRejectionIfNeeded(sheet, row, message) {
  const logCell = sheet.getRange(row, CONFIG.REJECT_LOG_COL);
  const prevLog = logCell.getValue().toString();
  if (prevLog === message) return;

  try {
    const email = getApplicantEmail(sheet, row);
    if (!email) {
      logCell.setValue(message + "（找不到有效 Email，未寄信）");
      return;
    }
    const origPerson = sheet.getRange(row, 2).getValue().toString().trim();

    MailApp.sendEmail({
      to: email,
      subject: "您的換班申請未通過（系統自動退件）",
      body:
`${origPerson} 您好，

您所提出的換班/代班申請經系統檢核後無法通過，原因如下：
${message}

請重新確認排班內容後再次提出申請，如有疑問請洽管理員。

（此為系統自動發送信件，請勿直接回覆）`
    });
    logCell.setValue(message);
  } catch (err) {
    logCell.setValue("退件通知寄送失敗: " + err.message);
  }
}

/**
 * N欄（審核通知）編輯事件的批次分派
 */
function handleApprovalEditRange(sheet, range) {
  const startRow = Math.max(2, range.getRow());
  const endRow = range.getLastRow();
  if (startRow > endRow) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;

  try {
    for (let r = startRow; r <= endRow; r++) {
      try {
        handleApprovalEdit(sheet, r);
      } catch (err) {
        console.error(`第 ${r} 列審核通知處理失敗: ${err.message}`);
        sheet.getRange(r, CONFIG.APPROVE_LOG_COL).setValue("處理失敗: " + err.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 核心審核通知邏輯：勾選 N 欄「審核通知」後寄送核准信
 */
function handleApprovalEdit(sheet, row) {
  const approveCheck = sheet.getRange(row, CONFIG.APPROVE_CHECK_COL);
  const approveLog = sheet.getRange(row, CONFIG.APPROVE_LOG_COL);
  const statusVal = sheet.getRange(row, CONFIG.STATUS_COL).getValue().toString();
  const isChecked = approveCheck.getValue() === true;

  if (!isChecked) return;

  const prevLog = approveLog.getValue().toString();
  if (prevLog.startsWith("已寄送核准通知")) return;

  if (!statusVal.startsWith("已更新日曆")) {
    approveCheck.setValue(false);
    approveLog.setValue("尚未成功執行換班（請先勾選L欄，並確認M欄狀態已顯示「已更新日曆」），無法寄送核准通知");
    return;
  }

  try {
    sendApprovalEmail(sheet, row);
    approveLog.setValue(`已寄送核准通知 ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")}`);
  } catch (err) {
    approveCheck.setValue(false);
    approveLog.setValue("寄信失敗: " + err.message);
  }
}

/**
 * 審核通過 → 寄送核准信（嚴格判定 G/H/I 欄，僅雙向換班時才通知互換對象）
 */
function sendApprovalEmail(sheet, row) {
  const email = getApplicantEmail(sheet, row);
  if (!email) throw new Error("找不到有效的申請人 Email");

  const origPerson = sheet.getRange(row, 2).getValue().toString().trim();
  const origDate = sheet.getRange(row, 3).getValue();
  const origStartStr = sheet.getRange(row, 4).getDisplayValue().toString().trim();
  const origEndStr = sheet.getRange(row, 5).getDisplayValue().toString().trim();
  const targetPerson = sheet.getRange(row, 6).getValue().toString().trim();

  const dateStr = origDate instanceof Date
    ? Utilities.formatDate(origDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : origDate;

  const swapNote =
    targetPerson &&
    targetPerson !== "請假" &&
    targetPerson !== "無"
      ? `（對象：${targetPerson}）`
      : "";

  // 寄給提出申請的人
  MailApp.sendEmail({
    to: email,
    subject: "【值班換班申請】已審核通過",
    body:
`${origPerson} 您好，

您於 ${dateStr} ${origStartStr}-${origEndStr} 提出的換班/代班申請${swapNote}已審核通過，並已完成排班異動，新的值班安排已同步至值班日曆。

請留意您的值班時間，如有任何問題歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`
  });

  // 只有真正的「雙向換班」才通知互換對象
  const swapDate = sheet.getRange(row, 7).getValue();
  const swapStartStr = sheet.getRange(row, 8).getDisplayValue().toString().trim();
  const swapEndStr = sheet.getRange(row, 9).getDisplayValue().toString().trim();

  const isSwap = Boolean(
    swapDate &&
    swapStartStr &&
    swapEndStr &&
    targetPerson &&
    targetPerson !== "請假" &&
    targetPerson !== "無"
  );

  if (isSwap) {
    const targetEmail = getStaffEmailByName(targetPerson);

    if (targetEmail) {
      MailApp.sendEmail({
        to: targetEmail,
        subject: "【值班換班申請】您的班表已完成互換",
        body:
`${targetPerson} 您好，

您與 ${origPerson} 的換班申請已審核通過，於 ${dateStr} ${origStartStr}-${origEndStr} 的值班已完成互換，新的值班安排已同步至值班日曆，請留意您的值班時間。

如對此次換班有疑問，歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`
      });
    }
  }
}
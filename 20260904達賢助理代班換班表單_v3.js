/**
 * 試算表編輯監聽器
 * 支援多列同時選取、批次拖曳填滿、Ctrl+Z/Y
 * 本版新增：換班原子性回復、必填欄位檢查、外層錯誤回饋表單、
 *           ROW_ID 加上 side 標記、批次執行時間/列數防護
 */

// 單次觸發最多處理的列數（避免一次貼上過多列導致執行超時）
const MAX_ROWS_PER_RUN = 30;
// Apps Script 單次執行上限約 6 分鐘，這裡抓 5 分鐘當安全煞車點（毫秒）
const TIME_BUDGET_MS = 5 * 60 * 1000;

function handleSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // 檢查編輯範圍是否有涵蓋到 K 欄（第 11 欄）
  const startCol = range.getColumn();
  const endCol = range.getLastColumn();
  if (startCol > 11 || endCol < 11) return;

  const startRow = Math.max(2, range.getRow()); // 略過標題列第 1 列
  let endRow = range.getLastRow();
  if (startRow > endRow) return;

  // 🟢 列數防護：單次觸發列數過多時，先只處理前 MAX_ROWS_PER_RUN 列，
  // 其餘列寫提示訊息，避免一次貼上大量列造成執行逾時或狀態混亂
  let truncated = false;
  if (endRow - startRow + 1 > MAX_ROWS_PER_RUN) {
    const realEndRow = endRow;
    endRow = startRow + MAX_ROWS_PER_RUN - 1;
    truncated = true;
    sheet.getRange(endRow + 1, 12, realEndRow - endRow, 1)
      .setValue(`尚未處理：單次批次上限為 ${MAX_ROWS_PER_RUN} 列，請稍後重新勾選此列（或分批操作）`);
  }

  // 使用 Lock 避免高頻編輯或並行衝突
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    sheet.getRange(startRow, 12, endRow - startRow + 1, 1).setValue("系統忙碌中，請稍候重試");
    return;
  }

  const runStart = Date.now();

  try {
    const calendars = CalendarApp.getCalendarsByName("達賢館創新組助理值班");
    if (calendars.length === 0) {
      sheet.getRange(startRow, 12, endRow - startRow + 1, 1).setValue("錯誤：找不到指定日曆");
      return;
    }
    const calendar = calendars[0];

    // 迴圈逐列處理（支援批次多列編輯）
    for (let r = startRow; r <= endRow; r++) {
      // 🟢 時間防護：接近執行時間上限就提前停止，剩餘列留給下次觸發處理
      if (Date.now() - runStart > TIME_BUDGET_MS) {
        sheet.getRange(r, 12, endRow - r + 1, 1)
          .setValue("尚未處理：本次執行時間已達上限，請重新勾選此列");
        truncated = true;
        break;
      }

      // 🟡 單列獨立 try/catch：單一列的非預期例外不會中斷其他列的處理
      try {
        processSingleRow(sheet, calendar, r);
      } catch (rowErr) {
        console.error(`第 ${r} 列處理發生非預期錯誤: ${rowErr.message}`);
        sheet.getRange(r, 12).setValue("執行失敗（非預期錯誤）: " + rowErr.message);
        const checkCell = sheet.getRange(r, 11);
        if (checkCell.getValue() === true) checkCell.setValue(false);
      }
    }
  } catch (err) {
    // 🟡 外層真正意外（例如日曆服務整體失敗）：回饋到整個受影響範圍的 L 欄，
    // 而不是只寫進執行紀錄讓使用者看不到
    console.error(err);
    sheet.getRange(startRow, 12, endRow - startRow + 1, 1)
      .setValue("系統發生錯誤，請聯絡管理員: " + err.message);
  } finally {
    lock.releaseLock();
  }

  if (truncated) {
    console.log(`本次觸發因批次列數或時間上限而部分列未處理（範圍 ${startRow}-${range.getLastRow()}）`);
  }
}

/**
 * 核心處理單一列資料
 */
function processSingleRow(sheet, calendar, row) {
  const checkCell = sheet.getRange(row, 11);
  const statusCell = sheet.getRange(row, 12);
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

  // 🟠 必填欄位檢查：不再靜默忽略，改為明確錯誤訊息並彈回 checkbox
  if (!origPerson || !origDate) {
    if (isChecked) {
      statusCell.setValue("錯誤：請填寫原值班人員與原值班日期");
      checkCell.setValue(false);
    }
    return;
  }
  if (isChecked && (!origStartStr || !origEndStr)) {
    statusCell.setValue("錯誤：請填寫原值班起訖時間");
    checkCell.setValue(false);
    return;
  }

  // 統一雙向換班判定標準（嚴格對稱）
  const isSwap = Boolean(
    swapDate && swapStartStr && swapEndStr && targetPerson &&
    targetPerson !== "請假" && targetPerson !== "無"
  );

  // 🟡 ROW_ID 加上 side 標記，換班兩邊分別用不同 token，避免日期相近時互相誤觸
  const tokenA = `${row}-A`; // 原班方
  const tokenB = `${row}-B`; // 換班對象方

  try {
    // ==========================================
    // 動作 1：取消勾選 (還原回原本班表)
    // ==========================================
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

    // ==========================================
    // 動作 2：勾選 TRUE (執行更動日曆)
    // ==========================================
    if (isChecked) {
      if (statusVal.includes("已更新日曆")) return;

      const origStartTime = combineDateTimeByStr(origDate, origStartStr);
      const origEndTime = combineDateTimeByStr(origDate, origEndStr);

      // 1. 檢查原值班行程
      const events1 = findOverlappingEvents(calendar, origStartTime, origEndTime, origPerson);
      if (events1.length === 0) {
        statusCell.setValue(`錯誤：找不到 ${origPerson} 的原值班行程`);
        checkCell.setValue(false);
        return;
      }
      if (events1.length > 1) {
        statusCell.setValue(`錯誤：${origPerson} 同時段有多筆行程，請手動確認日曆`);
        checkCell.setValue(false);
        return;
      }

      const check1 = validateTimeRange(events1[0], origStartTime, origEndTime);
      if (!check1.valid) {
        statusCell.setValue(`錯誤：${origPerson} 的填寫時段超出其原班表 (${check1.actualRange})`);
        checkCell.setValue(false);
        return;
      }

      // === 情況 A：雙向換班 ===
      if (isSwap) {
        const swapStartTime = combineDateTimeByStr(swapDate, swapStartStr);
        const swapEndTime = combineDateTimeByStr(swapDate, swapEndStr);
        const events2 = findOverlappingEvents(calendar, swapStartTime, swapEndTime, targetPerson);

        if (events2.length === 0) {
          statusCell.setValue(`錯誤：找不到 ${targetPerson} 的互換時段行程`);
          checkCell.setValue(false);
          return;
        }
        if (events2.length > 1) {
          statusCell.setValue(`錯誤：${targetPerson} 互換時段有多筆行程，請手動確認日曆`);
          checkCell.setValue(false);
          return;
        }

        const check2 = validateTimeRange(events2[0], swapStartTime, swapEndTime);
        if (!check2.valid) {
          statusCell.setValue(`錯誤：${targetPerson} 的填寫時段超出其原班表 (${check2.actualRange})`);
          checkCell.setValue(false);
          return;
        }

        const origDateStr = origDate instanceof Date ? Utilities.formatDate(origDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : origDate;
        const swapDateStr = swapDate instanceof Date ? Utilities.formatDate(swapDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : swapDate;

        // 🔴 換班原子性：兩筆改動視為一組操作，其中一筆失敗就復原另一筆已做的改動
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
          // 第二筆失敗，若第一筆已經改成功，立刻復原，確保不留孤兒行程
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
    statusCell.setValue("執行失敗: " + err.message);
    if (isChecked) checkCell.setValue(false);
  }
}

/**
 * 輔助函式：驗證填寫時間是否在原班表內
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
 * 核心函式：時段切割並寫入 ROW_ID（含 side 標記）識別
 */
function applyShiftChange(calendar, mainEvent, subStart, subEnd, newWorker, origWorker, descText, tag, rowToken) {
  const evStart = mainEvent.getStartTime();
  const evEnd = mainEvent.getEndTime();
  const floor = extractFloorSuffix(mainEvent.getTitle());
  const rowMeta = `[ROW_ID:${rowToken}]`;

  // 1. 完全吻合時段
  if (Math.abs(evStart.getTime() - subStart.getTime()) < 60000 && Math.abs(evEnd.getTime() - subEnd.getTime()) < 60000) {
    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    mainEvent.setTitle(newTitle);
    mainEvent.setDescription(`${rowMeta}\n${descText}\n--------------------\n` + cleanDescription(mainEvent.getDescription(), rowToken));
    return;
  }

  // 2. 時段切割（紀錄原起訖時間）
  const backupMeta = `[SPLIT_ORIG_TIME:${evStart.getTime()}-${evEnd.getTime()}]`;

  // 情況 A：切中間段
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

  // 情況 B：切後半段
  if (subStart.getTime() > evStart.getTime() && Math.abs(subEnd.getTime() - evEnd.getTime()) < 60000) {
    mainEvent.setTime(evStart, subStart);
    mainEvent.setDescription(`${rowMeta}\n${backupMeta}\n` + (mainEvent.getDescription() || ""));

    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    calendar.createEvent(newTitle, subStart, subEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n${descText}\n--------------------\n`
    });
    return;
  }

  // 情況 C：切前半段
  if (Math.abs(subStart.getTime() - evStart.getTime()) < 60000 && subEnd.getTime() < evEnd.getTime()) {
    mainEvent.setTime(subEnd, evEnd);
    mainEvent.setDescription(`${rowMeta}\n${backupMeta}\n` + (mainEvent.getDescription() || ""));

    const newTitle = tag.startsWith("【") ? `${tag}${newWorker}${floor}` : `${newWorker}${floor} ${tag}`;
    calendar.createEvent(newTitle, subStart, subEnd, {
      description: `[AUTO_SPLIT_CREATED]\n${rowMeta}\n${descText}\n--------------------\n`
    });
    return;
  }

  // 若未命中任何已知區間，拋出例外以避免「假成功」（外層會據此觸發換班回滾）
  throw new Error("無法計算時段切割，請確認起訖時間是否正確");
}

/**
 * 還原函式：根據 ROW_ID（含 side 標記）精準還原
 */
function revertEvents(calendar, rangeStart, rangeEnd, origWorker, rowToken) {
  const searchStart = new Date(rangeStart.getTime() - 24 * 60 * 60 * 1000);
  const searchEnd = new Date(rangeEnd.getTime() + 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(searchStart, searchEnd);
  const rowTokenTag = `[ROW_ID:${rowToken}]`;

  events.forEach(evt => {
    const desc = evt.getDescription() || "";

    // 只處理帶有這個 side 專屬 ROW_ID 的行程，兩邊互不影響，即使日期相近也不會誤觸
    if (!desc.includes(rowTokenTag)) return;

    // 1. 自動建出的子行程直接刪除
    if (desc.includes("[AUTO_SPLIT_CREATED]")) {
      evt.deleteEvent();
      return;
    }

    // 2. 被縮短過的原行程還原時間
    const splitMatch = desc.match(/\[SPLIT_ORIG_TIME:(\d+)-(\d+)\]/);
    if (splitMatch) {
      const origS = new Date(parseInt(splitMatch[1]));
      const origE = new Date(parseInt(splitMatch[2]));
      evt.setTime(origS, origE);
      evt.setDescription(desc.replace(/\[SPLIT_ORIG_TIME:\d+-\d+\]\n?/g, ""));
    }

    // 3. 還原標題文字
    const title = evt.getTitle();
    if (title.includes("[代班]") || title.includes("[換班]") || title.includes("【請假】")) {
      const floor = extractFloorSuffix(title);
      evt.setTitle(`${origWorker}${floor}`);
      evt.setDescription(cleanDescription(desc, rowToken));
    }
  });
}

/**
 * 輔助函式：從標題擷取括號後綴
 */
function extractFloorSuffix(title) {
  const pureTitle = title.replace(/\s*\[(換班|代班)\]/g, "").replace(/【請假】/g, "");
  const match = pureTitle.match(/(\([^\)]+\)|（[^）]+）)$/);
  return match ? match[0] : "";
}

/**
 * 輔助函式：清除特定 ROW_ID 的備忘文字
 */
function cleanDescription(desc, rowToken) {
  if (!desc) return "";
  // 用純字串 replaceAll 取代 RegExp，rowToken 未來若含特殊字元也不會造成解析錯誤
  let res = desc.replaceAll(`[ROW_ID:${rowToken}]\n`, "").replaceAll(`[ROW_ID:${rowToken}]`, "");
  res = res.replace(/【(換班|代班|請假)紀錄】[\s\S]*?--------------------\n?/g, "").trim();
  return res;
}

/**
 * 輔助函式：利用純文字時間字串精準解析
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
 * 輔助函式：維持原樣的姓名搜尋
 */
function findOverlappingEvents(calendar, startTime, endTime, personName) {
  const searchStart = new Date(startTime.getTime() - 12 * 60 * 60 * 1000);
  const searchEnd = new Date(endTime.getTime() + 12 * 60 * 60 * 1000);
  const events = calendar.getEvents(searchStart, searchEnd);

  const matched = [];
  for (let i = 0; i < events.length; i++) {
    const evt = events[i];
    // 維持原本的 includes 比對
    if (evt.getTitle().includes(personName)) {
      if (evt.getStartTime() < endTime && evt.getEndTime() > startTime) {
        matched.push(evt);
      }
    }
  }
  return matched;
}

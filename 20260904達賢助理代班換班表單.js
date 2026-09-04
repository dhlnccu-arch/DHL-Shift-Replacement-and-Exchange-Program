/**
 * 試算表編輯監聽器
 * 支援多列同時選取、批次拖曳填滿、Ctrl+Z/Y
 * 本版修復：長班部分時段換出時的「殘餘時段撞班檢測」
 *           徹底解決換班後同一人在兩處值班的分身漏洞
 */

// 單次觸發最多處理的列數
const MAX_ROWS_PER_RUN = 30;
// Apps Script 執行上限煞車點（毫秒）
const TIME_BUDGET_MS = 5 * 60 * 1000;

function handleSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  const startCol = range.getColumn();
  const endCol = range.getLastColumn();
  if (startCol > 11 || endCol < 11) return;

  const startRow = Math.max(2, range.getRow());
  let endRow = range.getLastRow();
  if (startRow > endRow) return;

  let truncated = false;
  if (endRow - startRow + 1 > MAX_ROWS_PER_RUN) {
    const realEndRow = endRow;
    endRow = startRow + MAX_ROWS_PER_RUN - 1;
    truncated = true;
    sheet.getRange(endRow + 1, 12, realEndRow - endRow, 1)
      .setValue(`尚未處理：單次批次上限為 ${MAX_ROWS_PER_RUN} 列，請稍後重新勾選此列（或分批操作）`);
  }

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

    for (let r = startRow; r <= endRow; r++) {
      if (Date.now() - runStart > TIME_BUDGET_MS) {
        sheet.getRange(r, 12, endRow - r + 1, 1)
          .setValue("尚未處理：本次執行時間已達上限，請重新勾選此列");
        truncated = true;
        break;
      }

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

  const isSwap = Boolean(
    swapDate && swapStartStr && swapEndStr && targetPerson &&
    targetPerson !== "請假" && targetPerson !== "無"
  );

  const tokenA = `${row}-A`;
  const tokenB = `${row}-B`;

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

      // 快取原時段的時間窗行程
      const eventsO = getEventsForWindow(calendar, origStartTime, origEndTime);

      // 1. 搜尋原值班行程
      const events1 = matchPersonEvents(eventsO, origStartTime, origEndTime, origPerson);
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
        
        // 快取互換時段的時間窗行程
        const eventsS = getEventsForWindow(calendar, swapStartTime, swapEndTime);
        const events2 = matchPersonEvents(eventsS, swapStartTime, swapEndTime, targetPerson);

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

        // 🔴 撞班檢測 1：配合換班人 (targetPerson) 到原時段 (origStartTime ~ origEndTime) 是否撞班？
        // 考慮 targetPerson 原本的班 (events2[0]) 在讓出 swap 時段後的「殘餘時段」
        const conflictForTarget = findRealConflicts(eventsO, origStartTime, origEndTime, targetPerson, events2[0], swapStartTime, swapEndTime);
        if (conflictForTarget.length > 0) {
          const timeRangeStr = formatEventTime(conflictForTarget[0]);
          statusCell.setValue(`錯誤退件：${targetPerson} 在原值班時段已有班 (${conflictForTarget[0].getTitle()} ${timeRangeStr})`);
          checkCell.setValue(false);
          return;
        }

        // 🔴 撞班檢測 2：申請換班人 (origPerson) 到互換時段 (swapStartTime ~ swapEndTime) 是否撞班？
        // 考慮 origPerson 原本的班 (events1[0]) 在讓出 orig 時段後的「殘餘時段」
        const conflictForOrig = findRealConflicts(eventsS, swapStartTime, swapEndTime, origPerson, events1[0], origStartTime, origEndTime);
        if (conflictForOrig.length > 0) {
          const timeRangeStr = formatEventTime(conflictForOrig[0]);
          statusCell.setValue(`錯誤退件：${origPerson} 在互換時段已有值班 (${conflictForOrig[0].getTitle()} ${timeRangeStr})`);
          checkCell.setValue(false);
          return;
        }

        const origDateStr = origDate instanceof Date ? Utilities.formatDate(origDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : origDate;
        const swapDateStr = swapDate instanceof Date ? Utilities.formatDate(swapDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : swapDate;

        // 原子性執行換班
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
        // 代班人在該時段是否有其他衝突值班
        const conflictForTarget = findRealConflicts(eventsO, origStartTime, origEndTime, targetPerson, null, null, null);
        if (conflictForTarget.length > 0) {
          const timeRangeStr = formatEventTime(conflictForTarget[0]);
          statusCell.setValue(`錯誤退件：${targetPerson} 在代班時段已有值班 (${conflictForTarget[0].getTitle()} ${timeRangeStr})`);
          checkCell.setValue(false);
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
    statusCell.setValue("執行失敗: " + err.message);
    if (isChecked) checkCell.setValue(false);
  }
}

/**
 * 輔助函式：抓取某時間窗前後 12 小時內的所有行程
 */
function getEventsForWindow(calendar, reqStart, reqEnd) {
  const searchStart = new Date(reqStart.getTime() - 12 * 60 * 60 * 1000);
  const searchEnd = new Date(reqEnd.getTime() + 12 * 60 * 60 * 1000);
  return calendar.getEvents(searchStart, searchEnd);
}

/**
 * 輔助函式：一般姓名重疊篩選（找原班行程使用）
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
 * 核心輔助函式：精確衝突計算
 * 若遇上即將被讓出的行程 (cedingEvent)，計算「讓出時段之外的殘餘時段」
 * 只要殘餘時段仍與目標時段重疊，即判定為撞班！
 */
function findRealConflicts(events, targetStart, targetEnd, personName, cedingEvent, cededStart, cededEnd) {
  const cedingId = cedingEvent ? cedingEvent.getId() : null;
  const conflicts = [];

  for (let i = 0; i < events.length; i++) {
    const evt = events[i];
    if (!evt.getTitle().includes(personName)) continue;

    const evStart = evt.getStartTime();
    const evEnd = evt.getEndTime();

    // 如果這筆就是準備要讓出/被切開的行程
    if (cedingId && evt.getId() === cedingId) {
      // 檢查前半殘餘段：evStart ~ cededStart 是否與目標時段重疊？
      if (evStart.getTime() < cededStart.getTime()) {
        if (targetStart < cededStart && evStart < targetEnd) {
          conflicts.push(evt);
          continue;
        }
      }
      // 檢查後半殘餘段：cededEnd ~ evEnd 是否與目標時段重疊？
      if (cededEnd.getTime() < evEnd.getTime()) {
        if (targetStart < evEnd && cededEnd < targetEnd) {
          conflicts.push(evt);
          continue;
        }
      }
    } else {
      // 一般其他行程：標準相鄰不衝突判定
      if (targetStart < evEnd && evStart < targetEnd) {
        conflicts.push(evt);
      }
    }
  }
  return conflicts;
}

/**
 * 輔助函式：格式化行程時間為 HH:mm-HH:mm
 */
function formatEventTime(evt) {
  const timeZone = Session.getScriptTimeZone();
  const s = Utilities.formatDate(evt.getStartTime(), timeZone, "HH:mm");
  const e = Utilities.formatDate(evt.getEndTime(), timeZone, "HH:mm");
  return `${s}-${e}`;
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
 * 核心函式：時段切割並寫入 ROW_ID
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

  // 2. 時段切割
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

  throw new Error("無法計算時段切割，請確認起訖時間是否正確");
}

/**
 * 還原函式：精準還原
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
 * 輔助函式：從標題擷取括號後綴
 */
function extractFloorSuffix(title) {
  const pureTitle = title.replace(/\s*\[(換班|代班)\]/g, "").replace(/【請假】/g, "");
  const match = pureTitle.match(/(\([^\)]+\)|（[^）]+）)$/);
  return match ? match[0] : "";
}

/**
 * 輔助函式：清除備忘文字
 */
function cleanDescription(desc, rowToken) {
  if (!desc) return "";
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
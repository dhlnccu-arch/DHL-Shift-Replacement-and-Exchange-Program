/**
 * 試算表編輯監聽器
 * 支援多列同時選取、批次拖曳填滿、Ctrl+Z/Y
 * 本版新增：全域撞班/重複值班 (Conflict) 自動防呆阻擋、
 *           排除本次待替換事件 ID、嚴格相鄰不重疊判定
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

      // 🟡 效能優化：同一個時間窗只打一次 calendar.getEvents()，
      // 「找原班行程」與後面的「撞班檢測」共用同一份結果在記憶體中篩選，不重複呼叫 API
      const eventsO = getEventsForWindow(calendar, origStartTime, origEndTime);

      // 1. 搜尋原值班行程
      const events1 = matchPersonEvents(eventsO, origStartTime, origEndTime, origPerson, []);
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

      const event1Id = events1[0].getId();

      // === 情況 A：雙向換班 ===
      if (isSwap) {
        const swapStartTime = combineDateTimeByStr(swapDate, swapStartStr);
        const swapEndTime = combineDateTimeByStr(swapDate, swapEndStr);
        // 同樣道理，swap 時間窗也只抓一次，供「找互換行程」與「撞班檢測」共用
        const eventsS = getEventsForWindow(calendar, swapStartTime, swapEndTime);
        const events2 = matchPersonEvents(eventsS, swapStartTime, swapEndTime, targetPerson, []);

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

        const event2Id = events2[0].getId();

        // 🔴 撞班檢測 1：targetPerson 接下 origStartTime ~ origEndTime，是否跟自己現存的其他行程衝突？
        // 必須排除 event2Id（若兩人在同一天換班，自己即將被換走的那筆舊班不算衝突）
        // 沿用上面已抓好的 eventsO，不再重新打一次 calendar.getEvents()
        const conflictForTarget = matchPersonEvents(eventsO, origStartTime, origEndTime, targetPerson, [event2Id]);
        if (conflictForTarget.length > 0) {
          const timeRangeStr = formatEventTime(conflictForTarget[0]);
          statusCell.setValue(`錯誤退件：${targetPerson} 在該時段已有值班 (${conflictForTarget[0].getTitle()} ${timeRangeStr})`);
          checkCell.setValue(false);
          return;
        }

        // 🔴 撞班檢測 2：origPerson 接下 swapStartTime ~ swapEndTime，是否跟自己現存的其他行程衝突？
        // 必須排除 event1Id（自己即將被換走的那筆舊班不算衝突）
        // 沿用上面已抓好的 eventsS，不再重新打一次 calendar.getEvents()
        const conflictForOrig = matchPersonEvents(eventsS, swapStartTime, swapEndTime, origPerson, [event1Id]);
        if (conflictForOrig.length > 0) {
          const timeRangeStr = formatEventTime(conflictForOrig[0]);
          statusCell.setValue(`錯誤退件：${origPerson} 在該時段已有值班 (${conflictForOrig[0].getTitle()} ${timeRangeStr})`);
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
        // 🔴 撞班檢測：代班人 targetPerson 在代班時段是否已有其他班？
        // 沿用上面已抓好的 eventsO（跟找原班行程同一個時間窗），不再重新打一次 calendar.getEvents()
        const conflictForTarget = matchPersonEvents(eventsO, origStartTime, origEndTime, targetPerson, []);
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
 * 輔助函式：衝突行程偵測（排除指定 ID、嚴格排除相鄰邊界）
 * 衝突公式：startA < endB && startB < endA
 */
/**
 * 輔助函式：抓取某時間窗前後 12 小時內的所有行程（原始、未篩選）
 * 供同一個時間窗的多種篩選（找原班行程 / 撞班檢測）共用同一次 API 呼叫，
 * 避免像先前版本一樣對同一個時間窗重複打 calendar.getEvents()
 */
function getEventsForWindow(calendar, reqStart, reqEnd) {
  const searchStart = new Date(reqStart.getTime() - 12 * 60 * 60 * 1000);
  const searchEnd = new Date(reqEnd.getTime() + 12 * 60 * 60 * 1000);
  return calendar.getEvents(searchStart, searchEnd);
}

/**
 * 輔助函式：從一批已抓好的行程中，篩出屬於某人、且與請求時段重疊的行程
 * 這是 findOverlappingEvents 與 findConflictingEvents 共用的唯一一份重疊判定邏輯，
 * 避免同一條公式在兩個函式裡各寫一份、日後容易改一邊漏改另一邊
 */
function matchPersonEvents(events, reqStart, reqEnd, personName, excludeIds) {
  const exclude = excludeIds || [];
  const matched = [];
  for (let i = 0; i < events.length; i++) {
    const evt = events[i];

    // 1. 排除本次交易中即將被取代的舊行程（找原班行程時 excludeIds 傳空陣列即可）
    if (exclude.includes(evt.getId())) continue;

    // 2. 姓名比對
    if (!evt.getTitle().includes(personName)) continue;

    // 3. 嚴格重疊定義（相鄰不算重疊）
    const evStart = evt.getStartTime();
    const evEnd = evt.getEndTime();
    if (reqStart < evEnd && evStart < reqEnd) {
      matched.push(evt);
    }
  }
  return matched;
}

/**
 * 撞班檢測：某人在指定時段是否已有其他（非本次交易將被取代的）行程
 */
function findConflictingEvents(calendar, reqStart, reqEnd, personName, excludeIds) {
  const events = getEventsForWindow(calendar, reqStart, reqEnd);
  return matchPersonEvents(events, reqStart, reqEnd, personName, excludeIds);
}

/**
 * 輔助函式：格式化行程起訖時間為 HH:mm-HH:mm 供錯誤訊息顯示
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

/**
 * 輔助函式：搜尋時間上有交集、且屬於原人員的行程
 */
function findOverlappingEvents(calendar, startTime, endTime, personName) {
  const events = getEventsForWindow(calendar, startTime, endTime);
  return matchPersonEvents(events, startTime, endTime, personName, []);
}
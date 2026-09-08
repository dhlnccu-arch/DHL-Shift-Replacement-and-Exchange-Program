/**
 * 助理代班／換班系統 v6（穩定版）
 *
 * 欄位架構對照：
 * A–J 欄：原 Google 表單資料
 * A:時間戳記
 * B:申請人
 * C:原班日
 * D:起時
 * E:迄時
 * F:配合人
 * G:換班日
 * H:換班起時
 * I:換班迄時
 * J:備註
 *
 * K 欄 (11)：申請人 Email
 * L 欄 (12)：☑ 審核確認 / 執行換班
 * M 欄 (13)：執行狀態
 * N 欄 (14)：☑ 審核通知
 * O 欄 (15)：核准通知記錄
 * P 欄 (16)：退件通知記錄
 * Q 欄 (17)：申請 ID（系統自動產生 UUID，可隱藏）
 *
 * Apps Script 觸發條件共 2 個：
 *
 * 1. handleSheetEdit
 *    事件來源：來自試算表
 *    事件類型：編輯時
 *
 * 2. onFormSubmitPrecheck
 *    事件來源：來自試算表
 *    事件類型：表單提交時
 *
 * 表單提交時：
 * - 只做「唯讀預檢」
 * - 不會修改 Google Calendar
 * - 預檢失敗 → 自動退件
 * - 預檢成功 → M 顯示待管理員審核
 *
 * 管理員勾 L 時：
 * - 重新讀取最新 Google Calendar
 * - 再執行一次同一套 analyzeRequest()
 * - 通過才真正修改 Calendar
 */

const MAX_ROWS_PER_RUN = 30;
const TIME_BUDGET_MS = 5 * 60 * 1000;


// ============================================================
// CONFIG
// ============================================================

const CONFIG = {

  EXECUTE_CHECK_COL: 12, // L
  STATUS_COL: 13,        // M

  APPROVE_CHECK_COL: 14, // N
  APPROVE_LOG_COL: 15,   // O
  REJECT_LOG_COL: 16,    // P

  REQUEST_ID_COL: 17,    // Q

  APPLICANT_EMAIL_COL: 11, // K

  ADMIN_EMAILS: [
    "dhl.nccu@gmail.com"
  ],

  STAFF_DIRECTORY_SHEET_NAME: "員工名冊",

  STAFF_NAME_COL: 1,
  STAFF_EMAIL_COL: 2
};


// ============================================================
// 表單提交 → 自動前置預檢
// ============================================================

function onFormSubmitPrecheck(e) {

  if (!e || !e.range) {

    console.error(
      "onFormSubmitPrecheck: 找不到 e.range，請確認觸發器為『來自試算表 / 表單提交時』"
    );

    return;
  }


  const sheet = e.range.getSheet();

  const row = e.range.getRow();

  const statusCell =
    sheet.getRange(
      row,
      CONFIG.STATUS_COL
    );

  const checkCell =
    sheet.getRange(
      row,
      CONFIG.EXECUTE_CHECK_COL
    );


  // 每一筆申請建立穩定 UUID
  ensureRequestId(
    sheet,
    row
  );


  const lock =
    LockService.getScriptLock();


  if (!lock.tryLock(15000)) {

    statusCell.setValue(
      "預檢暫時無法完成：系統忙碌，請管理員稍後確認"
    );

    return;
  }


  try {

    const calendars =
      CalendarApp.getCalendarsByName(
        "達賢館創新組助理值班"
      );


    if (calendars.length === 0) {

      statusCell.setValue(
        "預檢暫時無法完成：找不到指定日曆，請管理員確認"
      );

      return;
    }


    const result =
      analyzeRequest(
        sheet,
        calendars[0],
        row
      );


    if (result.ok) {

      checkCell.setValue(false);

      statusCell.setValue(
        "預檢通過，待管理員審核"
      );

      return;
    }


    checkCell.setValue(false);


    const message =
      `預檢未通過：${result.message}`;


    statusCell.setValue(
      message
    );


    notifyRejectionIfNeeded(
      sheet,
      row,
      message
    );


  } catch (err) {

    console.error(
      `第 ${row} 列前置預檢失敗: ${err.message}`
    );


    statusCell.setValue(
      "預檢暫時無法完成：" +
      err.message +
      "（請管理員確認）"
    );


  } finally {

    lock.releaseLock();

  }
}


// ============================================================
// 統一分析函式
// 預檢與正式審核都使用這一套
// ============================================================

function analyzeRequest(
  sheet,
  calendar,
  row
) {

  const origPerson =
    sheet.getRange(row, 2)
      .getValue()
      .toString()
      .trim();


  const origDate =
    sheet.getRange(row, 3)
      .getValue();


  const origStartStr =
    sheet.getRange(row, 4)
      .getDisplayValue()
      .toString()
      .trim();


  const origEndStr =
    sheet.getRange(row, 5)
      .getDisplayValue()
      .toString()
      .trim();


  const targetPerson =
    sheet.getRange(row, 6)
      .getValue()
      .toString()
      .trim();


  const swapDate =
    sheet.getRange(row, 7)
      .getValue();


  const swapStartStr =
    sheet.getRange(row, 8)
      .getDisplayValue()
      .toString()
      .trim();


  const swapEndStr =
    sheet.getRange(row, 9)
      .getDisplayValue()
      .toString()
      .trim();


  const base = {

    row,

    origPerson,
    origDate,
    origStartStr,
    origEndStr,

    targetPerson,

    swapDate,
    swapStartStr,
    swapEndStr
  };


  // -----------------------------
  // 基本資料檢查
  // -----------------------------

  if (!origPerson || !origDate) {

    return {
      ...base,
      ok: false,
      message:
        "請填寫原值班人員與原值班日期"
    };
  }


  if (!origStartStr || !origEndStr) {

    return {
      ...base,
      ok: false,
      message:
        "請填寫原值班起訖時間"
    };
  }


  const origStartTime =
    combineDateTimeByStr(
      origDate,
      origStartStr
    );


  const origEndTime =
    combineDateTimeByStr(
      origDate,
      origEndStr
    );


  if (
    !isValidTimeRange(
      origStartTime,
      origEndTime
    )
  ) {

    return {
      ...base,
      ok: false,
      message:
        "原值班起訖時間格式不正確"
    };
  }


  // -----------------------------
  // 找原班
  // -----------------------------

  const eventsO =
    getEventsForWindow(
      calendar,
      origStartTime,
      origEndTime
    );


  const events1 =
    matchPersonEvents(
      eventsO,
      origStartTime,
      origEndTime,
      origPerson
    );


  if (events1.length === 0) {

    return {
      ...base,
      ok: false,
      message:
        `找不到 ${origPerson} 的原值班行程`
    };
  }


  if (events1.length > 1) {

    return {
      ...base,
      ok: false,
      message:
        `${origPerson} 同時段有多筆行程，請聯絡管理員確認`
    };
  }


  const check1 =
    validateTimeRange(
      events1[0],
      origStartTime,
      origEndTime
    );


  if (!check1.valid) {

    return {
      ...base,
      ok: false,
      message:
        `${origPerson} 的填寫時段超出其原班表 (${check1.actualRange})`
    };
  }


  // -----------------------------
  // 判斷申請類型
  // -----------------------------

  const hasTarget =
    Boolean(
      targetPerson &&
      targetPerson !== "請假" &&
      targetPerson !== "無"
    );


  const hasAnySwapField =
    Boolean(
      swapDate ||
      swapStartStr ||
      swapEndStr
    );


  const hasAllSwapFields =
    Boolean(
      swapDate &&
      swapStartStr &&
      swapEndStr
    );


  // 有配合人，但換班資料只填一部分
  if (
    hasTarget &&
    hasAnySwapField &&
    !hasAllSwapFields
  ) {

    return {
      ...base,
      ok: false,
      message:
        "配合換班日期與起訖時間填寫不完整；若為單純代班，請將配合換班日期與時間全部留白"
    };
  }


  const isSwap =
    Boolean(
      hasTarget &&
      hasAllSwapFields
    );


  const type =
    isSwap
      ? "swap"
      : (
          hasTarget
            ? "sub"
            : "leave"
        );


  // ==========================================================
  // 雙向換班
  // ==========================================================

  if (isSwap) {

    const swapStartTime =
      combineDateTimeByStr(
        swapDate,
        swapStartStr
      );


    const swapEndTime =
      combineDateTimeByStr(
        swapDate,
        swapEndStr
      );


    if (
      !isValidTimeRange(
        swapStartTime,
        swapEndTime
      )
    ) {

      return {
        ...base,
        ok: false,
        message:
          "配合換班起訖時間格式不正確"
      };
    }


    const eventsS =
      getEventsForWindow(
        calendar,
        swapStartTime,
        swapEndTime
      );


    const events2 =
      matchPersonEvents(
        eventsS,
        swapStartTime,
        swapEndTime,
        targetPerson
      );


    if (events2.length === 0) {

      return {
        ...base,
        ok: false,
        message:
          `找不到 ${targetPerson} 的互換時段行程`
      };
    }


    if (events2.length > 1) {

      return {
        ...base,
        ok: false,
        message:
          `${targetPerson} 互換時段有多筆行程，請聯絡管理員確認`
      };
    }


    const check2 =
      validateTimeRange(
        events2[0],
        swapStartTime,
        swapEndTime
      );


    if (!check2.valid) {

      return {
        ...base,
        ok: false,
        message:
          `${targetPerson} 的填寫時段超出其原班表 (${check2.actualRange})`
      };
    }


    // 配合人 → 原申請人的時段
    const conflictForTarget =
      findRealConflicts(

        eventsO,

        origStartTime,
        origEndTime,

        targetPerson,

        events2[0],

        swapStartTime,
        swapEndTime
      );


    if (
      conflictForTarget.length > 0
    ) {

      const evt =
        conflictForTarget[0];


      return {
        ...base,
        ok: false,
        message:
          `${targetPerson} 在原值班時段已有班 (${evt.getTitle()} ${formatEventTime(evt)})`
      };
    }


    // 申請人 → 配合人的時段
    const conflictForOrig =
      findRealConflicts(

        eventsS,

        swapStartTime,
        swapEndTime,

        origPerson,

        events1[0],

        origStartTime,
        origEndTime
      );


    if (
      conflictForOrig.length > 0
    ) {

      const evt =
        conflictForOrig[0];


      return {
        ...base,
        ok: false,
        message:
          `${origPerson} 在互換時段已有值班 (${evt.getTitle()} ${formatEventTime(evt)})`
      };
    }


    return {

      ...base,

      ok: true,

      type,

      origStartTime,
      origEndTime,

      swapStartTime,
      swapEndTime,

      eventsO,
      eventsS,

      origEvent:
        events1[0],

      swapEvent:
        events2[0]
    };
  }


  // ==========================================================
  // 單向代班
  // ==========================================================

  if (hasTarget) {

    const conflictForTarget =
      findRealConflicts(

        eventsO,

        origStartTime,
        origEndTime,

        targetPerson,

        null,
        null,
        null
      );


    if (
      conflictForTarget.length > 0
    ) {

      const evt =
        conflictForTarget[0];


      return {
        ...base,
        ok: false,
        message:
          `${targetPerson} 在代班時段已有值班 (${evt.getTitle()} ${formatEventTime(evt)})`
      };
    }
  }


  // ==========================================================
  // 單向代班或純請假 → OK
  // ==========================================================

  return {

    ...base,

    ok: true,

    type,

    origStartTime,
    origEndTime,

    eventsO,

    origEvent:
      events1[0],

    swapEvent:
      null
  };
}


// ============================================================
// 試算表編輯監聽器
// ============================================================

function handleSheetEdit(e) {

  if (!e || !e.range) {
    return;
  }


  const range =
    e.range;


  const sheet =
    range.getSheet();


  const startCol =
    range.getColumn();


  const endCol =
    range.getLastColumn();


  const touchesExecuteCol =
    !(
      startCol >
        CONFIG.EXECUTE_CHECK_COL ||
      endCol <
        CONFIG.EXECUTE_CHECK_COL
    );


  const touchesApproveCol =
    !(
      startCol >
        CONFIG.APPROVE_CHECK_COL ||
      endCol <
        CONFIG.APPROVE_CHECK_COL
    );


  if (
    !touchesExecuteCol &&
    !touchesApproveCol
  ) {

    return;
  }


  // ==========================================================
  // 先處理 L
  // ==========================================================

  if (touchesExecuteCol) {

    const startRow =
      Math.max(
        2,
        range.getRow()
      );


    let endRow =
      range.getLastRow();


    if (startRow > endRow) {
      return;
    }


    let truncated =
      false;


    if (
      endRow -
      startRow +
      1 >
      MAX_ROWS_PER_RUN
    ) {

      const realEndRow =
        endRow;


      endRow =
        startRow +
        MAX_ROWS_PER_RUN -
        1;


      truncated =
        true;


      sheet.getRange(
        endRow + 1,
        CONFIG.STATUS_COL,
        realEndRow - endRow,
        1
      ).setValue(

        `尚未處理：單次批次上限為 ${MAX_ROWS_PER_RUN} 列，請稍後重新勾選此列（或分批操作）`

      );
    }


    const lock =
      LockService.getScriptLock();


    if (
      !lock.tryLock(15000)
    ) {

      sheet.getRange(

        startRow,
        CONFIG.STATUS_COL,
        endRow - startRow + 1,
        1

      ).setValue(

        "系統忙碌中，請稍候重試"

      );


      return;
    }


    const runStart =
      Date.now();


    try {

      const calendars =
        CalendarApp.getCalendarsByName(
          "達賢館創新組助理值班"
        );


      if (
        calendars.length === 0
      ) {

        sheet.getRange(

          startRow,
          CONFIG.STATUS_COL,
          endRow - startRow + 1,
          1

        ).setValue(

          "錯誤：找不到指定日曆"

        );


        return;
      }


      const calendar =
        calendars[0];


      for (
        let r = startRow;
        r <= endRow;
        r++
      ) {

        if (
          Date.now() -
          runStart >
          TIME_BUDGET_MS
        ) {

          sheet.getRange(

            r,
            CONFIG.STATUS_COL,
            endRow - r + 1,
            1

          ).setValue(

            "尚未處理：本次執行時間已達上限，請重新勾選此列"

          );


          truncated =
            true;


          break;
        }


        try {

          processSingleRow(
            sheet,
            calendar,
            r
          );


        } catch (rowErr) {

          console.error(
            `第 ${r} 列處理發生非預期錯誤: ${rowErr.message}`
          );


          sheet.getRange(
            r,
            CONFIG.STATUS_COL
          ).setValue(

            "執行失敗（非預期錯誤）: " +
            rowErr.message

          );


          const checkCell =
            sheet.getRange(
              r,
              CONFIG.EXECUTE_CHECK_COL
            );


          if (
            checkCell.getValue() ===
            true
          ) {

            checkCell.setValue(
              false
            );
          }
        }
      }


    } catch (err) {

      console.error(err);


      sheet.getRange(

        startRow,
        CONFIG.STATUS_COL,
        endRow - startRow + 1,
        1

      ).setValue(

        "系統發生錯誤，請聯絡管理員: " +
        err.message

      );


    } finally {

      lock.releaseLock();

    }


    if (truncated) {

      console.log(

        `本次觸發因批次列數或時間上限而部分列未處理（範圍 ${startRow}-${range.getLastRow()}）`

      );
    }
  }


  // ==========================================================
  // L 處理完成後再處理 N
  // ==========================================================

  if (touchesApproveCol) {

    handleApprovalEditRange(
      sheet,
      range
    );
  }
}


// ============================================================
// 正式處理單筆申請
// ============================================================

function processSingleRow(
  sheet,
  calendar,
  row
) {

  const checkCell =
    sheet.getRange(
      row,
      CONFIG.EXECUTE_CHECK_COL
    );


  const statusCell =
    sheet.getRange(
      row,
      CONFIG.STATUS_COL
    );


  const isChecked =
    checkCell.getValue() === true;


  const statusVal =
    statusCell
      .getValue()
      .toString();


  try {

    // ========================================================
    // L 被取消 → 還原班表
    // ========================================================

    if (!isChecked) {

      if (
        !statusVal.includes(
          "已更新日曆"
        )
      ) {

        return;
      }


      const origPerson =
        sheet.getRange(row, 2)
          .getValue()
          .toString()
          .trim();


      const origDate =
        sheet.getRange(row, 3)
          .getValue();


      const origStartStr =
        sheet.getRange(row, 4)
          .getDisplayValue()
          .toString()
          .trim();


      const origEndStr =
        sheet.getRange(row, 5)
          .getDisplayValue()
          .toString()
          .trim();


      const targetPerson =
        sheet.getRange(row, 6)
          .getValue()
          .toString()
          .trim();


      const swapDate =
        sheet.getRange(row, 7)
          .getValue();


      const swapStartStr =
        sheet.getRange(row, 8)
          .getDisplayValue()
          .toString()
          .trim();


      const swapEndStr =
        sheet.getRange(row, 9)
          .getDisplayValue()
          .toString()
          .trim();


      const tokenBase =
        getRequestTokenBase(
          sheet,
          row,
          true
        );


      const tokenA =
        `${tokenBase}-A`;


      const tokenB =
        `${tokenBase}-B`;


      const origStartTime =
        combineDateTimeByStr(
          origDate,
          origStartStr
        );


      const origEndTime =
        combineDateTimeByStr(
          origDate,
          origEndStr
        );


      revertEvents(

        calendar,

        origStartTime,
        origEndTime,

        origPerson,

        tokenA
      );


      const isSwap =
        Boolean(

          swapDate &&
          swapStartStr &&
          swapEndStr &&

          targetPerson &&

          targetPerson !==
            "請假" &&

          targetPerson !==
            "無"

        );


      if (isSwap) {

        const swapStartTime =
          combineDateTimeByStr(
            swapDate,
            swapStartStr
          );


        const swapEndTime =
          combineDateTimeByStr(
            swapDate,
            swapEndStr
          );


        revertEvents(

          calendar,

          swapStartTime,
          swapEndTime,

          targetPerson,

          tokenB
        );
      }


      statusCell.setValue(
        ""
      );


      return;
    }


    // 已執行過，不重跑
    if (
      statusVal.includes(
        "已更新日曆"
      )
    ) {

      return;
    }


    // ========================================================
    // 正式審核
    // 每次都重新讀最新 Calendar
    // ========================================================

    const analysis =
      analyzeRequest(
        sheet,
        calendar,
        row
      );


    if (!analysis.ok) {

      denyRow(

        sheet,
        row,
        checkCell,
        statusCell,

        `正式審核未通過：${analysis.message}`

      );


      return;
    }


    const tokenBase =
      getRequestTokenBase(
        sheet,
        row,
        false
      );


    const tokenA =
      `${tokenBase}-A`;


    const tokenB =
      `${tokenBase}-B`;


    // ========================================================
    // 雙向換班
    // ========================================================

    if (
      analysis.type === "swap"
    ) {

      const origDateStr =
        analysis.origDate
        instanceof Date
          ? Utilities.formatDate(

              analysis.origDate,

              Session.getScriptTimeZone(),

              "yyyy/MM/dd"

            )
          : analysis.origDate;


      const swapDateStr =
        analysis.swapDate
        instanceof Date
          ? Utilities.formatDate(

              analysis.swapDate,

              Session.getScriptTimeZone(),

              "yyyy/MM/dd"

            )
          : analysis.swapDate;


      let firstSideApplied =
        false;


      try {

        applyShiftChange(

          calendar,

          analysis.origEvent,

          analysis.origStartTime,
          analysis.origEndTime,

          analysis.targetPerson,
          analysis.origPerson,

          `【換班紀錄】
- 實際到勤：${analysis.targetPerson}
- 原定值班：${analysis.origPerson}
- 互換對象時段：${swapDateStr}`,

          "[換班]",

          tokenA
        );


        firstSideApplied =
          true;


        applyShiftChange(

          calendar,

          analysis.swapEvent,

          analysis.swapStartTime,
          analysis.swapEndTime,

          analysis.origPerson,
          analysis.targetPerson,

          `【換班紀錄】
- 實際到勤：${analysis.origPerson}
- 原定值班：${analysis.targetPerson}
- 互換對象時段：${origDateStr}`,

          "[換班]",

          tokenB
        );


      } catch (swapErr) {


        if (
          firstSideApplied
        ) {

          try {

            revertEvents(

              calendar,

              analysis.origStartTime,
              analysis.origEndTime,

              analysis.origPerson,

              tokenA
            );


          } catch (
            rollbackErr
          ) {

            console.error(

              `換班回滾失敗（第 ${row} 列）: ${rollbackErr.message}`

            );


            throw new Error(

              `換班失敗且自動回滾也失敗，請手動檢查日曆: ${swapErr.message}`

            );
          }
        }


        throw swapErr;
      }


      statusCell.setValue(
        "已更新日曆（雙向換班完成）"
      );


      return;
    }


    // ========================================================
    // 單向代班
    // ========================================================

    if (
      analysis.type === "sub"
    ) {


      const desc =
        `【代班紀錄】
- 實際到勤：${analysis.targetPerson}
- 原定值班：${analysis.origPerson}（請假由他人代班）`;


      applyShiftChange(

        calendar,

        analysis.origEvent,

        analysis.origStartTime,
        analysis.origEndTime,

        analysis.targetPerson,
        analysis.origPerson,

        desc,

        "[代班]",

        tokenA
      );


      statusCell.setValue(
        "已更新日曆（代班完成）"
      );


      return;
    }


    // ========================================================
    // 純請假
    // ========================================================

    const desc =
      `【請假紀錄】
- 原定值班：${analysis.origPerson}（請假無人代理）`;


    applyShiftChange(

      calendar,

      analysis.origEvent,

      analysis.origStartTime,
      analysis.origEndTime,

      analysis.origPerson,
      analysis.origPerson,

      desc,

      "【請假】",

      tokenA
    );


    statusCell.setValue(
      "已更新日曆（請假完成）"
    );


  } catch (err) {


    const failMsg =
      "執行失敗: " +
      err.message;


    statusCell.setValue(
      failMsg
    );


    if (isChecked) {

      checkCell.setValue(
        false
      );


      notifyRejectionIfNeeded(

        sheet,
        row,

        failMsg
      );
    }
  }
}


// ============================================================
// Calendar 查詢
// ============================================================

function getEventsForWindow(
  calendar,
  reqStart,
  reqEnd
) {

  const searchStart =
    new Date(

      reqStart.getTime() -

      12 *
      60 *
      60 *
      1000

    );


  const searchEnd =
    new Date(

      reqEnd.getTime() +

      12 *
      60 *
      60 *
      1000

    );


  return calendar.getEvents(
    searchStart,
    searchEnd
  );
}


// ============================================================
// 精準姓名比對
// ============================================================

function matchPersonEvents(
  events,
  reqStart,
  reqEnd,
  personName
) {

  const matched = [];


  for (
    let i = 0;
    i < events.length;
    i++
  ) {

    const evt =
      events[i];


    if (
      extractWorkerName(
        evt.getTitle()
      ) !== personName
    ) {

      continue;
    }


    const evStart =
      evt.getStartTime();


    const evEnd =
      evt.getEndTime();


    if (
      reqStart < evEnd &&
      evStart < reqEnd
    ) {

      matched.push(
        evt
      );
    }
  }


  return matched;
}


// ============================================================
// 撞班判斷
// ============================================================

function findRealConflicts(

  events,

  targetStart,
  targetEnd,

  personName,

  cedingEvent,

  cededStart,
  cededEnd

) {


  const cedingId =
    cedingEvent
      ? cedingEvent.getId()
      : null;


  const conflicts = [];


  for (
    let i = 0;
    i < events.length;
    i++
  ) {


    const evt =
      events[i];


    if (
      extractWorkerName(
        evt.getTitle()
      ) !== personName
    ) {

      continue;
    }


    const evStart =
      evt.getStartTime();


    const evEnd =
      evt.getEndTime();


    if (
      cedingId &&
      evt.getId() ===
        cedingId
    ) {


      // 讓出時段之前的殘餘
      if (
        evStart.getTime() <
        cededStart.getTime()
      ) {

        if (
          targetStart <
            cededStart &&
          evStart <
            targetEnd
        ) {

          conflicts.push(
            evt
          );

          continue;
        }
      }


      // 讓出時段之後的殘餘
      if (
        cededEnd.getTime() <
        evEnd.getTime()
      ) {

        if (
          targetStart <
            evEnd &&
          cededEnd <
            targetEnd
        ) {

          conflicts.push(
            evt
          );

          continue;
        }
      }


    } else if (

      targetStart <
        evEnd &&

      evStart <
        targetEnd

    ) {

      conflicts.push(
        evt
      );
    }
  }


  return conflicts;
}


// ============================================================
// 從日曆標題取得「真正的姓名」
// ============================================================

function extractWorkerName(
  title
) {

  let name =
    (title || "")
      .toString()
      .trim();


  // 【請假】王小明(3F)
  name =
    name.replace(
      /^【請假】\s*/,
      ""
    );


  // 王小明(3F) [換班]
  // 王小明(3F) [代班]
  name =
    name.replace(
      /\s*\[(換班|代班)\]\s*$/,
      ""
    );


  // 移除最後樓層括號
  name =
    name.replace(
      /\s*(\([^\)]*\)|（[^）]*）)\s*$/,
      ""
    );


  return name.trim();
}


// ============================================================
// 時間是否合法
// ============================================================

function isValidTimeRange(
  start,
  end
) {

  return (

    start instanceof Date &&

    end instanceof Date &&

    !isNaN(
      start.getTime()
    ) &&

    !isNaN(
      end.getTime()
    ) &&

    start < end

  );
}


// ============================================================
// UUID
// ============================================================

function ensureRequestId(
  sheet,
  row
) {

  const cell =
    sheet.getRange(
      row,
      CONFIG.REQUEST_ID_COL
    );


  let id =
    cell.getValue()
      .toString()
      .trim();


  if (!id) {

    id =
      Utilities.getUuid();


    cell.setValue(
      id
    );
  }


  return id;
}


// ============================================================
// Calendar ROW_ID 基底
// ============================================================

function getRequestTokenBase(
  sheet,
  row,
  legacyForExisting
) {

  const cell =
    sheet.getRange(
      row,
      CONFIG.REQUEST_ID_COL
    );


  const existing =
    cell.getValue()
      .toString()
      .trim();


  if (existing) {

    return existing;
  }


  // 相容以前以列號當 ROW_ID 的舊資料
  if (legacyForExisting) {

    return String(row);
  }


  return ensureRequestId(
    sheet,
    row
  );
}


// ============================================================
// 顯示 Calendar 時段
// ============================================================

function formatEventTime(
  evt
) {

  const timeZone =
    Session.getScriptTimeZone();


  const s =
    Utilities.formatDate(

      evt.getStartTime(),

      timeZone,

      "HH:mm"

    );


  const e =
    Utilities.formatDate(

      evt.getEndTime(),

      timeZone,

      "HH:mm"

    );


  return `${s}-${e}`;
}


// ============================================================
// 確認申請時段是否在原班表內
// ============================================================

function validateTimeRange(
  event,
  reqStart,
  reqEnd
) {

  const evStart =
    event.getStartTime()
      .getTime();


  const evEnd =
    event.getEndTime()
      .getTime();


  const s =
    reqStart.getTime();


  const e =
    reqEnd.getTime();


  const timeZone =
    Session.getScriptTimeZone();


  const actualStartStr =
    Utilities.formatDate(

      event.getStartTime(),

      timeZone,

      "HH:mm"

    );


  const actualEndStr =
    Utilities.formatDate(

      event.getEndTime(),

      timeZone,

      "HH:mm"

    );


  const actualRange =
    `${actualStartStr}-${actualEndStr}`;


  if (

    s <
      evStart -
      60000 ||

    e >
      evEnd +
      60000

  ) {

    return {

      valid: false,

      actualRange:
        actualRange

    };
  }


  return {

    valid: true,

    actualRange:
      actualRange

  };
}


// ============================================================
// 套用班表異動
// ============================================================

function applyShiftChange(

  calendar,

  mainEvent,

  subStart,
  subEnd,

  newWorker,
  origWorker,

  descText,

  tag,

  rowToken

) {


  const evStart =
    mainEvent.getStartTime();


  const evEnd =
    mainEvent.getEndTime();


  const floor =
    extractFloorSuffix(
      mainEvent.getTitle()
    );


  const rowMeta =
    `[ROW_ID:${rowToken}]`;


  // ==========================================================
  // 整段吻合
  // ==========================================================

  if (

    Math.abs(
      evStart.getTime() -
      subStart.getTime()
    ) < 60000 &&

    Math.abs(
      evEnd.getTime() -
      subEnd.getTime()
    ) < 60000

  ) {


    const newTitle =
      tag.startsWith("【")
        ? `${tag}${newWorker}${floor}`
        : `${newWorker}${floor} ${tag}`;


    mainEvent.setTitle(
      newTitle
    );


    mainEvent.setDescription(

      `${rowMeta}
${descText}
--------------------
` +

      cleanDescription(
        mainEvent.getDescription(),
        rowToken
      )

    );


    return;
  }


  const backupMeta =
    `[SPLIT_ORIG_TIME:${evStart.getTime()}-${evEnd.getTime()}]`;


  // ==========================================================
  // 切中間
  // ==========================================================

  if (

    subStart.getTime() >
      evStart.getTime() &&

    subEnd.getTime() <
      evEnd.getTime()

  ) {


    mainEvent.setTime(
      evStart,
      subStart
    );


    mainEvent.setDescription(

      `${rowMeta}
${backupMeta}
` +

      (
        mainEvent.getDescription() ||
        ""
      )

    );


    const midTitle =
      tag.startsWith("【")
        ? `${tag}${newWorker}${floor}`
        : `${newWorker}${floor} ${tag}`;


    calendar.createEvent(

      midTitle,

      subStart,
      subEnd,

      {
        description:
`[AUTO_SPLIT_CREATED]
${rowMeta}
${descText}
--------------------
`
      }

    );


    calendar.createEvent(

      `${origWorker}${floor}`,

      subEnd,
      evEnd,

      {
        description:
`[AUTO_SPLIT_CREATED]
${rowMeta}
【原班後半段】
--------------------
`
      }

    );


    return;
  }


  // ==========================================================
  // 切後半
  // ==========================================================

  if (

    subStart.getTime() >
      evStart.getTime() &&

    Math.abs(
      subEnd.getTime() -
      evEnd.getTime()
    ) < 60000

  ) {


    mainEvent.setTime(
      evStart,
      subStart
    );


    mainEvent.setDescription(

      `${rowMeta}
${backupMeta}
` +

      (
        mainEvent.getDescription() ||
        ""
      )

    );


    const newTitle =
      tag.startsWith("【")
        ? `${tag}${newWorker}${floor}`
        : `${newWorker}${floor} ${tag}`;


    calendar.createEvent(

      newTitle,

      subStart,
      subEnd,

      {
        description:
`[AUTO_SPLIT_CREATED]
${rowMeta}
${descText}
--------------------
`
      }

    );


    return;
  }


  // ==========================================================
  // 切前半
  // ==========================================================

  if (

    Math.abs(
      subStart.getTime() -
      evStart.getTime()
    ) < 60000 &&

    subEnd.getTime() <
      evEnd.getTime()

  ) {


    mainEvent.setTime(
      subEnd,
      evEnd
    );


    mainEvent.setDescription(

      `${rowMeta}
${backupMeta}
` +

      (
        mainEvent.getDescription() ||
        ""
      )

    );


    const newTitle =
      tag.startsWith("【")
        ? `${tag}${newWorker}${floor}`
        : `${newWorker}${floor} ${tag}`;


    calendar.createEvent(

      newTitle,

      subStart,
      subEnd,

      {
        description:
`[AUTO_SPLIT_CREATED]
${rowMeta}
${descText}
--------------------
`
      }

    );


    return;
  }


  throw new Error(
    "無法計算時段切割，請確認起訖時間是否正確"
  );
}


// ============================================================
// 還原 Calendar
// ============================================================

function revertEvents(

  calendar,

  rangeStart,
  rangeEnd,

  origWorker,

  rowToken

) {


  const searchStart =
    new Date(

      rangeStart.getTime() -

      24 *
      60 *
      60 *
      1000

    );


  const searchEnd =
    new Date(

      rangeEnd.getTime() +

      24 *
      60 *
      60 *
      1000

    );


  const events =
    calendar.getEvents(
      searchStart,
      searchEnd
    );


  const rowTokenTag =
    `[ROW_ID:${rowToken}]`;


  events.forEach(
    evt => {


      const desc =
        evt.getDescription() ||
        "";


      if (
        !desc.includes(
          rowTokenTag
        )
      ) {

        return;
      }


      if (
        desc.includes(
          "[AUTO_SPLIT_CREATED]"
        )
      ) {

        evt.deleteEvent();

        return;
      }


      const splitMatch =
        desc.match(
          /\[SPLIT_ORIG_TIME:(\d+)-(\d+)\]/
        );


      if (splitMatch) {


        const origS =
          new Date(
            parseInt(
              splitMatch[1]
            )
          );


        const origE =
          new Date(
            parseInt(
              splitMatch[2]
            )
          );


        evt.setTime(
          origS,
          origE
        );


        evt.setDescription(

          desc.replace(
            /\[SPLIT_ORIG_TIME:\d+-\d+\]\n?/g,
            ""
          )

        );
      }


      const title =
        evt.getTitle();


      if (

        title.includes(
          "[代班]"
        ) ||

        title.includes(
          "[換班]"
        ) ||

        title.includes(
          "【請假】"
        )

      ) {


        const floor =
          extractFloorSuffix(
            title
          );


        evt.setTitle(
          `${origWorker}${floor}`
        );


        evt.setDescription(

          cleanDescription(
            desc,
            rowToken
          )

        );
      }
    }
  );
}


// ============================================================
// 擷取樓層
// ============================================================

function extractFloorSuffix(
  title
) {

  const pureTitle =
    title

      .replace(
        /\s*\[(換班|代班)\]/g,
        ""
      )

      .replace(
        /【請假】/g,
        ""
      );


  const match =
    pureTitle.match(
      /(\([^\)]+\)|（[^）]+）)$/
    );


  return match
    ? match[0]
    : "";
}


// ============================================================
// 清除系統中繼資料
// ============================================================

function cleanDescription(
  desc,
  rowToken
) {

  if (!desc) {
    return "";
  }


  let res =
    desc

      .replaceAll(
        `[ROW_ID:${rowToken}]\n`,
        ""
      )

      .replaceAll(
        `[ROW_ID:${rowToken}]`,
        ""
      );


  res =
    res.replace(

      /【(換班|代班|請假)紀錄】[\s\S]*?--------------------\n?/g,

      ""

    ).trim();


  return res;
}


// ============================================================
// 日期＋時間
// ============================================================

function combineDateTimeByStr(
  dateVal,
  timeStr
) {


  const d =
    new Date(
      dateVal
    );


  let hours = 0;
  let minutes = 0;


  const str =
    (timeStr || "")
      .toString()
      .trim();


  const match =
    str.match(
      /(\d{1,2}):(\d{2})/
    );


  if (match) {


    hours =
      parseInt(
        match[1],
        10
      );


    minutes =
      parseInt(
        match[2],
        10
      );


    const isPM =
      str.includes(
        "下午"
      ) ||

      str
        .toUpperCase()
        .includes("PM");


    const isAM =
      str.includes(
        "上午"
      ) ||

      str
        .toUpperCase()
        .includes("AM");


    if (isPM) {


      if (
        hours < 12
      ) {

        hours += 12;
      }


    } else if (isAM) {


      if (
        hours === 12
      ) {

        hours = 0;
      }
    }
  }


  d.setHours(
    hours,
    minutes,
    0,
    0
  );


  return d;
}


// ============================================================
// 取得申請人 Email
// ============================================================

function getApplicantEmail(
  sheet,
  row
) {


  if (
    !CONFIG.APPLICANT_EMAIL_COL
  ) {

    console.error(
      "尚未設定 CONFIG.APPLICANT_EMAIL_COL"
    );

    return "";
  }


  const email =
    sheet.getRange(
      row,
      CONFIG.APPLICANT_EMAIL_COL
    )
      .getValue()
      .toString()
      .trim();


  if (!email) {
    return "";
  }


  const emailPattern =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;


  if (
    !emailPattern.test(
      email
    )
  ) {

    console.error(
      `第 ${row} 列 K欄內容並非有效 Email：${email}`
    );

    return "";
  }


  return email;
}


// ============================================================
// 依姓名找員工 Email
// ============================================================

function getStaffEmailByName(
  name
) {


  if (!name) {
    return "";
  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const dirSheet =
    ss.getSheetByName(
      CONFIG.STAFF_DIRECTORY_SHEET_NAME
    );


  if (!dirSheet) {

    console.error(
      `找不到員工名冊分頁「${CONFIG.STAFF_DIRECTORY_SHEET_NAME}」，請確認分頁名稱是否一致`
    );

    return "";
  }


  const lastRow =
    dirSheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return "";
  }


  const numCols =
    Math.max(
      CONFIG.STAFF_NAME_COL,
      CONFIG.STAFF_EMAIL_COL
    );


  const data =
    dirSheet.getRange(

      2,
      1,

      lastRow - 1,

      numCols

    ).getValues();


  const target =
    name
      .toString()
      .trim();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {


    const rowName =
      (
        data[i][
          CONFIG.STAFF_NAME_COL -
          1
        ] ||
        ""
      )
        .toString()
        .trim();


    if (
      rowName === target
    ) {

      return (

        data[i][
          CONFIG.STAFF_EMAIL_COL -
          1
        ] ||

        ""

      )
        .toString()
        .trim();
    }
  }


  console.error(
    `員工名冊中找不到姓名「${target}」對應的 Email`
  );


  return "";
}


// ============================================================
// 退件
// ============================================================

function denyRow(

  sheet,
  row,

  checkCell,
  statusCell,

  message

) {


  statusCell.setValue(
    message
  );


  checkCell.setValue(
    false
  );


  notifyRejectionIfNeeded(
    sheet,
    row,
    message
  );
}


// ============================================================
// 退件 Email
// ============================================================

function notifyRejectionIfNeeded(
  sheet,
  row,
  message
) {


  const logCell =
    sheet.getRange(
      row,
      CONFIG.REJECT_LOG_COL
    );


  const prevLog =
    logCell.getValue()
      .toString();


  if (
    prevLog === message
  ) {

    return;
  }


  try {


    const email =
      getApplicantEmail(
        sheet,
        row
      );


    if (!email) {


      logCell.setValue(

        message +
        "（找不到有效 Email，未寄信）"

      );


      return;
    }


    const origPerson =
      sheet.getRange(row, 2)
        .getValue()
        .toString()
        .trim();


    const adminBcc =
      (
        CONFIG.ADMIN_EMAILS &&

        CONFIG.ADMIN_EMAILS.length >
        0
      )

        ? CONFIG.ADMIN_EMAILS.join(",")

        : undefined;


    MailApp.sendEmail({

      to:
        email,

      bcc:
        adminBcc,

      subject:
        "您的換班申請未通過（系統自動退件）",

      body:
`${origPerson} 您好，

您所提出的換班/代班申請經系統檢核後無法通過，原因如下：
${message}

請重新確認排班內容後再次提出申請，如有疑問請洽管理員。

（此為系統自動發送信件，請勿直接回覆）`

    });


    logCell.setValue(
      message
    );


  } catch (err) {


    logCell.setValue(

      "退件通知寄送失敗: " +
      err.message

    );
  }
}


// ============================================================
// N 欄批次處理
// ============================================================

function handleApprovalEditRange(
  sheet,
  range
) {


  const startRow =
    Math.max(
      2,
      range.getRow()
    );


  const endRow =
    range.getLastRow();


  if (
    startRow > endRow
  ) {

    return;
  }


  const lock =
    LockService
      .getScriptLock();


  if (
    !lock.tryLock(15000)
  ) {

    return;
  }


  try {


    for (
      let r = startRow;
      r <= endRow;
      r++
    ) {


      try {


        handleApprovalEdit(
          sheet,
          r
        );


      } catch (err) {


        console.error(
          `第 ${r} 列審核通知處理失敗: ${err.message}`
        );


        sheet.getRange(
          r,
          CONFIG.APPROVE_LOG_COL
        ).setValue(

          "處理失敗: " +
          err.message

        );
      }
    }


  } finally {


    lock.releaseLock();

  }
}


// ============================================================
// N 欄 → 核准通知
// ============================================================

function handleApprovalEdit(
  sheet,
  row
) {


  const approveCheck =
    sheet.getRange(
      row,
      CONFIG.APPROVE_CHECK_COL
    );


  const approveLog =
    sheet.getRange(
      row,
      CONFIG.APPROVE_LOG_COL
    );


  const statusVal =
    sheet.getRange(
      row,
      CONFIG.STATUS_COL
    )
      .getValue()
      .toString();


  const isChecked =
    approveCheck.getValue() ===
    true;


  if (!isChecked) {

    return;
  }


  const prevLog =
    approveLog.getValue()
      .toString();


  if (
    prevLog.startsWith(
      "已寄送核准通知"
    )
  ) {

    return;
  }


  if (
    !statusVal.startsWith(
      "已更新日曆"
    )
  ) {


    approveCheck.setValue(
      false
    );


    approveLog.setValue(

      "尚未成功執行換班（請先勾選L欄，並確認M欄狀態已顯示「已更新日曆」），無法寄送核准通知"

    );


    return;
  }


  try {


    sendApprovalEmail(
      sheet,
      row
    );


    approveLog.setValue(

      `已寄送核准通知 ${

        Utilities.formatDate(

          new Date(),

          Session.getScriptTimeZone(),

          "yyyy/MM/dd HH:mm"

        )

      }`

    );


  } catch (err) {


    approveCheck.setValue(
      false
    );


    approveLog.setValue(

      "寄信失敗: " +
      err.message

    );
  }
}


// ============================================================
// 寄送核准信
// ============================================================

function sendApprovalEmail(
  sheet,
  row
) {


  const email =
    getApplicantEmail(
      sheet,
      row
    );


  if (!email) {

    throw new Error(
      "找不到有效的申請人 Email"
    );
  }


  const origPerson =
    sheet.getRange(row, 2)
      .getValue()
      .toString()
      .trim();


  const origDate =
    sheet.getRange(row, 3)
      .getValue();


  const origStartStr =
    sheet.getRange(row, 4)
      .getDisplayValue()
      .toString()
      .trim();


  const origEndStr =
    sheet.getRange(row, 5)
      .getDisplayValue()
      .toString()
      .trim();


  const targetPerson =
    sheet.getRange(row, 6)
      .getValue()
      .toString()
      .trim();


  const swapDate =
    sheet.getRange(row, 7)
      .getValue();


  const swapStartStr =
    sheet.getRange(row, 8)
      .getDisplayValue()
      .toString()
      .trim();


  const swapEndStr =
    sheet.getRange(row, 9)
      .getDisplayValue()
      .toString()
      .trim();


  const dateStr =
    origDate instanceof Date

      ? Utilities.formatDate(

          origDate,

          Session.getScriptTimeZone(),

          "yyyy/MM/dd"

        )

      : origDate;


  const adminBcc =
    (
      CONFIG.ADMIN_EMAILS &&

      CONFIG.ADMIN_EMAILS.length >
      0
    )

      ? CONFIG.ADMIN_EMAILS.join(",")

      : undefined;


  const isSwap =
    Boolean(

      swapDate &&
      swapStartStr &&
      swapEndStr &&

      targetPerson &&

      targetPerson !==
        "請假" &&

      targetPerson !==
        "無"

    );


  const isSub =
    Boolean(

      !isSwap &&

      targetPerson &&

      targetPerson !==
        "請假" &&

      targetPerson !==
        "無"

    );


  let swapNote =
    "";


  if (isSwap) {


    swapNote =
      `（換班對象：${targetPerson}）`;


  } else if (isSub) {


    swapNote =
      `（代班人：${targetPerson}）`;

  }


  // ==========================================================
  // 寄給申請人
  // ==========================================================

  MailApp.sendEmail({

    to:
      email,

    bcc:
      adminBcc,

    subject:
      "【值班換班申請】已審核通過",

    body:
`${origPerson} 您好，

您於 ${dateStr} ${origStartStr}-${origEndStr} 提出的換班/代班申請${swapNote}已審核通過，並已完成排班異動，新的值班安排已同步至值班日曆。

請留意您的值班時間，如有任何問題歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`

  });


  // ==========================================================
  // 雙向換班 → 通知配合人
  // ==========================================================

  if (isSwap) {


    const targetEmail =
      getStaffEmailByName(
        targetPerson
      );


    if (targetEmail) {


      const swapDateStr =
        swapDate instanceof Date

          ? Utilities.formatDate(

              swapDate,

              Session.getScriptTimeZone(),

              "yyyy/MM/dd"

            )

          : swapDate;


      MailApp.sendEmail({

        to:
          targetEmail,

        bcc:
          adminBcc,

        subject:
          "【值班換班申請】您的班表已完成互換",

        body:
`${targetPerson} 您好，

您與 ${origPerson} 的換班申請已審核通過：
- 原定班表：${swapDateStr} ${swapStartStr}-${swapEndStr}
- 互換至：${dateStr} ${origStartStr}-${origEndStr}

新的值班安排已同步至值班日曆，請留意您的值班時間。如有任何問題歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`

      });
    }


  } else if (isSub) {


    // ========================================================
    // 單向代班 → 通知代班人
    // ========================================================

    const targetEmail =
      getStaffEmailByName(
        targetPerson
      );


    if (targetEmail) {


      MailApp.sendEmail({

        to:
          targetEmail,

        bcc:
          adminBcc,

        subject:
          "【值班代班通知】代班已審核通過並排入班表",

        body:
`${targetPerson} 您好，

您協助 ${origPerson} 代班的申請已審核通過：
- 代班時段：${dateStr} ${origStartStr}-${origEndStr}

該時段班表已同步至值班日曆，請留意準時到勤。如有任何問題歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`

      });
    }
  }
}
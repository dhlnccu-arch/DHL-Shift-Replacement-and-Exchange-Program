/**
 * 助理代班／換班系統 v6.2.1（後補代班版）
 *
 * A–J：Google 表單資料
 * K(11)：申請人 Email
 * L(12)：審核確認／執行班表異動
 * M(13)：執行狀態
 * N(14)：審核通知
 * O(15)：核准通知記錄
 * P(16)：退件通知記錄
 * Q(17)：申請 ID（UUID，可隱藏）
 *
 * 安裝式觸發器共 2 個：
 * 1. handleSheetEdit      → 來自試算表 → 編輯時
 * 2. onFormSubmitPrecheck → 來自試算表 → 表單提交時
 *
 * v6.2.1 重點：
 * - 表單送出先唯讀預檢；管理員勾 L 時重新讀最新 Calendar 再正式執行。
 * - 「請假（暫時找不到代班人員）」等以「請假」開頭的選項，一律視為純請假。
 * - Calendar 標題統一：姓名(樓層) [代班]／[換班]／[請假]
 * - 支援「先請假，後來找到代班人」：再次送一筆單向代班申請即可。
 * - 後補代班若取消 L，會精確恢復成原本的 [請假] 狀態，而不是正常班。
 * - 使用 UUID 作為 Calendar 事件控制識別碼，避免列排序／插刪造成串資料。
 */

const MAX_ROWS_PER_RUN = 30;
const TIME_BUDGET_MS = 5 * 60 * 1000;

const CONFIG = {
  EXECUTE_CHECK_COL: 12, // L
  STATUS_COL: 13,        // M
  APPROVE_CHECK_COL: 14, // N
  APPROVE_LOG_COL: 15,   // O
  REJECT_LOG_COL: 16,    // P
  REQUEST_ID_COL: 17,    // Q
  APPLICANT_EMAIL_COL: 11,
  ADMIN_EMAILS: ["dhl.nccu@gmail.com"],
  STAFF_DIRECTORY_SHEET_NAME: "員工名冊",
  STAFF_NAME_COL: 1,
  STAFF_EMAIL_COL: 2
};


// ============================================================
// 表單提交 → 前置預檢
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
// F 欄是否為真正的代班／換班人員
// ============================================================

function hasRealTargetPerson(value) {

  const v =
    (value || "")
      .toString()
      .trim();

  if (!v) return false;

  if (v === "無") return false;

  // 「請假」
  // 「請假（暫時找不到代班人員）」
  // 任何以「請假」開頭的選項，都不是人名
  if (v.startsWith("請假")) return false;

  return true;
}


// ============================================================
// 統一分析函式
// 預檢與正式審核都使用這一套規則
//
// 注意：
// 每次呼叫都重新讀取當下 Calendar，
// 不會沿用之前預檢的舊結果。
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


  // ==========================================================
  // 基本資料
  // ==========================================================

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


  // ==========================================================
  // 找原班
  //
  // 如果 Calendar 是：
  // 王小明(4F) [請假]
  //
  // extractWorkerName() 仍然會得到：
  // 王小明
  //
  // 所以後續找到代班人時，可以再次申請代班。
  // ==========================================================

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


  const origEvent =
    events1[0];


  const check1 =
    validateTimeRange(
      origEvent,
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


  // ==========================================================
  // 判斷申請類型
  // ==========================================================

  const hasTarget =
    hasRealTargetPerson(
      targetPerson
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


  // 有人名，但 G/H/I 只填一部分
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


  // 請假時不應有 G/H/I
  if (
    !hasTarget &&
    hasAnySwapField
  ) {

    return {

      ...base,

      ok: false,

      message:
        "請假申請不需填寫配合換班日期與時間，請將 G/H/I 留白"

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


  // 記錄原 Calendar 事件現在是正常班、請假、代班或換班
  const origStatus =
    getEventStatus(
      origEvent.getTitle()
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


    if (
      events2.length === 0
    ) {

      return {

        ...base,

        ok: false,

        message:
          `找不到 ${targetPerson} 的互換時段行程`

      };
    }


    if (
      events2.length > 1
    ) {

      return {

        ...base,

        ok: false,

        message:
          `${targetPerson} 互換時段有多筆行程，請聯絡管理員確認`

      };
    }


    const swapEvent =
      events2[0];


    const check2 =
      validateTimeRange(
        swapEvent,
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


    // 配合人要移到申請人的原班
    const conflictForTarget =
      findRealConflicts(

        eventsO,

        origStartTime,
        origEndTime,

        targetPerson,

        swapEvent,

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


    // 申請人要移到配合人的班
    const conflictForOrig =
      findRealConflicts(

        eventsS,

        swapStartTime,
        swapEndTime,

        origPerson,

        origEvent,

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

      origStatus,

      origStartTime,
      origEndTime,

      swapStartTime,
      swapEndTime,

      eventsO,
      eventsS,

      origEvent,
      swapEvent

    };
  }


  // ==========================================================
  // 單向代班
  //
  // 包括：
  // 1. 一開始就有代班人
  // 2. 先請假，之後才找到代班人
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


  return {

    ...base,

    ok: true,

    type,

    origStatus,

    origStartTime,
    origEndTime,

    eventsO,

    origEvent,

    swapEvent:
      null

  };
}


// ============================================================
// 試算表編輯監聽器
//
// 如果 L 與 N 同時被編輯：
// 一定先處理 L，再處理 N。
// ============================================================

function handleSheetEdit(e) {

  if (
    !e ||
    !e.range
  ) {

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

  if (
    touchesExecuteCol
  ) {

    const startRow =
      Math.max(
        2,
        range.getRow()
      );


    let endRow =
      range.getLastRow();


    if (
      startRow >
      endRow
    ) {

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

        realEndRow -
        endRow,

        1

      ).setValue(

        `尚未處理：單次批次上限為 ${MAX_ROWS_PER_RUN} 列，請稍後重新勾選此列（或分批操作）`

      );
    }


    const lock =
      LockService
        .getScriptLock();


    if (
      !lock.tryLock(
        15000
      )
    ) {

      sheet.getRange(

        startRow,

        CONFIG.STATUS_COL,

        endRow -
        startRow +
        1,

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
        CalendarApp
          .getCalendarsByName(
            "達賢館創新組助理值班"
          );


      if (
        calendars.length === 0
      ) {

        sheet.getRange(

          startRow,

          CONFIG.STATUS_COL,

          endRow -
          startRow +
          1,

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

            endRow -
            r +
            1,

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

      console.error(
        err
      );


      sheet.getRange(

        startRow,

        CONFIG.STATUS_COL,

        endRow -
        startRow +
        1,

        1

      ).setValue(

        "系統發生錯誤，請聯絡管理員: " +
        err.message

      );


    } finally {

      lock.releaseLock();

    }


    if (
      truncated
    ) {

      console.log(

        `本次觸發因批次列數或時間上限而部分列未處理（範圍 ${startRow}-${range.getLastRow()}）`

      );
    }
  }


  // ==========================================================
  // L 完成後再處理 N
  // ==========================================================

  if (
    touchesApproveCol
  ) {

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
    checkCell.getValue() ===
    true;


  const statusVal =
    statusCell
      .getValue()
      .toString();


  try {

    // ========================================================
    // L 被取消
    // → 還原「這一筆申請」造成的 Calendar 異動
    // ========================================================

    if (
      !isChecked
    ) {

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

          hasRealTargetPerson(
            targetPerson
          ) &&

          swapDate &&
          swapStartStr &&
          swapEndStr

        );


      if (
        isSwap
      ) {

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


    // 已經執行成功過
    if (
      statusVal.includes(
        "已更新日曆"
      )
    ) {

      return;
    }


    // ========================================================
    // 正式審核
    //
    // 每次勾 L 都重新讀最新 Calendar。
    // ========================================================

    const analysis =
      analyzeRequest(
        sheet,
        calendar,
        row
      );


    if (
      !analysis.ok
    ) {

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
      analysis.type ===
      "swap"
    ) {

      const origDateStr =
        formatDateValue(
          analysis.origDate
        );


      const swapDateStr =
        formatDateValue(
          analysis.swapDate
        );


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
    //
    // 如果 origStatus === leave：
    // 代表原事件已經是：
    // 王小明(4F) [請假]
    //
    // 這次就是「後補代班」。
    // ========================================================

    if (
      analysis.type ===
      "sub"
    ) {

      const isLaterSubstitute =
        analysis.origStatus ===
        "leave";


      const desc =
        isLaterSubstitute

          ? `【代班紀錄】
- 實際到勤：${analysis.targetPerson}
- 原定值班：${analysis.origPerson}
- 原狀態：已請假，後續找到代班人`

          : `【代班紀錄】
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

        isLaterSubstitute

          ? "已更新日曆（後補代班完成）"

          : "已更新日曆（代班完成）"

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

      "[請假]",

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


    if (
      isChecked
    ) {

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

  const matched =
    [];


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
      ) !==
      personName
    ) {

      continue;
    }


    const evStart =
      evt.getStartTime();


    const evEnd =
      evt.getEndTime();


    if (
      reqStart <
      evEnd &&

      evStart <
      reqEnd
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


  const conflicts =
    [];


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
      ) !==
      personName
    ) {

      continue;
    }


    const evStart =
      evt.getStartTime();


    const evEnd =
      evt.getEndTime();


    // 即將讓出的行程
    if (
      cedingId &&
      evt.getId() ===
      cedingId
    ) {

      // 前半段殘餘
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


      // 後半段殘餘
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
// Calendar 標題處理
// ============================================================

function extractWorkerName(
  title
) {

  let name =
    (title || "")
      .toString()
      .trim();


  // 王小明(4F) [代班]
  // 王小明(4F) [換班]
  // 王小明(4F) [請假]
  name =
    name.replace(
      /\s*\[(換班|代班|請假)\]\s*$/,
      ""
    );


  // 移除樓層
  name =
    name.replace(
      /\s*(\([^\)]*\)|（[^）]*）)\s*$/,
      ""
    );


  return name.trim();
}


function extractFloorSuffix(
  title
) {

  const pureTitle =
    (title || "")
      .toString()
      .replace(
        /\s*\[(換班|代班|請假)\]\s*$/g,
        ""
      )
      .trim();


  const match =
    pureTitle.match(
      /(\([^\)]+\)|（[^）]+）)$/
    );


  return match
    ? match[0]
    : "";
}


function getEventStatus(
  title
) {

  const t =
    (title || "")
      .toString()
      .trim();


  if (
    /\[請假\]\s*$/.test(
      t
    )
  ) {

    return "leave";
  }


  if (
    /\[代班\]\s*$/.test(
      t
    )
  ) {

    return "sub";
  }


  if (
    /\[換班\]\s*$/.test(
      t
    )
  ) {

    return "swap";
  }


  return "normal";
}


function buildShiftTitle(
  worker,
  floor,
  tag
) {

  return `${worker}${floor} ${tag}`;
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

    start <
    end

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


  if (
    !id
  ) {

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


  if (
    existing
  ) {

    return existing;
  }


  // 舊資料相容
  if (
    legacyForExisting
  ) {

    return String(
      row
    );
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


function formatDateValue(
  value
) {

  return value instanceof Date

    ? Utilities.formatDate(

        value,

        Session.getScriptTimeZone(),

        "yyyy/MM/dd"

      )

    : value;
}


// ============================================================
// 確認申請時段是否在原事件內
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

      valid:
        false,

      actualRange

    };
  }


  return {

    valid:
      true,

    actualRange

  };
}


// ============================================================
// 原始狀態封存
//
// 這是 v6.2 最重要的新增設計之一。
// 把修改前完整的 Calendar title / description 封存起來。
//
// 因此：
//
// 王小明(4F) [請假]
// ↓ 後補代班
// 李小華(4F) [代班]
//
// 如果取消後補代班：
//
// 李小華(4F) [代班]
// ↓
// 王小明(4F) [請假]
//
// 可以精確恢復。
// ============================================================

function encodeMetaText(
  text
) {

  return Utilities.base64EncodeWebSafe(

    (text || "")
      .toString(),

    Utilities.Charset.UTF_8

  );
}


function decodeMetaText(
  encoded
) {

  if (
    !encoded
  ) {

    return "";
  }


  return Utilities.newBlob(

    Utilities.base64DecodeWebSafe(
      encoded
    )

  ).getDataAsString(
    "UTF-8"
  );
}


// 移除前一層事件的系統控制標記。
// 真正原始內容會另外完整封存在 ORIG_DESC_B64。
function stripControlMetadata(
  desc
) {

  return (desc || "")

    .toString()

    .replace(
      /^\[AUTO_SPLIT_CREATED\]\n?/gm,
      ""
    )

    .replace(
      /^\[ROW_ID:[^\]]+\]\n?/gm,
      ""
    )

    .replace(
      /^\[ORIG_TITLE_B64:[^\]]*\]\n?/gm,
      ""
    )

    .replace(
      /^\[ORIG_DESC_B64:[^\]]*\]\n?/gm,
      ""
    )

    .replace(
      /^\[SPLIT_ORIG_TIME:\d+-\d+\]\n?/gm,
      ""
    )

    .trim();
}


function buildControlHeader(

  rowToken,

  originalTitle,
  originalDescription,

  evStart,
  evEnd,

  includeSplitTime

) {

  const parts =
    [

      `[ROW_ID:${rowToken}]`,

      `[ORIG_TITLE_B64:${encodeMetaText(originalTitle)}]`,

      `[ORIG_DESC_B64:${encodeMetaText(originalDescription)}]`

    ];


  if (
    includeSplitTime
  ) {

    parts.push(

      `[SPLIT_ORIG_TIME:${evStart.getTime()}-${evEnd.getTime()}]`

    );
  }


  return parts.join(
    "\n"
  );
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


  const originalTitle =
    mainEvent.getTitle();


  const originalDescription =
    mainEvent.getDescription() ||
    "";


  const inheritedDescription =
    stripControlMetadata(
      originalDescription
    );


  const floor =
    extractFloorSuffix(
      originalTitle
    );


  const rowMeta =
    `[ROW_ID:${rowToken}]`;


  const newTitle =
    buildShiftTitle(
      newWorker,
      floor,
      tag
    );


  // ==========================================================
  // 整段吻合
  // ==========================================================

  if (

    Math.abs(
      evStart.getTime() -
      subStart.getTime()
    ) <
    60000 &&

    Math.abs(
      evEnd.getTime() -
      subEnd.getTime()
    ) <
    60000

  ) {

    const controlHeader =
      buildControlHeader(

        rowToken,

        originalTitle,
        originalDescription,

        evStart,
        evEnd,

        false

      );


    mainEvent.setTitle(
      newTitle
    );


    mainEvent.setDescription(

      `${controlHeader}
${descText}
--------------------
${inheritedDescription}`.trim()

    );


    return;
  }


  const controlHeader =
    buildControlHeader(

      rowToken,

      originalTitle,
      originalDescription,

      evStart,
      evEnd,

      true

    );


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

      `${controlHeader}
${inheritedDescription}`.trim()

    );


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
${inheritedDescription}`.trim()

      }

    );


    // 後半段維持原本狀態。
    //
    // 如果原本是：
    // 王小明(4F) [請假]
    //
    // 後半段就仍然是：
    // 王小明(4F) [請假]
    //
    // 不會變回正常班。
    calendar.createEvent(

      originalTitle,

      subEnd,
      evEnd,

      {

        description:

`[AUTO_SPLIT_CREATED]
${rowMeta}
${inheritedDescription}`.trim()

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
    ) <
    60000

  ) {

    mainEvent.setTime(
      evStart,
      subStart
    );


    mainEvent.setDescription(

      `${controlHeader}
${inheritedDescription}`.trim()

    );


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
${inheritedDescription}`.trim()

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
    ) <
    60000 &&

    subEnd.getTime() <
    evEnd.getTime()

  ) {

    mainEvent.setTime(
      subEnd,
      evEnd
    );


    mainEvent.setDescription(

      `${controlHeader}
${inheritedDescription}`.trim()

    );


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
${inheritedDescription}`.trim()

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
//
// 新版優先使用：
//
// ORIG_TITLE_B64
// ORIG_DESC_B64
//
// 精確恢復修改前的狀態。
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


  // ==========================================================
  // 先刪除這次切割產生的新事件
  // ==========================================================

  for (
    let i = events.length - 1;
    i >= 0;
    i--
  ) {

    const evt =
      events[i];


    const desc =
      evt.getDescription() ||
      "";


    if (
      !desc.includes(
        rowTokenTag
      )
    ) {

      continue;
    }


    if (
      desc.includes(
        "[AUTO_SPLIT_CREATED]"
      )
    ) {

      evt.deleteEvent();

    }
  }


  // ==========================================================
  // 再恢復原事件
  // ==========================================================

  const remainingEvents =
    calendar.getEvents(
      searchStart,
      searchEnd
    );


  remainingEvents.forEach(
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


      const titleMatch =
        desc.match(
          /\[ORIG_TITLE_B64:([^\]]*)\]/
        );


      const descMatch =
        desc.match(
          /\[ORIG_DESC_B64:([^\]]*)\]/
        );


      const splitMatch =
        desc.match(
          /\[SPLIT_ORIG_TIME:(\d+)-(\d+)\]/
        );


      // 還原原始完整時段
      if (
        splitMatch
      ) {

        evt.setTime(

          new Date(
            parseInt(
              splitMatch[1],
              10
            )
          ),

          new Date(
            parseInt(
              splitMatch[2],
              10
            )
          )

        );
      }


      // ======================================================
      // 新版：精確恢復修改前 title
      // ======================================================

      if (
        titleMatch
      ) {

        evt.setTitle(

          decodeMetaText(
            titleMatch[1]
          )

        );


      } else {

        // 舊 v6 fallback

        const floor =
          extractFloorSuffix(
            evt.getTitle()
          );


        evt.setTitle(
          `${origWorker}${floor}`
        );

      }


      // ======================================================
      // 新版：精確恢復修改前 description
      // ======================================================

      if (
        descMatch
      ) {

        evt.setDescription(

          decodeMetaText(
            descMatch[1]
          )

        );


      } else {

        // 舊 v6 fallback

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
// 舊 v6 fallback
// ============================================================

function cleanDescription(
  desc,
  rowToken
) {

  if (
    !desc
  ) {

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
      )

      .replace(
        /^\[AUTO_SPLIT_CREATED\]\n?/gm,
        ""
      )

      .replace(
        /^\[SPLIT_ORIG_TIME:\d+-\d+\]\n?/gm,
        ""
      )

      .replace(
        /^\[ORIG_TITLE_B64:[^\]]*\]\n?/gm,
        ""
      )

      .replace(
        /^\[ORIG_DESC_B64:[^\]]*\]\n?/gm,
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


  let hours =
    0;


  let minutes =
    0;


  const str =
    (timeStr || "")
      .toString()
      .trim();


  const match =
    str.match(
      /(\d{1,2}):(\d{2})/
    );


  if (
    match
  ) {

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
        .includes(
          "PM"
        );


    const isAM =
      str.includes(
        "上午"
      ) ||

      str
        .toUpperCase()
        .includes(
          "AM"
        );


    if (
      isPM
    ) {

      if (
        hours < 12
      ) {

        hours +=
          12;
      }


    } else if (
      isAM
    ) {

      if (
        hours === 12
      ) {

        hours =
          0;
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

  const email =
    sheet.getRange(
      row,
      CONFIG.APPLICANT_EMAIL_COL
    )
      .getValue()
      .toString()
      .trim();


  if (
    !email
  ) {

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
// 員工姓名 → Email
// ============================================================

function getStaffEmailByName(
  name
) {

  if (
    !name
  ) {

    return "";
  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const dirSheet =
    ss.getSheetByName(
      CONFIG.STAFF_DIRECTORY_SHEET_NAME
    );


  if (
    !dirSheet
  ) {

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
      rowName ===
      target
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
    logCell
      .getValue()
      .toString();


  if (
    prevLog ===
    message
  ) {

    return;
  }


  try {

    const email =
      getApplicantEmail(
        sheet,
        row
      );


    if (
      !email
    ) {

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

        ? CONFIG.ADMIN_EMAILS.join(
            ","
          )

        : undefined;


    MailApp.sendEmail({

      to:
        email,

      bcc:
        adminBcc,

      subject:
        "您的換班／代班／請假申請未通過（系統自動退件）",

      body:

`${origPerson} 您好，

您所提出的值班異動申請經系統檢核後無法通過，原因如下：
${message}

請重新確認排班內容後再次提出申請，如有疑問請洽管理員。

（此為系統自動發送信件，請勿直接回覆）`

    });


    logCell.setValue(
      message
    );


  } catch (
    err
  ) {

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
    startRow >
    endRow
  ) {

    return;
  }


  const lock =
    LockService
      .getScriptLock();


  if (
    !lock.tryLock(
      15000
    )
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


      } catch (
        err
      ) {

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


  if (
    !isChecked
  ) {

    return;
  }


  const prevLog =
    approveLog
      .getValue()
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

      "尚未成功執行班表異動（請先勾選L欄，並確認M欄狀態已顯示「已更新日曆」），無法寄送核准通知"

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


  } catch (
    err
  ) {

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


  if (
    !email
  ) {

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


  const statusVal =
    sheet.getRange(
      row,
      CONFIG.STATUS_COL
    )
      .getValue()
      .toString();


  const dateStr =
    formatDateValue(
      origDate
    );


  const adminBcc =
    (
      CONFIG.ADMIN_EMAILS &&

      CONFIG.ADMIN_EMAILS.length >
      0
    )

      ? CONFIG.ADMIN_EMAILS.join(
          ","
        )

      : undefined;


  const hasTarget =
    hasRealTargetPerson(
      targetPerson
    );


  const isSwap =
    Boolean(

      hasTarget &&

      swapDate &&
      swapStartStr &&
      swapEndStr

    );


  const isSub =
    Boolean(

      hasTarget &&
      !isSwap

    );


  const isLaterSubstitute =
    statusVal.includes(
      "後補代班"
    );


  let applicantSubject =
    "【值班申請】已審核通過";


  let applicantDetail =
    "";


  if (
    isSwap
  ) {

    applicantSubject =
      "【值班換班申請】已審核通過";


    applicantDetail =
      `（換班對象：${targetPerson}）`;


  } else if (
    isSub
  ) {

    applicantSubject =
      isLaterSubstitute

        ? "【值班後補代班】已審核通過"

        : "【值班代班申請】已審核通過";


    applicantDetail =
      `（代班人：${targetPerson}）`;


  } else {

    applicantSubject =
      "【值班請假申請】已審核通過";


    applicantDetail =
      "";
  }


  // ==========================================================
  // 申請人
  // ==========================================================

  MailApp.sendEmail({

    to:
      email,

    bcc:
      adminBcc,

    subject:
      applicantSubject,

    body:

`${origPerson} 您好，

您於 ${dateStr} ${origStartStr}-${origEndStr} 提出的值班申請${applicantDetail}已審核通過，並已完成排班異動，新的值班安排已同步至值班日曆。

請留意您的值班安排，如有任何問題歡迎與管理員聯繫。

（此為系統自動發送信件，請勿直接回覆）`

  });


  // ==========================================================
  // 雙向換班 → 通知配合人
  // ==========================================================

  if (
    isSwap
  ) {

    const targetEmail =
      getStaffEmailByName(
        targetPerson
      );


    if (
      targetEmail
    ) {

      const swapDateStr =
        formatDateValue(
          swapDate
        );


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


  } else if (
    isSub
  ) {

    // ========================================================
    // 單向代班／後補代班 → 通知代班人
    // ========================================================

    const targetEmail =
      getStaffEmailByName(
        targetPerson
      );


    if (
      targetEmail
    ) {

      MailApp.sendEmail({

        to:
          targetEmail,

        bcc:
          adminBcc,

        subject:

          isLaterSubstitute

            ? "【值班後補代班通知】代班已審核通過並排入班表"

            : "【值班代班通知】代班已審核通過並排入班表",

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
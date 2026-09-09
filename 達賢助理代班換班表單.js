/**
 * ============================================================
 * 達賢助理代班／換班系統 v6.5.1
 * ============================================================
 *
 * 【用途】
 * 處理達賢助理的：
 * - 請假
 * - 單向代班
 * - 後補代班
 * - 雙向換班
 * - 同一申請涵蓋同一人多段班表
 *
 *
 * ============================================================
 * 一、Google Sheet 欄位
 * ============================================================
 *
 * A–J：Google Form 原始回覆
 *
 * K（11）申請人 Email
 *
 * L（12）審核確認／執行班表異動
 *   管理員勾選後：
 *   1. 重新讀取最新 Calendar
 *   2. 再次完整檢核
 *   3. 成功後才真正修改 Calendar
 *
 * M（13）執行狀態
 *
 * N（14）審核通知
 *   班表異動成功後，管理員勾選此欄寄出核准通知。
 *
 * O（15）核准通知紀錄
 *   儲存格內容：人類可讀狀態
 *   儲存格註解：逐收件人的寄信狀態
 *
 * P（16）退件通知紀錄
 *   儲存格內容：人類可讀退件紀錄
 *   儲存格註解：逐收件人的寄信狀態
 *
 * Q（17）申請 ID
 *   儲存 UUID。
 *   Q 儲存格註解另保存 Calendar 交易快照。
 *
 * 請勿刪除 Q、O、P 的內容或註解。
 * 請勿只排序部分欄位。
 *
 *
 * ============================================================
 * 二、回覆分頁
 * ============================================================
 *
 * 固定使用：
 *
 * 表單回覆 1
 *
 * 若有人將分頁改名，觸發器會略過該分頁，
 * 並在 Apps Script 執行紀錄留下 console.warn。
 *
 *
 * ============================================================
 * 三、安裝式觸發器
 * ============================================================
 *
 * 必須保留兩個：
 *
 * 1. handleSheetEdit
 *    → 事件來源：試算表
 *    → 事件類型：編輯時
 *
 * 2. onFormSubmitPrecheck
 *    → 事件來源：試算表
 *    → 事件類型：提交表單時
 *
 *
 * ============================================================
 * 四、作業流程
 * ============================================================
 *
 * 表單提交
 *   ↓
 * onFormSubmitPrecheck
 *   ↓
 * 唯讀預檢 Calendar
 *   ├─ 不通過 → M 寫退件原因 + P 記錄 + 寄退件信
 *   └─ 通過   → M 顯示「預檢通過，待管理員審核」
 *
 * 管理員勾 L
 *   ↓
 * 重新讀取「當下最新」Calendar
 *   ↓
 * analyzeRequest()
 *   ├─ 不通過 → 正式退件
 *   └─ 通過   → 修改 Calendar
 *
 * Calendar 修改成功
 *   ↓
 * M 顯示「已更新日曆」
 *
 * 管理員勾 N
 *   ↓
 * 寄核准通知
 *   ↓
 * O 記錄每一位收件人的寄送結果
 *
 *
 * ============================================================
 * 五、交易狀態機
 * ============================================================
 *
 * 每一筆真正修改 Calendar 的申請，
 * Q 欄註解會保存一份 shift-tx 交易紀錄。
 *
 *
 * APPLYING
 *   正在執行 Calendar 異動。
 *
 *      │成功
 *      ▼
 *
 * APPLIED
 *   Calendar 異動已完成。
 *   此時 L 應為勾選。
 *
 *      │管理員取消 L
 *      ▼
 *
 * UNDOING
 *   正在還原 Calendar。
 *
 *      │成功
 *      ▼
 *
 * UNDONE
 *   已成功恢復異動前狀態。
 *   原申請保留歷史紀錄，不可再次重用。
 *
 *
 * 若 APPLYING 或 UNDOING 中途發生無法確認的錯誤：
 *
 * APPLYING / UNDOING
 *      ↓
 * RECOVER
 *      ↓
 * 嘗試依 Q 欄快照恢復
 *      ↓
 * ROLLED_BACK
 *
 *
 * RECOVER：
 *   代表目前 Calendar 狀態不能百分之百確認，
 *   系統會停止其他正式異動，避免錯誤擴散。
 *
 * ROLLED_BACK：
 *   本次失敗異動已恢復到執行前狀態。
 *
 *
 * ============================================================
 * 六、多段班
 * ============================================================
 *
 * 例如：
 *
 * 子平 10:00–12:00
 * 子平 13:00–18:00
 *
 * 表單可以填：
 *
 * 10:00–18:00
 *
 * 系統會辨識實際班段為：
 *
 * 10:00–12:00
 * 13:00–18:00
 *
 * 12:00–13:00 原本沒有班，不會自行新增班表。
 *
 *
 * ============================================================
 * 七、Calendar Tag
 * ============================================================
 *
 * 系統控制碼不顯示在人眼看到的 Calendar 說明欄。
 *
 * 使用 CalendarEvent Tag 保存：
 *
 * DHL_ROW_ID
 * DHL_PART_ID
 * DHL_META_VERSION
 *
 * 完整復原資料仍以 Q 欄註解為主。
 *
 *
 * ============================================================
 * 八、後補代班
 * ============================================================
 *
 * 第一次：
 *
 * 王小明(4F) [請假]
 *
 * 後來找到李小華：
 *
 * 重新送一筆「單向代班」申請。
 *
 * 核准後：
 *
 * 李小華(4F) [代班]
 *
 * 若取消後補代班：
 *
 * 系統恢復：
 *
 * 王小明(4F) [請假]
 *
 * 若仍有後補代班／後續異動存在，
 * 不允許先取消原始請假。
 *
 *
 * ============================================================
 * 九、寄信
 * ============================================================
 *
 * 每一位收件人都有自己的狀態：
 *
 * READY
 * SENDING
 * SENT
 * FAILED
 * UNKNOWN
 *
 * 已確認 SENT 的收件人不會因另一封信失敗而重寄。
 *
 *
 * ============================================================
 * 十、安全原則
 * ============================================================
 *
 * - 系統錯誤 ≠ 申請退件。
 * - Calendar API 結果不明時，不假裝成功。
 * - Calendar API 結果不明時，也不假裝已還原。
 * - 郵件結果不明時，不盲目重新寄送。
 * - 快照過大時停止執行，不另外導入 PropertiesService。
 *
 * ============================================================
 */


const CONFIG = {
  VERSION: "6.5.1",

  CALENDAR_NAME: "達賢館創新組助理值班",

  // 有同名日曆時才需要填入正確 Calendar ID。
  CALENDAR_ID: "",

  // 已確認的 Google Form 回覆分頁名稱。
  RESPONSE_SHEET_NAME: "表單回覆 1",

  TIME_ZONE: "Asia/Taipei",

  APPLICANT_EMAIL_COL: 11,
  EXECUTE_CHECK_COL: 12,
  STATUS_COL: 13,
  APPROVE_CHECK_COL: 14,
  APPROVE_LOG_COL: 15,
  REJECT_LOG_COL: 16,
  REQUEST_ID_COL: 17,

  ADMIN_EMAILS: [
    "dhl.nccu@gmail.com"
  ],

  STAFF_DIRECTORY_SHEET_NAME: "員工名冊",
  STAFF_NAME_COL: 1,
  STAFF_EMAIL_COL: 2,

  MAX_ROWS: 30,
  RUN_MS: 240000,

  // 留安全餘裕，不逼近 Sheets 註解容量上限。
  NOTE_LIMIT: 45000
};


const OPEN_PHASES = [
  "APPLYING",
  "UNDOING",
  "RECOVER"
];


const ACTIVE_PHASE = "APPLIED";


const EVENT_TAGS = {
  ROW_ID: "DHL_ROW_ID",
  PART_ID: "DHL_PART_ID",
  VERSION: "DHL_META_VERSION"
};


// ============================================================
// 管理員選單
// ============================================================

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("班表工具")

    .addItem(
      "重試此列未完成通知",
      "retrySelectedNotifications"
    )

    .addItem(
      "核對此列不明寄信結果",
      "resolveSelectedMail"
    )

    .addSeparator()

    .addItem(
      "復原此列中斷的班表異動",
      "recoverSelectedCalendar"
    )

    .addItem(
      "人工核對原班後解除此列鎖定",
      "confirmSelectedRecovery"
    )

    .addSeparator()

    .addItem(
      "清理所選申請的日曆顯示",
      "cleanSelectedCalendarDisplay"
    )

    .addToUi();
}


// ============================================================
// 判斷是否為表單回覆分頁
// ============================================================

function isResponseSheet_(sheet) {

  if (CONFIG.RESPONSE_SHEET_NAME) {

    return (
      sheet.getName() ===
      CONFIG.RESPONSE_SHEET_NAME
    );
  }


  // 以下僅作備援。
  if (
    sheet.getMaxColumns() < 17
  ) {

    return false;
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        17
      )
      .getDisplayValues()[0];


  return (

    String(
      headers[10]
    ).trim() ===
      "電子郵件地址"

    &&

    String(
      headers[11]
    ).includes(
      "審核確認"
    )

    &&

    String(
      headers[13]
    ).includes(
      "審核通知"
    )

  );
}


// ============================================================
// CONFIG 檢查
// ============================================================

function validateConfig_(sheet) {

  const cols = [

    CONFIG.APPLICANT_EMAIL_COL,
    CONFIG.EXECUTE_CHECK_COL,
    CONFIG.STATUS_COL,
    CONFIG.APPROVE_CHECK_COL,
    CONFIG.APPROVE_LOG_COL,
    CONFIG.REJECT_LOG_COL,
    CONFIG.REQUEST_ID_COL

  ];


  const invalid =
    cols.some(
      col =>
        !Number.isInteger(col) ||
        col < 1 ||
        col > sheet.getMaxColumns()
    );


  if (
    invalid ||
    new Set(cols).size !==
      cols.length
  ) {

    throw new Error(
      "CONFIG 欄位設定缺漏、重複或超出範圍，已停止。"
    );
  }


  CONFIG.ADMIN_EMAILS
    .forEach(
      email => {

        if (
          !validEmail_(email)
        ) {

          throw new Error(
            "管理員 Email 格式不正確。"
          );
        }
      }
    );
}


// ============================================================
// Script Lock
// ============================================================

function locked_(work) {

  const lock =
    LockService
      .getScriptLock();


  if (
    !lock.tryLock(
      20000
    )
  ) {

    throw new Error(
      "系統忙碌，本次未處理；稍後重新勾選或重試。"
    );
  }


  try {

    return work();

  } finally {

    try {

      SpreadsheetApp.flush();

    } finally {

      lock.releaseLock();
    }
  }
}


// ============================================================
// 取得 Calendar
// ============================================================

function getCalendar_() {

  if (
    CONFIG.CALENDAR_ID
  ) {

    const calendar =
      CalendarApp
        .getCalendarById(
          CONFIG.CALENDAR_ID
        );


    if (!calendar) {

      throw new Error(
        "找不到設定的日曆或沒有存取權限。"
      );
    }


    return calendar;
  }


  const calendars =
    CalendarApp
      .getCalendarsByName(
        CONFIG.CALENDAR_NAME
      );


  if (
    calendars.length !== 1
  ) {

    throw new Error(
      "指定名稱的日曆不是唯一一個，請確認名稱或設定 CALENDAR_ID。"
    );
  }


  return calendars[0];
}


// ============================================================
// 表單提交 → 唯讀預檢
// ============================================================

function onFormSubmitPrecheck(e) {

  if (
    !e ||
    !e.range
  ) {

    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    !isResponseSheet_(sheet)
  ) {

    console.warn(
      `onFormSubmitPrecheck 略過工作表「${sheet.getName()}」：` +
      `不是設定的表單回覆分頁「${CONFIG.RESPONSE_SHEET_NAME}」。`
    );

    return;
  }


  if (
    e.range.getRow() < 2
  ) {

    return;
  }


  locked_(
    () => {

      validateConfig_(
        sheet
      );


      const c =
        context_(
          sheet,
          e.range.getRow()
        );


      const st =
        state_(c);


      // 延遲或重複觸發時，
      // 不覆蓋已正式執行的案件。
      if (

        st

        ||

        text_(
          cell_(
            c,
            CONFIG.STATUS_COL
          ).getValue()
        ).startsWith(
          "已更新日曆"
        )

      ) {

        return;
      }


      try {

        assertNoPending_(
          sheet,
          ""
        );


        const analysis =
          analyzeRequest(
            sheet,
            getCalendar_(),
            row_(c)
          );


        if (
          !analysis.ok
        ) {

          reject_(
            c,
            analysis,
            "預檢"
          );

        } else {

          cell_(
            c,
            CONFIG.EXECUTE_CHECK_COL
          ).setValue(
            false
          );


          status_(
            c,
            "預檢通過（" +
            now_() +
            "），待管理員審核"
          );
        }

      } catch (err) {

        status_(
          c,
          "預檢暫未完成，請管理員確認：" +
          err.message
        );
      }
    }
  );
}


// ============================================================
// 試算表編輯事件
//
// 同一次編輯若同時碰到 L 與 N：
// 逐列先處理 L，再處理 N。
// ============================================================

function handleSheetEdit(e) {

  if (
    !e ||
    !e.range
  ) {

    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    !isResponseSheet_(sheet)
  ) {

    console.warn(
      `handleSheetEdit 略過工作表「${sheet.getName()}」：` +
      `不是設定的表單回覆分頁「${CONFIG.RESPONSE_SHEET_NAME}」。`
    );

    return;
  }


  const lo =
    e.range.getColumn();


  const hi =
    e.range.getLastColumn();


  const touchesExecute = (

    lo <=
      CONFIG.EXECUTE_CHECK_COL

    &&

    hi >=
      CONFIG.EXECUTE_CHECK_COL

  );


  const touchesMail = (

    lo <=
      CONFIG.APPROVE_CHECK_COL

    &&

    hi >=
      CONFIG.APPROVE_CHECK_COL

  );


  if (
    !touchesExecute &&
    !touchesMail
  ) {

    return;
  }


  locked_(
    () => {

      validateConfig_(
        sheet
      );


      const first =
        Math.max(
          2,
          e.range.getRow()
        );


      const last =
        Math.min(
          e.range.getLastRow(),
          sheet.getLastRow()
        );


      const deadline =
        Date.now() +
        CONFIG.RUN_MS;


      let calendar =
        null;


      for (
        let r = first;
        r <= last;
        r++
      ) {

        if (

          r - first >=
            CONFIG.MAX_ROWS

          ||

          Date.now() >
            deadline - 30000

        ) {

          sheet
            .getParent()
            .toast(
              "部分列尚未處理，請分批重新操作；未覆寫既有狀態。"
            );

          break;
        }


        if (
          !text_(
            sheet
              .getRange(
                r,
                2
              )
              .getValue()
          )
        ) {

          continue;
        }


        let c;


        try {

          c =
            context_(
              sheet,
              r
            );


          // L 一定先執行。
          if (
            touchesExecute
          ) {

            calendar =
              calendar ||
              getCalendar_();


            processSingleRow_(
              c,
              calendar
            );
          }


          // L 完成後才處理 N。
          if (

            touchesMail

            &&

            cell_(
              c,
              CONFIG.APPROVE_CHECK_COL
            ).getValue() ===
              true

          ) {

            approval_(
              c,
              deadline
            );
          }

        } catch (err) {

          console.error(
            err
          );


          sheet
            .getRange(
              r,
              CONFIG.EXECUTE_CHECK_COL
            )
            .setNote(
              "本次未完成：" +
              err.message
            );


          sheet
            .getParent()
            .toast(
              "第 " +
              r +
              " 列：" +
              err.message
            );
        }
      }
    }
  );
}


// ============================================================
// 共用文字工具
// ============================================================

function text_(value) {

  return (
    value == null
      ? ""
      : String(value).trim()
  );
}


function clone_(value) {

  return JSON.parse(
    JSON.stringify(
      value
    )
  );
}


function now_() {

  return Utilities.formatDate(
    new Date(),
    CONFIG.TIME_ZONE,
    "yyyy/MM/dd HH:mm:ss"
  );
}


function hash_(value) {

  return Utilities.base64EncodeWebSafe(

    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(value)
    )

  );
}


// ============================================================
// 儲存格註解 JSON
// ============================================================

function readNote_(
  cell,
  kind
) {

  const note =
    cell.getNote();


  if (!note) {

    return null;
  }


  let value;


  try {

    value =
      JSON.parse(
        note
      );

  } catch (_) {

    throw new Error(
      "系統註解不是有效的 JSON，請勿覆寫或刪除原紀錄。"
    );
  }


  if (
    value.kind !== kind
  ) {

    throw new Error(
      "系統註解類型不符，已停止以保護紀錄。"
    );
  }


  return value;
}


function writeNote_(
  cell,
  value
) {

  const note =
    JSON.stringify(
      value
    );


  if (
    note.length >
    CONFIG.NOTE_LIMIT
  ) {

    throw new Error(
      "此筆案件涉及的班表區段、復原快照或通知紀錄過多，" +
      "系統紀錄已超過安全容量，因此尚未繼續執行。 " +
      "請聯絡管理員，將這筆異動拆成較小的申請後再處理。"
    );
  }


  cell.setNote(
    note
  );


  // 交易快照必須在下一個 Calendar API 動作前可靠寫入。
  SpreadsheetApp.flush();
}


// ============================================================
// 建立申請 Context
// ============================================================

function context_(
  sheet,
  r
) {

  const q =
    sheet.getRange(
      r,
      CONFIG.REQUEST_ID_COL
    );


  let id =
    text_(
      q.getValue()
    );


  const oldStatus =
    text_(
      sheet
        .getRange(
          r,
          CONFIG.STATUS_COL
        )
        .getValue()
    );


  if (!id) {

    if (
      oldStatus.includes(
        "已更新日曆"
      )
    ) {

      throw new Error(
        "此舊案件沒有 UUID，不能猜測舊列號還原；請人工核對日曆。"
      );
    }


    id =
      Utilities.getUuid();


    q.setValue(
      id
    );


    SpreadsheetApp.flush();
  }


  const c = {
    s: sheet,
    r: r,
    id: id
  };


  // 第一次確認目前列號。
  row_(c);


  return c;
}


// ============================================================
// UUID → 實際列號
//
// v6.5.1 效能優化：
//
// 1. 先使用 c.r 快取列號。
// 2. 只讀取目前列 Q 的單一儲存格確認 UUID。
// 3. UUID 仍相符 → 直接使用。
// 4. 若列被排序／插入／移動 → 才重新掃描整個 Q 欄。
// ============================================================

function row_(c) {

  const lastRow =
    c.s.getLastRow();


  // ----------------------------------------------------------
  // 快速路徑：
  // 目前快取列仍存在，且 Q 欄 UUID 沒變。
  // ----------------------------------------------------------

  if (

    Number.isInteger(
      c.r
    )

    &&

    c.r >= 2

    &&

    c.r <= lastRow

  ) {

    const currentId =
      text_(
        c.s
          .getRange(
            c.r,
            CONFIG.REQUEST_ID_COL
          )
          .getValue()
      );


    if (
      currentId ===
      c.id
    ) {

      return c.r;
    }
  }


  // ----------------------------------------------------------
  // 慢速路徑：
  // 快取位置已不正確，才重新掃描 Q 欄。
  // ----------------------------------------------------------

  if (
    lastRow < 2
  ) {

    throw new Error(
      "找不到申請資料。"
    );
  }


  const values =
    c.s
      .getRange(
        2,
        CONFIG.REQUEST_ID_COL,
        lastRow - 1,
        1
      )
      .getValues();


  const matches =
    [];


  values.forEach(
    (value, index) => {

      if (
        text_(
          value[0]
        ) ===
        c.id
      ) {

        matches.push(
          index + 2
        );
      }
    }
  );


  if (
    matches.length !== 1
  ) {

    throw new Error(
      "申請 ID 遺失或重複，請勿複製 Q 欄。"
    );
  }


  c.r =
    matches[0];


  return c.r;
}


// ============================================================
// 取得 Context 指定欄位
// ============================================================

function cell_(
  c,
  col
) {

  return c.s.getRange(
    row_(c),
    col
  );
}


function status_(
  c,
  message
) {

  cell_(
    c,
    CONFIG.STATUS_COL
  ).setValue(
    message
  );
}


// ============================================================
// 交易狀態
// ============================================================

function state_(c) {

  const st =
    readNote_(
      cell_(
        c,
        CONFIG.REQUEST_ID_COL
      ),
      "shift-tx"
    );


  if (
    st &&
    st.id !== c.id
  ) {

    throw new Error(
      "申請 ID 與快照不一致。"
    );
  }


  return st;
}


function saveState_(
  c,
  st
) {

  st.updated =
    now_();


  writeNote_(
    cell_(
      c,
      CONFIG.REQUEST_ID_COL
    ),
    st
  );
}


// ============================================================
// 是否存在尚未復原的其他交易
// ============================================================

function assertNoPending_(
  sheet,
  ownId
) {

  if (
    sheet.getLastRow() < 2
  ) {

    return;
  }


  const notes =
    sheet
      .getRange(
        2,
        CONFIG.REQUEST_ID_COL,
        sheet.getLastRow() - 1,
        1
      )
      .getNotes();


  notes.forEach(
    (entry, index) => {

      if (
        !entry[0]
      ) {

        return;
      }


      let st;


      try {

        st =
          JSON.parse(
            entry[0]
          );

      } catch (_) {

        throw new Error(
          "第 " +
          (index + 2) +
          " 列 Q 註解異常，請先核對。"
        );
      }


      if (

        st.kind ===
          "shift-tx"

        &&

        OPEN_PHASES.includes(
          st.phase
        )

        &&

        st.id !==
          ownId

      ) {

        throw new Error(
          "第 " +
          (index + 2) +
          " 列異動尚待復原，暫停其他班表異動。"
        );
      }
    }
  );
}


// ============================================================
// 是否為真正的代班／換班人
// ============================================================

function hasRealTargetPerson(
  value
) {

  value =
    text_(
      value
    );


  return (

    !!value

    &&

    value !==
      "無"

    &&

    !value.startsWith(
      "請假"
    )

  );
}


// ============================================================
// 讀取表單資料
// ============================================================

function readRequest_(
  sheet,
  r
) {

  const values =
    sheet
      .getRange(
        r,
        1,
        1,
        11
      )
      .getValues()[0];


  const display =
    sheet
      .getRange(
        r,
        1,
        1,
        11
      )
      .getDisplayValues()[0];


  const date =
    value =>
      value instanceof Date

        ? Utilities.formatDate(
            value,
            CONFIG.TIME_ZONE,
            "yyyy/MM/dd"
          )

        : text_(
            value
          );


  return {

    person:
      text_(
        values[1]
      ),

    date:
      date(
        values[2]
      ),

    start:
      text_(
        display[3]
      ),

    end:
      text_(
        display[4]
      ),

    target:
      text_(
        values[5]
      ),

    swapDate:
      date(
        values[6]
      ),

    swapStart:
      text_(
        display[7]
      ),

    swapEnd:
      text_(
        display[8]
      ),

    memo:
      text_(
        values[9]
      )
  };
}


// ============================================================
// 日期＋時間解析
// ============================================================

function parseTime_(
  date,
  time
) {

  const dateMatch =
    text_(date)
      .match(
        /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
      );


  const timeMatch =
    text_(time)
      .replace(
        /：/g,
        ":"
      )
      .match(
        /^(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
      );


  if (
    !dateMatch ||
    !timeMatch
  ) {

    throw new Error(
      "日期或時間格式錯誤，請用日期及 HH:mm。"
    );
  }


  let hour =
    Number(
      timeMatch[2]
    );


  const minute =
    Number(
      timeMatch[3]
    );


  const second =
    Number(
      timeMatch[4] || 0
    );


  const ampm =
    text_(
      timeMatch[1] ||
      timeMatch[5]
    ).toUpperCase();


  if (

    hour > 23

    ||

    minute > 59

    ||

    second > 59

    ||

    (
      ampm &&
      (
        hour < 1 ||
        hour > 12
      )
    )

  ) {

    throw new Error(
      "時間超出合法範圍。"
    );
  }


  if (ampm) {

    hour =
      hour % 12 +

      (
        ampm === "PM" ||
        ampm === "下午"
          ? 12
          : 0
      );
  }


  const pad =
    value =>
      String(
        Number(value)
      ).padStart(
        2,
        "0"
      );


  const input =

    `${dateMatch[1]}/` +
    `${pad(dateMatch[2])}/` +
    `${pad(dateMatch[3])} ` +
    `${pad(hour)}:` +
    `${pad(minute)}:` +
    `${pad(second)}`;


  const result =
    Utilities.parseDate(
      input,
      CONFIG.TIME_ZONE,
      "yyyy/MM/dd HH:mm:ss"
    );


  if (

    Utilities.formatDate(
      result,
      CONFIG.TIME_ZONE,
      "yyyy/MM/dd HH:mm:ss"
    ) !==
    input

  ) {

    throw new Error(
      "日期不存在。"
    );
  }


  return result;
}


// ============================================================
// Calendar 標題解析
// ============================================================

function titleInfo_(
  title
) {

  let raw =
    text_(
      title
    );


  const tag =
    raw.match(
      /\s*\[(換班|代班|請假)\]\s*$/
    );


  const state =
    tag
      ? tag[1]
      : "";


  if (tag) {

    raw =
      raw
        .slice(
          0,
          tag.index
        )
        .trim();
  }


  const floor =
    raw.match(
      /(\([^)]*\)|（[^）]*）)$/
    );


  return {

    name:
      (
        floor
          ? raw.slice(
              0,
              floor.index
            )
          : raw
      ).trim(),

    floor:
      floor
        ? floor[0]
        : "",

    state:
      state
  };
}


// ============================================================
// 時段重疊
// ============================================================

function overlap_(
  start1,
  end1,
  start2,
  end2
) {

  return (

    start1 <
      end2

    &&

    start2 <
      end1

  );
}


// ============================================================
// 將同一人的多筆 Calendar Event 轉為實際班段
// ============================================================

function buildSegments_(
  events,
  name,
  start,
  end,
  label
) {

  const matched =
    events

      .filter(
        event =>

          titleInfo_(
            event.getTitle()
          ).name ===
            name

          &&

          overlap_(
            +start,
            +end,
            +event.getStartTime(),
            +event.getEndTime()
          )
      )

      .sort(
        (a, b) =>
          +a.getStartTime() -
          +b.getStartTime()
      );


  if (
    matched.length === 0
  ) {

    return {

      ok: false,

      message:
        `找不到 ${name} 的${label}。`
    };
  }


  if (
    matched.some(
      event =>
        event.isAllDayEvent()
    )
  ) {

    return {

      ok: false,

      message:
        `${name} 的${label}包含全天事件，請管理員確認。`
    };
  }


  // 可以有中間空檔，
  // 但申請起點及終點要落在實際班表範圍內。
  if (

    +start <
      +matched[0].getStartTime()

    ||

    +end >
      +matched[
        matched.length - 1
      ].getEndTime()

  ) {

    return {

      ok: false,

      message:
        `${name} 的申請起訖時間超出實際${label}範圍；` +
        "中間空檔可以保留，但起點與終點必須落在原班內。"
    };
  }


  // 多段班可以，但彼此不能真正重疊。
  for (
    let i = 1;
    i < matched.length;
    i++
  ) {

    if (

      +matched[i]
        .getStartTime()

      <

      +matched[i - 1]
        .getEndTime()

    ) {

      return {

        ok: false,

        message:
          `${name} 在申請範圍內有互相重疊的多筆行程，請管理員確認。`
      };
    }
  }


  return {

    ok: true,

    segments:
      matched.map(
        event => ({

          event:
            event,

          s:
            Math.max(
              +start,
              +event.getStartTime()
            ),

          e:
            Math.min(
              +end,
              +event.getEndTime()
            )

        })
      )
  };
}


// ============================================================
// Calendar occurrence 精準比對
// ============================================================

function sameCalendarOccurrence_(
  a,
  b
) {

  return (

    a &&
    b

    &&

    a.getId() ===
      b.getId()

    &&

    +a.getStartTime() ===
      +b.getStartTime()

    &&

    +a.getEndTime() ===
      +b.getEndTime()

  );
}


// ============================================================
// 某段衝突是否完全屬於即將讓出的班段
// ============================================================

function intervalCoveredByCeding_(
  event,
  overlapStart,
  overlapEnd,
  cedingSegments
) {

  const ranges =
    (
      cedingSegments ||
      []
    )

      .filter(
        segment =>
          sameCalendarOccurrence_(
            event,
            segment.event
          )
      )

      .map(
        segment => [
          segment.s,
          segment.e
        ]
      )

      .sort(
        (a, b) =>
          a[0] - b[0]
      );


  if (
    !ranges.length
  ) {

    return false;
  }


  let cursor =
    overlapStart;


  for (
    const range of ranges
  ) {

    if (
      range[1] <=
      cursor
    ) {

      continue;
    }


    if (
      range[0] >
      cursor
    ) {

      return false;
    }


    cursor =
      Math.max(
        cursor,
        range[1]
      );


    if (
      cursor >=
      overlapEnd
    ) {

      return true;
    }
  }


  return (
    cursor >=
    overlapEnd
  );
}


// ============================================================
// 多段班撞班檢查
// ============================================================

function findConflictForSegments_(
  events,
  personName,
  incomingSegments,
  cedingSegments
) {

  for (
    const event of events
  ) {

    if (

      titleInfo_(
        event.getTitle()
      ).name !==
        personName

    ) {

      continue;
    }


    const eventStart =
      +event.getStartTime();


    const eventEnd =
      +event.getEndTime();


    for (
      const incoming of incomingSegments
    ) {

      const overlapStart =
        Math.max(
          incoming.s,
          eventStart
        );


      const overlapEnd =
        Math.min(
          incoming.e,
          eventEnd
        );


      if (
        overlapStart >=
        overlapEnd
      ) {

        continue;
      }


      if (
        intervalCoveredByCeding_(
          event,
          overlapStart,
          overlapEnd,
          cedingSegments
        )
      ) {

        continue;
      }


      return event;
    }
  }


  return null;
}


// ============================================================
// 統一申請分析
//
// 預檢與正式審核都走同一函式。
// 每次都重新讀取當下最新 Calendar。
// ============================================================

function analyzeRequest(
  sheet,
  calendar,
  r
) {

  const req =
    readRequest_(
      sheet,
      r
    );


  const reject =
    (
      code,
      message
    ) => ({

      ok: false,
      code: code,
      message: message,
      req: req

    });


  const hasTarget =
    hasRealTargetPerson(
      req.target
    );


  const anySwap =
    !!(

      req.swapDate ||
      req.swapStart ||
      req.swapEnd

    );


  const allSwap =
    !!(

      req.swapDate &&
      req.swapStart &&
      req.swapEnd

    );


  if (

    !req.person ||
    !req.date ||
    !req.start ||
    !req.end

  ) {

    return reject(
      "MISSING",
      "請填寫申請人、原值班日期與起訖時間。"
    );
  }


  if (

    hasTarget &&
    req.person ===
      req.target

  ) {

    return reject(
      "SELF",
      "申請人與配合人不能是同一人。"
    );
  }


  if (

    hasTarget &&
    anySwap &&
    !allSwap

  ) {

    return reject(
      "PARTIAL",
      "雙向換班的 G/H/I 請完整填寫；單向代班請全部留白。"
    );
  }


  if (

    !hasTarget &&
    anySwap

  ) {

    return reject(
      "LEAVE_FIELDS",
      "純請假的 G/H/I 請全部留白。"
    );
  }


  let originalStart;
  let originalEnd;
  let swapStart;
  let swapEnd;


  try {

    originalStart =
      parseTime_(
        req.date,
        req.start
      );


    originalEnd =
      parseTime_(
        req.date,
        req.end
      );


    if (
      !(originalStart <
        originalEnd)
    ) {

      throw new Error(
        "原班結束時間必須晚於開始時間，不支援跨日填法。"
      );
    }


    if (
      hasTarget &&
      allSwap
    ) {

      swapStart =
        parseTime_(
          req.swapDate,
          req.swapStart
        );


      swapEnd =
        parseTime_(
          req.swapDate,
          req.swapEnd
        );


      if (
        !(swapStart <
          swapEnd)
      ) {

        throw new Error(
          "互換結束時間必須晚於開始時間。"
        );
      }
    }

  } catch (err) {

    return reject(
      "TIME",
      err.message
    );
  }


  // ----------------------------------------------------------
  // 原值班：允許多段。
  // ----------------------------------------------------------

  const originalEvents =
    calendar.getEvents(
      originalStart,
      originalEnd
    );


  const originalResult =
    buildSegments_(
      originalEvents,
      req.person,
      originalStart,
      originalEnd,
      "原值班行程"
    );


  if (
    !originalResult.ok
  ) {

    return reject(
      "ORIGINAL",
      originalResult.message
    );
  }


  const origSegments =
    originalResult.segments;


  const type =
    hasTarget

      ? (
          allSwap
            ? "swap"
            : "sub"
        )

      : "leave";


  let swapSegments =
    [];


  // ----------------------------------------------------------
  // 雙向換班
  // ----------------------------------------------------------

  if (
    type ===
    "swap"
  ) {

    const targetEvents =
      calendar.getEvents(
        swapStart,
        swapEnd
      );


    const targetResult =
      buildSegments_(
        targetEvents,
        req.target,
        swapStart,
        swapEnd,
        "互換時段行程"
      );


    if (
      !targetResult.ok
    ) {

      return reject(
        "TARGET",
        targetResult.message
      );
    }


    swapSegments =
      targetResult.segments;


    const targetConflict =
      findConflictForSegments_(
        originalEvents,
        req.target,
        origSegments,
        swapSegments
      );


    if (
      targetConflict
    ) {

      return reject(
        "CONFLICT_SWAP_TARGET",

        `${req.target} 在換到的原值班實際時段已有班 ` +
        `(${targetConflict.getTitle()} ` +
        `${formatEventTime_(targetConflict)})。`
      );
    }


    const originalConflict =
      findConflictForSegments_(
        targetEvents,
        req.person,
        swapSegments,
        origSegments
      );


    if (
      originalConflict
    ) {

      return reject(
        "CONFLICT_SWAP_ORIGINAL",

        `${req.person} 在換到的互換實際時段已有班 ` +
        `(${originalConflict.getTitle()} ` +
        `${formatEventTime_(originalConflict)})。`
      );
    }

  } else if (
    hasTarget
  ) {

    // --------------------------------------------------------
    // 單向代班／後補代班：
    // 只檢查真正有班的 origSegments。
    // 中間午休空檔不算代班時段。
    // --------------------------------------------------------

    const conflict =
      findConflictForSegments_(
        originalEvents,
        req.target,
        origSegments,
        []
      );


    if (
      conflict
    ) {

      return reject(
        "CONFLICT_SUB",

        `${req.target} 在實際代班時段已有值班／請假紀錄 ` +
        `(${conflict.getTitle()} ` +
        `${formatEventTime_(conflict)})。`
      );
    }
  }


  const origStates =
    origSegments.map(
      segment =>
        titleInfo_(
          segment.event.getTitle()
        ).state
    );


  return {

    ok: true,

    req:
      req,

    type:
      type,

    origSegments:
      origSegments,

    swapSegments:
      swapSegments,

    late:
      type === "sub" &&
      origStates.length > 0 &&
      origStates.every(
        state =>
          state === "請假"
      )
  };
}


// ============================================================
// Event 時間顯示
// ============================================================

function formatEventTime_(
  event
) {

  return (

    Utilities.formatDate(
      event.getStartTime(),
      CONFIG.TIME_ZONE,
      "HH:mm"
    )

    +

    "-"

    +

    Utilities.formatDate(
      event.getEndTime(),
      CONFIG.TIME_ZONE,
      "HH:mm"
    )

  );
}


// ============================================================
// 舊版 Base64 解碼
//
// enc_() 已於 v6.5.1 移除。
// 新版不再寫入 Base64 metadata。
// dec_() 必須保留，用於讀取舊 v6.3 Calendar。
// ============================================================

function dec_(
  text
) {

  return Utilities
    .newBlob(
      Utilities.base64DecodeWebSafe(
        text
      )
    )
    .getDataAsString(
      "UTF-8"
    );
}


// ============================================================
// 從舊版 description 取得 metadata
// ============================================================

function meta_(
  description,
  key
) {

  const match =
    String(
      description ||
      ""
    ).match(

      new RegExp(
        "^\\[" +
        key +
        ":([^\\]\\r\\n]*)\\]\\s*$",
        "m"
      )

    );


  return (
    match
      ? match[1]
      : null
  );
}


// ============================================================
// 取得舊版依賴鏈
// ============================================================

function owners_(
  description,
  depth
) {

  depth =
    depth || 0;


  if (
    depth > 20
  ) {

    throw new Error(
      "舊快照層數過多，請人工核對。"
    );
  }


  const result =
    [];


  const own =
    meta_(
      description,
      "ROW_ID"
    );


  if (own) {

    result.push(
      own
    );
  }


  const parents =
    meta_(
      description,
      "ANCESTORS_B64"
    );


  if (parents) {

    try {

      const decoded =
        JSON.parse(
          dec_(parents)
        );


      if (
        Array.isArray(
          decoded
        )
      ) {

        result.push(
          ...decoded
        );
      }

    } catch (_) {

      throw new Error(
        "舊版 ANCESTORS_B64 無法解析，請人工核對。"
      );
    }
  }


  const previous =
    meta_(
      description,
      "ORIG_DESC_B64"
    );


  if (previous) {

    result.push(
      ...owners_(
        dec_(previous),
        depth + 1
      )
    );
  }


  return [
    ...new Set(
      result
    )
  ];
}


// ============================================================
// Calendar 說明欄只保留人類可讀文字
// ============================================================

function visibleDesc_(
  description
) {

  let result =
    String(
      description ||
      ""
    )

      .replace(
        /^\[(?:ROW_ID|ANCESTORS_B64|PART_ID|ORIG_TITLE_B64|ORIG_DESC_B64|SPLIT_ORIG_TIME):[^\]\r\n]*\]\r?\n?/gm,
        ""
      )

      .replace(
        /^\[AUTO_SPLIT_CREATED\]\r?\n?/gm,
        ""
      )

      .replace(
        /^[ \t]*-[ \t]*申請[ \t]*ID[：:][^\r\n]*(?:\r?\n|$)/gmi,
        ""
      )

      .trim();


  result =
    result

      .replace(
        /\r?\n?-{10,}[ \t]*$/g,
        ""
      )

      .replace(
        /\n{3,}/g,
        "\n\n"
      )

      .trim();


  return result;
}


// ============================================================
// 合併新紀錄與既有人類可讀說明
// ============================================================

function composeDesc_(
  detail,
  previousHuman
) {

  const detailText =
    visibleDesc_(
      detail
    );


  const previousText =
    visibleDesc_(
      previousHuman
    );


  if (
    detailText &&
    previousText
  ) {

    return (

      detailText +

      "\n--------------------\n" +

      previousText

    ).trim();
  }


  return (
    detailText ||
    previousText
  ).trim();
}


// ============================================================
// 從舊 description 讀取 metadata
// ============================================================

function legacyMetaFromDesc_(
  description
) {

  const desc =
    String(
      description ||
      ""
    );


  return {

    rowId:
      meta_(
        desc,
        "ROW_ID"
      ) || "",

    partId:
      meta_(
        desc,
        "PART_ID"
      ) || "",

    splitChild:
      /^\[AUTO_SPLIT_CREATED\]\s*$/m
        .test(
          desc
        ),

    legacyOwners:
      owners_(
        desc
      )
  };
}


// ============================================================
// Event metadata
//
// 優先讀 Calendar Tag。
// 若尚未遷移，兼容舊 description metadata。
// ============================================================

function eventMeta_(
  event
) {

  const legacy =
    legacyMetaFromDesc_(
      event.getDescription() ||
      ""
    );


  const tagRow =
    event.getTag(
      EVENT_TAGS.ROW_ID
    );


  const tagPart =
    event.getTag(
      EVENT_TAGS.PART_ID
    );


  return {

    rowId:
      text_(tagRow) ||
      legacy.rowId,

    partId:
      text_(tagPart) ||
      legacy.partId,

    splitChild:
      Boolean(
        text_(tagPart)
      ) ||
      legacy.splitChild,

    legacyOwners:
      legacy.legacyOwners
  };
}


// ============================================================
// 快照標準化
// ============================================================

function normalizeSnapshot_(
  snapshot
) {

  const result =
    clone_(
      snapshot ||
      {}
    );


  const legacy =
    legacyMetaFromDesc_(
      result.desc ||
      ""
    );


  const explicit =
    result.meta ||
    {};


  result.desc =
    visibleDesc_(
      result.desc ||
      ""
    );


  result.meta = {

    rowId:
      text_(
        explicit.rowId
      ) ||
      legacy.rowId,

    partId:
      text_(
        explicit.partId
      ) ||
      legacy.partId,

    splitChild:
      Boolean(
        explicit.splitChild
      ) ||
      Boolean(
        text_(
          explicit.partId
        )
      ) ||
      legacy.splitChild
  };


  return result;
}


// ============================================================
// Calendar Event → 快照
// ============================================================

function snap_(
  event
) {

  return normalizeSnapshot_({

    id:
      event.getId(),

    s:
      +event.getStartTime(),

    e:
      +event.getEndTime(),

    title:
      event.getTitle(),

    desc:
      event.getDescription() ||
      "",

    location:
      event.getLocation() ||
      "",

    meta:
      eventMeta_(event)

  });
}


// ============================================================
// Metadata 比對
// ============================================================

function sameMeta_(
  a,
  b
) {

  a =
    a || {};


  b =
    b || {};


  return (

    text_(
      a.rowId
    ) ===
      text_(
        b.rowId
      )

    &&

    text_(
      a.partId
    ) ===
      text_(
        b.partId
      )

    &&

    Boolean(
      a.splitChild
    ) ===
      Boolean(
        b.splitChild
      )

  );
}


// ============================================================
// 完整快照比對
// ============================================================

function same_(
  a,
  b
) {

  a =
    normalizeSnapshot_(
      a
    );


  b =
    normalizeSnapshot_(
      b
    );


  return (

    a.s ===
      b.s

    &&

    a.e ===
      b.e

    &&

    a.title ===
      b.title

    &&

    a.desc ===
      b.desc

    &&

    a.location ===
      b.location

    &&

    sameMeta_(
      a.meta,
      b.meta
    )

  );
}


// ============================================================
// 中途失敗時判斷主 Event 是否仍可安全復原
// ============================================================

function compatible_(
  current,
  before,
  after
) {

  current =
    normalizeSnapshot_(
      current
    );


  before =
    normalizeSnapshot_(
      before
    );


  after =
    normalizeSnapshot_(
      after
    );


  const metaOkay =

    sameMeta_(
      current.meta,
      before.meta
    )

    ||

    sameMeta_(
      current.meta,
      after.meta
    );


  return (

    [
      before.title,
      after.title
    ].includes(
      current.title
    )

    &&

    [
      before.desc,
      after.desc
    ].includes(
      current.desc
    )

    &&

    (
      (
        current.s ===
          before.s

        &&

        current.e ===
          before.e
      )

      ||

      (
        current.s ===
          after.s

        &&

        current.e ===
          after.e
      )
    )

    &&

    current.location ===
      before.location

    &&

    metaOkay

  );
}


// ============================================================
// 寫入／移除 Calendar Tag
// ============================================================

function applyEventMeta_(
  event,
  meta
) {

  meta =
    meta ||
    {};


  const setOrDelete =
    (
      key,
      value
    ) => {

      const wanted =
        text_(
          value
        );


      const current =
        event.getTag(
          key
        );


      if (wanted) {

        if (
          current !==
          wanted
        ) {

          event.setTag(
            key,
            wanted
          );
        }

      } else if (
        current !== null
      ) {

        event.deleteTag(
          key
        );
      }
    };


  setOrDelete(
    EVENT_TAGS.ROW_ID,
    meta.rowId
  );


  setOrDelete(
    EVENT_TAGS.PART_ID,
    meta.partId
  );


  setOrDelete(
    EVENT_TAGS.VERSION,

    meta.rowId
      ? CONFIG.VERSION
      : ""
  );
}


// ============================================================
// 將 Event 恢復／套用到指定快照
// ============================================================

function putSnapshot_(
  event,
  target
) {

  const normalized =
    normalizeSnapshot_(
      target
    );


  if (

    +event.getStartTime() !==
      normalized.s

    ||

    +event.getEndTime() !==
      normalized.e

  ) {

    event.setTime(
      new Date(
        normalized.s
      ),
      new Date(
        normalized.e
      )
    );
  }


  if (
    event.getTitle() !==
    normalized.title
  ) {

    event.setTitle(
      normalized.title
    );
  }


  applyEventMeta_(
    event,
    normalized.meta
  );


  if (
    (
      event.getDescription() ||
      ""
    ) !==
    normalized.desc
  ) {

    event.setDescription(
      normalized.desc
    );
  }
}


// ============================================================
// 建立單一 Calendar Event 的修改 Plan
// ============================================================

function plan_(
  event,
  start,
  end,
  worker,
  tag,
  token,
  detail
) {

  const before =
    snap_(
      event
    );


  const after =
    clone_(
      before
    );


  const parentTokens =

    before.meta &&
    before.meta.rowId

      ? [
          before.meta.rowId
        ]

      : [];


  const human =
    before.desc;


  const title =

    `${worker}` +
    `${titleInfo_(before.title).floor}` +
    ` [${tag}]`;


  after.meta = {

    rowId:
      token,

    partId:
      "",

    splitChild:
      false
  };


  const additions =
    [];


  const add =
    (
      partStart,
      partEnd,
      partTitle,
      body
    ) => {

      const key =

        token +
        ":" +
        (
          additions.length +
          1
        );


      additions.push({

        key:
          key,

        stage:
          "NEW",

        id:
          "",

        want: {

          s:
            partStart,

          e:
            partEnd,

          title:
            partTitle,

          location:
            before.location,

          desc:
            visibleDesc_(
              body
            ),

          meta: {

            rowId:
              token,

            partId:
              key,

            splitChild:
              true
          }
        }
      });
    };


  // ----------------------------------------------------------
  // 整段異動
  // ----------------------------------------------------------

  if (

    start ===
      before.s

    &&

    end ===
      before.e

  ) {

    after.title =
      title;


    after.desc =
      composeDesc_(
        detail,
        human
      );


  } else {

    // --------------------------------------------------------
    // 部分時段切割
    // --------------------------------------------------------

    after.desc =
      human;


    if (
      start >
      before.s
    ) {

      // 保留前段作為主 Event。
      after.e =
        start;


      // 中間／後半的異動段。
      add(
        start,
        end,
        title,
        composeDesc_(
          detail,
          human
        )
      );


      // 若異動段不是到原 Event 結尾，
      // 再建立後半原班。
      if (
        end <
        before.e
      ) {

        add(
          end,
          before.e,
          before.title,
          human
        );
      }

    } else {

      // 異動前半，
      // 主 Event 保留後半。
      after.s =
        end;


      add(
        start,
        end,
        title,
        composeDesc_(
          detail,
          human
        )
      );
    }
  }


  return {

    token:
      token,

    parentTokens:
      parentTokens,

    before:
      before,

    after:
      after,

    additions:
      additions,

    touched:
      false,

    restored:
      false
  };
}


// ============================================================
// M 欄成功狀態文字
// ============================================================

function appliedText_(
  st
) {

  const label =

    st.type ===
      "swap"

      ? "雙向換班"

      : st.type ===
        "leave"

        ? "請假"

        : st.late

          ? "後補代班"

          : "代班";


  return (
    "已更新日曆（" +
    label +
    "完成）"
  );
}


// ============================================================
// 建立完整交易
// ============================================================

function newTransaction_(
  c,
  calendar,
  analysis
) {

  const tx =
    Utilities.getUuid();


  const base =
    c.id +
    "-" +
    tx;


  const tag =

    analysis.type ===
      "leave"

      ? "請假"

      : analysis.type ===
        "swap"

        ? "換班"

        : "代班";


  const firstWorker =

    analysis.type ===
      "leave"

      ? analysis.req.person

      : analysis.req.target;


  const plans =
    [];


  // ----------------------------------------------------------
  // 原值班人的多段班
  // ----------------------------------------------------------

  analysis.origSegments
    .forEach(
      (
        segment,
        index
      ) => {

        plans.push(

          plan_(
            segment.event,
            segment.s,
            segment.e,
            firstWorker,
            tag,

            `${base}-A${index + 1}`,

            `【${tag}紀錄】\n` +
            `- 原定值班：${analysis.req.person}\n` +
            `- 實際安排：${firstWorker}`
          )

        );
      }
    );


  // ----------------------------------------------------------
  // 雙向換班另一方的多段班
  // ----------------------------------------------------------

  if (
    analysis.type ===
    "swap"
  ) {

    analysis.swapSegments
      .forEach(
        (
          segment,
          index
        ) => {

          plans.push(

            plan_(
              segment.event,
              segment.s,
              segment.e,
              analysis.req.person,
              tag,

              `${base}-B${index + 1}`,

              `【換班紀錄】\n` +
              `- 原定值班：${analysis.req.target}\n` +
              `- 實際到勤：${analysis.req.person}`
            )

          );
        }
      );
  }


  return {

    kind:
      "shift-tx",

    version:
      CONFIG.VERSION,

    id:
      c.id,

    tx:
      tx,

    phase:
      "APPLYING",

    calendar:
      calendar.getId(),

    req:
      analysis.req,

    requestHash:
      hash_(
        analysis.req
      ),

    type:
      analysis.type,

    late:
      analysis.late,

    plans:
      plans,

    purpose:
      "rollback",

    error:
      ""
  };
}


// ============================================================
// 正式處理單筆 L
// ============================================================

function processSingleRow_(
  c,
  calendar
) {

  let st =
    state_(c);


  const checked =

    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).getValue() ===
      true;


  if (

    st &&
    OPEN_PHASES.includes(
      st.phase
    )

  ) {

    throw new Error(
      "此列有中斷異動，請用「班表工具→復原此列中斷的班表異動」。"
    );
  }


  assertNoPending_(
    c.s,
    c.id
  );


  const oldStatus =
    text_(
      cell_(
        c,
        CONFIG.STATUS_COL
      ).getValue()
    );


  // ----------------------------------------------------------
  // 舊版已成功案件：
  // 若只是保持 L 勾選，不需處理。
  // 若取消 L，才嘗試安全匯入。
  // ----------------------------------------------------------

  if (

    !st

    &&

    oldStatus.startsWith(
      "已更新日曆"
    )

  ) {

    if (checked) {

      return;
    }


    try {

      st =
        adoptLegacy_(
          c,
          calendar
        );

    } catch (err) {

      cell_(
        c,
        CONFIG.EXECUTE_CHECK_COL
      ).setValue(
        true
      );


      throw err;
    }
  }


  // ----------------------------------------------------------
  // 已成功 APPLIED
  // ----------------------------------------------------------

  if (

    st &&
    st.phase ===
      ACTIVE_PHASE

  ) {

    // L 保持勾選，不重複執行。
    if (checked) {

      status_(
        c,
        appliedText_(st)
      );

      return;
    }


    // L 被取消 → 嘗試完整還原。
    try {

      const events =
        eventsForTx_(
          calendar,
          st
        );


      preflightUndo_(
        c,
        st,
        events
      );


      st.phase =
        "UNDOING";


      st.purpose =
        "undo";


      saveState_(
        c,
        st
      );


      restoreTransaction_(
        c,
        calendar,
        st,
        events,
        false
      );


      finishRestore_(
        c,
        st
      );

    } catch (err) {

      if (
        st.phase ===
        "APPLIED"
      ) {

        // 還沒開始正式還原，
        // 所以 L 恢復勾選。
        cell_(
          c,
          CONFIG.EXECUTE_CHECK_COL
        ).setValue(
          true
        );


        cell_(
          c,
          CONFIG.EXECUTE_CHECK_COL
        ).setNote(
          "取消未執行：" +
          err.message
        );


        status_(
          c,
          appliedText_(st) +
          "；取消未執行：" +
          err.message
        );

      } else {

        // 還原途中出錯。
        markRecovery_(
          c,
          st,
          err
        );
      }
    }


    return;
  }


  if (!checked) {

    return;
  }


  if (

    st &&
    st.phase ===
      "UNDONE"

  ) {

    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).setValue(
      false
    );


    throw new Error(
      "這筆已取消並保留歷史紀錄；再次申請請另填一筆表單。"
    );
  }


  // ----------------------------------------------------------
  // 正式審核：
  // 再次讀取最新 Calendar。
  // ----------------------------------------------------------

  const analysis =
    analyzeRequest(
      c.s,
      calendar,
      row_(c)
    );


  if (
    !analysis.ok
  ) {

    reject_(
      c,
      analysis,
      "正式審核"
    );

    return;
  }


  st =
    newTransaction_(
      c,
      calendar,
      analysis
    );


  // 必須先成功保存交易快照，
  // 才能開始修改 Calendar。
  saveState_(
    c,
    st
  );


  const references = [

    ...analysis
      .origSegments
      .map(
        segment =>
          segment.event
      ),

    ...analysis
      .swapSegments
      .map(
        segment =>
          segment.event
      )

  ];


  try {

    for (
      let index = 0;
      index < st.plans.length;
      index++
    ) {

      const p =
        st.plans[index];


      const event =
        references[index];


      if (

        !same_(
          snap_(event),
          p.before
        )

      ) {

        throw new Error(
          "原事件已變動，停止本次異動。"
        );
      }


      p.touched =
        true;


      saveState_(
        c,
        st
      );


      putSnapshot_(
        event,
        p.after
      );


      // ------------------------------------------------------
      // 建立切割後新增 Event
      // ------------------------------------------------------

      for (
        const addition of
        p.additions
      ) {

        addition.stage =
          "CREATING";


        saveState_(
          c,
          st
        );


        let created;


        try {

          created =
            calendar.createEvent(
              addition.want.title,

              new Date(
                addition.want.s
              ),

              new Date(
                addition.want.e
              ),

              {
                description:
                  addition.want.desc,

                location:
                  addition.want.location
              }
            );

        } catch (err) {

          const definitelyRejected =

            /permission|authorization|quota|too many times|invalid (argument|date|time)/i
              .test(
                err.message
              );


          if (
            definitelyRejected
          ) {

            addition.stage =
              "NEW";


            saveState_(
              c,
              st
            );
          }


          throw err;
        }


        addition.id =
          created.getId();


        addition.stage =
          "CREATED_UNTAGGED";


        saveState_(
          c,
          st
        );


        putSnapshot_(
          created,
          addition.want
        );


        addition.stage =
          "CREATED";


        saveState_(
          c,
          st
        );
      }
    }


    st.phase =
      ACTIVE_PHASE;


    saveState_(
      c,
      st
    );


    status_(
      c,
      appliedText_(st)
    );


    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).setNote(
      ""
    );

  } catch (err) {

    st.error =
      err.message;


    st.purpose =
      "rollback";


    st.phase =
      "RECOVER";


    try {

      saveState_(
        c,
        st
      );


      restoreTransaction_(
        c,
        calendar,
        st,
        eventsForTx_(
          calendar,
          st
        ),
        false
      );


      finishRestore_(
        c,
        st
      );

    } catch (
      recoveryError
    ) {

      markRecovery_(
        c,
        st,
        recoveryError
      );
    }
  }
}


// ============================================================
// 取得交易相關 Calendar Events
// ============================================================

function eventsForTx_(
  calendar,
  st
) {

  if (
    calendar.getId() !==
    st.calendar
  ) {

    throw new Error(
      "日曆 ID 與原異動紀錄不同。"
    );
  }


  const points =
    st.plans.flatMap(
      p => [

        p.before.s,
        p.before.e,

        ...p.additions
          .flatMap(
            addition => [
              addition.want.s,
              addition.want.e
            ]
          )

      ]
    );


  return calendar.getEvents(

    new Date(
      Math.min(
        ...points
      ) -
      86400000
    ),

    new Date(
      Math.max(
        ...points
      ) +
      86400000
    )

  );
}


// ============================================================
// 找主 Event
// ============================================================

function mainEvent_(
  p,
  events
) {

  const before =
    normalizeSnapshot_(
      p.before
    );


  const after =
    normalizeSnapshot_(
      p.after
    );


  const candidates =
    events.filter(
      event =>

        event.getId() ===
          before.id

        &&

        (
          (
            +event.getStartTime() ===
              before.s

            ||

            +event.getStartTime() ===
              after.s
          )

          ||

          eventMeta_(
            event
          ).rowId ===
            p.token
        )
    );


  if (
    candidates.length !== 1
  ) {

    throw new Error(
      "原主事件不存在或不唯一，不能安全還原。"
    );
  }


  return candidates[0];
}


// ============================================================
// 找切割新增 Event
// ============================================================

function addedEvent_(
  addition,
  events
) {

  const wanted =
    normalizeSnapshot_(
      addition.want
    );


  const candidates =
    events.filter(
      event =>

        eventMeta_(
          event
        ).partId ===
          addition.key

        ||

        (
          addition.id

          &&

          event.getId() ===
            addition.id

          &&

          +event.getStartTime() ===
            wanted.s
        )
    );


  if (
    candidates.length > 1
  ) {

    throw new Error(
      "新增事件識別碼重複，請人工核對。"
    );
  }


  return (
    candidates[0] ||
    null
  );
}


// ============================================================
// Plan 的上游依賴 Token
// ============================================================

function planParentTokens_(
  plan
) {

  if (

    Array.isArray(
      plan.parentTokens
    )

    &&

    plan.parentTokens.length

  ) {

    return [

      ...new Set(

        plan.parentTokens
          .map(
            text_
          )
          .filter(
            Boolean
          )

      )

    ];
  }


  const before =
    normalizeSnapshot_(
      plan.before
    );


  return (

    before.meta &&
    before.meta.rowId

      ? [
          before.meta.rowId
        ]

      : []

  );
}


// ============================================================
// 檢查是否仍有後續申請依賴本交易
// ============================================================

function checkDependents_(
  c,
  tokens,
  events
) {

  const lastRow =
    c.s.getLastRow();


  // ----------------------------------------------------------
  // 第一層：從 Q 欄交易快照判斷。
  // ----------------------------------------------------------

  if (
    lastRow >= 2
  ) {

    const notes =
      c.s
        .getRange(
          2,
          CONFIG.REQUEST_ID_COL,
          lastRow - 1,
          1
        )
        .getNotes();


    notes.forEach(
      (
        entry,
        index
      ) => {

        if (
          !entry[0]
        ) {

          return;
        }


        let other;


        try {

          other =
            JSON.parse(
              entry[0]
            );

        } catch (_) {

          throw new Error(
            "第 " +
            (index + 2) +
            " 列 Q 註解異常，請先核對。"
          );
        }


        if (

          !other

          ||

          other.kind !==
            "shift-tx"

          ||

          other.id ===
            c.id

          ||

          !(
            other.phase ===
              ACTIVE_PHASE

            ||

            OPEN_PHASES.includes(
              other.phase
            )
          )

        ) {

          return;
        }


        const dependent =

          (
            other.plans ||
            []
          )
            .some(
              plan =>

                planParentTokens_(
                  plan
                )
                  .some(
                    parent =>
                      tokens.includes(
                        parent
                      )
                  )
            );


        if (
          dependent
        ) {

          throw new Error(
            "第 " +
            (index + 2) +
            " 列仍依賴這筆班表異動，請先取消後補代班／後續異動。"
          );
        }
      }
    );
  }


  // ----------------------------------------------------------
  // 第二層：
  // 相容舊 v6.3 Description dependency。
  // ----------------------------------------------------------

  for (
    const event of
    events
  ) {

    const raw =
      event.getDescription() ||
      "";


    const top =
      eventMeta_(
        event
      ).rowId;


    if (

      top

      &&

      !tokens.includes(
        top
      )

      &&

      owners_(
        raw
      )
        .some(
          token =>
            tokens.includes(
              token
            )
        )

    ) {

      throw new Error(
        "仍有後補代班／後續異動「" +
        event.getTitle() +
        "」，請先取消後續申請。"
      );
    }
  }
}


// ============================================================
// 還原後是否撞上後來新增的其他班
// ============================================================

function checkRestoreConflicts_(
  st,
  events
) {

  const tokens =
    st.plans.map(
      p =>
        p.token
    );


  const outside =
    events.filter(
      event => {

        if (
          tokens.includes(
            eventMeta_(
              event
            ).rowId
          )
        ) {

          return false;
        }


        return !st.plans
          .some(
            p => {

              const before =
                normalizeSnapshot_(
                  p.before
                );


              const after =
                normalizeSnapshot_(
                  p.after
                );


              return (

                (
                  event.getId() ===
                    before.id

                  &&

                  [
                    +before.s,
                    +after.s
                  ].includes(
                    +event.getStartTime()
                  )
                )

                ||

                p.additions.some(
                  addition => {

                    const wanted =
                      normalizeSnapshot_(
                        addition.want
                      );


                    return (

                      addition.id

                      &&

                      event.getId() ===
                        addition.id

                      &&

                      +event.getStartTime() ===
                        wanted.s

                    );
                  }
                )

              );
            }
          );
      }
    );


  for (
    const p of
    st.plans
  ) {

    const before =
      normalizeSnapshot_(
        p.before
      );


    const name =
      titleInfo_(
        before.title
      ).name;


    const conflict =
      outside.some(
        event =>

          titleInfo_(
            event.getTitle()
          ).name ===
            name

          &&

          overlap_(
            before.s,
            before.e,
            +event.getStartTime(),
            +event.getEndTime()
          )
      );


    if (
      conflict
    ) {

      throw new Error(
        "還原後 " +
        name +
        " 會與另一筆值班／請假紀錄衝突，請先協調。"
      );
    }
  }
}


// ============================================================
// 確認 APPLIED 的 Calendar 仍完全符合交易完成狀態
// ============================================================

function assertAppliedIntact_(
  c,
  st,
  events
) {

  checkDependents_(
    c,
    st.plans.map(
      p =>
        p.token
    ),
    events
  );


  for (
    const p of
    st.plans
  ) {

    if (

      !same_(
        snap_(
          mainEvent_(
            p,
            events
          )
        ),
        p.after
      )

    ) {

      throw new Error(
        "主事件已被其他操作修改，取消已攔阻。"
      );
    }


    for (
      const addition of
      p.additions
    ) {

      const event =
        addedEvent_(
          addition,
          events
        );


      if (

        !event

        ||

        !same_(
          snap_(event),
          addition.want
        )

      ) {

        throw new Error(
          "衍生事件已遺失或變動，取消已攔阻。"
        );
      }
    }
  }
}


// ============================================================
// 正式取消前完整檢查
// ============================================================

function preflightUndo_(
  c,
  st,
  events
) {

  assertAppliedIntact_(
    c,
    st,
    events
  );


  checkRestoreConflicts_(
    st,
    events
  );
}


// ============================================================
// 還原整筆交易
// ============================================================

function restoreTransaction_(
  c,
  calendar,
  st,
  events,
  confirmed
) {

  checkDependents_(
    c,
    st.plans.map(
      p =>
        p.token
    ),
    events
  );


  checkRestoreConflicts_(
    st,
    events
  );


  const errors =
    [];


  // 雙向換班時從最後處理的一側開始倒序還原。
  for (
    const p of
    [...st.plans]
      .reverse()
  ) {

    if (

      !p.touched ||
      p.restored

    ) {

      continue;
    }


    try {

      const main =
        mainEvent_(
          p,
          events
        );


      if (

        !compatible_(
          snap_(main),
          p.before,
          p.after
        )

      ) {

        throw new Error(
          "主事件出現非本次修改，未覆寫。"
        );
      }


      // ------------------------------------------------------
      // 先刪除本次建立的衍生 Event。
      // ------------------------------------------------------

      for (
        const addition of
        [...p.additions]
          .reverse()
      ) {

        if (

          addition.stage ===
            "NEW"

          ||

          addition.stage ===
            "DELETED"

        ) {

          continue;
        }


        const event =
          addedEvent_(
            addition,
            events
          );


        if (!event) {

          if (

            (
              addition.stage ===
                "CREATING"

              ||

              addition.stage ===
                "CREATED_UNTAGGED"
            )

            &&

            !addition.id

            &&

            !confirmed

          ) {

            throw new Error(
              "新增 API 結果不明：未能確認是否曾建立事件，須人工核對。"
            );
          }

        } else {

          if (

            !same_(
              snap_(event),
              addition.want
            )

          ) {

            throw new Error(
              "衍生事件已被修改，未刪除。"
            );
          }


          addition.stage =
            "DELETING";


          saveState_(
            c,
            st
          );


          event.deleteEvent();


          // 從同一次 getEvents() 的記憶體清單移除，
          // 不再重新 query Calendar。
          events.splice(
            events.indexOf(
              event
            ),
            1
          );
        }


        addition.stage =
          "DELETED";


        saveState_(
          c,
          st
        );
      }


      // ------------------------------------------------------
      // 再精確恢復主 Event。
      // ------------------------------------------------------

      putSnapshot_(
        main,
        p.before
      );


      p.restored =
        true;


      saveState_(
        c,
        st
      );

    } catch (err) {

      errors.push(
        err.message
      );
    }
  }


  if (
    errors.length
  ) {

    throw new Error(
      errors.join(
        "；"
      )
    );
  }
}


// ============================================================
// 完成 Undo / Rollback
// ============================================================

function finishRestore_(
  c,
  st
) {

  st.phase =

    st.purpose ===
      "undo"

      ? "UNDONE"

      : "ROLLED_BACK";


  saveState_(
    c,
    st
  );


  cell_(
    c,
    CONFIG.EXECUTE_CHECK_COL
  ).setValue(
    false
  );


  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(
    false
  );


  const priorMail =
    text_(
      cell_(
        c,
        CONFIG.APPROVE_LOG_COL
      ).getValue()
    );


  if (
    st.phase ===
    "UNDONE"
  ) {

    status_(
      c,

      "已還原班表"

      +

      (
        priorMail
          ? "；曾有通知紀錄，請另行通知相關人員"
          : ""
      )
    );

  } else {

    status_(
      c,

      "執行失敗，已還原原班表；可重新勾 L 重試：" +
      st.error
    );
  }
}


// ============================================================
// 標記為需人工復原
// ============================================================

function markRecovery_(
  c,
  st,
  err
) {

  st.phase =
    "RECOVER";


  st.error =
    [
      st.error,
      err.message
    ]
      .filter(
        Boolean
      )
      .join(
        "；"
      );


  try {

    saveState_(
      c,
      st
    );

  } catch (
    saveError
  ) {

    console.error(
      saveError
    );
  }


  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(
    false
  );


  status_(
    c,

    "需人工確認：日曆異動／還原尚未完成；暫停其他異動。" +
    st.error
  );
}


// ============================================================
// 舊版 v6.2 / v6.3 成功案件安全匯入
// ============================================================

function adoptLegacy_(
  c,
  calendar
) {

  const req =
    readRequest_(
      c.s,
      row_(c)
    );


  const originalStart =
    +parseTime_(
      req.date,
      req.start
    );


  const originalEnd =
    +parseTime_(
      req.date,
      req.end
    );


  const isSwap = (

    hasRealTargetPerson(
      req.target
    )

    &&

    req.swapDate

    &&

    req.swapStart

    &&

    req.swapEnd

  );


  const windows = [{

    s:
      originalStart,

    e:
      originalEnd,

    person:
      req.person,

    token:
      c.id +
      "-A"

  }];


  if (isSwap) {

    windows.push({

      s:
        +parseTime_(
          req.swapDate,
          req.swapStart
        ),

      e:
        +parseTime_(
          req.swapDate,
          req.swapEnd
        ),

      person:
        req.target,

      token:
        c.id +
        "-B"
    });
  }


  const events =
    calendar.getEvents(

      new Date(

        Math.min(
          ...windows.map(
            w => w.s
          )
        ) -
        86400000

      ),

      new Date(

        Math.max(
          ...windows.map(
            w => w.e
          )
        ) +
        86400000

      )

    );


  checkDependents_(
    c,
    windows.map(
      w =>
        w.token
    ),
    events
  );


  const plans =
    windows.map(
      w => {

        const group =
          events.filter(
            event =>
              eventMeta_(
                event
              ).rowId ===
                w.token
          );


        const mainEvents =
          group.filter(
            event =>
              !eventMeta_(
                event
              ).splitChild
          );


        if (
          mainEvents.length !== 1
        ) {

          throw new Error(
            "舊案件主事件不完整，請人工核對，不能自動還原。"
          );
        }


        const main =
          mainEvents[0];


        const rawDescription =
          main.getDescription() ||
          "";


        const originalTitle =
          meta_(
            rawDescription,
            "ORIG_TITLE_B64"
          );


        const originalDescription =
          meta_(
            rawDescription,
            "ORIG_DESC_B64"
          );


        if (

          originalTitle ===
            null

          ||

          originalDescription ===
            null

        ) {

          throw new Error(
            "舊案件缺少完整狀態快照，請人工處理。"
          );
        }


        const after =
          snap_(
            main
          );


        let before =
          clone_(
            after
          );


        before.title =
          dec_(
            originalTitle
          );


        before.desc =
          dec_(
            originalDescription
          );


        const split =
          meta_(
            rawDescription,
            "SPLIT_ORIG_TIME"
          );


        if (split) {

          const parts =
            split
              .split(
                "-"
              )
              .map(
                Number
              );


          before.s =
            parts[0];


          before.e =
            parts[1];
        }


        before =
          normalizeSnapshot_(
            before
          );


        if (

          titleInfo_(
            before.title
          ).name !==
            w.person

          ||

          w.s <
            before.s

          ||

          w.e >
            before.e

        ) {

          throw new Error(
            "舊申請資料與日曆快照不一致。"
          );
        }


        const expected =

          Number(
            w.s >
            before.s
          )

          +

          Number(
            w.e <
            before.e
          );


        if (
          group.length !==
          1 + expected
        ) {

          throw new Error(
            "舊案件衍生事件數量不符，已停止。"
          );
        }


        const additions =
          group

            .filter(
              event =>
                event !==
                main
            )

            .map(
              (
                event,
                index
              ) => ({

                key:
                  w.token +
                  ":legacy:" +
                  index,

                id:
                  event.getId(),

                stage:
                  "CREATED",

                want:
                  snap_(
                    event
                  )

              })
            );


        return {

          token:
            w.token,

          parentTokens:

            before.meta &&
            before.meta.rowId

              ? [
                  before.meta.rowId
                ]

              : [],

          before:
            before,

          after:
            after,

          additions:
            additions,

          touched:
            true,

          restored:
            false
        };
      }
    );


  const st = {

    kind:
      "shift-tx",

    version:
      CONFIG.VERSION,

    id:
      c.id,

    tx:
      "legacy-" +
      c.id,

    phase:
      ACTIVE_PHASE,

    calendar:
      calendar.getId(),

    req:
      req,

    requestHash:
      hash_(
        req
      ),

    type:

      isSwap

        ? "swap"

        : hasRealTargetPerson(
            req.target
          )

          ? "sub"

          : "leave",

    late:
      text_(
        cell_(
          c,
          CONFIG.STATUS_COL
        ).getValue()
      ).includes(
        "後補代班"
      ),

    plans:
      plans,

    purpose:
      "undo",

    error:
      ""
  };


  saveState_(
    c,
    st
  );


  return st;
}


// ============================================================
// Email 格式
// ============================================================

function validEmail_(
  email
) {

  return (

    /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/

      .test(
        text_(email)
      )

  );
}


// ============================================================
// 申請人 Email
// ============================================================

function getApplicantEmail(
  sheet,
  r
) {

  const col =
    CONFIG.APPLICANT_EMAIL_COL;


  if (

    !Number.isInteger(
      col
    )

    ||

    col < 1

    ||

    col >
      sheet.getMaxColumns()

  ) {

    throw new Error(
      "CONFIG.APPLICANT_EMAIL_COL 未正確設定。"
    );
  }


  const email =
    text_(
      sheet
        .getRange(
          r,
          col
        )
        .getValue()
    );


  if (
    !validEmail_(
      email
    )
  ) {

    throw new Error(
      "K 欄申請人 Email 空白或格式錯誤。"
    );
  }


  return email;
}


// ============================================================
// 員工姓名 → Email
// ============================================================

function staffEmail_(
  sheet,
  name
) {

  const directory =
    sheet
      .getParent()
      .getSheetByName(
        CONFIG.STAFF_DIRECTORY_SHEET_NAME
      );


  if (

    !directory

    ||

    directory.getLastRow() <
      2

  ) {

    throw new Error(
      "員工名冊不存在或沒有資料。"
    );
  }


  const data =
    directory.getRange(

      2,
      1,

      directory.getLastRow() -
      1,

      Math.max(
        CONFIG.STAFF_NAME_COL,
        CONFIG.STAFF_EMAIL_COL
      )

    ).getValues();


  const matches =
    data.filter(
      entry =>

        text_(
          entry[
            CONFIG.STAFF_NAME_COL -
            1
          ]
        ) ===
          name
    );


  if (
    matches.length !== 1
  ) {

    throw new Error(
      "員工名冊姓名「" +
      name +
      "」不存在或重複。"
    );
  }


  const email =
    text_(
      matches[0][
        CONFIG.STAFF_EMAIL_COL -
        1
      ]
    );


  if (
    !validEmail_(
      email
    )
  ) {

    throw new Error(
      "名冊「" +
      name +
      "」Email 格式不正確。"
    );
  }


  return email;
}


// ============================================================
// 通知紀錄 Store
// ============================================================

function noticeStore_(
  c,
  col
) {

  const store =

    readNote_(
      cell_(
        c,
        col
      ),
      "shift-mail"
    )

    ||

    {

      kind:
        "shift-mail",

      id:
        c.id,

      batches:
        {},

      latest:
        ""
    };


  if (
    store.id !==
    c.id
  ) {

    throw new Error(
      "通知註解屬於另一筆申請，已停止。"
    );
  }


  return store;
}


function saveMail_(
  c,
  col,
  store
) {

  if (
    store.id !==
    c.id
  ) {

    throw new Error(
      "通知紀錄與申請 ID 不一致。"
    );
  }


  writeNote_(
    cell_(
      c,
      col
    ),
    store
  );
}


// ============================================================
// 單一寄信 Task
// ============================================================

function task_(
  role,
  name,
  subject,
  body
) {

  return {

    role:
      role,

    name:
      name,

    subject:
      subject,

    body:
      body,

    status:
      "READY",

    email:
      "",

    at:
      "",

    error:
      ""
  };
}


// ============================================================
// N 欄核准通知
// ============================================================

function approval_(
  c,
  deadline
) {

  assertNoPending_(
    c.s,
    ""
  );


  const st =
    state_(c);


  const currentStatus =
    text_(
      cell_(
        c,
        CONFIG.STATUS_COL
      ).getValue()
    );


  if (

    (
      !st

      &&

      !currentStatus
        .startsWith(
          "已更新日曆"
        )
    )

    ||

    (
      st &&
      st.phase !==
        ACTIVE_PHASE
    )

    ||

    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).getValue() !==
      true

  ) {

    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(
      false
    );


    throw new Error(
      "班表尚未成功異動或已取消，不能寄核准信。"
    );
  }


  const req =

    st
      ? st.req
      : readRequest_(
          c.s,
          row_(c)
        );


  // 審核成功後，
  // A–J 原申請內容不得偷偷被修改。
  if (

    st

    &&

    hash_(
      readRequest_(
        c.s,
        row_(c)
      )
    ) !==
      st.requestHash

  ) {

    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(
      false
    );


    throw new Error(
      "審核後申請內容已被修改，不能寄送與日曆不一致的通知。"
    );
  }


  // 正式寄信前，
  // 再確認 Calendar 還維持核准完成狀態。
  if (st) {

    try {

      assertAppliedIntact_(
        c,
        st,
        eventsForTx_(
          getCalendar_(),
          st
        )
      );

    } catch (err) {

      cell_(
        c,
        CONFIG.APPROVE_CHECK_COL
      ).setValue(
        false
      );


      throw new Error(
        "班表已有後續變動，不寄出過期核准通知：" +
        err.message
      );
    }
  }


  const col =
    CONFIG.APPROVE_LOG_COL;


  const store =
    noticeStore_(
      c,
      col
    );


  const key =

    st
      ? st.tx
      : "legacy-" +
        c.id;


  const legacyText =
    text_(
      cell_(
        c,
        col
      ).getValue()
    );


  // 舊版已有成功紀錄時不重寄。
  if (

    !st

    &&

    !store.latest

    &&

    legacyText.startsWith(
      "已寄送核准通知"
    )

  ) {

    return;
  }


  if (
    !store.batches[key]
  ) {

    const isSwap =

      st
        ? st.type ===
          "swap"

        : (
            hasRealTargetPerson(
              req.target
            )

            &&

            !!req.swapDate
          );


    const isSub = (

      hasRealTargetPerson(
        req.target
      )

      &&

      !isSwap

    );


    const label =

      isSwap

        ? "換班"

        : isSub

          ? (
              st &&
              st.late

                ? "後補代班"

                : "代班"
            )

          : "請假";


    const original =

      `${req.date} ` +
      `${req.start}–` +
      `${req.end}`;


    const applicantBody =

      `${req.person} 您好，\n\n` +

      `原值班時段：${original}\n` +

      (
        isSwap

          ? (
              `您改至：${req.swapDate} ` +
              `${req.swapStart}–${req.swapEnd}\n` +
              `配合人員：${req.target}\n`
            )

          : isSub

            ? (
                `代班人員：${req.target}\n`
              )

            : (
                "此時段已核准請假，尚無代班人員。\n"
              )
      )

      +

      "\n班表已更新，請以最新值班日曆為準。";


    const tasks = [

      task_(
        "applicant",
        req.person,
        `【值班${label}申請】已審核通過`,
        applicantBody
      )

    ];


    if (
      isSwap ||
      isSub
    ) {

      const targetBody =

        `${req.target} 您好，\n\n` +

        `您配合 ${req.person} 的申請已通過。\n`

        +

        (
          isSwap

            ? (
                `原定時段：${req.swapDate} ` +
                `${req.swapStart}–${req.swapEnd}\n`
              )

            : ""
        )

        +

        `實際到勤時段：${original}\n\n` +

        "班表已更新，請留意準時到勤。";


      tasks.push(

        task_(
          "target",
          req.target,

          isSwap

            ? "【值班換班通知】班表已完成互換"

            : "【值班代班通知】代班已審核通過",

          targetBody
        )

      );
    }


    // 舊版只留下「寄信失敗」時，
    // 無法知道其中某人是否其實已收到。
    if (

      !st

      &&

      /失敗/.test(
        legacyText
      )

    ) {

      tasks.forEach(
        task => {

          task.status =
            "UNKNOWN";


          task.error =
            "舊版紀錄無法確認此收件人是否已寄。";
        }
      );
    }


    store.batches[key] = {

      label:
        "核准通知",

      tasks:
        tasks
    };


    store.latest =
      key;


    saveMail_(
      c,
      col,
      store
    );
  }


  sendBatch_(
    c,
    col,
    store,
    key,
    deadline
  );
}


// ============================================================
// 預檢／正式審核退件
// ============================================================

function reject_(
  c,
  analysis,
  stage
) {

  cell_(
    c,
    CONFIG.EXECUTE_CHECK_COL
  ).setValue(
    false
  );


  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(
    false
  );


  status_(
    c,

    stage +
    "未通過：" +
    analysis.message
  );


  const col =
    CONFIG.REJECT_LOG_COL;


  const store =
    noticeStore_(
      c,
      col
    );


  // 同一筆資料＋同一錯誤原因，
  // 不會只因「預檢／正式審核」文字不同而重寄。
  const key =
    hash_([

      analysis.req,
      analysis.code,
      analysis.message

    ]);


  if (
    !store.batches[key]
  ) {

    store.batches[key] = {

      label:
        analysis.message,

      requestHash:
        hash_(
          analysis.req
        ),

      tasks: [

        task_(
          "applicant",
          analysis.req.person,
          "【值班異動申請】系統檢核未通過",

          `${analysis.req.person} 您好，\n\n` +

          `原值班：${analysis.req.date} ` +
          `${analysis.req.start}–${analysis.req.end}\n` +

          `申請未通過，原因：\n` +
          `${analysis.message}\n\n` +

          "請確認資料後重新提出申請，或洽管理員。"
        )

      ]
    };
  }


  store.latest =
    key;


  saveMail_(
    c,
    col,
    store
  );


  sendBatch_(
    c,
    col,
    store,
    key,
    Date.now() +
    60000
  );
}


// ============================================================
// 逐收件人寄信
// ============================================================

function sendBatch_(
  c,
  col,
  store,
  key,
  deadline
) {

  const batch =
    store.batches[key];


  for (
    let index = 0;
    index < batch.tasks.length;
    index++
  ) {

    const task =
      batch.tasks[index];


    if (
      task.status ===
      "SENT"
    ) {

      continue;
    }


    // 上次停在 SENDING：
    // 有可能已經送出，但尚未寫入 SENT。
    if (
      task.status ===
      "SENDING"
    ) {

      task.status =
        "UNKNOWN";


      task.error =
        "上次交寄未留下完成紀錄，請先人工核對。";


      saveMail_(
        c,
        col,
        store
      );
    }


    if (
      task.status ===
      "UNKNOWN"
    ) {

      continue;
    }


    if (
      Date.now() >
      deadline - 10000
    ) {

      break;
    }


    try {

      task.email =

        task.role ===
          "applicant"

          ? getApplicantEmail(
              c.s,
              row_(c)
            )

          : staffEmail_(
              c.s,
              task.name
            );


      const admins = [

        ...new Set(
          CONFIG.ADMIN_EMAILS
        )

      ]
        .filter(
          email =>

            email
              .toLowerCase() !==
            task.email
              .toLowerCase()
        );


      if (

        MailApp
          .getRemainingDailyQuota()

        <

        1 +
        admins.length

      ) {

        throw new Error(
          "本日寄信額度不足，尚未寄出。"
        );
      }


      task.status =
        "SENDING";


      task.error =
        "";


      // 先記錄「準備交寄」。
      saveMail_(
        c,
        col,
        store
      );


      const message = {

        to:
          task.email,

        subject:
          task.subject,

        body:

          task.body +

          `\n\n申請 ID：${c.id}` +

          `\n通知識別：${key}:${index}` +

          "\n（系統通知，請勿直接回覆）"

      };


      if (
        admins.length
      ) {

        message.bcc =
          admins.join(
            ","
          );
      }


      try {

        MailApp.sendEmail(
          message
        );

      } catch (err) {

        const definitelyNotSent =

          /permission|authorization|required permissions|quota|too many times|invalid.*(email|recipient)|無權限|授權|額度/i

            .test(
              err.message
            );


        task.status =

          definitelyNotSent
            ? "FAILED"
            : "UNKNOWN";


        task.error =
          err.message;


        saveMail_(
          c,
          col,
          store
        );


        continue;
      }


      task.status =
        "SENT";


      task.at =
        now_();


      // 每成功一封就立即記錄。
      saveMail_(
        c,
        col,
        store
      );

    } catch (err) {

      if (

        task.status ===
          "SENDING"

        ||

        task.status ===
          "SENT"

      ) {

        task.status =
          "UNKNOWN";

      } else {

        task.status =
          "FAILED";
      }


      task.error =
        err.message;


      try {

        saveMail_(
          c,
          col,
          store
        );

      } catch (
        saveError
      ) {

        console.error(
          saveError
        );
      }
    }
  }


  const done =
    batch.tasks.every(
      task =>
        task.status ===
        "SENT"
    );


  const labels = {

    SENT:
      "已寄送",

    READY:
      "尚未寄",

    FAILED:
      "未完成，可重試",

    SENDING:
      "寄送結果待核對",

    UNKNOWN:
      "結果不明，須人工核對"

  };


  const heading =

    col ===
      CONFIG.APPROVE_LOG_COL

      ? (
          done

            ? "已寄送核准通知"

            : "核准通知尚未全數完成"
        )

      : (
          "退件原因：" +
          batch.label
        );


  const details =
    batch.tasks
      .map(
        task =>

          `${task.name}：` +

          `${labels[task.status]} ` +

          `${task.at || ""}` +

          (
            task.error
              ? "；" +
                task.error
              : ""
          )
      )
      .join(
        "\n"
      );


  cell_(
    c,
    col
  ).setValue(

    heading +
    "\n" +
    details

  );


  if (
    col ===
    CONFIG.APPROVE_LOG_COL
  ) {

    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(
      done
    );
  }
}


// ============================================================
// 管理員選取目前列
// ============================================================

function selected_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();


  const r =
    sheet
      .getActiveRange()
      .getRow();


  if (

    !isResponseSheet_(
      sheet
    )

    ||

    r < 2

  ) {

    throw new Error(
      `請先在「${CONFIG.RESPONSE_SHEET_NAME}」分頁選取一筆申請。`
    );
  }


  validateConfig_(
    sheet
  );


  return context_(
    sheet,
    r
  );
}


// ============================================================
// 班表工具：重試未完成通知
// ============================================================

function retrySelectedNotifications() {

  locked_(
    () => {

      const c =
        selected_();


      const st =
        state_(c);


      if (

        st

        &&

        (
          OPEN_PHASES.includes(
            st.phase
          )

          ||

          st.phase ===
            "UNDONE"
        )

      ) {

        throw new Error(
          "此列異動未完成或已取消，不補寄先前的核准／退件信。"
        );
      }


      const isApplied = (

        st &&
        st.phase ===
          ACTIVE_PHASE

      );


      const isLegacyApplied = (

        !st

        &&

        text_(
          cell_(
            c,
            CONFIG.STATUS_COL
          ).getValue()
        ).startsWith(
          "已更新日曆"
        )

      );


      if (

        isApplied ||
        isLegacyApplied

      ) {

        cell_(
          c,
          CONFIG.APPROVE_CHECK_COL
        ).setValue(
          true
        );


        approval_(
          c,
          Date.now() +
          60000
        );

      } else {

        const store =
          noticeStore_(
            c,
            CONFIG.REJECT_LOG_COL
          );


        if (
          !store.latest
        ) {

          throw new Error(
            "此列沒有可重試的退件通知。"
          );
        }


        if (

          store
            .batches[
              store.latest
            ]
            .requestHash

          !==

          hash_(
            readRequest_(
              c.s,
              row_(c)
            )
          )

        ) {

          throw new Error(
            "申請內容已變更，請重新檢核，不補寄過期退件理由。"
          );
        }


        sendBatch_(
          c,
          CONFIG.REJECT_LOG_COL,
          store,
          store.latest,
          Date.now() +
          60000
        );
      }
    }
  );
}


// ============================================================
// 班表工具：核對 UNKNOWN 寄信
// ============================================================

function resolveSelectedMail() {

  const c =
    locked_(
      () =>
        selected_()
    );


  const ui =
    SpreadsheetApp
      .getUi();


  for (
    const col of [

      CONFIG.APPROVE_LOG_COL,
      CONFIG.REJECT_LOG_COL

    ]
  ) {

    const store =
      locked_(
        () =>
          noticeStore_(
            c,
            col
          )
      );


    for (
      const key of
      Object.keys(
        store.batches
      )
    ) {

      const tasks =
        store
          .batches[key]
          .tasks;


      for (
        let index = 0;
        index < tasks.length;
        index++
      ) {

        if (

          ![
            "UNKNOWN",
            "SENDING"
          ].includes(
            tasks[index].status
          )

        ) {

          continue;
        }


        const task =
          tasks[index];


        // UI 等候期間不持有 Script Lock。
        const answer =
          ui.alert(

            "先核對信箱／管理員副本",

            `收件人：${task.name} ${task.email}\n` +

            `主旨：${task.subject}\n` +

            `通知識別：${key}:${index}\n\n` +

            "已確認寄出→「是」；" +

            "已確認未寄出→「否」；" +

            "不確定→「取消」。",

            ui.ButtonSet
              .YES_NO_CANCEL
          );


        if (
          answer ===
          ui.Button.CANCEL
        ) {

          return;
        }


        locked_(
          () => {

            const latest =
              noticeStore_(
                c,
                col
              );


            const current =
              latest
                .batches[key]
                .tasks[index];


            if (

              ![
                "UNKNOWN",
                "SENDING"
              ].includes(
                current.status
              )

            ) {

              return;
            }


            current.status =

              answer ===
                ui.Button.YES

                ? "SENT"

                : "READY";


            current.at =

              answer ===
                ui.Button.YES

                ? (
                    now_() +
                    "（人工核對）"
                  )

                : "";


            current.error =
              "";


            saveMail_(
              c,
              col,
              latest
            );
          }
        );
      }
    }
  }


  ui.alert(
    "核對紀錄已保存。請再用「重試此列未完成通知」更新顯示並補寄。"
  );
}


// ============================================================
// 班表工具：復原中斷交易
// ============================================================

function recoverSelectedCalendar() {

  locked_(
    () => {

      const c =
        selected_();


      const st =
        state_(c);


      if (

        !st

        ||

        !OPEN_PHASES.includes(
          st.phase
        )

      ) {

        throw new Error(
          "此列沒有待復原的中斷異動。"
        );
      }


      const calendar =
        getCalendar_();


      try {

        restoreTransaction_(
          c,
          calendar,
          st,
          eventsForTx_(
            calendar,
            st
          ),
          false
        );


        finishRestore_(
          c,
          st
        );

      } catch (err) {

        markRecovery_(
          c,
          st,
          err
        );
      }
    }
  );
}


// ============================================================
// 班表工具：人工核對後解除 RECOVER
// ============================================================

function confirmSelectedRecovery() {

  const ui =
    SpreadsheetApp
      .getUi();


  const answer =
    ui.alert(

      "僅限人工核對後",

      "請先到日曆核對：" +
      "此筆申請新增的事件均已移除，原班已恢復。" +
      "確定後才繼續；程式還會再次比對快照。",

      ui.ButtonSet
        .YES_NO
    );


  if (
    answer !==
    ui.Button.YES
  ) {

    return;
  }


  locked_(
    () => {

      const c =
        selected_();


      const st =
        state_(c);


      if (

        !st

        ||

        !OPEN_PHASES.includes(
          st.phase
        )

      ) {

        throw new Error(
          "沒有需要解除的中斷紀錄。"
        );
      }


      const events =
        eventsForTx_(
          getCalendar_(),
          st
        );


      checkDependents_(
        c,
        st.plans.map(
          p =>
            p.token
        ),
        events
      );


      for (
        const p of
        st.plans
      ) {

        const originalMatches =
          same_(
            snap_(
              mainEvent_(
                p,
                events
              )
            ),
            p.before
          );


        const hasRemainingAdditions =
          p.additions.some(
            addition =>
              addedEvent_(
                addition,
                events
              )
          );


        if (

          !originalMatches

          ||

          hasRemainingAdditions

        ) {

          throw new Error(
            "目前日曆仍與原班快照不符，不能解除。"
          );
        }


        p.restored =
          true;


        p.additions
          .forEach(
            addition => {

              addition.stage =
                "DELETED";
            }
          );
      }


      finishRestore_(
        c,
        st
      );
    }
  );
}


// ============================================================
// 班表工具：清理舊 v6.3 顯示中的控制碼
// ============================================================

function cleanSelectedCalendarDisplay() {

  locked_(
    () => {

      const c =
        selected_();


      const st =
        state_(c);


      if (

        !st

        ||

        st.phase !==
          ACTIVE_PHASE

      ) {

        throw new Error(
          "請選擇已成功更新日曆、尚未取消的申請列。"
        );
      }


      const calendar =
        getCalendar_();


      const events =
        eventsForTx_(
          calendar,
          st
        );


      // 清理前先確認 Calendar 沒被其他人修改。
      assertAppliedIntact_(
        c,
        st,
        events
      );


      const targets =
        [];


      for (
        const p of
        st.plans
      ) {

        targets.push(
          mainEvent_(
            p,
            events
          )
        );


        for (
          const addition of
          p.additions
        ) {

          const event =
            addedEvent_(
              addition,
              events
            );


          if (event) {

            targets.push(
              event
            );
          }
        }
      }


      // 同一 Event 只處理一次。
      const unique =
        [];


      const seen =
        new Set();


      targets.forEach(
        event => {

          const key =

            event.getId() +
            "|" +
            event
              .getStartTime()
              .getTime();


          if (
            !seen.has(
              key
            )
          ) {

            seen.add(
              key
            );


            unique.push(
              event
            );
          }
        }
      );


      let changed =
        0;


      for (
        const event of
        unique
      ) {

        // snap_() 會把舊 description metadata
        // 轉換成邏輯快照。
        const logicalBefore =
          snap_(
            event
          );


        const rawBefore =
          event.getDescription() ||
          "";


        const oldTagRow =
          event.getTag(
            EVENT_TAGS.ROW_ID
          );


        const oldTagPart =
          event.getTag(
            EVENT_TAGS.PART_ID
          );


        putSnapshot_(
          event,
          logicalBefore
        );


        if (

          rawBefore !==
            event.getDescription()

          ||

          oldTagRow !==
            event.getTag(
              EVENT_TAGS.ROW_ID
            )

          ||

          oldTagPart !==
            event.getTag(
              EVENT_TAGS.PART_ID
            )

        ) {

          changed++;
        }


        // 清理後，邏輯快照必須完全相同。
        if (

          !same_(
            snap_(event),
            logicalBefore
          )

        ) {

          throw new Error(
            "清理後核對失敗，已停止；請檢查日曆。"
          );
        }
      }


      c.s
        .getParent()
        .toast(

          changed > 0

            ? (
                "已清理 " +
                changed +
                " 個日曆行程的系統代碼；班表內容未改變。"
              )

            : (
                "這筆申請的日曆顯示已經是乾淨格式。"
              )
        );
    }
  );
}
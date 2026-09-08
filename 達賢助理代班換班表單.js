/**
 * 達賢助理代班／換班系統 v6.3
 * 異常復原與逐封通知版
 *
 * 取代整套舊程式，不要接在舊程式後面。
 *
 * A–K：原表單資料
 * L：審核確認
 * M：執行狀態
 * N：審核通知
 * O：核准通知紀錄
 * P：退件通知紀錄
 * Q：申請 ID
 *
 * Q 儲存格註解：保存異動快照
 * O／P 儲存格註解：保存逐封寄信紀錄
 *
 * 請勿刪除 Q、O、P 的內容或註解。
 * 請勿只排序部分欄位。
 *
 * 保留兩個安裝式觸發器：
 *
 * handleSheetEdit
 * → 試算表 → 編輯時
 *
 * onFormSubmitPrecheck
 * → 試算表 → 提交表單時
 *
 * 系統錯誤不是申請退件。
 * API 結果不明時停住等待人工核對，
 * 不假稱全部成功或全部還原。
 * 郵件結果不明時也不盲目重寄。
 */

const CONFIG = {
  VERSION: "6.3",

  CALENDAR_NAME: "達賢館創新組助理值班",

  // 有同名日曆時，才需要填入正確日曆 ID。
  CALENDAR_ID: "",

  // 空白時，依 K／L／N 欄標題辨識回覆分頁。
  RESPONSE_SHEET_NAME: "",

  TIME_ZONE: "Asia/Taipei",

  APPLICANT_EMAIL_COL: 11,
  EXECUTE_CHECK_COL: 12,
  STATUS_COL: 13,
  APPROVE_CHECK_COL: 14,
  APPROVE_LOG_COL: 15,
  REJECT_LOG_COL: 16,
  REQUEST_ID_COL: 17,

  ADMIN_EMAILS: ["dhl.nccu@gmail.com"],

  STAFF_DIRECTORY_SHEET_NAME: "員工名冊",
  STAFF_NAME_COL: 1,
  STAFF_EMAIL_COL: 2,

  MAX_ROWS: 30,
  RUN_MS: 240000,
  NOTE_LIMIT: 45000
};

const OPEN_PHASES = [
  "APPLYING",
  "UNDOING",
  "RECOVER"
];

const ACTIVE_PHASE = "APPLIED";


// ============================================================
// 選單與入口
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
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
    .addToUi();
}


function isResponseSheet_(s) {
  if (CONFIG.RESPONSE_SHEET_NAME) {
    return s.getName() === CONFIG.RESPONSE_SHEET_NAME;
  }

  if (s.getMaxColumns() < 17) return false;

  const h = s.getRange(
    1,
    1,
    1,
    17
  ).getDisplayValues()[0];

  return (
    String(h[10]).trim() === "電子郵件地址" &&
    String(h[11]).includes("審核確認") &&
    String(h[13]).includes("審核通知")
  );
}


function validateConfig_(s) {
  const cols = [
    CONFIG.APPLICANT_EMAIL_COL,
    CONFIG.EXECUTE_CHECK_COL,
    CONFIG.STATUS_COL,
    CONFIG.APPROVE_CHECK_COL,
    CONFIG.APPROVE_LOG_COL,
    CONFIG.REJECT_LOG_COL,
    CONFIG.REQUEST_ID_COL
  ];

  const invalid = cols.some(c =>
    !Number.isInteger(c) ||
    c < 1 ||
    c > s.getMaxColumns()
  );

  if (
    invalid ||
    new Set(cols).size !== cols.length
  ) {
    throw new Error(
      "CONFIG 欄位設定缺漏、重複或超出範圍，已停止。"
    );
  }

  CONFIG.ADMIN_EMAILS.forEach(email => {
    if (!validEmail_(email)) {
      throw new Error(
        "管理員 Email 格式不正確。"
      );
    }
  });
}


function locked_(work) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(20000)) {
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


function getCalendar_() {
  if (CONFIG.CALENDAR_ID) {
    const calendar = CalendarApp.getCalendarById(
      CONFIG.CALENDAR_ID
    );

    if (!calendar) {
      throw new Error(
        "找不到設定的日曆或沒有存取權限。"
      );
    }

    return calendar;
  }

  const list = CalendarApp.getCalendarsByName(
    CONFIG.CALENDAR_NAME
  );

  if (list.length !== 1) {
    throw new Error(
      "指定名稱的日曆不是唯一一個，請確認名稱或設定 CALENDAR_ID。"
    );
  }

  return list[0];
}


// ============================================================
// 表單提交：只做預檢，不改班表
// ============================================================

function onFormSubmitPrecheck(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (
    !isResponseSheet_(sheet) ||
    e.range.getRow() < 2
  ) {
    return;
  }

  locked_(() => {
    validateConfig_(sheet);

    const c = context_(
      sheet,
      e.range.getRow()
    );

    const st = state_(c);

    // 延遲或重複觸發，不得蓋掉已執行或已取消的結果。
    if (
      st ||
      text_(
        cell_(c, CONFIG.STATUS_COL).getValue()
      ).startsWith("已更新日曆")
    ) {
      return;
    }

    try {
      assertNoPending_(sheet, "");

      const analysis = analyzeRequest(
        sheet,
        getCalendar_(),
        row_(c)
      );

      if (!analysis.ok) {
        reject_(c, analysis, "預檢");

      } else {
        cell_(
          c,
          CONFIG.EXECUTE_CHECK_COL
        ).setValue(false);

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
  });
}


// ============================================================
// 編輯事件：同一把鎖，逐列 L 在前、N 在後
// ============================================================

function handleSheetEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (!isResponseSheet_(sheet)) return;

  const lo = e.range.getColumn();
  const hi = e.range.getLastColumn();

  const touchesExecute =
    lo <= CONFIG.EXECUTE_CHECK_COL &&
    hi >= CONFIG.EXECUTE_CHECK_COL;

  const touchesMail =
    lo <= CONFIG.APPROVE_CHECK_COL &&
    hi >= CONFIG.APPROVE_CHECK_COL;

  if (!touchesExecute && !touchesMail) return;

  locked_(() => {
    validateConfig_(sheet);

    const first = Math.max(
      2,
      e.range.getRow()
    );

    const last = Math.min(
      e.range.getLastRow(),
      sheet.getLastRow()
    );

    const deadline = Date.now() + CONFIG.RUN_MS;

    let calendar = null;

    for (let r = first; r <= last; r++) {
      if (
        r - first >= CONFIG.MAX_ROWS ||
        Date.now() > deadline - 30000
      ) {
        sheet.getParent().toast(
          "部分列尚未處理，請分批重新操作；未覆寫既有狀態。"
        );

        break;
      }

      if (
        !text_(
          sheet.getRange(r, 2).getValue()
        )
      ) {
        continue;
      }

      let c;

      try {
        c = context_(sheet, r);

        if (touchesExecute) {
          calendar = calendar || getCalendar_();

          processSingleRow_(
            c,
            calendar
          );
        }

        // L 執行完，才處理同列 N。
        if (
          touchesMail &&
          cell_(
            c,
            CONFIG.APPROVE_CHECK_COL
          ).getValue() === true
        ) {
          approval_(
            c,
            deadline
          );
        }

      } catch (err) {
        console.error(err);

        // 不把已成功狀態改成失敗。
        // 操作錯誤寫入 L 儲存格註解。
        sheet.getRange(
          r,
          CONFIG.EXECUTE_CHECK_COL
        ).setNote(
          "本次未完成：" + err.message
        );

        sheet.getParent().toast(
          "第 " + r + " 列：" + err.message
        );
      }
    }
  });
}


// ============================================================
// 共用工具：UUID、狀態、快照
// ============================================================

function text_(value) {
  return value == null
    ? ""
    : String(value).trim();
}


function clone_(value) {
  return JSON.parse(
    JSON.stringify(value)
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


function readNote_(cell, kind) {
  const text = cell.getNote();

  if (!text) return null;

  let value;

  try {
    value = JSON.parse(text);

  } catch (_) {
    throw new Error(
      "系統註解不是有效的 JSON，請勿覆寫或刪除原紀錄。"
    );
  }

  if (value.kind !== kind) {
    throw new Error(
      "系統註解類型不符，已停止以保護紀錄。"
    );
  }

  return value;
}


function writeNote_(cell, value) {
  const text = JSON.stringify(value);

  if (text.length > CONFIG.NOTE_LIMIT) {
    throw new Error(
      "快照／通知紀錄過大，請管理員封存處理。"
    );
  }

  cell.setNote(text);

  SpreadsheetApp.flush();
}


function context_(sheet, r) {
  const q = sheet.getRange(
    r,
    CONFIG.REQUEST_ID_COL
  );

  let id = text_(q.getValue());

  const oldStatus = text_(
    sheet.getRange(
      r,
      CONFIG.STATUS_COL
    ).getValue()
  );

  if (!id) {
    if (oldStatus.includes("已更新日曆")) {
      throw new Error(
        "此舊案件沒有 UUID，不能猜測舊列號還原；請人工核對日曆。"
      );
    }

    id = Utilities.getUuid();

    q.setValue(id);

    SpreadsheetApp.flush();
  }

  const c = {
    s: sheet,
    r: r,
    id: id
  };

  // 同時檢查 UUID 是否重複。
  row_(c);

  return c;
}


function row_(c) {
  const lastRow = c.s.getLastRow();

  const values = lastRow > 1
    ? c.s.getRange(
        2,
        CONFIG.REQUEST_ID_COL,
        lastRow - 1,
        1
      ).getValues()
    : [];

  const matches = [];

  values.forEach((value, index) => {
    if (text_(value[0]) === c.id) {
      matches.push(index + 2);
    }
  });

  if (matches.length !== 1) {
    throw new Error(
      "申請 ID 遺失或重複，請勿複製 Q 欄。"
    );
  }

  c.r = matches[0];

  return c.r;
}


function cell_(c, col) {
  return c.s.getRange(
    row_(c),
    col
  );
}


function status_(c, message) {
  cell_(
    c,
    CONFIG.STATUS_COL
  ).setValue(message);
}


function state_(c) {
  const st = readNote_(
    cell_(
      c,
      CONFIG.REQUEST_ID_COL
    ),
    "shift-tx"
  );

  if (st && st.id !== c.id) {
    throw new Error(
      "申請 ID 與快照不一致。"
    );
  }

  return st;
}


function saveState_(c, st) {
  st.updated = now_();

  writeNote_(
    cell_(
      c,
      CONFIG.REQUEST_ID_COL
    ),
    st
  );
}


// 若有其他未完成的日曆異動，暫停新的正式異動。
function assertNoPending_(sheet, ownId) {
  if (sheet.getLastRow() < 2) return;

  const notes = sheet.getRange(
    2,
    CONFIG.REQUEST_ID_COL,
    sheet.getLastRow() - 1,
    1
  ).getNotes();

  notes.forEach((entry, index) => {
    if (!entry[0]) return;

    let st;

    try {
      st = JSON.parse(entry[0]);

    } catch (_) {
      throw new Error(
        "第 " +
        (index + 2) +
        " 列 Q 註解異常，請先核對。"
      );
    }

    if (
      st.kind === "shift-tx" &&
      OPEN_PHASES.includes(st.phase) &&
      st.id !== ownId
    ) {
      throw new Error(
        "第 " +
        (index + 2) +
        " 列異動尚待復原，暫停其他班表異動。"
      );
    }
  });
}


// ============================================================
// 讀取申請與嚴格時間解析
// ============================================================

function hasRealTargetPerson(value) {
  value = text_(value);

  return (
    !!value &&
    value !== "無" &&
    !value.startsWith("請假")
  );
}


function readRequest_(sheet, r) {
  const values = sheet.getRange(
    r,
    1,
    1,
    11
  ).getValues()[0];

  const display = sheet.getRange(
    r,
    1,
    1,
    11
  ).getDisplayValues()[0];

  const date = value => value instanceof Date
    ? Utilities.formatDate(
        value,
        CONFIG.TIME_ZONE,
        "yyyy/MM/dd"
      )
    : text_(value);

  return {
    person: text_(values[1]),
    date: date(values[2]),
    start: text_(display[3]),
    end: text_(display[4]),
    target: text_(values[5]),
    swapDate: date(values[6]),
    swapStart: text_(display[7]),
    swapEnd: text_(display[8]),
    memo: text_(values[9])
  };
}


function parseTime_(date, time) {
  const dateMatch = text_(date).match(
    /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/
  );

  const timeMatch = text_(time)
    .replace(/：/g, ":")
    .match(
      /^(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
    );

  if (!dateMatch || !timeMatch) {
    throw new Error(
      "日期或時間格式錯誤，請用日期及 HH:mm。"
    );
  }

  let hour = Number(timeMatch[2]);
  const minute = Number(timeMatch[3]);
  const second = Number(timeMatch[4] || 0);

  const ampm = text_(
    timeMatch[1] || timeMatch[5]
  ).toUpperCase();

  if (
    hour > 23 ||
    minute > 59 ||
    second > 59 ||
    (ampm && (hour < 1 || hour > 12))
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

  const pad = value =>
    String(Number(value)).padStart(2, "0");

  const input =
    `${dateMatch[1]}/${pad(dateMatch[2])}/${pad(dateMatch[3])} ` +
    `${pad(hour)}:${pad(minute)}:${pad(second)}`;

  const result = Utilities.parseDate(
    input,
    CONFIG.TIME_ZONE,
    "yyyy/MM/dd HH:mm:ss"
  );

  if (
    Utilities.formatDate(
      result,
      CONFIG.TIME_ZONE,
      "yyyy/MM/dd HH:mm:ss"
    ) !== input
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

function titleInfo_(title) {
  let raw = text_(title);

  const tag = raw.match(
    /\s*\[(換班|代班|請假)\]\s*$/
  );

  const state = tag ? tag[1] : "";

  if (tag) {
    raw = raw.slice(
      0,
      tag.index
    ).trim();
  }

  const floor = raw.match(
    /(\([^)]*\)|（[^）]*）)$/
  );

  return {
    name: (
      floor
        ? raw.slice(0, floor.index)
        : raw
    ).trim(),

    floor: floor
      ? floor[0]
      : "",

    state: state
  };
}


function overlap_(start1, end1, start2, end2) {
  return (
    start1 < end2 &&
    start2 < end1
  );
}


// ============================================================
// 共用檢核：每次重新查當下 Calendar
// ============================================================

function analyzeRequest(sheet, calendar, r) {
  const req = readRequest_(
    sheet,
    r
  );

  const reject = (code, message) => ({
    ok: false,
    code: code,
    message: message,
    req: req
  });

  const hasTarget = hasRealTargetPerson(
    req.target
  );

  const anySwap = !!(
    req.swapDate ||
    req.swapStart ||
    req.swapEnd
  );

  const allSwap = !!(
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
    req.person === req.target
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
    originalStart = parseTime_(
      req.date,
      req.start
    );

    originalEnd = parseTime_(
      req.date,
      req.end
    );

    if (!(originalStart < originalEnd)) {
      throw new Error(
        "原班結束時間必須晚於開始時間，不支援跨日填法。"
      );
    }

    if (hasTarget && allSwap) {
      swapStart = parseTime_(
        req.swapDate,
        req.swapStart
      );

      swapEnd = parseTime_(
        req.swapDate,
        req.swapEnd
      );

      if (!(swapStart < swapEnd)) {
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

  const findPerson = (
    events,
    name,
    start,
    end
  ) => events.filter(event =>
    titleInfo_(
      event.getTitle()
    ).name === name &&
    overlap_(
      +start,
      +end,
      +event.getStartTime(),
      +event.getEndTime()
    )
  );

  const originalEvents = calendar.getEvents(
    originalStart,
    originalEnd
  );

  const originalMatches = findPerson(
    originalEvents,
    req.person,
    originalStart,
    originalEnd
  );

  if (originalMatches.length !== 1) {
    return reject(
      "ORIGINAL",
      originalMatches.length
        ? req.person + " 同時段有多筆行程，請管理員確認。"
        : "找不到 " + req.person + " 的原值班行程。"
    );
  }

  const contained = (
    event,
    start,
    end
  ) => (
    !event.isAllDayEvent() &&
    +start >= +event.getStartTime() &&
    +end <= +event.getEndTime()
  );

  const originalEvent = originalMatches[0];

  if (
    !contained(
      originalEvent,
      originalStart,
      originalEnd
    )
  ) {
    return reject(
      "RANGE",
      "申請時段超出原班，或原班是全天事件。"
    );
  }

  const type = hasTarget
    ? (allSwap ? "swap" : "sub")
    : "leave";

  let swapEvent = null;

  if (type === "swap") {
    const targetEvents = calendar.getEvents(
      swapStart,
      swapEnd
    );

    const targetMatches = findPerson(
      targetEvents,
      req.target,
      swapStart,
      swapEnd
    );

    if (targetMatches.length !== 1) {
      return reject(
        "TARGET",
        "配合人互換時段的班表不存在或不唯一。"
      );
    }

    swapEvent = targetMatches[0];

    if (
      !contained(
        swapEvent,
        swapStart,
        swapEnd
      )
    ) {
      return reject(
        "TARGET_RANGE",
        "互換時段超出配合人原班。"
      );
    }

    const targetConflict = hasConflict_(
      originalEvents,
      req.target,
      originalStart,
      originalEnd,
      swapEvent,
      swapStart,
      swapEnd
    );

    const originalConflict = hasConflict_(
      targetEvents,
      req.person,
      swapStart,
      swapEnd,
      originalEvent,
      originalStart,
      originalEnd
    );

    if (
      targetConflict ||
      originalConflict
    ) {
      return reject(
        "CONFLICT_SWAP",
        "雙向換班後與既有值班／請假紀錄或殘餘時段衝突。"
      );
    }

  } else if (
    hasTarget &&
    hasConflict_(
      originalEvents,
      req.target,
      originalStart,
      originalEnd,
      null,
      null,
      null
    )
  ) {
    return reject(
      "CONFLICT_SUB",
      req.target + " 在代班時段已有值班／請假紀錄。"
    );
  }

  return {
    ok: true,
    req: req,
    type: type,
    orig: originalEvent,
    swap: swapEvent,

    os: +originalStart,
    oe: +originalEnd,

    ss: swapStart
      ? +swapStart
      : null,

    se: swapEnd
      ? +swapEnd
      : null,

    late:
      type === "sub" &&
      titleInfo_(
        originalEvent.getTitle()
      ).state === "請假"
  };
}


function hasConflict_(
  events,
  name,
  start,
  end,
  cedingEvent,
  cededStart,
  cededEnd
) {
  return events.some(event => {
    if (
      titleInfo_(
        event.getTitle()
      ).name !== name
    ) {
      return false;
    }

    const eventStart = +event.getStartTime();
    const eventEnd = +event.getEndTime();

    // 重複行程的 iCalUID 可能相同，
    // 所以還要比對該次行程的開始／結束時間。
    const sameEvent = (
      cedingEvent &&
      event.getId() === cedingEvent.getId() &&
      eventStart === +cedingEvent.getStartTime() &&
      eventEnd === +cedingEvent.getEndTime()
    );

    if (!sameEvent) {
      return overlap_(
        +start,
        +end,
        eventStart,
        eventEnd
      );
    }

    const frontConflict = (
      eventStart < +cededStart &&
      overlap_(
        +start,
        +end,
        eventStart,
        +cededStart
      )
    );

    const backConflict = (
      +cededEnd < eventEnd &&
      overlap_(
        +start,
        +end,
        +cededEnd,
        eventEnd
      )
    );

    return frontConflict || backConflict;
  });
}


// ============================================================
// 中繼資料、快照與修改計畫
// ============================================================

function enc_(text) {
  return Utilities.base64EncodeWebSafe(
    text,
    Utilities.Charset.UTF_8
  );
}


function dec_(text) {
  return Utilities.newBlob(
    Utilities.base64DecodeWebSafe(text)
  ).getDataAsString("UTF-8");
}


function meta_(description, key) {
  const match = String(
    description || ""
  ).match(
    new RegExp(
      "^\\[" +
      key +
      ":([^\\]\\r\\n]*)\\]\\s*$",
      "m"
    )
  );

  return match ? match[1] : null;
}


// 找出目前事件依賴的前序申請。
// 同時讀取 v6.2 的 Base64 快照鏈。
function owners_(description, depth) {
  depth = depth || 0;

  if (depth > 20) {
    throw new Error(
      "舊快照層數過多，請人工核對。"
    );
  }

  const result = [];

  const own = meta_(
    description,
    "ROW_ID"
  );

  if (own) result.push(own);

  const parents = meta_(
    description,
    "ANCESTORS_B64"
  );

  if (parents) {
    result.push(
      ...JSON.parse(
        dec_(parents)
      )
    );
  }

  const previous = meta_(
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

  return [...new Set(result)];
}


function humanDesc_(description) {
  return String(description || "")
    .replace(
      /^\[(?:ROW_ID|ANCESTORS_B64|PART_ID|ORIG_TITLE_B64|ORIG_DESC_B64|SPLIT_ORIG_TIME):[^\]\r\n]*\]\r?\n?/gm,
      ""
    )
    .replace(
      /^\[AUTO_SPLIT_CREATED\]\r?\n?/gm,
      ""
    )
    .trim();
}


function snap_(event) {
  return {
    id: event.getId(),
    s: +event.getStartTime(),
    e: +event.getEndTime(),
    title: event.getTitle(),
    desc: event.getDescription() || "",
    location: event.getLocation() || ""
  };
}


function same_(a, b) {
  return (
    a.s === b.s &&
    a.e === b.e &&
    a.title === b.title &&
    a.desc === b.desc &&
    a.location === b.location
  );
}


// 中途 setter 失敗時，各欄位可能停在修改前或修改後。
// 只允許本次修改可解釋的組合，避免覆寫外部手動修改。
function compatible_(current, before, after) {
  return (
    [before.title, after.title].includes(current.title) &&
    [before.desc, after.desc].includes(current.desc) &&
    (
      (
        current.s === before.s &&
        current.e === before.e
      ) ||
      (
        current.s === after.s &&
        current.e === after.e
      )
    ) &&
    current.location === before.location
  );
}


function putSnapshot_(event, target) {
  if (
    +event.getStartTime() !== target.s ||
    +event.getEndTime() !== target.e
  ) {
    event.setTime(
      new Date(target.s),
      new Date(target.e)
    );
  }

  if (event.getTitle() !== target.title) {
    event.setTitle(target.title);
  }

  // 說明最後寫，保留控制標記直到其他屬性完成。
  if (
    (event.getDescription() || "") !== target.desc
  ) {
    event.setDescription(target.desc);
  }
}


function plan_(
  event,
  start,
  end,
  worker,
  tag,
  token,
  detail
) {
  const before = snap_(event);
  const after = clone_(before);

  const ancestry = owners_(before.desc);

  const header =
    `[ROW_ID:${token}]\n` +
    `[ANCESTORS_B64:${enc_(JSON.stringify(ancestry))}]`;

  const human = humanDesc_(before.desc);

  const title =
    `${worker}${titleInfo_(before.title).floor} [${tag}]`;

  const additions = [];

  const add = (
    partStart,
    partEnd,
    partTitle,
    body
  ) => {
    const key =
      token + ":" + (additions.length + 1);

    additions.push({
      key: key,
      stage: "NEW",
      id: "",

      want: {
        s: partStart,
        e: partEnd,
        title: partTitle,
        location: before.location,

        desc:
          `[AUTO_SPLIT_CREATED]\n` +
          `${header}\n` +
          `[PART_ID:${key}]\n` +
          body
      }
    });

    additions[
      additions.length - 1
    ].want.desc = additions[
      additions.length - 1
    ].want.desc.trim();
  };

  // 整段異動
  if (
    start === before.s &&
    end === before.e
  ) {
    after.title = title;

    after.desc =
      `${header}\n` +
      `${detail}\n` +
      `--------------------\n` +
      human;

    after.desc = after.desc.trim();

  } else {
    after.desc = `${header}\n${human}`.trim();

    // 切中間或後半
    if (start > before.s) {
      after.e = start;

      add(
        start,
        end,
        title,
        detail + "\n" + human
      );

      if (end < before.e) {
        add(
          end,
          before.e,
          before.title,
          human
        );
      }

    } else {
      // 切前半
      after.s = end;

      add(
        start,
        end,
        title,
        detail + "\n" + human
      );
    }
  }

  return {
    token: token,
    before: before,
    after: after,
    additions: additions,
    touched: false,
    restored: false
  };
}


function appliedText_(st) {
  const label =
    st.type === "swap"
      ? "雙向換班"
      : st.type === "leave"
        ? "請假"
        : st.late
          ? "後補代班"
          : "代班";

  return "已更新日曆（" + label + "完成）";
}


function newTransaction_(c, calendar, analysis) {
  const tx = Utilities.getUuid();

  const base = c.id + "-" + tx;

  const tag =
    analysis.type === "leave"
      ? "請假"
      : analysis.type === "swap"
        ? "換班"
        : "代班";

  const firstWorker =
    analysis.type === "leave"
      ? analysis.req.person
      : analysis.req.target;

  const plans = [
    plan_(
      analysis.orig,
      analysis.os,
      analysis.oe,
      firstWorker,
      tag,
      base + "-A",

      `【${tag}紀錄】\n` +
      `- 原定值班：${analysis.req.person}\n` +
      `- 實際安排：${firstWorker}\n` +
      `- 申請 ID：${c.id}`
    )
  ];

  if (analysis.type === "swap") {
    plans.push(
      plan_(
        analysis.swap,
        analysis.ss,
        analysis.se,
        analysis.req.person,
        tag,
        base + "-B",

        `【換班紀錄】\n` +
        `- 原定值班：${analysis.req.target}\n` +
        `- 實際到勤：${analysis.req.person}\n` +
        `- 申請 ID：${c.id}`
      )
    );
  }

  return {
    kind: "shift-tx",
    version: CONFIG.VERSION,

    id: c.id,
    tx: tx,

    phase: "APPLYING",
    calendar: calendar.getId(),

    req: analysis.req,
    requestHash: hash_(analysis.req),

    type: analysis.type,
    late: analysis.late,

    plans: plans,

    purpose: "rollback",
    error: ""
  };
}


// ============================================================
// 正式異動與失敗補償
// ============================================================

function processSingleRow_(c, calendar) {
  let st = state_(c);

  const checked = cell_(
    c,
    CONFIG.EXECUTE_CHECK_COL
  ).getValue() === true;

  if (
    st &&
    OPEN_PHASES.includes(st.phase)
  ) {
    throw new Error(
      "此列有中斷異動，請用「班表工具→復原此列中斷的班表異動」。"
    );
  }

  assertNoPending_(
    c.s,
    c.id
  );

  const oldStatus = text_(
    cell_(
      c,
      CONFIG.STATUS_COL
    ).getValue()
  );

  // 已執行的舊版案件：只有取消時才嘗試安全匯入。
  if (
    !st &&
    oldStatus.startsWith("已更新日曆")
  ) {
    if (checked) return;

    try {
      st = adoptLegacy_(
        c,
        calendar
      );

    } catch (err) {
      cell_(
        c,
        CONFIG.EXECUTE_CHECK_COL
      ).setValue(true);

      throw err;
    }
  }

  // 已執行成功
  if (
    st &&
    st.phase === ACTIVE_PHASE
  ) {
    if (checked) {
      status_(
        c,
        appliedText_(st)
      );

      return;
    }

    // 取消 L：先完整檢查，通過才做任何刪除／還原。
    try {
      const events = eventsForTx_(
        calendar,
        st
      );

      preflightUndo_(
        st,
        events
      );

      st.phase = "UNDOING";
      st.purpose = "undo";

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
      if (st.phase === "APPLIED") {
        cell_(
          c,
          CONFIG.EXECUTE_CHECK_COL
        ).setValue(true);

        cell_(
          c,
          CONFIG.EXECUTE_CHECK_COL
        ).setNote(
          "取消未執行：" + err.message
        );

        status_(
          c,
          appliedText_(st) +
          "；取消未執行：" +
          err.message
        );

      } else {
        markRecovery_(
          c,
          st,
          err
        );
      }
    }

    return;
  }

  if (!checked) return;

  // 已成功後又取消的案件，不重用原列。
  if (
    st &&
    st.phase === "UNDONE"
  ) {
    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).setValue(false);

    throw new Error(
      "這筆已取消並保留歷史紀錄；再次申請請另填一筆表單。"
    );
  }

  const analysis = analyzeRequest(
    c.s,
    calendar,
    row_(c)
  );

  if (!analysis.ok) {
    reject_(
      c,
      analysis,
      "正式審核"
    );

    return;
  }

  st = newTransaction_(
    c,
    calendar,
    analysis
  );

  // 必須先成功保存快照，才開始改 Calendar。
  saveState_(
    c,
    st
  );

  const references = [
    analysis.orig,
    analysis.swap
  ];

  try {
    for (
      let index = 0;
      index < st.plans.length;
      index++
    ) {
      const p = st.plans[index];
      const event = references[index];

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

      p.touched = true;

      saveState_(
        c,
        st
      );

      putSnapshot_(
        event,
        p.after
      );

      for (const addition of p.additions) {
        addition.stage = "CREATING";

        saveState_(
          c,
          st
        );

        let created;

        try {
          created = calendar.createEvent(
            addition.want.title,
            new Date(addition.want.s),
            new Date(addition.want.e),
            {
              description: addition.want.desc,
              location: addition.want.location
            }
          );

        } catch (err) {
          // 明確被拒絕的呼叫，沒有建立事件。
          // 其他錯誤保留 CREATING，等待核對。
          const definitelyRejected =
            /permission|authorization|quota|too many times|invalid (argument|date|time)/i
              .test(err.message);

          if (definitelyRejected) {
            addition.stage = "NEW";

            saveState_(
              c,
              st
            );
          }

          throw err;
        }

        addition.id = created.getId();
        addition.stage = "CREATED";

        saveState_(
          c,
          st
        );
      }
    }

    st.phase = ACTIVE_PHASE;

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
    ).setNote("");

  } catch (err) {
    st.error = err.message;
    st.purpose = "rollback";
    st.phase = "RECOVER";

    try {
      saveState_(
        c,
        st
      );

      restoreTransaction_(
        c,
        calendar,
        st,
        eventsForTx_(calendar, st),
        false
      );

      finishRestore_(
        c,
        st
      );

    } catch (recoveryError) {
      markRecovery_(
        c,
        st,
        recoveryError
      );
    }
  }
}


// ============================================================
// 還原前檢查與復原
// ============================================================

function eventsForTx_(calendar, st) {
  if (calendar.getId() !== st.calendar) {
    throw new Error(
      "日曆 ID 與原異動紀錄不同。"
    );
  }

  const points = st.plans.flatMap(p => [
    p.before.s,
    p.before.e,

    ...p.additions.flatMap(addition => [
      addition.want.s,
      addition.want.e
    ])
  ]);

  return calendar.getEvents(
    new Date(
      Math.min(...points) - 86400000
    ),
    new Date(
      Math.max(...points) + 86400000
    )
  );
}


function mainEvent_(p, events) {
  const candidates = events.filter(event =>
    event.getId() === p.before.id &&
    (
      (
        +event.getStartTime() === p.before.s ||
        +event.getStartTime() === p.after.s
      ) ||
      meta_(
        event.getDescription(),
        "ROW_ID"
      ) === p.token
    )
  );

  if (candidates.length !== 1) {
    throw new Error(
      "原主事件不存在或不唯一，不能安全還原。"
    );
  }

  return candidates[0];
}


function addedEvent_(addition, events) {
  const candidates = events.filter(event =>
    meta_(
      event.getDescription(),
      "PART_ID"
    ) === addition.key ||
    (
      addition.id &&
      event.getId() === addition.id &&
      +event.getStartTime() === addition.want.s
    )
  );

  if (candidates.length > 1) {
    throw new Error(
      "新增事件識別碼重複，請人工核對。"
    );
  }

  return candidates[0] || null;
}


// 有後續依賴時，不允許先取消原申請。
function checkDependents_(tokens, events) {
  for (const event of events) {
    const description =
      event.getDescription() || "";

    const top = meta_(
      description,
      "ROW_ID"
    );

    if (
      top &&
      !tokens.includes(top) &&
      owners_(
        description
      ).some(token => tokens.includes(token))
    ) {
      throw new Error(
        "仍有後補代班／後續異動「" +
        event.getTitle() +
        "」，請先取消後續申請。"
      );
    }
  }
}


// 還原也要避免與後來新增的其他班表撞班。
function checkRestoreConflicts_(st, events) {
  const tokens = st.plans.map(
    p => p.token
  );

  const outside = events.filter(event => {
    if (
      tokens.includes(
        meta_(
          event.getDescription(),
          "ROW_ID"
        )
      )
    ) {
      return false;
    }

    return !st.plans.some(p =>
      (
        event.getId() === p.before.id &&
        [
          +p.before.s,
          +p.after.s
        ].includes(
          +event.getStartTime()
        )
      ) ||
      p.additions.some(addition =>
        addition.id &&
        event.getId() === addition.id &&
        +event.getStartTime() === addition.want.s
      )
    );
  });

  for (const p of st.plans) {
    const name = titleInfo_(
      p.before.title
    ).name;

    const conflict = outside.some(event =>
      titleInfo_(
        event.getTitle()
      ).name === name &&
      overlap_(
        p.before.s,
        p.before.e,
        +event.getStartTime(),
        +event.getEndTime()
      )
    );

    if (conflict) {
      throw new Error(
        "還原後 " +
        name +
        " 會與另一筆值班／請假紀錄衝突，請先協調。"
      );
    }
  }
}


// 確認本次異動目前仍完整存在，未被後續操作改掉。
function assertAppliedIntact_(st, events) {
  checkDependents_(
    st.plans.map(p => p.token),
    events
  );

  for (const p of st.plans) {
    if (
      !same_(
        snap_(
          mainEvent_(p, events)
        ),
        p.after
      )
    ) {
      throw new Error(
        "主事件已被其他操作修改，取消已攔阻。"
      );
    }

    for (const addition of p.additions) {
      const event = addedEvent_(
        addition,
        events
      );

      if (
        !event ||
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


function preflightUndo_(st, events) {
  assertAppliedIntact_(
    st,
    events
  );

  checkRestoreConflicts_(
    st,
    events
  );
}


function restoreTransaction_(
  c,
  calendar,
  st,
  events,
  confirmed
) {
  checkDependents_(
    st.plans.map(p => p.token),
    events
  );

  checkRestoreConflicts_(
    st,
    events
  );

  const errors = [];

  // 雙向換班從後處理的一邊開始復原。
  for (const p of [...st.plans].reverse()) {
    if (
      !p.touched ||
      p.restored
    ) {
      continue;
    }

    try {
      const main = mainEvent_(
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

      // 先刪衍生事件，再恢復主事件。
      for (
        const addition of [...p.additions].reverse()
      ) {
        if (
          addition.stage === "NEW" ||
          addition.stage === "DELETED"
        ) {
          continue;
        }

        const event = addedEvent_(
          addition,
          events
        );

        if (!event) {
          if (
            addition.stage === "CREATING" &&
            !addition.id &&
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

          addition.stage = "DELETING";

          saveState_(
            c,
            st
          );

          event.deleteEvent();

          // 從同一批清單移除，
          // 後續不再讀取或修改已刪除物件。
          events.splice(
            events.indexOf(event),
            1
          );
        }

        addition.stage = "DELETED";

        saveState_(
          c,
          st
        );
      }

      putSnapshot_(
        main,
        p.before
      );

      p.restored = true;

      saveState_(
        c,
        st
      );

    } catch (err) {
      errors.push(err.message);
    }
  }

  if (errors.length) {
    throw new Error(
      errors.join("；")
    );
  }
}


function finishRestore_(c, st) {
  st.phase = st.purpose === "undo"
    ? "UNDONE"
    : "ROLLED_BACK";

  saveState_(
    c,
    st
  );

  cell_(
    c,
    CONFIG.EXECUTE_CHECK_COL
  ).setValue(false);

  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(false);

  const priorMail = text_(
    cell_(
      c,
      CONFIG.APPROVE_LOG_COL
    ).getValue()
  );

  if (st.phase === "UNDONE") {
    status_(
      c,
      "已還原班表" +
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


function markRecovery_(c, st, err) {
  st.phase = "RECOVER";

  st.error = [
    st.error,
    err.message
  ].filter(Boolean).join("；");

  try {
    saveState_(
      c,
      st
    );

  } catch (saveError) {
    console.error(saveError);
  }

  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(false);

  status_(
    c,
    "需人工確認：日曆異動／還原尚未完成；暫停其他異動。" +
    st.error
  );
}


// ============================================================
// 舊 v6.2 案件安全匯入
//
// 必須有完整快照，才允許自動還原。
// 沒有快照或沒有 UUID，就不猜測。
// ============================================================

function adoptLegacy_(c, calendar) {
  const req = readRequest_(
    c.s,
    row_(c)
  );

  const originalStart = +parseTime_(
    req.date,
    req.start
  );

  const originalEnd = +parseTime_(
    req.date,
    req.end
  );

  const isSwap = (
    hasRealTargetPerson(req.target) &&
    req.swapDate &&
    req.swapStart &&
    req.swapEnd
  );

  const windows = [{
    s: originalStart,
    e: originalEnd,
    person: req.person,
    token: c.id + "-A"
  }];

  if (isSwap) {
    windows.push({
      s: +parseTime_(
        req.swapDate,
        req.swapStart
      ),

      e: +parseTime_(
        req.swapDate,
        req.swapEnd
      ),

      person: req.target,
      token: c.id + "-B"
    });
  }

  const events = calendar.getEvents(
    new Date(
      Math.min(
        ...windows.map(w => w.s)
      ) - 86400000
    ),

    new Date(
      Math.max(
        ...windows.map(w => w.e)
      ) + 86400000
    )
  );

  checkDependents_(
    windows.map(w => w.token),
    events
  );

  const plans = windows.map(w => {
    const group = events.filter(event =>
      meta_(
        event.getDescription(),
        "ROW_ID"
      ) === w.token
    );

    const mainEvents = group.filter(event =>
      !/^\[AUTO_SPLIT_CREATED\]\s*$/m.test(
        event.getDescription()
      )
    );

    if (mainEvents.length !== 1) {
      throw new Error(
        "舊案件主事件不完整，請人工核對，不能自動還原。"
      );
    }

    const after = snap_(
      mainEvents[0]
    );

    const originalTitle = meta_(
      after.desc,
      "ORIG_TITLE_B64"
    );

    const originalDescription = meta_(
      after.desc,
      "ORIG_DESC_B64"
    );

    if (
      originalTitle === null ||
      originalDescription === null
    ) {
      throw new Error(
        "舊案件缺少完整狀態快照，請人工處理。"
      );
    }

    const before = clone_(after);

    before.title = dec_(
      originalTitle
    );

    before.desc = dec_(
      originalDescription
    );

    const split = meta_(
      after.desc,
      "SPLIT_ORIG_TIME"
    );

    if (split) {
      const parts = split
        .split("-")
        .map(Number);

      before.s = parts[0];
      before.e = parts[1];
    }

    if (
      titleInfo_(before.title).name !== w.person ||
      w.s < before.s ||
      w.e > before.e
    ) {
      throw new Error(
        "舊申請資料與日曆快照不一致。"
      );
    }

    const expected =
      Number(w.s > before.s) +
      Number(w.e < before.e);

    if (
      group.length !== 1 + expected
    ) {
      throw new Error(
        "舊案件衍生事件數量不符，已停止。"
      );
    }

    const additions = group
      .filter(event => event !== mainEvents[0])
      .map((event, index) => ({
        key: w.token + ":legacy:" + index,
        id: event.getId(),
        stage: "CREATED",
        want: snap_(event)
      }));

    return {
      token: w.token,
      before: before,
      after: after,
      additions: additions,
      touched: true,
      restored: false
    };
  });

  const st = {
    kind: "shift-tx",
    version: CONFIG.VERSION,

    id: c.id,
    tx: "legacy-" + c.id,

    phase: ACTIVE_PHASE,
    calendar: calendar.getId(),

    req: req,
    requestHash: hash_(req),

    type: isSwap
      ? "swap"
      : hasRealTargetPerson(req.target)
        ? "sub"
        : "leave",

    late: text_(
      cell_(
        c,
        CONFIG.STATUS_COL
      ).getValue()
    ).includes("後補代班"),

    plans: plans,
    purpose: "undo",
    error: ""
  };

  saveState_(
    c,
    st
  );

  return st;
}


// ============================================================
// Email 設定與收件人檢查
// ============================================================

function validEmail_(email) {
  return /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/
    .test(text_(email));
}


function getApplicantEmail(sheet, r) {
  const col = CONFIG.APPLICANT_EMAIL_COL;

  if (
    !Number.isInteger(col) ||
    col < 1 ||
    col > sheet.getMaxColumns()
  ) {
    throw new Error(
      "CONFIG.APPLICANT_EMAIL_COL 未正確設定。"
    );
  }

  const email = text_(
    sheet.getRange(
      r,
      col
    ).getValue()
  );

  if (!validEmail_(email)) {
    throw new Error(
      "K 欄申請人 Email 空白或格式錯誤。"
    );
  }

  return email;
}


function staffEmail_(sheet, name) {
  const directory = sheet.getParent()
    .getSheetByName(
      CONFIG.STAFF_DIRECTORY_SHEET_NAME
    );

  if (
    !directory ||
    directory.getLastRow() < 2
  ) {
    throw new Error(
      "員工名冊不存在或沒有資料。"
    );
  }

  const data = directory.getRange(
    2,
    1,
    directory.getLastRow() - 1,
    Math.max(
      CONFIG.STAFF_NAME_COL,
      CONFIG.STAFF_EMAIL_COL
    )
  ).getValues();

  const matches = data.filter(entry =>
    text_(
      entry[CONFIG.STAFF_NAME_COL - 1]
    ) === name
  );

  if (matches.length !== 1) {
    throw new Error(
      "員工名冊姓名「" +
      name +
      "」不存在或重複。"
    );
  }

  const email = text_(
    matches[0][
      CONFIG.STAFF_EMAIL_COL - 1
    ]
  );

  if (!validEmail_(email)) {
    throw new Error(
      "名冊「" +
      name +
      "」Email 格式不正確。"
    );
  }

  return email;
}


// ============================================================
// 通知紀錄
// ============================================================

function noticeStore_(c, col) {
  const store = readNote_(
    cell_(c, col),
    "shift-mail"
  ) || {
    kind: "shift-mail",
    id: c.id,
    batches: {},
    latest: ""
  };

  if (store.id !== c.id) {
    throw new Error(
      "通知註解屬於另一筆申請，已停止。"
    );
  }

  return store;
}


function saveMail_(c, col, store) {
  if (store.id !== c.id) {
    throw new Error(
      "通知紀錄與申請 ID 不一致。"
    );
  }

  writeNote_(
    cell_(c, col),
    store
  );
}


function task_(
  role,
  name,
  subject,
  body
) {
  return {
    role: role,
    name: name,
    subject: subject,
    body: body,

    status: "READY",
    email: "",
    at: "",
    error: ""
  };
}


// ============================================================
// 核准通知
// ============================================================

function approval_(c, deadline) {
  assertNoPending_(
    c.s,
    ""
  );

  const st = state_(c);

  const currentStatus = text_(
    cell_(
      c,
      CONFIG.STATUS_COL
    ).getValue()
  );

  if (
    (
      !st &&
      !currentStatus.startsWith("已更新日曆")
    ) ||
    (
      st &&
      st.phase !== ACTIVE_PHASE
    ) ||
    cell_(
      c,
      CONFIG.EXECUTE_CHECK_COL
    ).getValue() !== true
  ) {
    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(false);

    throw new Error(
      "班表尚未成功異動或已取消，不能寄核准信。"
    );
  }

  const req = st
    ? st.req
    : readRequest_(
        c.s,
        row_(c)
      );

  if (
    st &&
    hash_(
      readRequest_(
        c.s,
        row_(c)
      )
    ) !== st.requestHash
  ) {
    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(false);

    throw new Error(
      "審核後申請內容已被修改，不能寄送與日曆不一致的通知。"
    );
  }

  // 防止後續班表已經改動，卻補寄過期核准內容。
  if (st) {
    try {
      assertAppliedIntact_(
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
      ).setValue(false);

      throw new Error(
        "班表已有後續變動，不寄出過期核准通知：" +
        err.message
      );
    }
  }

  const col = CONFIG.APPROVE_LOG_COL;

  const store = noticeStore_(
    c,
    col
  );

  const key = st
    ? st.tx
    : "legacy-" + c.id;

  const legacyText = text_(
    cell_(c, col).getValue()
  );

  // 舊版已有完整成功紀錄：不再重寄。
  if (
    !st &&
    !store.latest &&
    legacyText.startsWith("已寄送核准通知")
  ) {
    return;
  }

  if (!store.batches[key]) {
    const isSwap = st
      ? st.type === "swap"
      : (
          hasRealTargetPerson(req.target) &&
          !!req.swapDate
        );

    const isSub = (
      hasRealTargetPerson(req.target) &&
      !isSwap
    );

    const label = isSwap
      ? "換班"
      : isSub
        ? (
            st && st.late
              ? "後補代班"
              : "代班"
          )
        : "請假";

    const original =
      `${req.date} ${req.start}–${req.end}`;

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
            ? `代班人員：${req.target}\n`
            : "此時段已核准請假，尚無代班人員。\n"
      ) +
      "\n班表已更新，請以最新值班日曆為準。";

    const tasks = [
      task_(
        "applicant",
        req.person,
        `【值班${label}申請】已審核通過`,
        applicantBody
      )
    ];

    if (isSwap || isSub) {
      const targetBody =
        `${req.target} 您好，\n\n` +
        `您配合 ${req.person} 的申請已通過。\n` +
        (
          isSwap
            ? (
                `原定時段：${req.swapDate} ` +
                `${req.swapStart}–${req.swapEnd}\n`
              )
            : ""
        ) +
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

    // 舊版只記「失敗」時，可能第一封其實已寄出。
    // 不猜測哪位收件人收到，先要求人工核對。
    if (
      !st &&
      /失敗/.test(legacyText)
    ) {
      tasks.forEach(task => {
        task.status = "UNKNOWN";

        task.error =
          "舊版紀錄無法確認此收件人是否已寄。";
      });
    }

    store.batches[key] = {
      label: "核准通知",
      tasks: tasks
    };

    store.latest = key;

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
// 系統檢核退件
// ============================================================

function reject_(c, analysis, stage) {
  cell_(
    c,
    CONFIG.EXECUTE_CHECK_COL
  ).setValue(false);

  cell_(
    c,
    CONFIG.APPROVE_CHECK_COL
  ).setValue(false);

  status_(
    c,
    stage +
    "未通過：" +
    analysis.message
  );

  const col = CONFIG.REJECT_LOG_COL;

  const store = noticeStore_(
    c,
    col
  );

  // 不把預檢／正式審核前綴放進識別碼。
  // 同一筆資料、同一理由不會只因階段不同而重寄。
  const key = hash_([
    analysis.req,
    analysis.code,
    analysis.message
  ]);

  if (!store.batches[key]) {
    store.batches[key] = {
      label: analysis.message,

      requestHash: hash_(
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
          `申請未通過，原因：\n${analysis.message}\n\n` +
          "請確認資料後重新提出申請，或洽管理員。"
        )
      ]
    };
  }

  store.latest = key;

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
    Date.now() + 60000
  );
}


// ============================================================
// 逐封寄送
//
// READY：尚未寄
// SENDING：開始交寄，但未寫入完成紀錄
// SENT：API 回傳成功並完成紀錄
// FAILED：明確未完成，可以重試
// UNKNOWN：結果不明，必須先人工核對
// ============================================================

function sendBatch_(
  c,
  col,
  store,
  key,
  deadline
) {
  const batch = store.batches[key];

  for (
    let index = 0;
    index < batch.tasks.length;
    index++
  ) {
    const task = batch.tasks[index];

    if (task.status === "SENT") continue;

    // 上次可能在寄出後中斷，不能直接重寄。
    if (task.status === "SENDING") {
      task.status = "UNKNOWN";

      task.error =
        "上次交寄未留下完成紀錄，請先人工核對。";

      saveMail_(
        c,
        col,
        store
      );
    }

    if (task.status === "UNKNOWN") continue;

    if (
      Date.now() > deadline - 10000
    ) {
      break;
    }

    try {
      task.email = task.role === "applicant"
        ? getApplicantEmail(
            c.s,
            row_(c)
          )
        : staffEmail_(
            c.s,
            task.name
          );

      const admins = [
        ...new Set(CONFIG.ADMIN_EMAILS)
      ].filter(email =>
        email.toLowerCase() !==
        task.email.toLowerCase()
      );

      if (
        MailApp.getRemainingDailyQuota() <
        1 + admins.length
      ) {
        throw new Error(
          "本日寄信額度不足，尚未寄出。"
        );
      }

      task.status = "SENDING";
      task.error = "";

      // 先保存交寄中狀態，再呼叫 MailApp。
      saveMail_(
        c,
        col,
        store
      );

      const message = {
        to: task.email,
        subject: task.subject,

        body:
          task.body +
          `\n\n申請 ID：${c.id}` +
          `\n通知識別：${key}:${index}` +
          "\n（系統通知，請勿直接回覆）"
      };

      if (admins.length) {
        message.bcc = admins.join(",");
      }

      try {
        MailApp.sendEmail(message);

      } catch (err) {
        // 明確的授權／額度／收件地址拒絕才直接允許重試。
        // 一般連線錯誤視為結果不明。
        const definitelyNotSent =
          /permission|authorization|required permissions|quota|too many times|invalid.*(email|recipient)|無權限|授權|額度/i
            .test(err.message);

        task.status = definitelyNotSent
          ? "FAILED"
          : "UNKNOWN";

        task.error = err.message;

        saveMail_(
          c,
          col,
          store
        );

        continue;
      }

      task.status = "SENT";
      task.at = now_();

      // 每一封立即記錄，不等另一位收件人寄完。
      saveMail_(
        c,
        col,
        store
      );

    } catch (err) {
      // API 成功但寫入紀錄失敗，也不能當成未寄。
      if (
        task.status === "SENDING" ||
        task.status === "SENT"
      ) {
        task.status = "UNKNOWN";

      } else {
        task.status = "FAILED";
      }

      task.error = err.message;

      try {
        saveMail_(
          c,
          col,
          store
        );

      } catch (saveError) {
        console.error(saveError);
      }
    }
  }

  const done = batch.tasks.every(
    task => task.status === "SENT"
  );

  const labels = {
    SENT: "已寄送",
    READY: "尚未寄",
    FAILED: "未完成，可重試",
    SENDING: "寄送結果待核對",
    UNKNOWN: "結果不明，須人工核對"
  };

  const heading =
    col === CONFIG.APPROVE_LOG_COL
      ? (
          done
            ? "已寄送核准通知"
            : "核准通知尚未全數完成"
        )
      : "退件原因：" + batch.label;

  const details = batch.tasks.map(task =>
    `${task.name}：${labels[task.status]} ` +
    `${task.at || ""}` +
    (
      task.error
        ? "；" + task.error
        : ""
    )
  ).join("\n");

  cell_(
    c,
    col
  ).setValue(
    heading + "\n" + details
  );

  if (col === CONFIG.APPROVE_LOG_COL) {
    cell_(
      c,
      CONFIG.APPROVE_CHECK_COL
    ).setValue(done);
  }
}


// ============================================================
// 人工工具：選取一列後操作，不必修改程式或刪除紀錄
// ============================================================

function selected_() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveSheet();

  const r = sheet
    .getActiveRange()
    .getRow();

  if (
    !isResponseSheet_(sheet) ||
    r < 2
  ) {
    throw new Error(
      "請先在表單回覆分頁選取一筆申請。"
    );
  }

  validateConfig_(sheet);

  return context_(
    sheet,
    r
  );
}


// ============================================================
// 重試尚未完成的通知
// ============================================================

function retrySelectedNotifications() {
  locked_(() => {
    const c = selected_();
    const st = state_(c);

    if (
      st &&
      (
        OPEN_PHASES.includes(st.phase) ||
        st.phase === "UNDONE"
      )
    ) {
      throw new Error(
        "此列異動未完成或已取消，不補寄先前的核准／退件信。"
      );
    }

    const isApplied = (
      st &&
      st.phase === ACTIVE_PHASE
    );

    const isLegacyApplied = (
      !st &&
      text_(
        cell_(
          c,
          CONFIG.STATUS_COL
        ).getValue()
      ).startsWith("已更新日曆")
    );

    if (
      isApplied ||
      isLegacyApplied
    ) {
      cell_(
        c,
        CONFIG.APPROVE_CHECK_COL
      ).setValue(true);

      approval_(
        c,
        Date.now() + 60000
      );

    } else {
      const store = noticeStore_(
        c,
        CONFIG.REJECT_LOG_COL
      );

      if (!store.latest) {
        throw new Error(
          "此列沒有可重試的退件通知。"
        );
      }

      if (
        store.batches[
          store.latest
        ].requestHash !== hash_(
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
        Date.now() + 60000
      );
    }
  });
}


// ============================================================
// 人工核對不明寄信結果
//
// 必須先實際確認收件人信箱／管理員副本。
// 「是」：確認已寄，標記完成
// 「否」：確認未寄，允許下次重試
// 「取消」：仍不確定，不改狀態
// ============================================================

function resolveSelectedMail() {
  const c = locked_(
    () => selected_()
  );

  const ui = SpreadsheetApp.getUi();

  for (
    const col of [
      CONFIG.APPROVE_LOG_COL,
      CONFIG.REJECT_LOG_COL
    ]
  ) {
    const store = locked_(
      () => noticeStore_(c, col)
    );

    for (
      const key of Object.keys(store.batches)
    ) {
      const tasks = store.batches[key].tasks;

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

        const task = tasks[index];

        // UI 等候期間不占用 Script Lock。
        const answer = ui.alert(
          "先核對信箱／管理員副本",

          `收件人：${task.name} ${task.email}\n` +
          `主旨：${task.subject}\n` +
          `通知識別：${key}:${index}\n\n` +
          "已確認寄出→「是」；" +
          "已確認未寄出→「否」；" +
          "不確定→「取消」。",

          ui.ButtonSet.YES_NO_CANCEL
        );

        if (answer === ui.Button.CANCEL) {
          return;
        }

        locked_(() => {
          const latest = noticeStore_(
            c,
            col
          );

          const current = latest
            .batches[key]
            .tasks[index];

          if (
            ![
              "UNKNOWN",
              "SENDING"
            ].includes(current.status)
          ) {
            return;
          }

          current.status =
            answer === ui.Button.YES
              ? "SENT"
              : "READY";

          current.at =
            answer === ui.Button.YES
              ? now_() + "（人工核對）"
              : "";

          current.error = "";

          saveMail_(
            c,
            col,
            latest
          );
        });
      }
    }
  }

  ui.alert(
    "核對紀錄已保存。請再用「重試此列未完成通知」更新顯示並補寄。"
  );
}


// ============================================================
// 復原中斷的日曆異動
// ============================================================

function recoverSelectedCalendar() {
  locked_(() => {
    const c = selected_();
    const st = state_(c);

    if (
      !st ||
      !OPEN_PHASES.includes(st.phase)
    ) {
      throw new Error(
        "此列沒有待復原的中斷異動。"
      );
    }

    const calendar = getCalendar_();

    try {
      restoreTransaction_(
        c,
        calendar,
        st,
        eventsForTx_(calendar, st),
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
  });
}


// ============================================================
// 人工已恢復原班後，驗證快照並解除鎖定
//
// 這不是略過檢查。
// 原班仍不符合快照，就不允許解除。
// ============================================================

function confirmSelectedRecovery() {
  const ui = SpreadsheetApp.getUi();

  const answer = ui.alert(
    "僅限人工核對後",

    "請先到日曆核對：" +
    "此筆申請新增的事件均已移除，原班已恢復。" +
    "確定後才繼續；程式還會再次比對快照。",

    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  locked_(() => {
    const c = selected_();
    const st = state_(c);

    if (
      !st ||
      !OPEN_PHASES.includes(st.phase)
    ) {
      throw new Error(
        "沒有需要解除的中斷紀錄。"
      );
    }

    const events = eventsForTx_(
      getCalendar_(),
      st
    );

    checkDependents_(
      st.plans.map(p => p.token),
      events
    );

    for (const p of st.plans) {
      const originalMatches = same_(
        snap_(
          mainEvent_(p, events)
        ),
        p.before
      );

      const hasRemainingAdditions =
        p.additions.some(addition =>
          addedEvent_(
            addition,
            events
          )
        );

      if (
        !originalMatches ||
        hasRemainingAdditions
      ) {
        throw new Error(
          "目前日曆仍與原班快照不符，不能解除。"
        );
      }

      p.restored = true;

      p.additions.forEach(addition => {
        addition.stage = "DELETED";
      });
    }

    finishRestore_(
      c,
      st
    );
  });
}
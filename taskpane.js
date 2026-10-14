/*
 * KJST 도번 자동 대문자 (Office Add-in, shared runtime)
 *
 * 도면관리대장(01_도면MASTER 시트, DrawingMasterTable 표)의 '도면 번호' 열에
 * 값이 입력되면 즉시 공백을 지우고 대문자로 바꿉니다.
 *   예) 30ms-1025-r3  ->  30MS-1025-R3
 *
 * 처음 한 번 리본의 [도번 대문자] 버튼을 누르면 "이 파일을 열 때 자동 실행"이
 * 파일에 저장되어, 그 뒤로는 파일만 열면 백그라운드에서 동작합니다.
 */

var SHEET_NAME = "01_도면MASTER";
var TABLE_NAME = "DrawingMasterTable";
var COLUMN_NAME = "도면 번호";

var registeredFor = null; // 이벤트를 등록한 시트 (중복 등록 방지)

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    setStatus("엑셀에서만 동작합니다.", "warn");
    return;
  }
  enableAutoLoad();
  register();
  var btn = document.getElementById("retry");
  if (btn) {
    btn.onclick = function () {
      registeredFor = null;
      register();
    };
  }
});

/* 파일을 열 때 추가 기능이 자동으로 켜지도록 이 문서에 설정 저장 */
function enableAutoLoad() {
  try {
    if (Office.addin && Office.addin.setStartupBehavior) {
      Office.addin.setStartupBehavior(Office.StartupBehavior.load).catch(function (e) {
        log("자동 실행 설정 실패: " + e.message);
      });
    }
  } catch (e) {
    log("자동 실행 설정 미지원: " + e.message);
  }
}

/* 01_도면MASTER 시트에 변경 이벤트 등록 */
function register() {
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
    sheet.load("id");
    return context.sync().then(function () {
      if (sheet.isNullObject) {
        setStatus("이 파일에는 '" + SHEET_NAME + "' 시트가 없어 동작하지 않습니다.", "warn");
        return;
      }
      if (registeredFor === sheet.id) {
        setStatus("켜짐 — 도면 번호를 입력하면 자동으로 대문자로 바뀝니다.", "ok");
        return;
      }
      sheet.onChanged.add(onChanged);
      return context.sync().then(function () {
        registeredFor = sheet.id;
        setStatus("켜짐 — 도면 번호를 입력하면 자동으로 대문자로 바뀝니다.", "ok");
      });
    });
  }).catch(function (e) {
    setStatus("연결 실패: " + e.message, "warn");
  });
}

/* 입력값 정리: 공백 제거 + 대문자 */
function normalize(v) {
  if (typeof v !== "string") {
    return v;
  }
  return v.replace(/\s+/g, "").toUpperCase();
}

function onChanged(event) {
  // 다른 사람(공동 편집)의 변경은 그 사람의 엑셀이 처리하므로 내 변경만 처리
  if (event.source && event.source !== Excel.EventSource.local) {
    return Promise.resolve();
  }
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem(SHEET_NAME);
    var column = sheet.tables.getItem(TABLE_NAME).columns.getItem(COLUMN_NAME).getDataBodyRange();
    var target = sheet.getRange(event.address).getIntersectionOrNullObject(column);
    target.load("values");
    return context.sync().then(function () {
      if (target.isNullObject) {
        return;
      }
      var dirty = false;
      var out = target.values.map(function (row) {
        return row.map(function (v) {
          var n = normalize(v);
          if (n !== v) {
            dirty = true;
          }
          return n;
        });
      });
      if (!dirty) {
        return; // 이미 대문자면 아무것도 하지 않음 (무한 반복 방지)
      }
      target.values = out;
      return context.sync();
    });
  }).catch(function (e) {
    log("변환 실패 (" + event.address + "): " + e.message);
  });
}

function setStatus(text, kind) {
  var el = document.getElementById("status");
  if (el) {
    el.textContent = text;
    el.className = "status " + (kind || "");
  }
}

function log(text) {
  var el = document.getElementById("log");
  if (el) {
    var line = document.createElement("div");
    line.textContent = new Date().toLocaleTimeString() + "  " + text;
    el.insertBefore(line, el.firstChild);
  }
  if (window.console) {
    console.log("[KJST]", text);
  }
}

if (typeof module !== "undefined") {
  module.exports = { normalize: normalize };
}

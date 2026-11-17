/*
 * KJST-ADDIN (Office Add-in, shared runtime)
 *
 * 도면관리대장(01_도면MASTER 시트, DrawingMasterTable 표)의 A~AA열에
 * 값이 입력되면 즉시 정리합니다.
 *   - 모든 칸: 영어 소문자 -> 대문자 ('비고' 칸은 제외, 입력한 그대로)
 *   - '도면 번호' 칸: 공백도 함께 제거        예) 30ms-1025-r3   ->  30MS-1025-R3
 *   - '도면명' 칸: 작은따옴표·큰따옴표 제거    예) base "plate"   ->  BASE PLATE
 *   - 수식이 들어 있는 칸(자동 칸)은 건드리지 않음
 *
 * 처음 한 번 리본의 [KJST-ADDIN] 버튼을 누르면 "이 파일을 열 때 자동 실행"이
 * 파일에 저장되어, 그 뒤로는 파일만 열면 백그라운드에서 동작합니다.
 */

var SHEET_NAME = "01_도면MASTER";
var TABLE_NAME = "DrawingMasterTable";
var COLUMN_RANGE = "A:AA";          // 정리 대상 열 범위
var NUMBER_COLUMN = "도면 번호";      // 공백 제거 + 대문자
var NAME_COLUMN = "도면명";           // 따옴표 제거 + 대문자
var SKIP_COLUMNS = ["비고"];          // 자동 변환하지 않는 열 (입력한 그대로 둠)

// 제거할 따옴표: ' " 와 워드/한글에서 붙여 넣을 때 생기는 둥근 따옴표 ‘ ’ “ ”
var QUOTES = /['"‘’“”]/g;

var registeredFor = null; // 이벤트를 등록한 시트 (중복 등록 방지)
var ON_TEXT = "켜짐 — A~AA열의 영어는 대문자로, 도면명의 따옴표는 자동으로 지워집니다.";

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
        setStatus(ON_TEXT, "ok");
        return;
      }
      sheet.onChanged.add(onChanged);
      return context.sync().then(function () {
        registeredFor = sheet.id;
        setStatus(ON_TEXT, "ok");
      });
    });
  }).catch(function (e) {
    setStatus("연결 실패: " + e.message, "warn");
  });
}

/* 열 이름에 따라 입력값 정리 */
function normalize(v, columnName) {
  if (typeof v !== "string" || SKIP_COLUMNS.indexOf(columnName) !== -1) {
    return v;
  }
  if (columnName === NUMBER_COLUMN) {
    return v.replace(/\s+/g, "").toUpperCase();
  }
  if (columnName === NAME_COLUMN) {
    return v.replace(QUOTES, "").trim().toUpperCase();
  }
  return v.toUpperCase();
}

function onChanged(event) {
  // 다른 사람(공동 편집)의 변경은 그 사람의 엑셀이 처리하므로 내 변경만 처리
  if (event.source && event.source !== Excel.EventSource.local) {
    return Promise.resolve();
  }
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem(SHEET_NAME);
    var table = sheet.tables.getItem(TABLE_NAME);
    var header = table.getHeaderRowRange();
    header.load("values, columnIndex");

    // 표 데이터 영역 ∩ A~AA열 ∩ 방금 바뀐 칸
    var area = table.getDataBodyRange().getIntersection(sheet.getRange(COLUMN_RANGE));
    var target = area.getIntersectionOrNullObject(sheet.getRange(event.address));
    target.load("values, formulas, columnIndex");

    return context.sync().then(function () {
      if (target.isNullObject) {
        return;
      }
      var names = header.values[0];
      var changed = 0;
      for (var i = 0; i < target.values.length; i++) {
        for (var j = 0; j < target.values[i].length; j++) {
          var f = target.formulas[i][j];
          if (typeof f === "string" && f.charAt(0) === "=") {
            continue; // 자동(수식) 칸은 그대로
          }
          var v = target.values[i][j];
          var name = names[target.columnIndex + j - header.columnIndex];
          var n = normalize(v, name);
          if (n !== v) {
            target.getCell(i, j).values = [[n]];
            changed++;
          }
        }
      }
      if (!changed) {
        return; // 바꿀 것이 없으면 아무것도 하지 않음 (무한 반복 방지)
      }
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

/*
 * KJST-ADDIN (Office Add-in, shared runtime)
 *
 * 도면관리대장(01_도면MASTER 시트, DrawingMasterTable 표)의 A~AA열에
 * 값이 입력되면 즉시 정리합니다.
 *   - 모든 칸: 영어 소문자 -> 대문자 ('비고' 칸은 제외, 입력한 그대로)
 *   - '도면 번호' 칸: 공백도 함께 제거        예) 30ms-1025-r3   ->  30MS-1025-R3
 *   - '도면 이름' 칸: 작은따옴표·큰따옴표 제거  예) base "plate"   ->  BASE PLATE
 *   - 수식이 들어 있는 칸(자동 칸)은 건드리지 않음
 *   - '도면 유형' 칸: 값이 지워졌거나, 도면 번호가 비었거나 P 도면이 아닌데 값이 남아 있으면
 *     원래 수식으로 되돌림 (P 도면에서 고른 제관도/주물제작도는 그대로)
 *
 * 처음 한 번 리본의 [KJST-ADDIN] 버튼을 누르면 "이 파일을 열 때 자동 실행"이
 * 파일에 저장되어, 그 뒤로는 파일만 열면 백그라운드에서 동작합니다.
 */

var SHEET_NAME = "01_도면MASTER";
var TABLE_NAME = "DrawingMasterTable";
var COLUMN_RANGE = "A:AA";          // 정리 대상 열 범위
var VERSION = "1.1.3";              // 작업창 제목 옆에 표시 (어떤 버전이 동작하는지 확인용)
// 열 이름은 띄어쓰기를 무시하고 비교합니다 ("도면 이름" = "도면이름")
var NUMBER_COLUMNS = ["도면번호"];          // 공백 제거 + 대문자
var NAME_COLUMNS = ["도면이름", "도면명"];  // 따옴표 제거 + 대문자
var SKIP_COLUMNS = ["비고"];                // 자동 변환하지 않는 열 (입력한 그대로 둠)
var TYPE_COLUMNS = ["도면유형"];            // 수식 칸. P 도면만 제관도/주물제작도를 골라 값으로 둠
// 도면 유형 수식 (R1C1). 표 안의 다른 '도면 유형' 칸에서 수식을 못 찾을 때만 사용
var TYPE_FORMULA_R1C1 = '=IF(RC7="","",IF(MID(RC7,6+(LEFT(RC7,1)="K"),1)="A","ASSY",' +
  'IF(MID(RC7,6+(LEFT(RC7,1)="K"),1)="P","제관/주물 선택","PART")))';

// 제거할 따옴표: ' " 와 워드/한글에서 붙여 넣을 때 생기는 둥근 따옴표 ‘ ’ “ ”
var QUOTES = /['"‘’“”]/g;

var registeredFor = null; // 이벤트를 등록한 시트 (중복 등록 방지)
var ON_TEXT = "켜짐 (v" + VERSION + ") — A~AA열의 영어는 대문자로, 도면 이름의 따옴표는 자동으로 지워집니다.";

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    setStatus("엑셀에서만 동작합니다.", "warn");
    return;
  }
  var ver = document.getElementById("ver");
  if (ver) {
    ver.textContent = "v" + VERSION;
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
function colKey(name) {
  return String(name == null ? "" : name).replace(/\s+/g, "");
}

function normalize(v, columnName) {
  var key = colKey(columnName);
  if (typeof v !== "string" || SKIP_COLUMNS.indexOf(key) !== -1) {
    return v;
  }
  if (NUMBER_COLUMNS.indexOf(key) !== -1) {
    return v.replace(/\s+/g, "").toUpperCase();
  }
  if (NAME_COLUMNS.indexOf(key) !== -1) {
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
    target.load("values, formulas, columnIndex, columnCount, rowIndex, rowCount");

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
      var after = changed ? context.sync() : Promise.resolve();
      return after.then(function () {
        return fixType(context, sheet, table, names, header.columnIndex, target);
      });
    });
  }).catch(function (e) {
    log("변환 실패 (" + event.address + "): " + e.message);
  });
}

/* P 도면인지 (도면 번호의 기종·UNIT 다음 글자가 P) */
function isP(no) {
  var s = String(no == null ? "" : no).replace(/\s+/g, "").toUpperCase();
  var k = s.charAt(0) === "K" ? 1 : 0;
  return s.charAt(5 + k) === "P";
}

/* 도면 유형 칸이 수식이어야 하는지 (true = 원래 수식으로 되돌림) */
function needTypeFormula(no, typeFormula) {
  var isFormula = typeof typeFormula === "string" && typeFormula.charAt(0) === "=";
  if (isFormula) {
    return false;
  }
  var empty = typeFormula === "" || typeFormula == null;
  var noEmpty = String(no == null ? "" : no).replace(/\s+/g, "") === "";
  return empty || noEmpty || !isP(no);
}

/* 도면 번호 / 도면 유형이 바뀐 행의 도면 유형 칸 정리 */
function fixType(context, sheet, table, names, headerCol, target) {
  var numIdx = -1, typeIdx = -1;
  for (var c = 0; c < names.length; c++) {
    var key = colKey(names[c]);
    if (NUMBER_COLUMNS.indexOf(key) !== -1) numIdx = c;
    if (TYPE_COLUMNS.indexOf(key) !== -1) typeIdx = c;
  }
  if (numIdx < 0 || typeIdx < 0) {
    return Promise.resolve();
  }
  var numCol = headerCol + numIdx, typeCol = headerCol + typeIdx;
  var c0 = target.columnIndex, c1 = target.columnIndex + target.columnCount - 1;
  var touches = function (col) { return col >= c0 && col <= c1; };
  if (!touches(numCol) && !touches(typeCol)) {
    return Promise.resolve();
  }
  var nums = sheet.getRangeByIndexes(target.rowIndex, numCol, target.rowCount, 1);
  var types = sheet.getRangeByIndexes(target.rowIndex, typeCol, target.rowCount, 1);
  var column = table.columns.getItemAt(typeIdx).getDataBodyRange();
  nums.load("values");
  types.load("formulas");
  column.load("formulasR1C1");
  return context.sync().then(function () {
    var r1c1 = TYPE_FORMULA_R1C1;
    var all = column.formulasR1C1;
    for (var k = 0; k < all.length; k++) {
      var f = all[k][0];
      if (typeof f === "string" && f.indexOf("=IF(") === 0 && f.indexOf("PART") !== -1) {
        r1c1 = f;
        break;
      }
    }
    var fixed = 0;
    for (var i = 0; i < target.rowCount; i++) {
      if (needTypeFormula(nums.values[i][0], types.formulas[i][0])) {
        types.getCell(i, 0).formulasR1C1 = [[r1c1]];
        fixed++;
      }
    }
    if (fixed) {
      log("도면 유형 " + fixed + "칸을 원래 수식으로 되돌림");
      return context.sync();
    }
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
  module.exports = { normalize: normalize, needTypeFormula: needTypeFormula, isP: isP, onChanged: onChanged, VERSION: VERSION };
}

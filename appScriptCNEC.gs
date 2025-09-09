// Step 0a:from var sourceSheet = ss.getSheetByName("總教師時間表");
// Obtain teacher code sheer from column C6 to C60 and skip empty row
function getTeacherCodeSheer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  if (!sourceSheet) return [];

  // Get values from C6:C60
  var rawValues = sourceSheet.getRange(6, 3, 55, 1).getValues(); // 55 rows, 1 column
  var teacherCodes = rawValues
    .map(function (row) { return row[0]; }) // flatten
    .filter(function (code) { return code && code.toString().trim() !== ""; }); // skip empty

  return teacherCodes;
}

// Step 0b:from var sourceSheet = ss.getSheetByName("總教師時間表");
// Obtain teacher name sheer from column B6 to B60 and skip empty row
function getTeacherNamesSheer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  if (!sourceSheet) return [];

  // Get values from B6:B60
  var rawValues = sourceSheet.getRange(6, 2, 55, 1).getValues(); // 55 rows, 1 column
  var teacherNames = rawValues
    .map(function (row) { return row[0]; }) // flatten
    .filter(function (code) { return code && code.toString().trim() !== ""; }); // skip empty

  return teacherNames;
}

// Step 0c: get the column digit
function getColumnDigit(column_name) {
  var col = 0;
  for (var i = 0; i < column_name.length; i++) {
    col *= 26;
    col += column_name.charCodeAt(i) - 64; // 'A'.charCodeAt(0) is 65
  }
  return col;
}

// Step 1: Create new sheet for each teacher
function createNewSheetWithTEACHERCODE() {
  // Obtain teacher code sheer from column C6 to C60 and skip empty row
  const TEACHER_CODE_SHEER = getTeacherCodeSheer();
  // Make sure you have your spreadsheet object
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const tcode of TEACHER_CODE_SHEER) {
    let nameSheet = ss.getSheetByName(tcode);
    if (nameSheet) {
      console.log(`CreateNewSheet() Skip Processing: -${tcode}`);
    } else {
      ss.insertSheet(tcode);
      console.log(`Created sheet for: ${tcode}`);
    }
  }
}

// Step 2: Copy Template to each sheer with teacher code only
function copyTemplateToSheer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teacherNames = getTeacherNamesSheer();
  const TEACHER_CODE_SHEER = getTeacherCodeSheer();

  var templateSheet = ss.getSheetByName('template');
  var templateData = templateSheet.getDataRange().getValues();

  for (var i = 0; i < TEACHER_CODE_SHEER.length; i++) {
    var code = TEACHER_CODE_SHEER[i];
    var sheet = ss.getSheetByName(code);
    if (!sheet) {
      sheet = ss.insertSheet(code);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, templateData.length, templateData[0].length).setValues(templateData);
    // Add teacher name to C1
    sheet.getRange(1, 3).setValue(teacherNames[i]);
    sheet.getRange(2, 3).setValue(teacherNames[i]);
    Logger.log(TEACHER_CODE_SHEER[i]);
  }
}

// Step 3: Copy lesson info for Monday
// From Source D to T = 4	5	6	7	8	9	10	11	12	13	14	15	16	17	18	19	20
function copyLessonInfoForMonday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  var TEACHER_NAME_ORDER_INDIVIDUAL_SHEER = getTeacherCodeSheer();

  var MAX_LESSON_COUNT = 17;
  var SKIP_COUNT = 6;

  var BEGIN_COLUMN_INDEX = getColumnDigit('D');//Corresponding to column D

  for (var i = 0; i < TEACHER_NAME_ORDER_INDIVIDUAL_SHEER.length; i++) {
    var teacherSheetName = TEACHER_NAME_ORDER_INDIVIDUAL_SHEER[i];
    var destSheet = ss.getSheetByName(teacherSheetName);

    if (sourceSheet && destSheet) {
      var sourceRow = SKIP_COUNT + i;
      var sourceRange = sourceSheet.getRange(sourceRow, BEGIN_COLUMN_INDEX, 1, MAX_LESSON_COUNT); // D to T
      var rowValues = sourceRange.getValues()[0];
      var transformedColValues = [];
      var lessonIdx = 0;
      for (var j = 0; j < MAX_LESSON_COUNT; j++) {
        transformedColValues.push([rowValues[lessonIdx++]]);
      }
      destSheet.getRange(4, getColumnDigit('D'), MAX_LESSON_COUNT, 1).setValues(transformedColValues); // D4:D20
    }
  }
}

//Step 3: Copy lesson info for Tuesday
//From Source U to AK = 21	22	23	24	25	26	27	28	29	30	31	32	33	34	35	36	37
function copyLessonInfoForTuesday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  var TEACHER_NAME_ORDER_INDIVIDUAL_SHEER = getTeacherCodeSheer();

  var MAX_LESSON_COUNT = 17;
  var SKIP_COUNT = 6;

  var BEGIN_COLUMN_INDEX = getColumnDigit('U');//Corresponding to column U

  for (var i = 0; i < TEACHER_NAME_ORDER_INDIVIDUAL_SHEER.length; i++) {
    var teacherSheetName = TEACHER_NAME_ORDER_INDIVIDUAL_SHEER[i];
    var destSheet = ss.getSheetByName(teacherSheetName);

    if (sourceSheet && destSheet) {
      var sourceRow = SKIP_COUNT + i;
      var sourceRange = sourceSheet.getRange(sourceRow, BEGIN_COLUMN_INDEX, 1, MAX_LESSON_COUNT); // U to AK
      var rowValues = sourceRange.getValues()[0];
      var transformedColValues = [];
      var lessonIdx = 0;
      for (var j = 0; j < MAX_LESSON_COUNT; j++) {
        transformedColValues.push([rowValues[lessonIdx++]]);
      }
      destSheet.getRange(4, getColumnDigit('E'), MAX_LESSON_COUNT, 1).setValues(transformedColValues); // E4:E20
    }
  }
}

// Step 3: Copy lesson info for Wednesday
// From Source AL to BA = 38	39	40	41	42	43	44	45	46	47	48	49	50	51	52	53
function copyLessonInfoForWednesday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  var TEACHER_NAME_ORDER_INDIVIDUAL_SHEER = getTeacherCodeSheer();

  var MAX_LESSON_COUNT = 17;
  var SKIP_COUNT = 6;

  var BEGIN_COLUMN_INDEX = getColumnDigit('AL');//Corresponding to column AL

  for (var i = 0; i < TEACHER_NAME_ORDER_INDIVIDUAL_SHEER.length; i++) {
    var teacherSheetName = TEACHER_NAME_ORDER_INDIVIDUAL_SHEER[i];
    var destSheet = ss.getSheetByName(teacherSheetName);

    if (sourceSheet && destSheet) {
      var sourceRow = SKIP_COUNT + i;
      var sourceRange = sourceSheet.getRange(sourceRow, BEGIN_COLUMN_INDEX, 1, MAX_LESSON_COUNT); // AL to BA
      var rowValues = sourceRange.getValues()[0];
      // var transformedColValues = [];
      // var lessonIdx = 0;
      // for (var j = 0; j < MAX_LESSON_COUNT; j++) {
      //   transformedColValues.push([rowValues[lessonIdx++]]);
      // }
      // destSheet.getRange(4, getColumnDigit('G'), MAX_LESSON_COUNT, 1).setValues(transformedColValues); // F4:F20
      var insertBlankIndexes = [15];
      var paddedColValues = [];
      var lessonIdx = 0;
      for (var j = 0; j < MAX_LESSON_COUNT; j++) {
        if (insertBlankIndexes.indexOf(j) !== -1) {
          paddedColValues.push([""]);
        } else {
          paddedColValues.push([rowValues[lessonIdx++]]);
        }
      }
      destSheet.getRange(4, getColumnDigit('F'), MAX_LESSON_COUNT, 1).setValues(paddedColValues); // F4:F20
    }
  }
}

// Step 3: Copy lesson info for Thursday
// From Source BB to BR = 54	55	56	57	58	59	60	61	62	63	64	65	66	67	68	69	70
function copyLessonInfoForThursday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  var TEACHER_NAME_ORDER_INDIVIDUAL_SHEER = getTeacherCodeSheer();

  var MAX_LESSON_COUNT = 17;
  var SKIP_COUNT = 6;

  var BEGIN_COLUMN_INDEX = getColumnDigit('BB');//Corresponding to column BC

  for (var i = 0; i < TEACHER_NAME_ORDER_INDIVIDUAL_SHEER.length; i++) {
    var teacherSheetName = TEACHER_NAME_ORDER_INDIVIDUAL_SHEER[i];
    var destSheet = ss.getSheetByName(teacherSheetName);

    if (sourceSheet && destSheet) {
      var sourceRow = SKIP_COUNT + i;
      var sourceRange = sourceSheet.getRange(sourceRow, BEGIN_COLUMN_INDEX, 1, MAX_LESSON_COUNT); // BB to BR
      var rowValues = sourceRange.getValues()[0];
      var transformedColValues = [];
      var lessonIdx = 0;
      for (var j = 0; j < MAX_LESSON_COUNT; j++) {
        transformedColValues.push([rowValues[lessonIdx++]]);
      }
      destSheet.getRange(4, getColumnDigit('G'), MAX_LESSON_COUNT, 1).setValues(transformedColValues); // G4:G20
    }
  }
}

//Step 3: Copy lesson info for Friday
//From Source BS to CH = 71	72	73	74	75	76	77	78	79	80	81	82	83	84	85	86
function copyLessonInfoForFriday() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("總教師時間表");
  var TEACHER_NAME_ORDER_INDIVIDUAL_SHEER = getTeacherCodeSheer();

  var MAX_LESSON_COUNT = 17;
  var SKIP_COUNT = 6;

  var BEGIN_COLUMN_INDEX = getColumnDigit('BS');//Corresponding to column BS

  for (var i = 0; i < TEACHER_NAME_ORDER_INDIVIDUAL_SHEER.length; i++) {
    var teacherSheetName = TEACHER_NAME_ORDER_INDIVIDUAL_SHEER[i];
    var destSheet = ss.getSheetByName(teacherSheetName);

    if (sourceSheet && destSheet) {
      var sourceRow = SKIP_COUNT + i;
      var sourceRange = sourceSheet.getRange(sourceRow, BEGIN_COLUMN_INDEX, 1, MAX_LESSON_COUNT); // BS to CH
      var rowValues = sourceRange.getValues()[0]; // [val1, val2, ..., val11]
      // var transformedColValues = [];
      // var lessonIdx = 0;
      // for (var j = 0; j < MAX_LESSON_COUNT; j++) {
      //   transformedColValues.push([rowValues[lessonIdx++]]);
      // }
      // destSheet.getRange(4, getColumnDigit('H'), MAX_LESSON_COUNT, 1).setValues(transformedColValues); // H4:H20
      var insertBlankIndexes = [14];
      var paddedColValues = [];
      var lessonIdx = 0;
      for (var j = 0; j < MAX_LESSON_COUNT; j++) {
        if (insertBlankIndexes.indexOf(j) !== -1) {
          paddedColValues.push([""]);
        } else {
          paddedColValues.push([rowValues[lessonIdx++]]);
        }
      }
      destSheet.getRange(4, getColumnDigit('H'), MAX_LESSON_COUNT, 1).setValues(paddedColValues); // H4:H20
    }
  }
}

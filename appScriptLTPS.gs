const TEACHERCODE = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', 'A', 'K', 'S'];
var OLD_TEACHER_LIST = ['姚寶芬', '林鐃程', '莊鳳儀', '黃雋之', '李美姍', '溫婉慈', '楊曉怡', '陳藹儀', '鄧紫晴', '張頴儀', '王麗蓮', '劉雅恩', '黃俊基', '陳晉芝', '尹兆恒', '王翠儀', '王彥敏', '容淑儀', '李志豪', '李海言', '李珈琳', '林善嬴', '梁美詩', '楊麗芳', '潘國光', '潘濬仁', '羅宇琪', '葉兆彤', '葉曉文', '譚曉嵐', '陳振秋', '顏詠詩', '馮碧珊', '黃慧萍', '黃旭俊', 'Allyson Holder'];
var NEW_TEACHER_LIST = ['羅宇琪', '楊麗芳', '姚寶芬', '林鐃程', '葉兆彤', '鄧海玲', '梁思晴', '李美姍', '楊曉怡', '林善嬴', '王彥敏', '黃雋之', '王麗蓮', '劉雅恩', '黃俊基', '陳晉芝', '譚曉嵐', '鄧紫晴', '陳藹儀', '李珈琳', '尹兆恒', '馮碧珊', '顏詠詩', '王翠儀', '朱明霞', '葉曉文', '李海言', '潘國光', '李志豪', '潘濬仁', '容淑儀', '陳振秋', '黃旭俊', '梁美詩', 'ALLYSON HOLDER', 'ＫＡＲＹ', 'ＳＥＮＴＡ'];
var MAPPING_TABLE = [
  { '羅宇琪': '琪' },
  { '楊麗芳': '楊' },
  { '姚寶芬': '芬' },
  { '林鐃程': '程' },
  { '葉兆彤': '彤' },
  { '鄧海玲': '鄧' },
  { '梁思晴': '梁' },
  { '李美姍': '美' },
  { '楊曉怡': '曉' },
  { '林善嬴': '善' },
  { '王彥敏': '敏' },
  { '黃雋之': '雋' },
  { '王麗蓮': '麗' },
  { '劉雅恩': '雅' },
  { '黃俊基': '黃' },
  { '陳晉芝': '芝' },
  { '譚曉嵐': '譚' },
  { '鄧紫晴': '紫' },
  { '陳藹儀': '藹' },
  { '李珈琳': '珈' },
  { '尹兆恒': '尹' },
  { '馮碧珊': '馮' },
  { '顏詠詩': '顏' },
  { '王翠儀': '翠' },
  { '朱明霞': '朱' },
  { '葉曉文': '葉' },
  { '李海言': '言' },
  { '潘國光': '潘' },
  { '李志豪': '豪' },
  { '潘濬仁': '仁' },
  { '容淑儀': '儀' },
  { '陳振秋': '陳' },
  { '黃旭俊': '旭' },
  { '梁美詩': '詩' },
  { 'ALLYSON HOLDER': 'A' },
  { 'ＫＡＲＹ': 'K' },
  { 'ＳＥＮＴＡ': 'S' }
];

//Step 1: get all the teacher codes
function getTeacherNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var results = {};
  var teacherNames = [];
  var normalizedSet = new Set();

  sheets.forEach(function (sheet) {
    var lastCol = sheet.getLastColumn();
    var row1 = [];
    if (lastCol > 0) {
      var values = sheet.getRange(1, 1, 1, lastCol).getValues();
      if (values && values.length > 0) {
        row1 = values[0];
      }
    }
    var filtered = row1.filter(function (value) {
      return value !== "" && value.trim() !== "";
    });
    results[sheet.getName()] = filtered;
    filtered.forEach(function (name) {
      var normalized = normalizeName(name);
      if (normalized && !normalizedSet.has(normalized)) {
        teacherNames.push(normalized); 
        // Store original string
        normalizedSet.add(normalized);
        // Logger.log('Raw: ' + name + ', Normalized: ' + normalized);
      }
    });
  });
  Logger.log('getTeacherNames() - Results: ' + teacherNames);
  return teacherNames;
}

function getTeacherNamesAndCode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var results = {};
  var teacherNames = [];
  var normalizedSet = new Set();

  sheets.forEach(function (sheet) {
    var lastCol = sheet.getLastColumn();
    var row1 = [];
    if (lastCol > 0) {
      var values = sheet.getRange(1, 1, 1, lastCol).getValues();
      if (values && values.length > 0) {
        row1 = values[0];
      }
    }
    var filtered = row1.filter(function (value) {
      return value !== "" && value.trim() !== "";
    });
    results[sheet.getName()] = filtered;
    filtered.forEach(function (name) {
      var normalized = normalizeName(name);
      if (normalized && !normalizedSet.has(normalized)) {
        teacherNames.push(name); 
        // Store original string
        normalizedSet.add(normalized);
        // Logger.log('Raw: ' + name + ', Normalized: ' + normalized);
      }
    });
  });
  Logger.log('getTeacherNames() - Results: ' + teacherNames);
  return teacherNames;
}

//Step 2: create sheet to from all the teacher codes
function createSheetsForAllTeacherCodes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var teacherCode = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', 'A', 'K', 'S'];
  var teacherCode = extractTeacherCodes();
  var existingSheets = ss.getSheets().map(function (sheet) { return sheet.getName(); });
  Logger.log('Existing Sheets: ' + existingSheets);

  for (var i = 0; i < teacherCode.length; i++) {
    var name = teacherCode[i];
    if (existingSheets.indexOf(name) !== -1) {
      // Skip if sheet already exists
      Logger.log('Skipped: ' + name);
      continue;
    }
    ss.insertSheet(name);
    existingSheets.push(name); // Update the list as we add new sheets
    Logger.log('Created: ' + name);
  }
}

//Step 3: Copy Template to each sheer with teacher code only
function copyTemplateToSheer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var teacherNames = ['羅宇(琪)(1A)', '(楊)麗芳(1A)', '姚寶(芬)(1B)', '林鐃(程)(1B)', '葉兆(彤)(2A)', '(鄧)海玲(2B)', '(梁)思晴(3A)', '李(美)姍(3B)', '楊(曉)怡(4A)', '林(善)嬴(4B)', '王彥(敏)(5A)', '黃(雋)之(5B)', '王(麗)蓮(6A)', '劉(雅)恩(6B)', '(黃)俊基(6C)', '陳晉(芝)(6D)', '(譚)曉嵐', '鄧(紫)晴', '陳(藹)儀', '李(珈)琳', '(尹)兆恒', '(馮)碧珊', '(顏)詠詩', '王(翠)儀', '(朱)明霞', '(葉)曉文', '李海(言)', '(潘)國光', '李志(豪)', '潘濬仁(仁)', '容淑(儀)', '(陳)振秋', '黃(旭)俊', '梁美(詩)', 'ALLYSON HOLDER(NET)', 'KARY', 'SENTA (NCS)'];
  // var teacherCode = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', 'A', 'K', 'S'];
  var teacherNames = extractTeacherNames();
  var teacherCode = extractTeacherCodes();
  var templateSheet = ss.getSheetByName('Template1');
  var templateData = templateSheet.getDataRange().getValues();

  for (var i = 0; i < teacherCode.length; i++) {
    var code = teacherCode[i];
    var sheet = ss.getSheetByName(code);
    if (!sheet) {
      sheet = ss.insertSheet(code);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, templateData.length, templateData[0].length).setValues(templateData);
    // Add teacher name to C1
    sheet.getRange(1, 3).setValue(teacherNames[i]);
    Logger.log(teacherCode[i]);
  }
}

//Step 4: Copy the lesson info from original sheer from teacher names
function copyLessonInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var teacherCode = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', 'A', 'K', 'S'];
  // var teacherSheers = ['羅宇琪', '楊麗芳', '姚寶芬', '林鐃程', '葉兆彤', '鄧海玲', '梁思晴', '李美姍', '楊曉怡', '林善嬴', '王彥敏', '黃雋之', '王麗蓮', '劉雅恩', '黃俊基', '陳晉芝', '譚曉嵐', '鄧紫晴', '陳藹儀', '李珈琳', '尹兆恒', '馮碧珊', '顏詠詩', '王翠儀', '朱明霞', '葉曉文', '李海言', '潘國光', '李志豪', '潘濬仁', '容淑儀', '陳振秋', '黃旭俊', '梁美詩', 'ＮＥＴ', 'ＫＡＲＹ', 'ＳＥＮＴＡ'];
  var teacherSheers = extractTeacherNames();
  var teacherCode = extractTeacherCodes();

  for (var i = 0; i < teacherCode.length; i++) {
    var sourceSheet = ss.getSheetByName(teacherSheers[i]);
    var destSheet = ss.getSheetByName(teacherCode[i]);
    if (sourceSheet && destSheet) {
      var sourceRange = sourceSheet.getRange("C3:H18");
      var destRange = destSheet.getRange("D5:I20");
      var values = sourceRange.getValues();
      var backgrounds = sourceRange.getBackgrounds();
      destRange.setValues(values);
      destRange.setBackgrounds(backgrounds); // Copy cell backgrounds
      Logger.log("Copied for: " + teacherSheers[i] + " -> " + teacherCode[i]);
    } else {
      Logger.log("Missing sheet for: " + teacherSheers[i] + " or " + teacherCode[i]);
    }
  }
}


//Step 5: Update the lesson info to remove all newline characters from teacher code
function updateLessonInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < TEACHERCODE.length; i++) {
    var destSheet = ss.getSheetByName(TEACHERCODE[i]);
    if (destSheet) {
      var destRange = destSheet.getRange("D5:I20");
      var values = destRange.getValues();

      // Process each cell: convert multiline to single line using custom logic
      for (var r = 0; r < values.length; r++) {
        for (var c = 0; c < values[0].length; c++) {
          var cell = values[r][c];
          if (typeof cell === 'string') {
            // Split into parts on newline, filter out empty lines, trim each part
            var parts = cell.replace(/\r/g, '').split('\n').filter(Boolean).map(function (p) { return p.trim(); });

            // Rebuild value
            var result = '';
            for (var j = 0; j < parts.length; j++) {
              if (/^\d+$/.test(parts[j]) && j === parts.length - 1) {
                // If last part is digits only, add a space before it
                result += ' ' + parts[j];
              } else {
                result += parts[j];
              }
            }
            values[r][c] = result;
          }
        }
      }
      destSheet.getRange("D5:I20").setValues(values);
      Logger.log("Cleaned, joined, and copied for: " + TEACHERCODE[i]);
    } else {
      Logger.log("Missing sheet for: " + TEACHERCODE[i]);
    }
  }
}

//Step 6: add 班主任 prefix
function add班主任() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var i = 0; i < TEACHERCODE.length; i++) {
    var destSheet = ss.getSheetByName(TEACHERCODE[i]);

    if (destSheet) {
      var destRange = destSheet.getRange("D7:I7");
      var values = destRange.getValues();

      // Process each cell in the range
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          // Check if cell is not empty
          if (values[row][col] !== "") {
            values[row][col] = values[row][col] + " 班主任";
          }
        }
      }

      // Set the updated values back to the range
      destRange.setValues(values);
    }
  }
}

// Step 7: add 當值 prefix
function add當值ToRows(rowList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var columnStart = 4; // Column D is 4
  var columnEnd = 9;   // Column I is 9

  for (var i = 0; i < TEACHERCODE.length; i++) {
    var destSheet = ss.getSheetByName(TEACHERCODE[i]);
    if (!destSheet) continue;

    for (var r = 0; r < rowList.length; r++) {
      var rowNum = rowList[r];
      var destRange = destSheet.getRange(rowNum, columnStart, 1, columnEnd - columnStart + 1);
      var values = destRange.getValues();

      // Process each cell in the row
      for (var col = 0; col < values[0].length; col++) {
        if (values[0][col] !== "") {
          values[0][col] = "當值 " + values[0][col];
        }
      }
      destRange.setValues(values);
    }
  }
}

//Step 8: add 當值 5, 11, 14, 17, 18
function add當值s() {
  var rowList = [5, 11, 14, 17, 18];
  add當值ToRows(rowList);
}

/* ----------------Test Cases Below---------------- */
function extractTeacherNames() {
  var teacherNames = getTeacherNames(); // Assumes array of name strings

  // Remove duplicates
  var resultArray = Array.from(new Set(teacherNames));
  resultArray = resultArray.map(function (name) {
    if (name === "潘濬仁仁") return "潘濬仁";
    if (name === "ALLYSONHOLDERNET") return "ALLYSON HOLDER";
    if (name === "SENTANCS") return "SENTA";
    return name;
  });
  Logger.log('extractTeacherNames() - Results: ' + resultArray);
  Logger.log('Is Array: ' + Array.isArray(resultArray));
  return resultArray;
}

function extractTeacherCodes() {
  var teacherNames = getTeacherNamesAndCode(); // Array of name strings

  // Collect codes: find all Chinese chars in parentheses, skip class or staff codes
  var codeSet = new Set(
    teacherNames
      .map(function (name) {
        // Only extract Chinese character codes inside parentheses, e.g. (美), (善)
        var codeMatch = name.match(/\(([\u4e00-\u9fa5]{1,3})\)/);
        // Return with single quotes, e.g. '美'
        return codeMatch ? codeMatch[1] : null;
      })
      .filter(function (code) { return code !== null; })
  );

  // Add 'A', 'K'
  codeSet.add("A");
  codeSet.add("K");

  var results = Array.from(codeSet);
  Logger.log('Teacher codes: ' + results);
  Logger.log('Is Array: ' + Array.isArray(results));
  return results;
}
//琪,楊,芬,程,彤,鄧,梁,美,曉,善,敏,雋,麗,雅,黃,芝,譚,紫,藹,珈,尹,馮,顏,朱,葉,言,潘,豪,仁,容,陳,旭,詩,A,K,
//楊麗芳,姚寶芬,林鐃程,葉兆彤,鄧海玲,梁思晴,李美姍,楊曉怡,林善嬴,王彥敏,黃雋之,王麗蓮,劉雅恩,黃俊基,陳晉芝,譚曉嵐,鄧紫晴,陳藹儀,李珈琳,尹兆恒,馮碧珊,顏詠詩,王（翠）儀,朱明霞,葉曉文,李海言,潘國光,李志豪,潘濬仁,容淑儀,陳振秋,黃旭俊,梁美詩,ALLYSON HOLDER,KARY

// Test Case 1:
function compareTeacherCodes() {
  var teacherCode1 = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', '容'];
  var teacherCode2 = ['琪', '楊', '芬', '程', '彤', '鄧', '梁', '美', '曉', '善', '敏', '雋', '麗', '雅', '黃', '芝', '譚', '紫', '藹', '珈', '尹', '馮', '顏', '翠', '朱', '葉', '言', '潘', '豪', '仁', '儀', '陳', '旭', '詩', 'A', 'K', 'S'];

  console.log('teacherCode1:', teacherCode1);
  console.log('teacherCode2:', teacherCode2);

  // Find codes in teacherCode1 but not in teacherCode2
  var onlyIn1 = teacherCode1.filter(x => !teacherCode2.includes(x));

  // Find codes in teacherCode2 but not in teacherCode1
  var onlyIn2 = teacherCode2.filter(x => !teacherCode1.includes(x));

  // Codes present in both
  var inBoth = teacherCode1.filter(x => teacherCode2.includes(x));

  console.log('Only in teacherCode1:', onlyIn1);
  console.log('Only in teacherCode2:', onlyIn2);
  console.log('In both:', inBoth);
}


function generateTeacherMappingTable(rawArray) {
  var mapping = {};
  rawArray.forEach(function (item) {
    item = item.trim();

    // English name handling: e.g., ALLYSON HOLDER(NET), KARY, SENTA (NCS)
    var enMatch = item.match(/^([A-Z\s]+)(?:\((\w+)\))?/i);
    if (enMatch && enMatch[1] && /^[A-Z\s]+$/.test(enMatch[1])) {
      var name = enMatch[1].trim();
      var code = enMatch[2] ? enMatch[2][0].toUpperCase() : name[0].toUpperCase();
      mapping[name] = code;
      return;
    }

    // Step 1: Remove trailing class code (with or without parentheses, with/without space)
    // Matches: (1A), (6B), (6D), 1A, 6B, 6D at the end, possibly with space
    var noClass = item.replace(/\s*\(?\d[A-D]\)?\s*$/i, '');

    // Step 2: Remove all parentheses (keeping contents) (for Chinese names)
    var zhName = noClass.replace(/[()]/g, '');

    // Step 3: Remove all whitespace
    zhName = zhName.replace(/\s+/g, '');

    // Grab code from the first parentheses
    var codeMatch = noClass.match(/\(([^\)]+)\)/);

    if (zhName && codeMatch) {
      mapping[zhName] = codeMatch[1];
    }
  });
  return mapping;
}

function getTeacherMappingTable() {
  var mapping = {};
  mapping = generateTeacherMappingTable(getTeacherNames());

  // Build array of single-key objects
  const arr = Object.entries(mapping).map(([key, value]) => ({ [key]: value }));

  Logger.log("var MAPPING_TABLE = " + JSON.stringify(arr, null, 2) + ";");

  // Return a pretty JSON string that matches your desired format
  return "var MAPPING_TABLE = " + JSON.stringify(arr, null, 2) + ";";
}

function normalizeName(name) {
  // Logger.log("Input name: " + name);
  // Remove class codes in parens at the end
  var noClass = name.replace(/\s*\(?\d[A-D]\)?\s*$/i, '');
  // Remove all parentheses (but keep inside)
  var cleaned = noClass.replace(/[()]/g, '');
  // Remove all spaces
  cleaned = cleaned.replace(/\s+/g, '');
  // Logger.log("Cleaned name: " + cleaned);
  return cleaned;
}

// Remark all teachers from last year not existing in this year
function getOldTeacherList() {
  var missingTeachers = [];

  for (var i = 0; i < OLD_TEACHER_LIST.length; i++) {
    var teacher = OLD_TEACHER_LIST[i];
    if (!NEW_TEACHER_LIST.includes(teacher)) {
      missingTeachers.push(teacher);
    }
  }
  Logger.log(missingTeachers);
}

// Remark all teachers from this year not existing in last year
function getNewTeacherList() {
  var missingTeachers = [];

  for (var i = 0; i < NEW_TEACHER_LIST.length; i++) {
    var teacher = NEW_TEACHER_LIST[i];
    if (!OLD_TEACHER_LIST.includes(teacher)) {
      missingTeachers.push(teacher);
    }
  }
  Logger.log(missingTeachers);
}

//Test Case 2: Find the new teacher list teacher code
function getNewTeacherCode() {
  var NEW_TEACHER_LIST = ['姚寶芬', '林鐃程', 'KARY', '黃雋之', '李美姍', 'SENTA', '楊曉怡', '陳藹儀', '鄧紫晴', '梁思晴', '王麗蓮', '劉雅恩', '黃俊基', '陳晉芝', '尹兆恒', '王翠儀', '王彥敏', '容淑儀', '朱明霞', '李志豪', '李海言', '李珈琳', '林善嬴', '梁美詩', '楊麗芳', '潘國光', '潘濬仁', '羅宇琪', '葉兆彤', '葉曉文', '譚曉嵐', '陳振秋', '顏詠詩', '馮碧珊', '鄧海玲', '黃旭俊', 'ALLYSON HOLDER'];

  function normalize(str) {
    return str ? str.toUpperCase().replace(/[Ａ-Ｚ]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 65248); }) : str;
  }

  var newTeacherCode = [];
  for (var i = 0; i < NEW_TEACHER_LIST.length; i++) {
    var teacher = NEW_TEACHER_LIST[i];
    var normTeacher = normalize(teacher);
    var value = null;
    for (var j = 0; j < MAPPING_TABLE.length; j++) {
      var key = Object.keys(MAPPING_TABLE[j])[0];
      var normKey = normalize(key);
      if (normTeacher === normKey) {
        value = MAPPING_TABLE[j][key];
        break;
      }
    }
    newTeacherCode.push({ [teacher]: value });
  }
  Logger.log(newTeacherCode);
}
/* ----------------Test Cases End---------------- */
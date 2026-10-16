var SS_ID = "1HefVEeRj8eX9VSk628OgXMX2KG471yK999O9PC_ONSM";

function doGet(e) {
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutputFromFile('Bridge')
      .setTitle('Phoenix Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("ระบบจัดการโครงสร้างองค์กร - สถาบันฟินิกซ์ (Phoenix Foundation)")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * เวอร์ชันข้อมูลสำหรับซิงก์แบบเรียลไทม์ฝั่ง Browser
 * ใช้เวลาที่ไฟล์ Spreadsheet ถูกแก้ไขล่าสุด โดยไม่ต้องรีเฟรชหน้าเว็บ
 */
function getRealtimeVersion() {
  try {
    return String(DriveApp.getFileById(SS_ID).getLastUpdated().getTime());
  } catch (error) {
    return String(new Date().getTime());
  }
}

function getDefaultHeaders() {
  return [
    "ID",
    "NameTH",
    "NameEN",
    "CmdLevel",
    "Level",
    "SecurityLevel",
    "Locations",
    "PrimaryLocation",
    "ParentID",
    "Description"
  ];
}

function getMainSheet() {
  var ss = secSS_();
  var sheets = ss.getSheets();

  if (!sheets || sheets.length === 0) {
    throw new Error("ไม่พบ Sheet ใน Spreadsheet");
  }

  return sheets[0];
}

/**
 * เตรียม Header โดยจะเขียนกลับลง Sheet เฉพาะเมื่อมีการเปลี่ยนแปลงจริง
 */
function ensureHeaders(sheet) {
  var defaultHeaders = getDefaultHeaders();
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastRow === 0 || lastColumn === 0) {
    sheet
      .getRange(1, 1, 1, defaultHeaders.length)
      .setValues([defaultHeaders]);

    return defaultHeaders;
  }

  var headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  if (!headers[0] || String(headers[0]).trim() === "") {
    sheet
      .getRange(1, 1, 1, defaultHeaders.length)
      .setValues([defaultHeaders]);

    return defaultHeaders;
  }

  var changed = false;

  defaultHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      changed = true;
    }
  });

  if (changed) {
    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  return headers;
}

function normalizeCmdLevel(value) {
  if (
    value === undefined ||
    value === null ||
    String(value).trim() === ""
  ) {
    return "";
  }

  var parsed = parseInt(value, 10);

  if (isNaN(parsed) || parsed < 0 || parsed > 8) {
    return "";
  }

  return String(parsed);
}

function validateCmdLevel(value) {
  if (
    value === undefined ||
    value === null ||
    String(value).trim() === ""
  ) {
    return "";
  }

  var parsed = parseInt(value, 10);

  if (isNaN(parsed) || parsed < 0 || parsed > 8) {
    throw new Error(
      "ระดับบัญชาการต้องเป็น 1-5 หรือเว้นว่างได้"
    );
  }

  return String(parsed);
}

function getOrgData() {
  try {
    var sheet = getMainSheet();
    var headers = ensureHeaders(sheet);
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) return [];

    var values = sheet
      .getRange(2, 1, lastRow - 1, headers.length)
      .getValues();

    return values
      .map(function(row) {
        var obj = {};

        headers.forEach(function(header, index) {
          obj[header] = row[index];
        });

        if (
          (obj.CmdLevel === undefined ||
            obj.CmdLevel === null ||
            String(obj.CmdLevel).trim() === "") &&
          obj.CommandLevel !== undefined
        ) {
          obj.CmdLevel = obj.CommandLevel;
        }

        obj.CmdLevel = normalizeCmdLevel(obj.CmdLevel);

        return obj;
      })
      .filter(function(obj) {
        return (
          String(obj.ID || "").trim() !== "" ||
          String(obj.NameTH || "").trim() !== "" ||
          String(obj.NameEN || "").trim() !== ""
        );
      });

  } catch (error) {
    throw new Error(
      "ไม่สามารถโหลดข้อมูลได้: " + error.message
    );
  }
}

function saveOrgData(data) {
  try {
    if (!data) {
      throw new Error("ไม่พบข้อมูลที่ต้องการบันทึก");
    }

    var sheet = getMainSheet();
    var headers = ensureHeaders(sheet);
    var idColumnIndex = headers.indexOf("ID");

    if (idColumnIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ ID ใน Sheet");
    }

    if (
      (!data.CmdLevel ||
        String(data.CmdLevel).trim() === "") &&
      data.CommandLevel !== undefined
    ) {
      data.CmdLevel = data.CommandLevel;
    }

    data.CmdLevel = validateCmdLevel(data.CmdLevel);

    var id = data.ID
      ? String(data.ID).trim()
      : "";

    var lastRow = sheet.getLastRow();
    var targetRow = -1;

    if (id !== "" && lastRow >= 2) {
      var idValues = sheet
        .getRange(
          2,
          idColumnIndex + 1,
          lastRow - 1,
          1
        )
        .getValues();

      for (var i = 0; i < idValues.length; i++) {
        if (
          String(idValues[i][0]).trim() === id
        ) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow === -1) {
      data.ID = Utilities.getUuid();
      targetRow = sheet.getLastRow() + 1;
    } else {
      data.ID = id;
    }

    var rowData = headers.map(function(header) {
      return data[header] !== undefined &&
        data[header] !== null
        ? data[header]
        : "";
    });

    sheet
      .getRange(
        targetRow,
        1,
        1,
        headers.length
      )
      .setValues([rowData]);

    return {
      success: true,
      mode: id ? "update" : "insert",
      message: id
        ? "อัปเดตข้อมูลเรียบร้อยแล้ว"
        : "บันทึกข้อมูลเรียบร้อยแล้ว",
      id: data.ID,
      data: data
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถบันทึกข้อมูลได้: " + error.message
    );
  }
}

/**
 * บันทึกลำดับทั้งหมดด้วยการเขียนครั้งเดียว
 */
function saveReorderedData(data) {
  try {
    if (!Array.isArray(data)) {
      throw new Error(
        "รูปแบบข้อมูลสำหรับจัดลำดับไม่ถูกต้อง"
      );
    }

    var sheet = getMainSheet();
    var headers = ensureHeaders(sheet);
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        message: "ไม่มีข้อมูลสำหรับจัดลำดับ"
      };
    }

    data.forEach(function(item) {
      if (!item) return;

      if (
        (!item.CmdLevel ||
          String(item.CmdLevel).trim() === "") &&
        item.CommandLevel !== undefined
      ) {
        item.CmdLevel = item.CommandLevel;
      }

      item.CmdLevel =
        validateCmdLevel(item.CmdLevel);
    });

    var rowsToWrite = data.map(function(item) {
      return headers.map(function(header) {
        return item[header] !== undefined &&
          item[header] !== null
          ? item[header]
          : "";
      });
    });

    var existingRowCount = lastRow - 1;
    var maxColumns = Math.max(
      sheet.getLastColumn(),
      headers.length
    );

    if (existingRowCount > 0) {
      sheet
        .getRange(
          2,
          1,
          existingRowCount,
          maxColumns
        )
        .clearContent();
    }

    if (rowsToWrite.length > 0) {
      sheet
        .getRange(
          2,
          1,
          rowsToWrite.length,
          headers.length
        )
        .setValues(rowsToWrite);
    }

    return {
      success: true,
      message: "บันทึกลำดับข้อมูลเรียบร้อยแล้ว"
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถบันทึกลำดับข้อมูลได้: " +
      error.message
    );
  }
}

function deleteOrgData(id) {
  try {
    if (!id) {
      throw new Error("ไม่พบ ID ที่ต้องการลบ");
    }

    id = String(id).trim();

    var sheet = getMainSheet();
    var headers = ensureHeaders(sheet);
    var lastRow = sheet.getLastRow();
    var idColumnIndex = headers.indexOf("ID");

    if (lastRow < 2) {
      throw new Error("ไม่พบข้อมูลที่ต้องการลบ");
    }

    if (idColumnIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ ID ใน Sheet");
    }

    var idValues = sheet
      .getRange(
        2,
        idColumnIndex + 1,
        lastRow - 1,
        1
      )
      .getValues();

    var targetRow = -1;

    for (var i = 0; i < idValues.length; i++) {
      if (
        String(idValues[i][0]).trim() === id
      ) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(
        "ไม่พบข้อมูล ID: " + id
      );
    }

    sheet.deleteRow(targetRow);

    return {
      success: true,
      message: "ลบข้อมูลเรียบร้อยแล้ว",
      id: id
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถลบข้อมูลได้: " +
      error.message
    );
  }
}

function testConnection() {
  try {
    var ss = secSS_();
    var sheet = ss.getSheets()[0];

    return {
      success: true,
      spreadsheetName: ss.getName(),
      sheetName: sheet.getName(),
      message: "เชื่อมต่อ Spreadsheet สำเร็จ"
    };

  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}


/* =========================================================
   ข้อมูลระดับสถาบัน / Organization Profile
========================================================= */

function getOrganizationProfileSheet_() {
  var ss = secSS_();
  var sheet =
    ss.getSheetByName("OrganizationInfo");

  if (!sheet) {
    sheet = ss.insertSheet("OrganizationInfo");

    sheet
      .getRange(1, 1, 1, 8)
      .setValues([[
        "NameTH",
        "NameEN",
        "Role",
        "Mission",
        "Responsibilities",
        "Vision",
        "Scope",
        "Additional"
      ]]);
  }

  return sheet;
}

function getOrganizationProfile() {
  try {
    var ss = secSS_();
    var sheet =
      ss.getSheetByName("OrganizationInfo");

    var defaultProfile = {
      NameTH: "สถาบันฟินิกซ์",
      NameEN:
        "Phoenix Foundation Executive Command",
      Role: "",
      Mission: "",
      Responsibilities: "",
      Vision: "",
      Scope: "",
      Additional: ""
    };

    if (!sheet || sheet.getLastRow() < 2) {
      return defaultProfile;
    }

    var values = sheet
      .getRange(2, 1, 1, 8)
      .getValues()[0];

    var keys = [
      "NameTH",
      "NameEN",
      "Role",
      "Mission",
      "Responsibilities",
      "Vision",
      "Scope",
      "Additional"
    ];

    keys.forEach(function(key, i) {
      defaultProfile[key] =
        values[i] === null ||
        values[i] === undefined
          ? ""
          : String(values[i]);
    });

    if (!defaultProfile.NameTH) {
      defaultProfile.NameTH =
        "สถาบันฟินิกซ์";
    }

    if (!defaultProfile.NameEN) {
      defaultProfile.NameEN =
        "Phoenix Foundation Executive Command";
    }

    return defaultProfile;

  } catch (error) {
    throw new Error(
      "ไม่สามารถโหลดรายละเอียดสถาบันได้: " +
      error.message
    );
  }
}

function saveOrganizationProfile(data) {
  try {
    if (!data) {
      throw new Error(
        "ไม่พบข้อมูลรายละเอียดสถาบัน"
      );
    }

    var sheet =
      getOrganizationProfileSheet_();

    var row = [[
      data.NameTH || "สถาบันฟินิกซ์",
      data.NameEN ||
        "Phoenix Foundation Executive Command",
      data.Role || "",
      data.Mission || "",
      data.Responsibilities || "",
      data.Vision || "",
      data.Scope || "",
      data.Additional || ""
    ]];

    if (sheet.getLastRow() < 2) {
      sheet
        .getRange(2, 1, 1, 8)
        .setValues(row);
    } else {
      sheet
        .getRange(2, 1, 1, 8)
        .setValues(row);
    }

    return {
      success: true,
      message:
        "บันทึกรายละเอียดสถาบันเรียบร้อยแล้ว",
      data: {
        NameTH: row[0][0],
        NameEN: row[0][1],
        Role: row[0][2],
        Mission: row[0][3],
        Responsibilities: row[0][4],
        Vision: row[0][5],
        Scope: row[0][6],
        Additional: row[0][7]
      }
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถบันทึกรายละเอียดสถาบันได้: " +
      error.message
    );
  }
}


/* =========================================================
   ระบบจัดการตำแหน่ง
   Sheet: Positions
========================================================= */

function getPositionsSheet_() {
  var ss = secSS_();
  var sheet =
    ss.getSheetByName("Positions");

  if (!sheet) {
    sheet = ss.insertSheet("Positions");

    sheet
      .getRange(1, 1, 1, 6)
      .setValues([[
        "PositionID",
        "OrganizationID",
        "LineOfWork",
        "PositionName",
        "Tier",
        "SortOrder"
      ]]);
  }

  return sheet;
}

// Tier ของเจ้าหน้าที่รองรับตั้งแต่ Tier 0 ถึง Tier 8
function normalizeTier_(value) {
  if (
    value === undefined ||
    value === null ||
    String(value).trim() === ""
  ) {
    return "";
  }

  var parsed = parseInt(value, 10);

  if (
    isNaN(parsed) ||
    parsed < 0 ||
    parsed > 8
  ) {
    return "";
  }

  return String(parsed);
}

function validateTier_(value) {
  // Tier 0 เป็นค่าที่ถูกต้อง และต้องไม่ถูกมองเป็นค่าว่าง
  if (
    value === undefined ||
    value === null ||
    String(value).trim() === ""
  ) {
    throw new Error("Tier ต้องเป็น 0-8");
  }

  var parsed = parseInt(value, 10);

  if (
    isNaN(parsed) ||
    parsed < 0 ||
    parsed > 8
  ) {
    throw new Error("Tier ต้องเป็น 0-8");
  }

  return String(parsed);
}


/**
 * สายงานที่ระบบอนุญาต
 */
function validateLineOfWork_(value) {
  var allowed = [
    "สายบริหารระดับสูง",
    "ฝ่ายบริหาร",
    "เจ้าหน้าที่ปฏิบัติการ",
    "เจ้าหน้าที่พิเศษ"
  ];

  var text =
    value === undefined ||
    value === null
      ? ""
      : String(value).trim();

  if (text === "") return "";

  if (allowed.indexOf(text) === -1) {
    throw new Error("สายงานไม่ถูกต้อง");
  }

  return text;
}

function getPositionsByOrganization(
  organizationId
) {
  try {
    if (!organizationId) return [];

    var sheet =
      getPositionsSheet_();

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) return [];

    var values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getValues();

    return values
      .map(function(row) {
        return {
          PositionID:
            String(row[0] || ""),
          OrganizationID:
            String(row[1] || ""),
          LineOfWork:
            String(row[2] || ""),
          PositionName:
            String(row[3] || ""),
          Tier:
            normalizeTier_(row[4]),
          SortOrder:
            Number(row[5]) || 0
        };
      })
      .filter(function(item) {
        return (
          item.OrganizationID ===
          String(organizationId).trim()
        );
      })
      .sort(function(a, b) {
        return a.SortOrder - b.SortOrder;
      });

  } catch (error) {
    throw new Error(
      "ไม่สามารถโหลดตำแหน่งได้: " +
      error.message
    );
  }
}

function savePositions(
  organizationId,
  positions
) {
  try {
    if (!organizationId) {
      throw new Error(
        "ไม่พบ Organization ID"
      );
    }

    if (!Array.isArray(positions)) {
      throw new Error(
        "รูปแบบข้อมูลตำแหน่งไม่ถูกต้อง"
      );
    }

    var sheet =
      getPositionsSheet_();

    var lastRow =
      sheet.getLastRow();

    var orgId =
      String(organizationId).trim();

    /*
     * ลบตำแหน่งเดิมของหน่วยงานนี้
     */
    if (lastRow >= 2) {
      var values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            6
          )
          .getValues();

      for (
        var i = values.length - 1;
        i >= 0;
        i--
      ) {
        if (
          String(
            values[i][1] || ""
          ).trim() === orgId
        ) {
          sheet.deleteRow(i + 2);
        }
      }
    }

    if (positions.length === 0) {
      return {
        success: true,
        message:
          "ลบตำแหน่งของหน่วยงานเรียบร้อยแล้ว",
        count: 0
      };
    }

    var rows =
      positions.map(function(
        position,
        index
      ) {
        var line =
          validateLineOfWork_(
            position.LineOfWork
          );

        var name =
          String(
            position.PositionName || ""
          ).trim();

        if (!name) {
          throw new Error(
            "กรุณาระบุชื่อตำแหน่ง"
          );
        }

        var tier =
          validateTier_(
            position.Tier
          );

        var positionId =
          String(
            position.PositionID || ""
          ).trim() ||
          Utilities.getUuid();

        return [
          positionId,
          orgId,
          line,
          name,
          tier,
          index + 1
        ];
      });

    sheet
      .getRange(
        sheet.getLastRow() + 1,
        1,
        rows.length,
        6
      )
      .setValues(rows);

    return {
      success: true,
      message:
        "บันทึกตำแหน่งเรียบร้อยแล้ว",
      count: rows.length
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถบันทึกตำแหน่งได้: " +
      error.message
    );
  }
}

function deletePositionsByOrganization(
  organizationId
) {
  try {
    if (!organizationId) {
      return {
        success: true,
        count: 0
      };
    }

    var sheet =
      getPositionsSheet_();

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        count: 0
      };
    }

    var values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getValues();

    var target =
      String(organizationId).trim();

    var deleted = 0;

    for (
      var i = values.length - 1;
      i >= 0;
      i--
    ) {
      if (
        String(
          values[i][1] || ""
        ).trim() === target
      ) {
        sheet.deleteRow(i + 2);
        deleted++;
      }
    }

    return {
      success: true,
      count: deleted
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถลบตำแหน่งได้: " +
      error.message
    );
  }
}


/* =========================================================
   สรุปตำแหน่งทั้งหมด
   อ่านเฉพาะ Sheet Positions
========================================================= */

function getAllPositionsSummary() {
  try {
    var ss =
      secSS_();

    var positionSheet =
      ss.getSheetByName("Positions");

    if (!positionSheet) {
      return [];
    }

    var lastRow =
      positionSheet.getLastRow();

    if (lastRow < 2) {
      return [];
    }

    var values =
      positionSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getValues();

    return values
      .map(function(row) {
        return {
          PositionID:
            String(row[0] || "").trim(),

          OrganizationID:
            String(row[1] || "").trim(),

          LineOfWork:
            String(row[2] || "").trim(),

          PositionName:
            String(row[3] || "").trim(),

          Tier:
            normalizeTier_(row[4]),

          SortOrder:
            Number(row[5]) || 0,

          OrganizationNameTH: "",

          OrganizationNameEN: ""
        };
      })
      .filter(function(item) {
        return (
          item.PositionName !== "" ||
          item.OrganizationID !== ""
        );
      })
      .sort(function(a, b) {
        return a.PositionName.localeCompare(
          b.PositionName,
          "th"
        );
      });

  } catch (error) {
    throw new Error(
      "ไม่สามารถโหลดสรุปตำแหน่งได้: " +
      error.message
    );
  }
}


/* =========================================================
   ฝ่ายบริหาร
   Management Dashboard
========================================================= */

/**
 * จำแนกตำแหน่งฝ่ายบริหาร
 *
 * สำคัญ:
 * - ตรวจจาก "คำขึ้นต้นของชื่อตำแหน่ง"
 * - รองผู้อำนวยการ ต้องตรวจสอบก่อน ผู้อำนวยการ
 * - เสนาธิการ ตรวจจากคำขึ้นต้นเช่นกัน
 */
function classifyManagementPosition_(positionName) {
  var name =
    String(positionName || "").trim();

  if (name.indexOf("รองผู้อำนวยการ") === 0) {
    return "รองผู้อำนวยการ";
  }

  if (name.indexOf("ผู้อำนวยการ") === 0) {
    return "ผู้อำนวยการ";
  }

  if (name.indexOf("เสนาธิการ") === 0) {
    return "เสนาธิการ";
  }

  return "";
}


/**
 * สร้างข้อมูลสรุปฝ่ายบริหาร
 *
 * รวม:
 * 1. ฝ่ายบริหาร
 * 2. สายบริหารระดับสูง
 *
 * ฝ่ายบริหาร:
 * - ผู้อำนวยการ
 * - รองผู้อำนวยการ
 * - เสนาธิการ
 *
 * สายบริหารระดับสูง:
 * - สรุปชื่อตำแหน่งจริงทั้งหมด
 * - นับจำนวนแต่ละชื่อ
 */
function getManagementSummary() {
  try {
    var ss =
      secSS_();

    var positionSheet =
      ss.getSheetByName("Positions");

    if (!positionSheet) {
      return {
        success: true,

        totalManagement: 0,
        totalManagementLine: 0,
        totalSeniorManagementLine: 0,

        management: {
          total: 0,
          director: 0,
          deputyDirector: 0,
          chiefOfStaff: 0,
          positions: [],
          positionSummary: []
        },

        seniorManagement: {
          total: 0,
          positions: [],
          positionSummary: []
        }
      };
    }

    var lastRow =
      positionSheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,

        totalManagement: 0,
        totalManagementLine: 0,
        totalSeniorManagementLine: 0,

        management: {
          total: 0,
          director: 0,
          deputyDirector: 0,
          chiefOfStaff: 0,
          positions: [],
          positionSummary: []
        },

        seniorManagement: {
          total: 0,
          positions: [],
          positionSummary: []
        }
      };
    }

    var values =
      positionSheet
        .getRange(
          2,
          1,
          lastRow - 1,
          6
        )
        .getValues();

    var managementPositions = [];
    var seniorManagementPositions = [];

    values.forEach(function(row) {
      var positionId =
        String(row[0] || "").trim();

      var organizationId =
        String(row[1] || "").trim();

      var lineOfWork =
        String(row[2] || "").trim();

      var positionName =
        String(row[3] || "").trim();

      var tier =
        normalizeTier_(row[4]);

      var sortOrder =
        Number(row[5]) || 0;

      if (!positionName) {
        return;
      }

      /*
       * -----------------------------------------
       * ฝ่ายบริหาร
       * -----------------------------------------
       */
      if (lineOfWork === "ฝ่ายบริหาร") {
        var managementType =
          classifyManagementPosition_(
            positionName
          );

        /*
         * เฉพาะตำแหน่งที่ขึ้นต้นด้วย
         * ผู้อำนวยการ / รองผู้อำนวยการ / เสนาธิการ
         */
        if (managementType !== "") {
          managementPositions.push({
            PositionID: positionId,
            OrganizationID: organizationId,
            LineOfWork: lineOfWork,
            PositionName: positionName,
            Tier: tier,
            SortOrder: sortOrder,
            ManagementType: managementType
          });
        }
      }

      /*
       * -----------------------------------------
       * สายบริหารระดับสูง
       * -----------------------------------------
       */
      if (lineOfWork === "สายบริหารระดับสูง") {
        seniorManagementPositions.push({
          PositionID: positionId,
          OrganizationID: organizationId,
          LineOfWork: lineOfWork,
          PositionName: positionName,
          Tier: tier,
          SortOrder: sortOrder
        });
      }
    });


    /* =====================================================
       สรุปฝ่ายบริหาร
    ===================================================== */

    var directorCount =
      managementPositions.filter(function(item) {
        return item.ManagementType ===
          "ผู้อำนวยการ";
      }).length;

    var deputyDirectorCount =
      managementPositions.filter(function(item) {
        return item.ManagementType ===
          "รองผู้อำนวยการ";
      }).length;

    var chiefOfStaffCount =
      managementPositions.filter(function(item) {
        return item.ManagementType ===
          "เสนาธิการ";
      }).length;


    /*
     * สรุปชื่อของตำแหน่งฝ่ายบริหาร
     */
    var managementSummaryMap = {};

    managementPositions.forEach(function(item) {
      var key =
        item.PositionName;

      if (!managementSummaryMap[key]) {
        managementSummaryMap[key] = {
          PositionName: key,
          Count: 0,
          ManagementType:
            item.ManagementType
        };
      }

      managementSummaryMap[key].Count++;
    });

    var managementPositionSummary =
      Object.keys(
        managementSummaryMap
      )
      .map(function(key) {
        return managementSummaryMap[key];
      })
      .sort(function(a, b) {
        return a.PositionName.localeCompare(
          b.PositionName,
          "th"
        );
      });


    /* =====================================================
       สรุปสายบริหารระดับสูง
    ===================================================== */

    var seniorSummaryMap = {};

    seniorManagementPositions.forEach(function(item) {
      var key =
        item.PositionName;

      if (!seniorSummaryMap[key]) {
        seniorSummaryMap[key] = {
          PositionName: key,
          Count: 0
        };
      }

      seniorSummaryMap[key].Count++;
    });

    var seniorPositionSummary =
      Object.keys(
        seniorSummaryMap
      )
      .map(function(key) {
        return seniorSummaryMap[key];
      })
      .sort(function(a, b) {
        return a.PositionName.localeCompare(
          b.PositionName,
          "th"
        );
      });


    /*
     * เรียงรายการตำแหน่งจริง
     */
    managementPositions.sort(function(a, b) {
      return (
        a.PositionName.localeCompare(
          b.PositionName,
          "th"
        ) ||
        a.SortOrder - b.SortOrder
      );
    });

    seniorManagementPositions.sort(function(a, b) {
      return (
        a.PositionName.localeCompare(
          b.PositionName,
          "th"
        ) ||
        a.SortOrder - b.SortOrder
      );
    });


    /*
     * จำนวนรวมฝ่ายบริหาร
     * = ฝ่ายบริหาร + สายบริหารระดับสูง
     */
    var totalManagementLine =
      managementPositions.length;

    var totalSeniorManagementLine =
      seniorManagementPositions.length;

    var totalManagement =
      totalManagementLine +
      totalSeniorManagementLine;


    return {
      success: true,

      /*
       * รวมทั้งหมด
       */
      totalManagement:
        totalManagement,

      /*
       * เฉพาะฝ่ายบริหาร
       */
      totalManagementLine:
        totalManagementLine,

      /*
       * เฉพาะสายบริหารระดับสูง
       */
      totalSeniorManagementLine:
        totalSeniorManagementLine,

      /*
       * ฝ่ายบริหาร
       */
      management: {
        total:
          totalManagementLine,

        director:
          directorCount,

        deputyDirector:
          deputyDirectorCount,

        chiefOfStaff:
          chiefOfStaffCount,

        positions:
          managementPositions,

        positionSummary:
          managementPositionSummary
      },

      /*
       * สายบริหารระดับสูง
       */
      seniorManagement: {
        total:
          totalSeniorManagementLine,

        positions:
          seniorManagementPositions,

        positionSummary:
          seniorPositionSummary
      }
    };

  } catch (error) {
    throw new Error(
      "ไม่สามารถโหลดสรุปฝ่ายบริหารได้: " +
      error.message
    );
  }
}

/* =========================================================
   SECRETARIAT SYSTEM — TEAM / MEMBER / TIER / BUDGET
   ใช้ Spreadsheet เดียวกับระบบโครงสร้างองค์กร แต่แยก Sheet
========================================================= */

var SEC_SHEETS = {
  TEAMS: 'SEC_TEAMS',
  MEMBERS: 'SEC_MEMBERS',
  TIER_HISTORY: 'SEC_TIER_HISTORY',
  MEMBER_HISTORY: 'SEC_MEMBER_HISTORY',
  BUDGETS: 'SEC_BUDGETS',
  BUDGET_UNITS: 'SEC_BUDGET_UNITS',
  BUDGET_CONFIG: 'SEC_BUDGET_CONFIG',
  BUDGET_SOURCES: 'SEC_BUDGET_SOURCES',
  AUDIT: 'SEC_AUDIT_LOG'
};

function secSheet_(name, headers) {
  var ss = secSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    var existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    var changed = false;
    headers.forEach(function(h){ if(existing.indexOf(h)===-1){ existing.push(h); changed=true; }});
    if(changed) sh.getRange(1,1,1,existing.length).setValues([existing]);
  }
  return sh;
}

var SEC_SS_CACHE_=null;
function secSS_(){
  if(!SEC_SS_CACHE_) SEC_SS_CACHE_=SpreadsheetApp.openById(SS_ID);
  return SEC_SS_CACHE_;
}
const SEC_CENTRAL_BUDGET_SHEET='SEC_CENTRAL_BUDGETS';
const SEC_EMERGENCY_EXPENSE_SHEET='SEC_EMERGENCY_EXPENSES';
var SEC_SHEETS_INITIALIZED_=false;
function secInitSheets() {
  if(SEC_SHEETS_INITIALIZED_) return true;
  secSheet_(SEC_SHEETS.TEAMS,['TeamID','Code','Organization','OrganizationID','OrganizationParentID','Name','Head','Tier','Status','Description','CreatedAt','UpdatedAt']);
  secSheet_(SEC_SHEETS.MEMBERS,['MemberID','Code','Rank','Name','TeamID','MemberOrder','Organization','OrganizationID','OrganizationParentID','Role','Tier','Status','IndividualCapability','Note','JoinedAt','UpdatedAt']);
  secSheet_(SEC_SHEETS.TIER_HISTORY,['HistoryID','MemberID','MemberName','OldTier','NewTier','Reason','ApprovedBy','ChangedAt']);
  secSheet_(SEC_SHEETS.MEMBER_HISTORY,['HistoryID','MemberID','MemberName','ChangeType','Field','OldValue','NewValue','ChangedAt']);
  secSheet_(SEC_SHEETS.BUDGETS,['BudgetID','FiscalYear','Organization','Item','Amount','Status','Requester','Note','CreatedAt','UpdatedAt']);
  secSheet_(SEC_SHEETS.BUDGET_UNITS,['BudgetID','FiscalYear','OrganizationID','ParentID','Organization','ParentOrganization','BudgetAmount','ReserveAmount','Note','CreatedAt','UpdatedAt']);
  secSheet_(SEC_SHEETS.BUDGET_CONFIG,['ConfigID','FiscalYear','TotalBudget','ReserveBudget','UpdatedAt']);
  secSheet_(SEC_SHEETS.BUDGET_SOURCES,['SourceID','FiscalYear','SourceName','Amount','BudgetPart','Note','CreatedAt','UpdatedAt']);
  secSheet_(SEC_CENTRAL_BUDGET_SHEET,['CentralID','FiscalYear','Item','MonthlyAmount','Months','Amount','Note','CreatedAt','UpdatedAt']);
  secSheet_(SEC_EMERGENCY_EXPENSE_SHEET,['ExpenseID','FiscalYear','Item','Amount','BudgetPart','Detail','CreatedAt','UpdatedAt']);
  secSheet_(SEC_SHEETS.AUDIT,['AuditID','Actor','Action','EntityType','EntityID','Detail','CreatedAt']);
  SEC_SHEETS_INITIALIZED_=true;
  return true;
}
function secRows_(name) {
  secInitSheets();
  var sh=secSS_().getSheetByName(name);
  if(!sh) throw new Error('ไม่พบ Sheet: '+name);
  var last=sh.getLastRow();
  if(last<2) return [];

  var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var v=sh.getRange(2,1,last-1,h.length).getValues();

  return v
    .filter(function(r){
      return r.some(function(x){return String(x===null||x===undefined?'':x).trim()!=='';});
    })
    .map(function(r){
      var o={};
      h.forEach(function(k,i){
        var value=r[i];

        /*
         * สำคัญสำหรับ google.script.run:
         * Date object ไม่สามารถส่งกลับไปยัง HTML Service ได้โดยตรง
         * จึงแปลงวันที่ทั้งหมดเป็นข้อความ ISO ก่อนส่งให้หน้าเว็บ
         */
        if(value instanceof Date){
          o[k]=value.toISOString();
        }else{
          o[k]=value;
        }
      });
      return o;
    });
}
function secId_(prefix){return prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase();}
function secAppend_(name,obj){
  secInitSheets();
  var sh=secSS_().getSheetByName(name);
  if(!sh) throw new Error('ไม่พบ Sheet: '+name);
  var lastColumn=sh.getLastColumn();
  if(lastColumn<1) throw new Error('Sheet '+name+' ยังไม่มีหัวตาราง');
  var h=sh.getRange(1,1,1,lastColumn).getValues()[0];
  sh.appendRow(h.map(function(k){return obj[k]!==undefined?obj[k]:'';}));
  // ไม่ flush ทุก append: Apps Script จะ commit การเปลี่ยนแปลงเมื่อจบ execution อยู่แล้ว
  return true;
}
function secAudit_(action,type,id,detail){
  try {
    secAppend_(SEC_SHEETS.AUDIT,{AuditID:secId_('AUD'),Actor:'สภาผู้บริหาร',Action:action,EntityType:type,EntityID:id,Detail:detail||'',CreatedAt:new Date()});
  } catch(e) {
    console.warn('SEC audit log failed: '+e.message);
  }
}
function secUpdateRow_(name,idField,id,obj){var sh=secSS_().getSheetByName(name),h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),h.length).getValues();for(var i=0;i<vals.length;i++){if(String(vals[i][h.indexOf(idField)])===String(id)){h.forEach(function(k,j){if(obj[k]!==undefined) vals[i][j]=obj[k];});sh.getRange(i+2,1,1,h.length).setValues([vals[i]]);return true;}}return false;}

function secGetDashboard(){
  var teams=secRows_(SEC_SHEETS.TEAMS), members=secRows_(SEC_SHEETS.MEMBERS), budgets=secRows_(SEC_SHEETS.BUDGETS);
  var total=budgets.reduce(function(s,b){return s+Number(b.Amount||0);},0), pending=budgets.filter(function(b){return b.Status==='รออนุมัติ'||!b.Status;}).length;
  return {teams:teams.length,members:members.length,budget:total,pending:pending,recentTeams:teams.slice(-6).reverse(),recentBudgets:budgets.slice(-8).reverse()};
}
function secGetTeams(){
  try {
    return secRows_(SEC_SHEETS.TEAMS);
  } catch(e) {
    throw new Error('ไม่สามารถโหลดข้อมูลทีมจาก SEC_TEAMS: '+e.message);
  }
}
function secGetMembers(){return secRows_(SEC_SHEETS.MEMBERS);}
function secGetTierHistory(){return secRows_(SEC_SHEETS.TIER_HISTORY).reverse();}
function secGetBudgets(){return secRows_(SEC_SHEETS.BUDGETS).reverse();}
function secGetTeamDetail(id){var t=secGetTeams().find(function(x){return String(x.TeamID)===String(id);});if(!t)return null;t.Members=secGetMembers().filter(function(m){return String(m.TeamID)===String(id);});return t;}
function secGetMemberDetail(id){var m=secGetMembers().find(function(x){return String(x.MemberID)===String(id);});if(!m)return null;m.TeamName=(secGetTeams().find(function(t){return String(t.TeamID)===String(m.TeamID);})||{}).Name||'-';m.History=secRows_(SEC_SHEETS.TIER_HISTORY).filter(function(h){return String(h.MemberID)===String(id);}).reverse();return m;}

function secGetOrganizationOptions(){
  try {
    var data=getOrgData();
    return data.map(function(x){
      return {
        ID:String(x.ID||''),
        NameTH:String(x.NameTH||''),
        NameEN:String(x.NameEN||''),
        Level:String(x.Level||''),
        ParentID:String(x.ParentID||''),
        Locations:String(x.Locations||'')
      };
    }).filter(function(x){return x.ID && x.NameTH;});
  } catch(e) {
    throw new Error('ไม่สามารถโหลดข้อมูลโครงสร้างองค์กรสำหรับสำนักเลขา: '+e.message);
  }
}

function secGetBudgetOrganizationTree(){
  var data=secGetOrganizationOptions();
  var roots=data.filter(function(x){
    var level=String(x.Level||'').trim();
    return level==='หน่วยงานหลัก' || level==='ส่วนปฏิบัติการพิเศษ';
  });
  return roots.map(function(root){
    var children=data.filter(function(x){
      return String(x.Level||'').trim()==='หน่วยงานย่อย' &&
             String(x.ParentID||'').trim()===String(root.ID).trim();
    });
    return {Root:root,Children:children};
  });
}

function secGetPositionOptions(){
  try{
    var sh=getPositionsSheet_(), last=sh.getLastRow();
    if(last<2)return [];
    var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var rows=sh.getRange(2,1,last-1,h.length).getValues();
    return rows.filter(function(r){return r.some(function(v){return String(v||'').trim()!=='';});}).map(function(r){
      var o={};h.forEach(function(k,i){o[k]=r[i];});
      return {PositionID:String(o.PositionID||''),OrganizationID:String(o.OrganizationID||''),LineOfWork:String(o.LineOfWork||''),PositionName:String(o.PositionName||''),Tier:String(o.Tier||''),SortOrder:Number(o.SortOrder||0)};
    }).filter(function(x){return x.OrganizationID&&x.PositionName;});
  }catch(e){throw new Error('ไม่สามารถโหลดตำแหน่งจากโครงสร้างองค์กร: '+e.message);}
}

function secDeleteTeam(id){
  secInitSheets();
  id=String(id||'').trim();
  if(!id)throw new Error('ไม่พบรหัสทีม');
  var ss=secSS_();
  var teamSh=ss.getSheetByName(SEC_SHEETS.TEAMS), memSh=ss.getSheetByName(SEC_SHEETS.MEMBERS), histSh=ss.getSheetByName(SEC_SHEETS.TIER_HISTORY);
  var team=secGetTeams().find(function(x){return String(x.TeamID)===id;});
  if(!team)throw new Error('ไม่พบทีมที่ต้องการลบ');
  var memberIds=secGetMembers().filter(function(m){return String(m.TeamID)===id;}).map(function(m){return String(m.MemberID);});
  function deleteBy_(sh,field,values){
    if(!sh||sh.getLastRow()<2||!values.length)return;
    var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], idx=h.indexOf(field);if(idx<0)return;
    var vals=sh.getRange(2,1,sh.getLastRow()-1,h.length).getValues(), keep=[];
    vals.forEach(function(r){if(values.indexOf(String(r[idx]))===-1)keep.push(r);});
    sh.getRange(2,1,sh.getLastRow()-1,h.length).clearContent();
    if(keep.length)sh.getRange(2,1,keep.length,h.length).setValues(keep);
  }
  deleteBy_(memSh,'TeamID',[id]);
  deleteBy_(histSh,'MemberID',memberIds);
  deleteBy_(teamSh,'TeamID',[id]);
  secAudit_('ลบทีม','TEAM',id,team.Name+' / สมาชิก '+memberIds.length+' คน');
  SpreadsheetApp.flush();
  return {success:true,TeamID:id,deletedMembers:memberIds.length};
}

function secSaveTeam(data){
  secInitSheets();
  data=data||{};
  var now=new Date();
  if(!String(data.Name||'').trim()) throw new Error('กรุณาระบุชื่อทีม');

  data.Name=String(data.Name).trim();
  data.Code='';
  data.Organization=String(data.Organization||'').trim();
  data.OrganizationID=String(data.OrganizationID||'').trim();
  data.OrganizationParentID=String(data.OrganizationParentID||'').trim();
  data.Tier=String(data.Tier||'1');
  data.Status=String(data.Status||'ใช้งาน');
  data.Description=String(data.Description||'').trim();

  if(data.TeamID){
    data.UpdatedAt=now;
    if(!secUpdateRow_(SEC_SHEETS.TEAMS,'TeamID',data.TeamID,data)) throw new Error('ไม่พบทีมที่ต้องการแก้ไข');
    secAudit_('แก้ไขทีม','TEAM',data.TeamID,data.Name);
    SpreadsheetApp.flush();
    return {success:true,action:'update',TeamID:data.TeamID};
  }

  data.TeamID=secId_('TEAM');
  data.CreatedAt=now;
  data.UpdatedAt=now;
  secAppend_(SEC_SHEETS.TEAMS,data);
  secAudit_('สร้างทีม','TEAM',data.TeamID,data.Name);
  SpreadsheetApp.flush();
  return {success:true,action:'create',TeamID:data.TeamID};
}
function secSaveMember(data){
  secInitSheets();
  data=data||{};
  var now=new Date();
  data.Rank=String(data.Rank||'').trim();
  data.Name=String(data.Name||'').trim();
  data.Code=String(data.Code||'').trim();
  data.TeamID=String(data.TeamID||'').trim();
  data.MemberOrder=Number(data.MemberOrder||0);
  data.Organization=String(data.Organization||'').trim();
  data.OrganizationID=String(data.OrganizationID||'').trim();
  data.OrganizationParentID=String(data.OrganizationParentID||'').trim();
  data.Role=String(data.Role||'').trim();
  data.Tier=String(data.Tier||'1');
  data.Status=String(data.Status||'ปฏิบัติงาน');
  data.IndividualCapability=String(data.IndividualCapability||'').trim();
  data.Note=String(data.Note||'').trim();
  if(!data.Name) throw new Error('กรุณาระบุชื่อสมาชิก');
  if(!data.TeamID) throw new Error('กรุณาระบุทีมปฏิบัติการของสมาชิก');
  if(data.MemberID){
    var oldMember=secGetMembers().find(function(x){return String(x.MemberID)===String(data.MemberID);});
    if(!oldMember) throw new Error('ไม่พบสมาชิกที่ต้องการแก้ไข');
    data.UpdatedAt=now;
    if(!secUpdateRow_(SEC_SHEETS.MEMBERS,'MemberID',data.MemberID,data)) throw new Error('ไม่พบสมาชิกที่ต้องการแก้ไข');
    var tracked=[
      ['Rank','ยศ'],['Name','ชื่อ-นามสกุล'],['TeamID','ทีมปฏิบัติการ'],['MemberOrder','ลำดับสมาชิก'],
      ['Organization','หน่วยงาน'],['OrganizationID','หน่วยงาน ID'],['OrganizationParentID','หน่วยงานแม่/หลัก ID'],
      ['Role','ตำแหน่ง'],['Tier','Tier'],['Status','สถานะ'],['IndividualCapability','ความสามารถรายบุคคล'],['Note','หมายเหตุ']
    ];
    tracked.forEach(function(pair){var field=pair[0],label=pair[1],before=String(oldMember[field]||''),after=String(data[field]||'');if(before!==after){secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:data.MemberID,MemberName:data.Name,ChangeType:'แก้ไขข้อมูลสมาชิก',Field:label,OldValue:before||'-',NewValue:after||'-',ChangedAt:now});}});
    if(String(oldMember.Tier||'')!==String(data.Tier||'')){secAppend_(SEC_SHEETS.TIER_HISTORY,{HistoryID:secId_('TIER'),MemberID:data.MemberID,MemberName:data.Name,OldTier:String(oldMember.Tier||'-'),NewTier:String(data.Tier||'-'),Reason:'แก้ไขข้อมูลรายบุคคล',ApprovedBy:'สภาผู้บริหาร',ChangedAt:now});}
    secAudit_('แก้ไขสมาชิก','MEMBER',data.MemberID,data.Name);
    SpreadsheetApp.flush();
    return {success:true,action:'update',MemberID:data.MemberID,TeamID:data.TeamID,MemberOrder:data.MemberOrder};
  }
  if(!data.MemberOrder){
    var existing=secGetMembers().filter(function(m){return String(m.TeamID)===data.TeamID;});
    data.MemberOrder=existing.reduce(function(max,m){return Math.max(max,Number(m.MemberOrder||0));},0)+1;
  }
  data.MemberID=secId_('MEM');
  data.JoinedAt=now;
  data.UpdatedAt=now;
  secAppend_(SEC_SHEETS.MEMBERS,data);
  secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:data.MemberID,MemberName:data.Name,ChangeType:'สร้างสมาชิก',Field:'ข้อมูลสมาชิก',OldValue:'-',NewValue:'สร้างสมาชิกใหม่',ChangedAt:now});
  secAudit_('สร้างสมาชิก','MEMBER',data.MemberID,data.Name);
  SpreadsheetApp.flush();
  return {success:true,action:'create',MemberID:data.MemberID,TeamID:data.TeamID,MemberOrder:data.MemberOrder};
}

function secChangeMemberPersonal(data){
  secInitSheets(); data=data||{};
  var id=String(data.MemberID||'').trim(); if(!id)throw new Error('ไม่พบสมาชิก');
  var old=secGetMembers().find(function(x){return String(x.MemberID)===id;}); if(!old)throw new Error('ไม่พบสมาชิกที่ต้องการแก้ไข');
  data.TeamID=String(data.TeamID||old.TeamID||'').trim();
  data.Rank=String(data.Rank||'').trim(); data.Name=String(data.Name||'').trim();
  data.Organization=String(data.Organization||'').trim(); data.OrganizationID=String(data.OrganizationID||'').trim(); data.OrganizationParentID=String(data.OrganizationParentID||'').trim();
  data.Role=String(data.Role||'').trim(); data.Tier=String(data.Tier||'1'); data.Status=String(data.Status||'ปฏิบัติงาน'); data.IndividualCapability=String(data.IndividualCapability||'').trim();
  if(!data.Name)throw new Error('กรุณาระบุชื่อสมาชิก');
  var now=new Date();
  var fields=[
    ['Rank','ยศ'],['Name','ชื่อ-นามสกุล'],['TeamID','ทีมปฏิบัติการ'],['MemberOrder','ลำดับสมาชิก'],
    ['Organization','หน่วยงาน'],['OrganizationID','หน่วยงาน ID'],['OrganizationParentID','หน่วยงานแม่/หลัก ID'],
    ['Role','ตำแหน่ง'],['Tier','Tier'],['Status','สถานะ'],['IndividualCapability','ความสามารถรายบุคคล'],['Note','หมายเหตุ']
  ];
  fields.forEach(function(pair){var f=pair[0],label=pair[1],before=String(old[f]||''),after=String(data[f]||'');if(before!==after){secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:id,MemberName:data.Name,ChangeType:'แก้ไขข้อมูลส่วนบุคคล',Field:label,OldValue:before||'-',NewValue:after||'-',ChangedAt:now});}});
  if(String(old.Tier||'')!==String(data.Tier||'')){secAppend_(SEC_SHEETS.TIER_HISTORY,{HistoryID:secId_('TIER'),MemberID:id,MemberName:data.Name,OldTier:String(old.Tier||'-'),NewTier:String(data.Tier||'-'),Reason:'แก้ไขข้อมูลส่วนบุคคล',ApprovedBy:'สภาผู้บริหาร',ChangedAt:now});}
  var update={Rank:data.Rank,Name:data.Name,Organization:data.Organization,OrganizationID:data.OrganizationID,OrganizationParentID:data.OrganizationParentID,Role:data.Role,Tier:data.Tier,Status:data.Status,IndividualCapability:data.IndividualCapability,UpdatedAt:now};
  if(!secUpdateRow_(SEC_SHEETS.MEMBERS,'MemberID',id,update))throw new Error('ไม่สามารถบันทึกข้อมูลสมาชิกได้');
  secAudit_('แก้ไขข้อมูลส่วนบุคคล','MEMBER',id,data.Name);
  SpreadsheetApp.flush();
  return {success:true,MemberID:id,TeamID:data.TeamID};
}
function secGetMemberHistory(id){
  secInitSheets();
  return secRows_(SEC_SHEETS.MEMBER_HISTORY).filter(function(h){return String(h.MemberID)===String(id);}).reverse();
}
function secChangeMemberPosition(id,newRole){
  secInitSheets();
  id=String(id||'').trim(); newRole=String(newRole||'').trim();
  if(!id)throw new Error('ไม่พบสมาชิก');
  if(!newRole)throw new Error('กรุณาระบุตำแหน่งใหม่');
  var m=secGetMembers().find(function(x){return String(x.MemberID)===id;});
  if(!m)throw new Error('ไม่พบสมาชิกที่ต้องการแก้ไข');
  var oldRole=String(m.Role||'').trim();
  if(oldRole===newRole)throw new Error('ตำแหน่งใหม่ต้องแตกต่างจากตำแหน่งเดิม');
  // ตำแหน่งเป็นข้อมูลแบบอิสระ: ไม่บังคับให้มีอยู่ในชีตตำแหน่ง
  // และไม่บังคับให้สมาชิกต้องมีหน่วยงานก่อนเปลี่ยนตำแหน่ง
  var now=new Date();
  if(!secUpdateRow_(SEC_SHEETS.MEMBERS,'MemberID',id,{Role:newRole,UpdatedAt:now}))throw new Error('ไม่สามารถบันทึกตำแหน่งใหม่ได้');
  secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:id,MemberName:m.Name,ChangeType:'เปลี่ยนตำแหน่ง',Field:'Role',OldValue:oldRole,NewValue:newRole,ChangedAt:now});
  secAudit_('เปลี่ยนตำแหน่งสมาชิก','MEMBER',id,oldRole+' → '+newRole);
  SpreadsheetApp.flush();
  return {success:true,MemberID:id,OldRole:oldRole,NewRole:newRole};
}
function secDeleteMember(id){
  secInitSheets(); id=String(id||'').trim(); if(!id)throw new Error('ไม่พบสมาชิก');
  var m=secGetMembers().find(function(x){return String(x.MemberID)===id;}); if(!m)throw new Error('ไม่พบสมาชิกที่ต้องการลบ');
  var sh=secSS_().getSheetByName(SEC_SHEETS.MEMBERS), h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], idx=h.indexOf('MemberID');
  if(idx<0)throw new Error('ไม่พบคอลัมน์ MemberID');
  for(var r=sh.getLastRow();r>=2;r--){if(String(sh.getRange(r,idx+1).getValue())===id)sh.deleteRow(r);}
  secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:id,MemberName:m.Name,ChangeType:'ลบสมาชิก',Field:'ข้อมูลสมาชิก',OldValue:'สมาชิกในระบบ',NewValue:'ลบสมาชิก',ChangedAt:new Date()});
  secAudit_('ลบสมาชิก','MEMBER',id,(m.Rank?'['+m.Rank+'] ':'')+m.Name);
  SpreadsheetApp.flush(); return {success:true,MemberID:id,TeamID:m.TeamID};
}
function secReorderMembers(teamId,orderedMemberIds){
  secInitSheets();
  teamId=String(teamId||'').trim();
  if(!teamId)throw new Error('ไม่พบทีมปฏิบัติการ');
  if(!Array.isArray(orderedMemberIds))throw new Error('ข้อมูลลำดับสมาชิกไม่ถูกต้อง');
  var members=secGetMembers().filter(function(m){return String(m.TeamID)===teamId;});
  var allowed={}; members.forEach(function(m){allowed[String(m.MemberID)]=true;});
  var seen={};
  orderedMemberIds.forEach(function(id){id=String(id);if(!allowed[id])throw new Error('พบสมาชิกที่ไม่อยู่ในทีม');if(seen[id])throw new Error('พบสมาชิกซ้ำในลำดับ');seen[id]=true;});
  var sh=secSS_().getSheetByName(SEC_SHEETS.MEMBERS);
  var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], idIdx=h.indexOf('MemberID'), orderIdx=h.indexOf('MemberOrder');
  if(idIdx<0||orderIdx<0)throw new Error('ไม่พบคอลัมน์ลำดับสมาชิก');
  var vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),h.length).getValues();
  var orderMap={}; orderedMemberIds.forEach(function(id,i){orderMap[String(id)]=i+1;});
  var now=new Date();
  var memberMap={}; members.forEach(function(m){memberMap[String(m.MemberID)]=m;});
  vals.forEach(function(r){
    var id=String(r[idIdx]||'');
    if(orderMap[id]!==undefined){
      var oldOrder=Number(memberMap[id]&&memberMap[id].MemberOrder||0), newOrder=Number(orderMap[id]);
      if(oldOrder!==newOrder){
        secAppend_(SEC_SHEETS.MEMBER_HISTORY,{HistoryID:secId_('MHIST'),MemberID:id,MemberName:(memberMap[id]||{}).Name||'-',ChangeType:'เปลี่ยนลำดับสมาชิก',Field:'ลำดับสมาชิก',OldValue:String(oldOrder||'-'),NewValue:String(newOrder),ChangedAt:now});
      }
      r[orderIdx]=newOrder;
    }
  });
  if(vals.length)sh.getRange(2,1,vals.length,h.length).setValues(vals);
  SpreadsheetApp.flush();
  secAudit_('จัดลำดับสมาชิก','TEAM',teamId,orderedMemberIds.join(','));
  return {success:true,TeamID:teamId,count:orderedMemberIds.length};
}
function secPromoteMember(id,newTier,reason,approvedBy){var m=secGetMembers().find(function(x){return String(x.MemberID)===String(id);});if(!m)throw new Error('ไม่พบสมาชิก');var old=String(m.Tier||'1'), next=String(newTier);if(!/^[0-8]$/.test(next))throw new Error('Tier ต้องเป็น 0-8');if(old===next)throw new Error('Tier ใหม่ต้องแตกต่างจาก Tier เดิม');secUpdateRow_(SEC_SHEETS.MEMBERS,'MemberID',id,{Tier:next,UpdatedAt:new Date()});secAppend_(SEC_SHEETS.TIER_HISTORY,{HistoryID:secId_('TIER'),MemberID:id,MemberName:m.Name,OldTier:old,NewTier:next,Reason:reason||'',ApprovedBy:approvedBy||'สภาผู้บริหาร',ChangedAt:new Date()});secAudit_('เลื่อน Tier','MEMBER',id,old+' → '+next+' : '+(reason||''));return true;}

function secGetBudgetSources(fiscalYear){
  secInitSheets();
  var year=String(fiscalYear||'').trim();
  if(!year) return [];

  var rows=secRows_(SEC_SHEETS.BUDGET_SOURCES).filter(function(x){
    return String(x.FiscalYear||'')===year;
  });

  /*
   * นำข้อมูล SEC_BUDGET_CONFIG เดิมเข้ามาเป็นแหล่งงบเพียงครั้งเดียวต่อปี
   * เพื่อไม่ให้รายการที่ผู้ใช้ลบแล้วถูกสร้างกลับมาอีก
   */
  var props=PropertiesService.getScriptProperties();
  var markerKey='SEC_BUDGET_SOURCE_MIGRATED_'+year;
  var migrated=props.getProperty(markerKey)==='1';

  if(!migrated){
    var legacy=secRows_(SEC_SHEETS.BUDGET_CONFIG).find(function(x){
      return String(x.FiscalYear||'')===year;
    });

    if(legacy && rows.length===0){
      var total=Number(legacy.TotalBudget||0);
      var reserve=Number(legacy.ReserveBudget||0);
      var now=new Date();

      if(total!==0){
        secAppend_(SEC_SHEETS.BUDGET_SOURCES,{
          SourceID:secId_('BSRC'),
          FiscalYear:year,
          SourceName:'ยอดงบประมาณเดิมจากระบบ',
          Amount:total,
          BudgetPart:'งบประมาณหลัก',
          Note:'นำเข้าจาก SEC_BUDGET_CONFIG อัตโนมัติครั้งเดียว',
          CreatedAt:now,
          UpdatedAt:now
        });
      }

      if(reserve!==0){
        secAppend_(SEC_SHEETS.BUDGET_SOURCES,{
          SourceID:secId_('BSRC'),
          FiscalYear:year,
          SourceName:'ยอดงบสำรองเดิมจากระบบ',
          Amount:reserve,
          BudgetPart:'งบสำรอง',
          Note:'นำเข้าจาก SEC_BUDGET_CONFIG อัตโนมัติครั้งเดียว',
          CreatedAt:now,
          UpdatedAt:now
        });
      }

      SpreadsheetApp.flush();
      rows=secRows_(SEC_SHEETS.BUDGET_SOURCES).filter(function(x){
        return String(x.FiscalYear||'')===year;
      });
    }

    // ตั้ง marker แม้ยอดเดิมเป็น 0 เพื่อป้องกันการพยายาม migrate ซ้ำ
    props.setProperty(markerKey,'1');
  }

  return rows;
}
function secDeleteRow_(name,idField,id){
  var sh=secSS_().getSheetByName(name);
  if(!sh||sh.getLastRow()<2)return false;
  var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], idx=h.indexOf(idField);
  if(idx<0)return false;
  var vals=sh.getRange(2,1,sh.getLastRow()-1,h.length).getValues();
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][idx])===String(id)){sh.deleteRow(i+2);return true;}
  }
  return false;
}
function secGetBudgetConfig(fiscalYear){
  secInitSheets();
  var year=String(fiscalYear||'').trim();
  var sources=secGetBudgetSources(year);

  if(sources.length){
    var total=0,reserve=0;
    sources.forEach(function(x){
      var amount=Number(x.Amount||0);
      if(String(x.BudgetPart||'')==='งบสำรอง') reserve+=amount;
      else total+=amount;
    });
    return {
      ConfigID:'SOURCE-CALC-'+year,
      FiscalYear:year,
      TotalBudget:total,
      ReserveBudget:reserve,
      SourceCount:sources.length
    };
  }

  /*
   * ถ้า migration ของปีนี้เกิดขึ้นแล้วแต่ผู้ใช้ลบแหล่งงบจนหมด
   * ให้ถือว่างบของปีนั้นเป็น 0 ไม่ย้อนกลับไปใช้ค่า legacy
   */
  var props=PropertiesService.getScriptProperties();
  if(props.getProperty('SEC_BUDGET_SOURCE_MIGRATED_'+year)==='1'){
    return {
      ConfigID:'SOURCE-CALC-'+year,
      FiscalYear:year,
      TotalBudget:0,
      ReserveBudget:0,
      SourceCount:0
    };
  }

  var rows=secRows_(SEC_SHEETS.BUDGET_CONFIG);
  var found=rows.find(function(x){return String(x.FiscalYear||'')===year;});
  return found||{
    ConfigID:'',
    FiscalYear:year,
    TotalBudget:0,
    ReserveBudget:0,
    SourceCount:0
  };
}
function secSaveBudgetSource(data){
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  var sourceName=String(data.SourceName||'').trim();
  var part=String(data.BudgetPart||'').trim();
  var amount=Number(data.Amount||0);
  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  if(!sourceName)throw new Error('กรุณาระบุแหล่งที่มาของงบประมาณ');
  if(part!=='งบประมาณหลัก'&&part!=='งบสำรอง')throw new Error('กรุณาเลือกส่วนของงบ');
  if(!isFinite(amount)||amount<=0)throw new Error('จำนวนเงินต้องมากกว่า 0');
  var now=new Date();
  var payload={FiscalYear:year,SourceName:sourceName,Amount:amount,BudgetPart:part,Note:String(data.Note||'').trim(),UpdatedAt:now};
  var id=String(data.SourceID||'').trim();
  if(id){
    payload.SourceID=id;
    if(!secUpdateRow_(SEC_SHEETS.BUDGET_SOURCES,'SourceID',id,payload))throw new Error('ไม่สามารถแก้ไขรายการแหล่งงบได้');
    secAudit_('แก้ไขแหล่งงบประมาณ','BUDGET_SOURCE',id,sourceName+' | '+part+' | '+amount);
  }else{
    payload.SourceID=secId_('BSRC');
    payload.CreatedAt=now;
    secAppend_(SEC_SHEETS.BUDGET_SOURCES,payload);
    secAudit_('เพิ่มแหล่งงบประมาณ','BUDGET_SOURCE',payload.SourceID,sourceName+' | '+part+' | '+amount);
  }
  SpreadsheetApp.flush();
  // google.script.run ไม่สามารถส่ง Date object กลับไปยัง HTML Service ได้
  // จึงส่งเฉพาะข้อมูลที่ serializable และให้หน้าเว็บใช้ local row ที่สร้างไว้แล้ว
  return {success:true,SourceID:payload.SourceID,FiscalYear:year,Config:secGetBudgetConfig(year)};
}
function secDeleteBudgetSource(sourceId){
  secInitSheets();
  var id=String(sourceId||'').trim();
  if(!id)throw new Error('ไม่พบรายการแหล่งงบ');
  var sh=secSS_().getSheetByName(SEC_SHEETS.BUDGET_SOURCES);
  if(!sh||sh.getLastRow()<2)throw new Error('ไม่พบรายการแหล่งงบในระบบ');
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idCol=headers.indexOf('SourceID')+1;
  if(!idCol)throw new Error('ไม่พบคอลัมน์ SourceID');
  var last=sh.getLastRow();
  var ids=sh.getRange(2,idCol,last-1,1).getDisplayValues();
  var rowNumber=0;
  for(var i=0;i<ids.length;i++){
    if(String(ids[i][0]||'').trim()===id){rowNumber=i+2;break;}
  }
  if(!rowNumber)throw new Error('ไม่พบรายการแหล่งงบที่ต้องการลบ');
  var row=sh.getRange(rowNumber,1,1,sh.getLastColumn()).getValues()[0];
  var obj={};headers.forEach(function(h,j){obj[h]=row[j];});
  sh.deleteRow(rowNumber);
  SpreadsheetApp.flush();
  try{secAudit_('ลบแหล่งงบประมาณ','BUDGET_SOURCE',id,String(obj.SourceName||'')+' | '+String(obj.BudgetPart||'')+' | '+String(obj.Amount||''));}catch(e){}
  var nextConfig=secGetBudgetConfig(String(obj.FiscalYear||''));
  return {success:true,FiscalYear:String(obj.FiscalYear||''),Config:nextConfig};
}
function secSaveBudgetConfig(data){
  // รองรับข้อมูลเดิม เพื่อไม่ให้ระบบเก่าหรือข้อมูลเก่าสูญหาย
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  var total=Number(data.TotalBudget||0), reserve=Number(data.ReserveBudget||0);
  if(!isFinite(total)||!isFinite(reserve))throw new Error('จำนวนเงินไม่ถูกต้อง');
  var existing=secGetBudgetConfig(year), now=new Date();
  var payload={FiscalYear:year,TotalBudget:total,ReserveBudget:reserve,UpdatedAt:now};
  if(existing.ConfigID && String(existing.ConfigID).indexOf('SOURCE-CALC-')!==0){
    payload.ConfigID=existing.ConfigID;
    if(!secUpdateRow_(SEC_SHEETS.BUDGET_CONFIG,'ConfigID',existing.ConfigID,payload))throw new Error('ไม่สามารถบันทึกงบรวมได้');
  }else{
    payload.ConfigID=secId_('BCFG'); payload.CreatedAt=now; secAppend_(SEC_SHEETS.BUDGET_CONFIG,payload);
  }
  try { secAudit_('กำหนดงบประมาณรวม/งบสำรอง','BUDGET_CONFIG',payload.ConfigID,'งบทั้งหมด '+total+' / สำรอง '+reserve); } catch(auditError) {}
  SpreadsheetApp.flush();
  return payload;
}
function secTransferBudgetReserve(data){
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  var amount=Number(data.Amount||0);
  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  if(!isFinite(amount)||amount===0)throw new Error('กรุณาระบุจำนวนเงินที่ต้องการโยก');
  var cfg=secGetBudgetConfig(year);
  var total=Number(cfg.TotalBudget||0), reserve=Number(cfg.ReserveBudget||0);
  if(amount>reserve)throw new Error('จำนวนเงินที่โยกมากกว่างบสำรองของสถาบัน');
  // ระบบใหม่เก็บการโยกเป็นรายการปรับปรุงแหล่งงบ เพื่อให้ประวัติแหล่งเงินยังตรวจสอบได้
  var now=new Date();
  secAppend_(SEC_SHEETS.BUDGET_SOURCES,{SourceID:secId_('BSRC'),FiscalYear:year,SourceName:'โอนจากงบสำรอง',Amount:-amount,BudgetPart:'งบสำรอง',Note:'โอนเข้ารวมเป็นงบประมาณหลัก',CreatedAt:now,UpdatedAt:now});
  secAppend_(SEC_SHEETS.BUDGET_SOURCES,{SourceID:secId_('BSRC'),FiscalYear:year,SourceName:'โอนเข้างบประมาณหลัก',Amount:amount,BudgetPart:'งบประมาณหลัก',Note:'รับโอนจากงบสำรอง',CreatedAt:now,UpdatedAt:now});
  SpreadsheetApp.flush();
  var next=secGetBudgetConfig(year);
  secAudit_('โยกงบสำรองเข้างบทั้งหมด','BUDGET_SOURCE','-', 'โยก '+amount+' | งบทั้งหมด '+next.TotalBudget+' | งบสำรอง '+next.ReserveBudget);
  return {FiscalYear:year,TotalBudget:next.TotalBudget,ReserveBudget:next.ReserveBudget,Transferred:amount};
}

function secGetCentralBudgets(fiscalYear){
  secInitSheets();
  var year=String(fiscalYear||'').trim();
  return secRows_(SEC_CENTRAL_BUDGET_SHEET).filter(function(x){return !year||String(x.FiscalYear||'')===year;});
}
function secSaveCentralBudget(data){
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  var item=String(data.Item||'').trim();
  var monthly=Number(data.MonthlyAmount||0);
  var months=Number(data.Months||12);
  var amount=Number(data.Amount||0);
  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  if(!item)throw new Error('กรุณาระบุรายการค่าใช้จ่ายส่วนกลาง');
  if(!isFinite(monthly)||monthly<0)throw new Error('งบต่อเดือนไม่ถูกต้อง');
  if(!isFinite(months)||months<1||months>12)throw new Error('จำนวนเดือนต้องอยู่ระหว่าง 1-12');
  if(!isFinite(amount)||amount<=0)throw new Error('งบส่วนกลางทั้งปีต้องมากกว่า 0');
  var now=new Date();
  var payload={FiscalYear:year,Item:item,MonthlyAmount:monthly,Months:months,Amount:amount,Note:String(data.Note||'').trim(),UpdatedAt:now};
  var id=String(data.CentralID||'').trim();
  if(id){
    payload.CentralID=id;
    if(!secUpdateRow_(SEC_CENTRAL_BUDGET_SHEET,'CentralID',id,payload))throw new Error('ไม่สามารถแก้ไขงบส่วนกลางได้');
    secAudit_('แก้ไขงบส่วนกลาง','CENTRAL_BUDGET',id,item+' | '+amount);
  }else{
    payload.CentralID=secId_('CBUD');payload.CreatedAt=now;secAppend_(SEC_CENTRAL_BUDGET_SHEET,payload);
    secAudit_('เพิ่มงบส่วนกลาง','CENTRAL_BUDGET',payload.CentralID,item+' | '+amount);
  }
  return {success:true,CentralID:payload.CentralID,FiscalYear:year};
}
function secDeleteCentralBudget(id){
  secInitSheets();
  var key=String(id||'').trim();
  if(!key)throw new Error('ไม่พบรายการงบส่วนกลาง');
  var rows=secRows_(SEC_CENTRAL_BUDGET_SHEET);
  var obj=rows.find(function(x){return String(x.CentralID||'')===key;});
  if(!obj)throw new Error('ไม่พบรายการงบส่วนกลางที่ต้องการลบ');
  if(!secDeleteRow_(SEC_CENTRAL_BUDGET_SHEET,'CentralID',key))throw new Error('ไม่สามารถลบรายการงบส่วนกลางได้');
  try{secAudit_('ลบงบส่วนกลาง','CENTRAL_BUDGET',key,String(obj.Item||'')+' | '+String(obj.Amount||''));}catch(e){}
  return {success:true,FiscalYear:String(obj.FiscalYear||'')};
}


function secGetEmergencyExpenses(fiscalYear){
  secInitSheets();
  var year=String(fiscalYear||'').trim();
  return secRows_(SEC_EMERGENCY_EXPENSE_SHEET).filter(function(x){
    return !year||String(x.FiscalYear||'')===year;
  });
}

function secEmergencyBaseAvailable_(year, part, excludeId){
  var cfg=secGetBudgetConfig(year);
  var srcTotal=Number(cfg.TotalBudget||0);
  var srcReserve=Number(cfg.ReserveBudget||0);

  // ใช้หลักเดียวกับหน้าเว็บ: ถ้ามีหน่วยงานย่อย ให้รวมงบย่อยแทนงบของหน่วยงานแม่
  var orgs=secGetOrganizationOptions();
  var units=secGetBudgetUnits(year);
  var roots=orgs.filter(function(x){
    var lv=String(x.Level||'').trim();
    return lv==='หน่วยงานหลัก'||lv==='ส่วนปฏิบัติการพิเศษ';
  });

  function children(rootId){
    return orgs.filter(function(x){
      return String(x.Level||'').trim()==='หน่วยงานย่อย' &&
             String(x.ParentID||'').trim()===String(rootId).trim();
    });
  }
  function unit(id){
    return units.find(function(u){
      return String(u.OrganizationID||'')===String(id||'');
    })||null;
  }

  var mainAllocated=0, reserveAllocated=0;
  roots.forEach(function(root){
    var kids=children(root.ID);
    if(kids.length){
      kids.forEach(function(k){
        var u=unit(k.ID);
        if(u){
          mainAllocated+=Number(u.BudgetAmount||0);
          reserveAllocated+=Number(u.ReserveAmount||0);
        }
      });
    }else{
      var u=unit(root.ID);
      if(u){
        mainAllocated+=Number(u.BudgetAmount||0);
        reserveAllocated+=Number(u.ReserveAmount||0);
      }
    }
  });

  // งบส่วนกลางถือเป็นงบประมาณหลัก
  var central=secRows_(SEC_CENTRAL_BUDGET_SHEET).filter(function(x){
    return String(x.FiscalYear||'')===String(year);
  }).reduce(function(s,x){return s+Number(x.Amount||0);},0);

  var expenses=secRows_(SEC_EMERGENCY_EXPENSE_SHEET).filter(function(x){
    return String(x.FiscalYear||'')===String(year) &&
           String(x.ExpenseID||'')!==String(excludeId||'');
  });

  var expenseMain=expenses.reduce(function(s,x){
    return String(x.BudgetPart||'')==='งบสำรอง'?s:s+Number(x.Amount||0);
  },0);
  var expenseReserve=expenses.reduce(function(s,x){
    return String(x.BudgetPart||'')==='งบสำรอง'?s+Number(x.Amount||0):s;
  },0);

  if(part==='งบสำรอง'){
    return srcReserve-reserveAllocated-expenseReserve;
  }
  return srcTotal-mainAllocated-central-expenseMain;
}

function secSaveEmergencyExpense(data){
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  var item=String(data.Item||'').trim();
  var amount=Number(data.Amount||0);
  var part=String(data.BudgetPart||'').trim();
  var detail=String(data.Detail||'').trim();

  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  if(!item)throw new Error('กรุณาระบุชื่อรายการ');
  if(!isFinite(amount)||amount<=0)throw new Error('จำนวนเงินต้องมากกว่า 0');
  if(part!=='งบประมาณหลัก'&&part!=='งบสำรอง')throw new Error('กรุณาเลือกส่วนของงบ');
  if(!detail)throw new Error('กรุณาระบุรายละเอียดการนำไปใช้');

  var id=String(data.ExpenseID||'').trim();
  var available=secEmergencyBaseAvailable_(year,part,id);
  if(amount>available){
    throw new Error(
      'งบ'+part+'คงเหลือไม่เพียงพอสำหรับรายการนี้ (คงเหลือ '+available.toLocaleString('th-TH')+' บาท)'
    );
  }

  var now=new Date();
  var payload={
    FiscalYear:year,
    Item:item,
    Amount:amount,
    BudgetPart:part,
    Detail:detail,
    UpdatedAt:now
  };

  if(id){
    payload.ExpenseID=id;
    if(!secUpdateRow_(SEC_EMERGENCY_EXPENSE_SHEET,'ExpenseID',id,payload)){
      throw new Error('ไม่สามารถแก้ไขค่าใช้จ่ายพิเศษ/ฉุกเฉินได้');
    }
    secAudit_('แก้ไขค่าใช้จ่ายพิเศษ/ฉุกเฉิน','EMERGENCY_EXPENSE',id,item+' | '+amount+' | '+part);
  }else{
    payload.ExpenseID=secId_('EMG');
    payload.CreatedAt=now;
    secAppend_(SEC_EMERGENCY_EXPENSE_SHEET,payload);
    secAudit_('เพิ่มค่าใช้จ่ายพิเศษ/ฉุกเฉิน','EMERGENCY_EXPENSE',payload.ExpenseID,item+' | '+amount+' | '+part);
  }

  SpreadsheetApp.flush();
  return {
    success:true,
    ExpenseID:payload.ExpenseID,
    FiscalYear:year,
    Item:item,
    Amount:amount,
    BudgetPart:part,
    Detail:detail
  };
}

function secDeleteEmergencyExpense(id){
  secInitSheets();
  var key=String(id||'').trim();
  if(!key)throw new Error('ไม่พบรายการค่าใช้จ่ายพิเศษ/ฉุกเฉิน');
  var rows=secRows_(SEC_EMERGENCY_EXPENSE_SHEET);
  var obj=rows.find(function(x){return String(x.ExpenseID||'')===key;});
  if(!obj)throw new Error('ไม่พบรายการค่าใช้จ่ายพิเศษ/ฉุกเฉินที่ต้องการลบ');
  if(!secDeleteRow_(SEC_EMERGENCY_EXPENSE_SHEET,'ExpenseID',key)){
    throw new Error('ไม่สามารถลบรายการค่าใช้จ่ายพิเศษ/ฉุกเฉินได้');
  }
  try{
    secAudit_('ลบค่าใช้จ่ายพิเศษ/ฉุกเฉิน','EMERGENCY_EXPENSE',key,
      String(obj.Item||'')+' | '+String(obj.Amount||'')+' | '+String(obj.BudgetPart||''));
  }catch(e){}
  SpreadsheetApp.flush();
  return {success:true,FiscalYear:String(obj.FiscalYear||'')};
}

function secGetBudgetUnits(fiscalYear){
  secInitSheets();
  var year=String(fiscalYear||'');
  return secRows_(SEC_SHEETS.BUDGET_UNITS).filter(function(x){return !year||String(x.FiscalYear||'')===year;});
}

function secSaveBudgetUnit(data){
  secInitSheets(); data=data||{};
  var year=String(data.FiscalYear||'').trim();
  var orgId=String(data.OrganizationID||'').trim();
  if(!year)throw new Error('กรุณาระบุปีงบประมาณ');
  if(!orgId)throw new Error('ไม่พบหน่วยงาน');
  var amount=Number(data.BudgetAmount||0), reserve=Number(data.ReserveAmount||0);
  if(amount<0||reserve<0)throw new Error('จำนวนเงินต้องไม่ติดลบ');
  var org=secGetOrganizationOptions().find(function(x){return String(x.ID)===orgId;});
  if(!org)throw new Error('ไม่พบหน่วยงานที่เลือก');
  var now=new Date();
  var existing=secGetBudgetUnits(year).find(function(x){return String(x.OrganizationID)===orgId;});
  var payload={FiscalYear:year,OrganizationID:orgId,ParentID:String(data.ParentID||org.ParentID||''),Organization:String(org.NameTH||data.Organization||''),ParentOrganization:String(data.ParentOrganization||''),BudgetAmount:amount,ReserveAmount:reserve,Note:String(data.Note||'').trim(),UpdatedAt:now};
  if(existing){
    payload.BudgetID=existing.BudgetID;
    if(!secUpdateRow_(SEC_SHEETS.BUDGET_UNITS,'BudgetID',existing.BudgetID,payload))throw new Error('ไม่สามารถแก้ไขงบหน่วยงานได้');
    secAudit_('แก้ไขงบหน่วยงาน','BUDGET_UNIT',existing.BudgetID,org.NameTH+' : '+amount+' / สำรอง '+reserve);
  }else{
    payload.BudgetID=secId_('BUD'); payload.CreatedAt=now; secAppend_(SEC_SHEETS.BUDGET_UNITS,payload);
    secAudit_('กำหนดงบหน่วยงาน','BUDGET_UNIT',payload.BudgetID,org.NameTH+' : '+amount+' / สำรอง '+reserve);
  }
  SpreadsheetApp.flush();
  return {success:true,BudgetID:payload.BudgetID,OrganizationID:orgId};
}
function secSaveBudget(data){secInitSheets();var amount=Number(data.Amount||0);if(!data.Item)throw new Error('กรุณาระบุรายการงบประมาณ');if(amount<0)throw new Error('จำนวนเงินไม่ถูกต้อง');var now=new Date();if(data.BudgetID){data.UpdatedAt=now;if(!secUpdateRow_(SEC_SHEETS.BUDGETS,'BudgetID',data.BudgetID,data))throw new Error('ไม่พบรายการงบ');secAudit_('แก้ไขงบประมาณ','BUDGET',data.BudgetID,data.Item);return data.BudgetID;}data.BudgetID=secId_('BUD');data.Status='รออนุมัติ';data.CreatedAt=now;data.UpdatedAt=now;secAppend_(SEC_SHEETS.BUDGETS,data);secAudit_('สร้างคำของบ','BUDGET',data.BudgetID,data.Item);return data.BudgetID;}
function secApproveBudget(id,status,actor,note){var allowed=['อนุมัติ','อนุมัติบางส่วน','ไม่อนุมัติ','รออนุมัติ'];if(allowed.indexOf(status)<0)throw new Error('สถานะไม่ถูกต้อง');if(!secUpdateRow_(SEC_SHEETS.BUDGETS,'BudgetID',id,{Status:status,UpdatedAt:new Date(),Note:note||''}))throw new Error('ไม่พบรายการงบ');secAudit_('เปลี่ยนสถานะงบประมาณ','BUDGET',id,status+' / '+(actor||'สภาผู้บริหาร'));return true;}
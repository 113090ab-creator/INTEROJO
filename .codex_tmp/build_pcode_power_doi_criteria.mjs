import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const dataPath = path.join(root, ".codex_tmp", "c_site_pcode_power_doi_criteria.json");
const outputDir = path.join(root, "outputs", "c_site_pcode_power_doi_criteria_20260805");
const outputPath = path.join(outputDir, "C관_P코드_파워별_DOI_기준.xlsx");
const data = JSON.parse(await fs.readFile(dataPath, "utf8"));

const workbook = Workbook.create();

const detailColumns = [
  "생산코드",
  "파워",
  "제품명",
  "사출코드",
  "분리코드",
  "거래처그룹",
  "거래처",
  "이니셜목록",
  "이니셜수",
  "수요라인수",
  "최초납기",
  "오더수량",
  "생산부족수량",
  "사출부족수량",
  "사출재고",
  "분리재고",
  "검사접착재고",
  "누수규격검사재고",
  "공정재고합계",
  "완제품재고",
  "DOI기준오더",
  "현재DOI",
  "DOI상태",
  "신제품여부",
  "기준등급",
  "기준DOI하한",
  "기준DOI상한",
  "상태",
  "판단메모",
  "부족대비공정재고율",
  "부족대비완제품재고율",
  "우선순위점수",
];

const numberColumns = new Set([
  "이니셜수",
  "수요라인수",
  "오더수량",
  "생산부족수량",
  "사출부족수량",
  "사출재고",
  "분리재고",
  "검사접착재고",
  "누수규격검사재고",
  "공정재고합계",
  "완제품재고",
  "DOI기준오더",
  "기준DOI하한",
  "기준DOI상한",
  "우선순위점수",
]);

const decimalColumns = new Set(["현재DOI", "부족대비공정재고율", "부족대비완제품재고율"]);
const textIdentifierColumns = new Set(["생산코드", "파워", "사출코드", "분리코드"]);

function rowsToMatrix(rows, columns) {
  return [
    columns,
    ...rows.map((row) =>
      columns.map((column) => {
        const value = row[column];
        if (value === undefined || value === null) return null;
        return value;
      }),
    ),
  ];
}

function applySheetDefaults(sheet) {
  sheet.showGridLines = false;
}

function writeTitle(sheet, title, subtitle, lastColumnIndex) {
  const titleRange = sheet.getRangeByIndexes(0, 0, 1, Math.max(1, lastColumnIndex + 1));
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: "#10275C",
    font: { bold: true, color: "#FFFFFF", size: 16 },
  };
  titleRange.format.rowHeight = 30;

  const subtitleRange = sheet.getRangeByIndexes(1, 0, 1, Math.max(1, lastColumnIndex + 1));
  subtitleRange.merge();
  subtitleRange.values = [[subtitle]];
  subtitleRange.format = {
    fill: "#EAF0FF",
    font: { color: "#243B6B", size: 10 },
  };
  subtitleRange.format.rowHeight = 24;
}

function applyTableStyle(sheet, rowCount, columns) {
  const colCount = columns.length;
  const headerRange = sheet.getRangeByIndexes(3, 0, 1, colCount);
  headerRange.format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
    borders: { preset: "all", style: "thin", color: "#C9D4E5" },
  };
  headerRange.format.rowHeight = 24;

  const bodyRange = sheet.getRangeByIndexes(4, 0, Math.max(1, rowCount), colCount);
  bodyRange.format = {
    borders: {
      insideHorizontal: { style: "thin", color: "#E5E7EB" },
      insideVertical: { style: "thin", color: "#E5E7EB" },
      bottom: { style: "thin", color: "#E5E7EB" },
    },
  };

  for (let idx = 0; idx < columns.length; idx += 1) {
    const column = columns[idx];
    const columnRange = sheet.getRangeByIndexes(4, idx, Math.max(1, rowCount), 1);
    if (numberColumns.has(column)) {
      columnRange.format.numberFormat = "#,##0";
    } else if (decimalColumns.has(column)) {
      columnRange.format.numberFormat = "#,##0.0";
    } else if (textIdentifierColumns.has(column)) {
      columnRange.format.numberFormat = "@";
    } else {
      columnRange.format.wrapText = false;
    }
  }

  const used = sheet.getRangeByIndexes(0, 0, rowCount + 4, colCount);
  used.format.font = { name: "맑은 고딕", size: 9 };
  used.format.autofitColumns();

  const widths = {
    생산코드: 18,
    파워: 9,
    제품명: 34,
    사출코드: 18,
    분리코드: 18,
    거래처그룹: 18,
    거래처: 18,
    이니셜목록: 26,
    판단메모: 34,
    상태: 13,
    DOI상태: 13,
    신제품여부: 10,
    기준등급: 10,
  };
  for (let idx = 0; idx < columns.length; idx += 1) {
    if (widths[columns[idx]]) {
      sheet.getRangeByIndexes(0, idx, rowCount + 4, 1).format.columnWidth = widths[columns[idx]];
    }
  }
  sheet.freezePanes.freezeRows(4);
}

function addStatusFormatting(sheet, rowCount, columns) {
  const statusIndex = columns.indexOf("상태");
  if (statusIndex < 0 || rowCount < 1) return;
  const statusRange = sheet.getRangeByIndexes(4, statusIndex, rowCount, 1);
  statusRange.conditionalFormats.add("containsText", {
    text: "생산최우선",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  statusRange.conditionalFormats.add("containsText", {
    text: "생산우선",
    format: { fill: "#FFEDD5", font: { bold: true, color: "#C2410C" } },
  });
  statusRange.conditionalFormats.add("containsText", {
    text: "생산지양",
    format: { fill: "#E0E7FF", font: { bold: true, color: "#3730A3" } },
  });
  statusRange.conditionalFormats.add("containsText", {
    text: "생산조정",
    format: { fill: "#FEF3C7", font: { bold: true, color: "#92400E" } },
  });

  const doiStatusIndex = columns.indexOf("DOI상태");
  if (doiStatusIndex >= 0 && rowCount >= 1) {
    const doiStatusRange = sheet.getRangeByIndexes(4, doiStatusIndex, rowCount, 1);
    doiStatusRange.conditionalFormats.add("containsText", {
      text: "신제품",
      format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
    });
    doiStatusRange.conditionalFormats.add("containsText", {
      text: "재고없음",
      format: { fill: "#FEE2E2", font: { bold: true, color: "#991B1B" } },
    });
    doiStatusRange.conditionalFormats.add("containsText", {
      text: "확인필요",
      format: { fill: "#F3F4F6", font: { color: "#374151" } },
    });
  }
}

function writeDataSheet(sheetName, title, subtitle, rows) {
  const sheet = workbook.worksheets.add(sheetName);
  applySheetDefaults(sheet);
  writeTitle(sheet, title, subtitle, detailColumns.length - 1);
  const matrix = rowsToMatrix(rows, detailColumns);
  sheet.getRangeByIndexes(3, 0, matrix.length, detailColumns.length).values = matrix;
  const powerIndex = detailColumns.indexOf("파워");
  if (powerIndex >= 0 && rows.length > 0) {
    sheet.getRangeByIndexes(4, powerIndex, rows.length, 1).formulas = rows.map((row) => {
      const power = String(row["파워"] ?? "").replaceAll('"', '""');
      return [`="${power}"`];
    });
  }
  applyTableStyle(sheet, rows.length, detailColumns);
  addStatusFormatting(sheet, rows.length, detailColumns);
  return sheet;
}

function writeSummarySheet() {
  const sheet = workbook.worksheets.add("요약");
  applySheetDefaults(sheet);
  writeTitle(sheet, "C관 P코드+파워별 DOI 기준 요약", "APS C관 스냅샷 기준으로 P코드 단위 생산 우선순위와 DOI 기준을 요약합니다.", 8);

  const summaryRows = [
    ["전체 P코드+파워 라인", data.summary.rows],
    ["오더수량 합계", data.summary.order_total],
    ["생산부족수량 합계", data.summary.shortage_total],
    ["사출부족수량 합계", data.summary.inj_shortage_total],
    ["완제품재고 합계", data.summary.fg_total],
    ["공정재고 합계", data.summary.wip_total],
  ];
  sheet.getRange("A4:B10").values = [["항목", "값"], ...summaryRows];
  sheet.getRange("A4:B4").format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
  };
  sheet.getRange("B5:B9").format.numberFormat = "#,##0";

  const statusRows = Object.entries(data.summary.status_counts);
  sheet.getRangeByIndexes(3, 3, statusRows.length + 1, 2).values = [["상태", "건수"], ...statusRows];
  sheet.getRange("D4:E4").format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
  };
  sheet.getRangeByIndexes(4, 4, statusRows.length, 1).format.numberFormat = "#,##0";

  const gradeRows = Object.entries(data.summary.grade_counts);
  sheet.getRangeByIndexes(3, 6, gradeRows.length + 1, 2).values = [["기준등급", "건수"], ...gradeRows];
  sheet.getRange("G4:H4").format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
  };
  sheet.getRangeByIndexes(4, 7, gradeRows.length, 1).format.numberFormat = "#,##0";

  const doiStatusRows = Object.entries(data.summary.doi_status_counts ?? {});
  sheet.getRangeByIndexes(12, 0, doiStatusRows.length + 1, 2).values = [["DOI상태", "건수"], ...doiStatusRows];
  sheet.getRange("A13:B13").format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
  };
  if (doiStatusRows.length > 0) {
    sheet.getRangeByIndexes(13, 1, doiStatusRows.length, 1).format.numberFormat = "#,##0";
  }

  const used = sheet.getRange("A1:H14");
  used.format.font = { name: "맑은 고딕", size: 10 };
  used.format.autofitColumns();
  return sheet;
}

function writeCriteriaSheet() {
  const sheet = workbook.worksheets.add("기준표");
  applySheetDefaults(sheet);
  writeTitle(sheet, "P코드+파워 DOI 기준표", "기준은 C관 P코드 분포에 맞춰 자동 분류합니다. 필요하면 이 값만 조정해서 운영 기준을 바꿀 수 있습니다.", 6);
  const rows = [
    ["A급", "오더수량 >= 1,500 또는 생산/사출부족 >= 17,000", "30~60일", "상위 10% 수준의 수요 또는 부족"],
    ["B급", "오더수량 >= 500 또는 생산/사출부족 >= 5,000", "20~60일", "반복 수요 또는 중간 부족"],
    ["C급", "그 외", "0~30일", "저회전/소량 품목"],
    ["신제품", "BAGUMORE, Burn Sugar, Viva Boom, 중국_축고정/축고정", "과거 DOI 제외", "DOI상태만 신제품으로 표시하고 상태/판단메모는 빈칸"],
  ];
  sheet.getRange("A4:D8").values = [["기준등급", "분류조건", "추천 DOI", "메모"], ...rows];
  sheet.getRange("A4:D4").format = {
    fill: "#DDE6F6",
    font: { bold: true, color: "#172554" },
  };
  sheet.getRange("A5:D8").format = {
    borders: {
      insideHorizontal: { style: "thin", color: "#E5E7EB" },
      insideVertical: { style: "thin", color: "#E5E7EB" },
      bottom: { style: "thin", color: "#E5E7EB" },
    },
  };
  sheet.getRange("A1:D8").format.font = { name: "맑은 고딕", size: 10 };
  sheet.getRange("A1:D8").format.autofitColumns();
  sheet.getRange("B:B").format.columnWidth = 46;
  sheet.getRange("D:D").format.columnWidth = 32;
  return sheet;
}

writeSummarySheet();
writeCriteriaSheet();
writeDataSheet("C관_P코드기준", "C관 P코드+파워별 DOI 기준", "한 줄의 판단 단위는 생산코드(P코드)+파워입니다.", data.rows);
writeDataSheet("신제품", "신제품 P코드", "신제품은 DOI상태만 표시하고 상태/판단메모는 비워둡니다.", data.new_products ?? []);
writeDataSheet("생산우선", "생산우선 P코드", "생산최우선/생산우선 상태만 모은 시트입니다.", data.priority);
writeDataSheet("생산조정", "생산조정 P코드", "생산지양/생산조정 상태만 모은 시트입니다.", data.adjust);

await fs.mkdir(outputDir, { recursive: true });

for (const sheetName of ["요약", "기준표", "C관_P코드기준", "신제품", "생산우선", "생산조정"]) {
  const preview = await workbook.render({
    sheetName,
    range: "A1:J20",
    scale: 1,
    format: "png",
  });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), bytes);
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath }, null, 2));
process.exitCode = 0;

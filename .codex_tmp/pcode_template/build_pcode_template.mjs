import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("C:/Users/유현아/Documents/GitHub/INTEROJO/outputs/pcode_template");
const outputPath = path.join(outputDir, "전체_P코드_거래처별_양식.xlsx");
const previewPath = path.join(outputDir, "전체_P코드_거래처별_양식_미리보기.png");

const workbook = Workbook.create();

const colors = {
  navy: "#172554",
  blue: "#2563EB",
  lightBlue: "#DBEAFE",
  gray: "#F8FAFC",
  gray2: "#E2E8F0",
  text: "#0F172A",
  muted: "#64748B",
  border: "#CBD5E1",
  red: "#DC2626",
  lightRed: "#FEE2E2",
  amber: "#B45309",
  lightAmber: "#FEF3C7",
  green: "#15803D",
  lightGreen: "#DCFCE7",
};

const customers = [
  { sheet: "PIA", label: "PIA" },
  { sheet: "Sincere", label: "Sincere" },
  { sheet: "OPTICAL SUPPLIES", label: "OPTICAL SUPPLIES" },
  { sheet: "MAXVUE_OPTIMAX", label: "MAXVUE/OPTIMAX" },
  { sheet: "ALENSA", label: "ALENSA" },
  { sheet: "FEEL GOOD", label: "FEEL GOOD" },
  { sheet: "CHINA_IRIS", label: "CHINA/IRIS" },
  { sheet: "T-Garden", label: "T-Garden" },
  { sheet: "HAPA_PPB", label: "HAPA/PPB" },
  { sheet: "국내", label: "국내" },
  { sheet: "기타", label: "기타 거래처" },
  { sheet: "거래처 미지정", label: "거래처 미지정" },
];

const headers = [
  "생산코드",
  "분리코드",
  "사출코드",
  "제품명",
  "파워",
  "오더수량1",
  "오더수량2",
  "오더수량3",
  "오더수량4",
  "오더수량5",
  "오더합계",
  "제품부족수량",
  "사출부족수량",
  "사출재고",
  "분리재고",
  "검사접착재고",
  "누수규격검사재고",
  "공정재고 합계",
  "DOI",
  "상태",
];

const firstDataRow = 6;
const reservedRows = 300;
const lastDataRow = firstDataRow + reservedRows - 1;
const tableStartRow = 5;
const tableEndRow = lastDataRow;

function colLetter(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function address(row, col, rowCount, colCount) {
  return `${colLetter(col)}${row}:${colLetter(col + colCount - 1)}${row + rowCount - 1}`;
}

function setWidths(sheet, widths) {
  widths.forEach((width, i) => {
    sheet.getRange(`${colLetter(i)}:${colLetter(i)}`).format.columnWidth = width;
  });
}

function setBaseFont(sheet) {
  // Intentionally left empty. Specific title/header/data ranges receive font styles
  // directly so later global formatting does not overwrite title/header emphasis.
}

function titleBlock(sheet, title, subtitle, lastColIndex) {
  const lastCol = colLetter(lastColIndex);
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: colors.navy,
    font: { name: "맑은 고딕", bold: true, color: "#FFFFFF", size: 15 },
    horizontalAlignment: "left",
  };

  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    fill: colors.gray,
    font: { name: "맑은 고딕", color: colors.muted, size: 10 },
    wrapText: true,
  };

  sheet.getRange(`A1:${lastCol}2`).format.borders = {
    preset: "outside",
    style: "thin",
    color: colors.border,
  };
}

function addTable(sheet, tableName, rows) {
  const tableRange = address(tableStartRow, 0, tableEndRow - tableStartRow + 1, headers.length);
  const blanks = Array.from({ length: reservedRows }, () => Array(headers.length).fill(null));
  rows.forEach((row, i) => {
    blanks[i] = row;
  });
  sheet.getRange(tableRange).values = [headers, ...blanks];

  const headerRange = sheet.getRange(address(tableStartRow, 0, 1, headers.length));
  headerRange.format = {
    fill: colors.lightBlue,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  headerRange.format.borders = { preset: "all", style: "thin", color: colors.border };

  const dataRange = sheet.getRange(address(firstDataRow, 0, reservedRows, headers.length));
  dataRange.format.font = { name: "맑은 고딕", size: 10, color: colors.text };
  dataRange.format.borders = {
    insideHorizontal: { style: "thin", color: colors.gray2 },
    insideVertical: { style: "thin", color: colors.gray2 },
    bottom: { style: "thin", color: colors.border },
    left: { style: "thin", color: colors.border },
    right: { style: "thin", color: colors.border },
  };

  try {
    const table = sheet.tables.add(tableRange, true, tableName);
    table.showFilterButton = true;
    table.showBandedColumns = false;
  } catch (error) {
    // The plain formatted range remains usable if table creation is unavailable.
  }
}

function applyFormulas(sheet) {
  sheet.getRange(`K${firstDataRow}`).formulas = [["=SUM(F6:J6)"]];
  sheet.getRange(`R${firstDataRow}`).formulas = [["=SUM(N6:Q6)"]];
  sheet.getRange(`T${firstDataRow}`).formulas = [[
    '=IF(COUNTA(A6:J6,L6:Q6,S6)=0,"",IF(M6>0,"사출부족",IF(L6>0,"제품부족",IF(AND(K6=0,R6>0),"수요없음재고",IF(AND(ISNUMBER(S6),S6<7),"DOI주의","정상")))))',
  ]];
  sheet.getRange(`K${firstDataRow}:K${lastDataRow}`).fillDown();
  sheet.getRange(`R${firstDataRow}:R${lastDataRow}`).fillDown();
  sheet.getRange(`T${firstDataRow}:T${lastDataRow}`).fillDown();
}

function formatColumns(sheet) {
  setWidths(sheet, [
    18, 18, 18, 34, 10, 12, 12, 12, 12, 12,
    13, 15, 15, 13, 13, 15, 17, 15, 10, 14,
  ]);

  const numericCols = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
  numericCols.forEach((col) => {
    const letter = colLetter(col);
    sheet.getRange(`${letter}${firstDataRow}:${letter}${lastDataRow}`).format.numberFormat = "#,##0";
    sheet.getRange(`${letter}${firstDataRow}:${letter}${lastDataRow}`).format.horizontalAlignment = "right";
  });

  sheet.getRange(`S${firstDataRow}:S${lastDataRow}`).format.numberFormat = "0.0";
  sheet.getRange(`S${firstDataRow}:S${lastDataRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A${firstDataRow}:E${lastDataRow}`).format.numberFormat = "@";
  sheet.getRange(`A${firstDataRow}:E${lastDataRow}`).format.horizontalAlignment = "left";
  sheet.getRange(`T${firstDataRow}:T${lastDataRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`D${firstDataRow}:D${lastDataRow}`).format.wrapText = true;
}

function applyConditionalFormatting(sheet) {
  const shortageRange = sheet.getRange(`L${firstDataRow}:M${lastDataRow}`);
  const statusRange = sheet.getRange(`T${firstDataRow}:T${lastDataRow}`);
  try {
    shortageRange.conditionalFormats.add("cellIs", {
      operator: "greaterThan",
      formula: 0,
      format: { fill: colors.lightRed, font: { color: colors.red, bold: true } },
    });
    statusRange.conditionalFormats.add("containsText", {
      text: "사출부족",
      format: { fill: colors.lightRed, font: { color: colors.red, bold: true } },
    });
    statusRange.conditionalFormats.add("containsText", {
      text: "제품부족",
      format: { fill: colors.lightAmber, font: { color: colors.amber, bold: true } },
    });
    statusRange.conditionalFormats.add("containsText", {
      text: "수요없음재고",
      format: { fill: colors.lightBlue, font: { color: colors.blue, bold: true } },
    });
    statusRange.conditionalFormats.add("containsText", {
      text: "정상",
      format: { fill: colors.lightGreen, font: { color: colors.green, bold: true } },
    });
  } catch (error) {
    // Conditional formatting is optional for file creation.
  }
}

function customerRows(customerLabel) {
  if (customerLabel !== "PIA") return [];
  return [
    [
      "P1061A-03.50CRW3",
      "Q1061-03.50CRW3",
      "R1061-03.50CRW3",
      "PIA D_Cherish Brown_55% 1-DAY",
      "-03.50",
      100,
      200,
      null,
      null,
      null,
      null,
      1500,
      800,
      0,
      250,
      100,
      0,
      null,
      4.2,
      null,
    ],
    [
      "P1061A-04.00CRW3",
      "Q1061-04.00CRW3",
      "R1061-04.00CRW3",
      "PIA D_Cherish Brown_55% 1-DAY",
      "-04.00",
      null,
      null,
      null,
      null,
      null,
      null,
      0,
      0,
      0,
      800,
      0,
      0,
      null,
      18.6,
      null,
    ],
  ];
}

function createCustomerSheet(customer, index) {
  const sheet = workbook.worksheets.add(customer.sheet);
  sheet.showGridLines = false;
  titleBlock(
    sheet,
    `${customer.label} 전체 P코드 원장`,
    "생산코드(P) 기준으로 수요가 없는 품목까지 공정재고와 DOI를 함께 확인하는 거래처별 시트입니다.",
    headers.length - 1,
  );
  sheet.getRange("A3:T3").merge();
  sheet.getRange("A3").values = [[
    "오더수량1~5는 이니셜별 오더 수량 입력/자동반영 영역입니다. 이니셜이 5개를 넘으면 오더수량 열을 추가하고 오더합계 수식 범위만 확장하면 됩니다.",
  ]];
  sheet.getRange("A3").format = {
    fill: colors.lightAmber,
    font: { name: "맑은 고딕", color: colors.amber, bold: true, size: 10 },
    wrapText: true,
  };
  addTable(sheet, `T_${index + 1}_${customer.sheet.replace(/[^A-Za-z0-9가-힣]/g, "_")}`, customerRows(customer.label));
  applyFormulas(sheet);
  formatColumns(sheet);
  applyConditionalFormatting(sheet);
  sheet.freezePanes.freezeRows(5);
  sheet.freezePanes.freezeColumns(4);
  setBaseFont(sheet);
  return sheet;
}

function createGuideSheet() {
  const sheet = workbook.worksheets.add("안내");
  sheet.showGridLines = false;
  titleBlock(sheet, "전체 P코드 거래처별 양식", "대시보드에서 전체를 직접 펼치기 어려울 때 별도 엑셀 원장으로 관리하는 기본 양식입니다.", 7);
  const rows = [
    ["기준", "설명"],
    ["시트 구성", "거래처별 1개 시트로 구성했습니다. 각 시트는 동일한 열 구조를 사용합니다."],
    ["생산코드", "A열 생산코드는 T코드가 아니라 P로 시작하는 생산코드 기준입니다."],
    ["오더수량", "오더수량1~5는 이니셜별 오더 수량 영역입니다. 오더합계는 자동 합산됩니다."],
    ["공정재고", "사출/분리/검사접착/누수규격검사 재고를 합산해 공정재고 합계를 계산합니다."],
    ["상태", "사출부족, 제품부족, 수요없음재고, DOI주의, 정상 순서로 자동 표시합니다."],
    ["운영 방식", "대시보드는 요약/검색 중심, 이 파일은 전체 P코드 확인과 대체 검토용으로 사용합니다."],
  ];
  sheet.getRange("A4:B10").values = rows;
  sheet.getRange("A4:B4").format = {
    fill: colors.lightBlue,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
  };
  sheet.getRange("A4:B10").format.borders = { preset: "all", style: "thin", color: colors.border };
  sheet.getRange("A5:A10").format = {
    fill: colors.gray,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
  };
  setWidths(sheet, [20, 95, 12, 12, 12, 12, 12, 12]);
  sheet.freezePanes.freezeRows(4);
  setBaseFont(sheet);
}

function createIndexSheet() {
  const sheet = workbook.worksheets.add("시트목록");
  sheet.showGridLines = false;
  titleBlock(sheet, "거래처 시트 목록", "거래처별 원장 탭과 사용 기준입니다.", 5);
  const rows = [["거래처", "시트명", "제품분류 기준", "비고"]];
  customers.forEach((customer) => {
    rows.push([customer.label, customer.sheet, "1-DAY / FRP / 기타", "동일 양식"]);
  });
  sheet.getRange(address(4, 0, rows.length, 4)).values = rows;
  sheet.getRange("A4:D4").format = {
    fill: colors.lightBlue,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
    horizontalAlignment: "center",
  };
  sheet.getRange(address(4, 0, rows.length, 4)).format.borders = { preset: "all", style: "thin", color: colors.border };
  setWidths(sheet, [24, 24, 24, 30, 12, 12]);
  sheet.freezePanes.freezeRows(4);
  setBaseFont(sheet);
}

createGuideSheet();
createIndexSheet();
customers.forEach(createCustomerSheet);

const preview = await workbook.render({
  sheetName: "PIA",
  range: "A1:T12",
  scale: 1,
  format: "png",
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});

const overview = await workbook.inspect({
  kind: "sheet",
  include: "name",
  maxChars: 2500,
  tableMaxRows: 4,
  tableMaxCols: 8,
});

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(JSON.stringify({ outputPath, previewPath }));
console.log(errors.ndjson);
console.log(overview.ndjson);

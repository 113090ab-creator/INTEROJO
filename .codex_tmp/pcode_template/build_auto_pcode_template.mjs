import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workDir = path.resolve("C:/Users/유현아/Documents/GitHub/INTEROJO/.codex_tmp/pcode_template");
const outputDir = path.resolve("C:/Users/유현아/Documents/GitHub/INTEROJO/outputs/pcode_template");
const inputJson = path.join(workDir, "pcode_auto_rows.json");
const outputPath = path.join(outputDir, "전체_P코드_자동조회_거래처별_양식.xlsx");
const previewInputPath = path.join(outputDir, "전체_P코드_자동조회_입력시트_미리보기.png");
const previewCustomerPath = path.join(outputDir, "전체_P코드_자동조회_거래처시트_미리보기.png");

const payload = JSON.parse(await fs.readFile(inputJson, "utf8"));
const rows = payload.rows;
const customers = payload.customers;

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

const inputHeaders = [
  "거래처그룹",
  "제품분류",
  "생산코드",
  "분리코드",
  "사출코드",
  "제품명",
  "파워",
  "오더수량1",
  "오더수량2",
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
  "납기일",
  "이니셜",
  "신호",
  "기존상태",
  "비고",
];

const customerHeaders = inputHeaders.slice(2, 19);
const firstDataRow = 6;
const inputHeaderRow = 5;
const lastInputRow = firstDataRow + rows.length - 1;
const inputLastCol = inputHeaders.length - 1;
const customerLastCol = customerHeaders.length - 1;

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

function rangeAddress(row, col, rowCount, colCount) {
  return `${colLetter(col)}${row}:${colLetter(col + colCount - 1)}${row + rowCount - 1}`;
}

function safeSheetName(raw, used) {
  let name = String(raw || "거래처 미지정")
    .replace(/[\\/?*[\]:]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  if (!name) name = "거래처 미지정";
  name = name.slice(0, 31);
  const base = name;
  let suffix = 2;
  while (used.has(name)) {
    const tail = `_${suffix}`;
    name = `${base.slice(0, 31 - tail.length)}${tail}`;
    suffix += 1;
  }
  used.add(name);
  return name;
}

function setWidths(sheet, widths) {
  widths.forEach((width, i) => {
    sheet.getRange(`${colLetter(i)}:${colLetter(i)}`).format.columnWidth = width;
  });
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

function styleHeader(sheet, row, colCount) {
  const range = sheet.getRange(rangeAddress(row, 0, 1, colCount));
  range.format = {
    fill: colors.lightBlue,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  range.format.borders = { preset: "all", style: "thin", color: colors.border };
}

function chunked(values, chunkSize) {
  const chunks = [];
  for (let i = 0; i < values.length; i += chunkSize) {
    chunks.push([i, values.slice(i, i + chunkSize)]);
  }
  return chunks;
}

function inputMatrix(records) {
  return records.map((row) => [
    row["거래처그룹"] || "거래처 미지정",
    row["제품분류"] || "기타",
    row["생산코드"] || "",
    row["분리코드"] || "",
    row["사출코드"] || "",
    row["제품명"] || "",
    row["파워"] || "",
    Number(row["오더수량1"] || 0),
    Number(row["오더수량2"] || 0),
    null,
    Number(row["제품부족수량"] || 0),
    Number(row["사출부족수량"] || 0),
    Number(row["사출재고"] || 0),
    Number(row["분리재고"] || 0),
    Number(row["검사접착재고"] || 0),
    Number(row["누수규격검사재고"] || 0),
    null,
    Number(row["DOI"] || 0),
    null,
    row["납기일"] || "",
    row["이니셜"] || "",
    row["신호"] || "",
    row["기존상태"] || "",
    row["비고"] || "",
  ]);
}

function applyNumberFormats(sheet) {
  const numericCols = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
  for (const idx of numericCols) {
    const col = colLetter(idx);
    sheet.getRange(`${col}${firstDataRow}:${col}${lastInputRow}`).format.numberFormat = "#,##0";
    sheet.getRange(`${col}${firstDataRow}:${col}${lastInputRow}`).format.horizontalAlignment = "right";
  }
  sheet.getRange(`R${firstDataRow}:R${lastInputRow}`).format.numberFormat = "0.0";
  sheet.getRange(`A${firstDataRow}:G${lastInputRow}`).format.numberFormat = "@";
  sheet.getRange(`T${firstDataRow}:X${lastInputRow}`).format.numberFormat = "@";
}

function applyConditionalFormatting(sheet, statusCol, rowStart, rowEnd) {
  const statusLetter = colLetter(statusCol);
  const statusRange = sheet.getRange(`${statusLetter}${rowStart}:${statusLetter}${rowEnd}`);
  const shortageRange = sheet.getRange(`K${rowStart}:L${rowEnd}`);
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
    // Conditional formatting is not required for the workbook to function.
  }
}

function createGuide(customerSheetMap) {
  const sheet = workbook.worksheets.add("안내");
  sheet.showGridLines = false;
  titleBlock(
    sheet,
    "전체 P코드 자동조회 양식",
    "전체코드_입력 시트에 값을 입력하면 거래처별 시트가 자동으로 해당 거래처 행만 조회합니다.",
    7,
  );
  const rowsGuide = [
    ["구분", "사용 방법"],
    ["전체코드_입력", "전체 P코드 기준 원장입니다. 거래처그룹, 오더수량, 부족수량, 재고, DOI 값을 여기에서 관리합니다."],
    ["거래처별 시트", "각 거래처 시트는 수식으로 자동 조회됩니다. 직접 수정하지 말고 전체코드_입력 값을 수정합니다."],
    ["오더합계", "오더수량1 + 오더수량2 수식입니다. 이니셜이 더 필요하면 오더수량 열을 추가한 뒤 합계 범위를 확장합니다."],
    ["공정재고 합계", "사출재고 + 분리재고 + 검사접착재고 + 누수규격검사재고 수식입니다."],
    ["상태", "사출부족, 제품부족, 수요없음재고, DOI주의, 수요있음, 수요없음 순서로 자동 표시합니다."],
    ["기준 데이터", `cloud_snapshots/all_item_status_snapshot.csv.gz 기준 ${rows.length.toLocaleString("ko-KR")}개 P코드, 거래처그룹 ${customers.length.toLocaleString("ko-KR")}개`],
  ];
  sheet.getRange(rangeAddress(4, 0, rowsGuide.length, 2)).values = rowsGuide;
  styleHeader(sheet, 4, 2);
  sheet.getRange(rangeAddress(5, 0, rowsGuide.length - 1, 2)).format.borders = {
    preset: "all",
    style: "thin",
    color: colors.border,
  };
  setWidths(sheet, [22, 110, 12, 12, 12, 12, 12, 12]);

  const listStart = 13;
  const listRows = [["거래처그룹", "시트명", "행수"]];
  for (const item of customerSheetMap) {
    listRows.push([item.customer, item.sheetName, item.count]);
  }
  sheet.getRange(rangeAddress(listStart, 0, listRows.length, 3)).values = listRows;
  styleHeader(sheet, listStart, 3);
  sheet.getRange(rangeAddress(listStart + 1, 0, listRows.length - 1, 3)).format.borders = {
    preset: "all",
    style: "thin",
    color: colors.border,
  };
  sheet.getRange(`C${listStart + 1}:C${listStart + listRows.length - 1}`).format.numberFormat = "#,##0";
  sheet.freezePanes.freezeRows(4);
}

function createInputSheet() {
  const sheet = workbook.worksheets.add("전체코드_입력");
  sheet.showGridLines = false;
  titleBlock(
    sheet,
    "전체코드 입력",
    "이 시트가 원본입니다. 여기 값을 수정하면 거래처별 시트가 자동으로 갱신됩니다.",
    inputLastCol,
  );
  sheet.getRange("A3:X3").merge();
  sheet.getRange("A3").values = [[
    "주의: 거래처별 시트는 수식 조회 결과입니다. 품목 정보, 수요, 재고, DOI는 이 시트에서만 수정하세요.",
  ]];
  sheet.getRange("A3").format = {
    fill: colors.lightAmber,
    font: { name: "맑은 고딕", color: colors.amber, bold: true, size: 10 },
    wrapText: true,
  };
  sheet.getRange(rangeAddress(inputHeaderRow, 0, 1, inputHeaders.length)).values = [inputHeaders];
  styleHeader(sheet, inputHeaderRow, inputHeaders.length);

  const matrix = inputMatrix(rows);
  for (const [offset, part] of chunked(matrix, 5000)) {
    sheet.getRange(rangeAddress(firstDataRow + offset, 0, part.length, inputHeaders.length)).values = part;
  }

  sheet.getRange(`J${firstDataRow}`).formulas = [[`=SUM(H${firstDataRow}:I${firstDataRow})`]];
  sheet.getRange(`Q${firstDataRow}`).formulas = [[`=SUM(M${firstDataRow}:P${firstDataRow})`]];
  sheet.getRange(`S${firstDataRow}`).formulas = [[
    `=IF(A${firstDataRow}="","",IF(L${firstDataRow}>0,"사출부족",IF(K${firstDataRow}>0,"제품부족",IF(AND(J${firstDataRow}=0,Q${firstDataRow}>0),"수요없음재고",IF(AND(ISNUMBER(R${firstDataRow}),R${firstDataRow}>0,R${firstDataRow}<7),"DOI주의",IF(J${firstDataRow}>0,"수요있음","수요없음"))))))`,
  ]];
  sheet.getRange(`J${firstDataRow}:J${lastInputRow}`).fillDown();
  sheet.getRange(`Q${firstDataRow}:Q${lastInputRow}`).fillDown();
  sheet.getRange(`S${firstDataRow}:S${lastInputRow}`).fillDown();

  setWidths(sheet, [
    20, 13, 18, 18, 18, 34, 10, 12, 12, 13, 15, 15,
    13, 13, 15, 18, 15, 10, 14, 13, 12, 10, 20, 26,
  ]);
  applyNumberFormats(sheet);
  applyConditionalFormatting(sheet, 18, firstDataRow, lastInputRow);
  sheet.freezePanes.freezeRows(inputHeaderRow);
  sheet.freezePanes.freezeColumns(6);
}

function createCustomerSheet(customer, sheetName) {
  const sheet = workbook.worksheets.add(sheetName);
  sheet.showGridLines = false;
  titleBlock(
    sheet,
    `${customer} 자동조회`,
    "전체코드_입력 시트에서 거래처그룹이 이 거래처와 같은 행만 자동 표시합니다.",
    customerLastCol,
  );
  sheet.getRange("A3:B3").values = [["조회 거래처", customer]];
  sheet.getRange("A3").format = {
    fill: colors.lightBlue,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
    horizontalAlignment: "center",
  };
  sheet.getRange("B3").format = {
    fill: colors.gray,
    font: { name: "맑은 고딕", bold: true, color: colors.text },
  };
  sheet.getRange("A5:Q5").values = [customerHeaders];
  styleHeader(sheet, 5, customerHeaders.length);
  sheet.getRange("A6").formulas = [[
    `=FILTER('전체코드_입력'!$C$${firstDataRow}:$S$${lastInputRow},'전체코드_입력'!$A$${firstDataRow}:$A$${lastInputRow}=$B$3,"")`,
  ]];
  setWidths(sheet, [18, 18, 18, 34, 10, 12, 12, 13, 15, 15, 13, 13, 15, 18, 15, 10, 14]);
  sheet.getRange("H6:Q2000").format.numberFormat = "#,##0";
  sheet.getRange("P6:P2000").format.numberFormat = "0.0";
  applyConditionalFormatting(sheet, 16, 6, 2000);
  sheet.freezePanes.freezeRows(5);
  sheet.freezePanes.freezeColumns(4);
}

const customerCounts = new Map();
for (const row of rows) {
  const key = row["거래처그룹"] || "거래처 미지정";
  customerCounts.set(key, (customerCounts.get(key) || 0) + 1);
}

const usedSheetNames = new Set(["안내", "전체코드_입력"]);
const customerSheetMap = customers.map((customer) => ({
  customer,
  sheetName: safeSheetName(customer, usedSheetNames),
  count: customerCounts.get(customer) || 0,
}));

createGuide(customerSheetMap);
createInputSheet();
for (const item of customerSheetMap) {
  createCustomerSheet(item.customer, item.sheetName);
}

const inputCheck = await workbook.inspect({
  kind: "region,formula",
  sheetId: "전체코드_입력",
  range: "A5:X12",
  maxChars: 4500,
  tableMaxRows: 8,
  tableMaxCols: 24,
});

const previewInput = await workbook.render({
  sheetName: "전체코드_입력",
  range: "A1:S14",
  scale: 1,
  format: "png",
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewInputPath, new Uint8Array(await previewInput.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(JSON.stringify({
  outputPath,
  previewInputPath,
  rows: rows.length,
  customers: customers.length,
}, null, 2));
console.log(inputCheck.ndjson);

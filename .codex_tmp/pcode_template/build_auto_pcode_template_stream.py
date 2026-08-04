# -*- coding: utf-8 -*-
import json
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parents[2]
WORK_DIR = BASE_DIR / ".codex_tmp" / "pcode_template"
OUTPUT_DIR = BASE_DIR / "outputs" / "pcode_template"
INPUT_JSON = WORK_DIR / "pcode_auto_rows.json"
OUTPUT_PATH = OUTPUT_DIR / "전체_P코드_자동조회_거래처별_양식.xlsx"

INPUT_HEADERS = [
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
]
CUSTOMER_HEADERS = INPUT_HEADERS[2:19]
FIRST_DATA_ROW = 6
HEADER_ROW = 5

NAVY = "172554"
LIGHT_BLUE = "DBEAFE"
GRAY = "F8FAFC"
LIGHT_AMBER = "FEF3C7"
TEXT = "0F172A"
MUTED = "64748B"
BORDER = "CBD5E1"
RED = "DC2626"

thin_border = Border(
    left=Side(style="thin", color=BORDER),
    right=Side(style="thin", color=BORDER),
    top=Side(style="thin", color=BORDER),
    bottom=Side(style="thin", color=BORDER),
)


def make_cell(ws, value, fill=None, font=None, align=None, border=False, number_format=None):
    cell = WriteOnlyCell(ws, value=value)
    if fill:
        cell.fill = fill
    if font:
        cell.font = font
    if align:
        cell.alignment = align
    if border:
        cell.border = thin_border
    if number_format:
        cell.number_format = number_format
    return cell


def safe_sheet_name(raw, used):
    invalid = '\\/*?:[]'
    name = "".join("_" if ch in invalid else ch for ch in str(raw or "거래처 미지정")).strip()
    name = " ".join(name.split()) or "거래처 미지정"
    name = name[:31]
    base = name
    idx = 2
    while name in used:
        suffix = f"_{idx}"
        name = f"{base[:31-len(suffix)]}{suffix}"
        idx += 1
    used.add(name)
    return name


def set_widths(ws, widths):
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def append_title(ws, title, subtitle, col_count):
    title_row = [
        make_cell(
            ws,
            title if i == 0 else "",
            fill=PatternFill("solid", fgColor=NAVY),
            font=Font(name="맑은 고딕", bold=True, color="FFFFFF", size=15),
            align=Alignment(horizontal="left"),
        )
        for i in range(col_count)
    ]
    subtitle_row = [
        make_cell(
            ws,
            subtitle if i == 0 else "",
            fill=PatternFill("solid", fgColor=GRAY),
            font=Font(name="맑은 고딕", color=MUTED, size=10),
            align=Alignment(wrap_text=True),
        )
        for i in range(col_count)
    ]
    ws.append(title_row)
    ws.append(subtitle_row)


def styled_header(ws, headers):
    fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    font = Font(name="맑은 고딕", bold=True, color=TEXT)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.append([make_cell(ws, h, fill=fill, font=font, align=align, border=True) for h in headers])


def input_row(ws, rec, row_idx):
    values = [
        rec.get("거래처그룹") or "거래처 미지정",
        rec.get("제품분류") or "기타",
        rec.get("생산코드") or "",
        rec.get("분리코드") or "",
        rec.get("사출코드") or "",
        rec.get("제품명") or "",
        rec.get("파워") or "",
        rec.get("오더수량1") or 0,
        rec.get("오더수량2") or 0,
        f"=SUM(H{row_idx}:I{row_idx})",
        rec.get("제품부족수량") or 0,
        rec.get("사출부족수량") or 0,
        rec.get("사출재고") or 0,
        rec.get("분리재고") or 0,
        rec.get("검사접착재고") or 0,
        rec.get("누수규격검사재고") or 0,
        f"=SUM(M{row_idx}:P{row_idx})",
        rec.get("DOI") or 0,
        (
            f'=IF(A{row_idx}="","",IF(L{row_idx}>0,"사출부족",'
            f'IF(K{row_idx}>0,"제품부족",IF(AND(J{row_idx}=0,Q{row_idx}>0),"수요없음재고",'
            f'IF(AND(ISNUMBER(R{row_idx}),R{row_idx}>0,R{row_idx}<7),"DOI주의",'
            f'IF(J{row_idx}>0,"수요있음","수요없음"))))))'
        ),
        rec.get("납기일") or "",
        rec.get("이니셜") or "",
        rec.get("신호") or "",
        rec.get("기존상태") or "",
        rec.get("비고") or "",
    ]
    row = []
    for col_idx, value in enumerate(values, start=1):
        number_format = None
        if col_idx in {8, 9, 10, 11, 12, 13, 14, 15, 16, 17}:
            number_format = "#,##0"
        elif col_idx == 18:
            number_format = "0.0"
        row.append(make_cell(ws, value, font=Font(name="맑은 고딕", size=10, color=TEXT), number_format=number_format))
    return row


def build_workbook():
    payload = json.loads(INPUT_JSON.read_text(encoding="utf-8"))
    rows = payload["rows"]
    customers = payload["customers"]
    row_count = len(rows)
    last_row = FIRST_DATA_ROW + row_count - 1

    counts = {}
    for rec in rows:
        key = rec.get("거래처그룹") or "거래처 미지정"
        counts[key] = counts.get(key, 0) + 1

    wb = Workbook(write_only=True)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    used_names = {"안내", "전체코드_입력"}
    sheet_map = [(cust, safe_sheet_name(cust, used_names), counts.get(cust, 0)) for cust in customers]

    guide = wb.create_sheet("안내")
    append_title(
        guide,
        "전체 P코드 자동조회 양식",
        "전체코드_입력 시트에 값을 입력하면 거래처별 시트가 자동으로 해당 거래처 행만 조회합니다.",
        4,
    )
    guide.append([])
    styled_header(guide, ["구분", "사용 방법", "비고", ""])
    guide_rows = [
        ["전체코드_입력", "전체 P코드 기준 원장입니다. 값은 이 시트에서만 수정합니다.", f"{row_count:,}개 P코드", ""],
        ["거래처별 시트", "각 거래처 시트는 FILTER 수식으로 자동 조회됩니다.", f"{len(customers):,}개 거래처그룹", ""],
        ["오더합계", "오더수량1 + 오더수량2 수식입니다.", "필요 시 오더수량 열 추가 가능", ""],
        ["공정재고 합계", "사출재고 + 분리재고 + 검사접착재고 + 누수규격검사재고 수식입니다.", "", ""],
        ["상태", "사출부족, 제품부족, 수요없음재고, DOI주의 등을 자동 표시합니다.", "", ""],
    ]
    for row in guide_rows:
        guide.append([make_cell(guide, v, font=Font(name="맑은 고딕", size=10), border=True) for v in row])
    guide.append([])
    styled_header(guide, ["거래처그룹", "시트명", "행수", ""])
    for cust, sheet_name, count in sheet_map:
        guide.append([
            make_cell(guide, cust, font=Font(name="맑은 고딕", size=10), border=True),
            make_cell(guide, sheet_name, font=Font(name="맑은 고딕", size=10), border=True),
            make_cell(guide, count, font=Font(name="맑은 고딕", size=10), border=True, number_format="#,##0"),
            make_cell(guide, "", border=True),
        ])
    set_widths(guide, [28, 70, 18, 12])

    source = wb.create_sheet("전체코드_입력")
    append_title(source, "전체코드 입력", "이 시트가 원본입니다. 여기 값을 수정하면 거래처별 시트가 자동 갱신됩니다.", len(INPUT_HEADERS))
    source.append([
        make_cell(
            source,
            "주의: 거래처별 시트는 수식 조회 결과입니다. 품목 정보, 수요, 재고, DOI는 이 시트에서만 수정하세요.",
            fill=PatternFill("solid", fgColor=LIGHT_AMBER),
            font=Font(name="맑은 고딕", color="B45309", bold=True),
            align=Alignment(wrap_text=True),
        ),
        *[make_cell(source, "", fill=PatternFill("solid", fgColor=LIGHT_AMBER)) for _ in range(len(INPUT_HEADERS) - 1)],
    ])
    source.append([])
    styled_header(source, INPUT_HEADERS)
    for idx, rec in enumerate(rows, start=FIRST_DATA_ROW):
        source.append(input_row(source, rec, idx))
    source.freeze_panes = "G6"
    source.auto_filter.ref = f"A{HEADER_ROW}:X{last_row}"
    set_widths(source, [20, 13, 18, 18, 18, 34, 10, 12, 12, 13, 15, 15, 13, 13, 15, 18, 15, 10, 14, 13, 12, 10, 20, 26])

    for cust, sheet_name, _count in sheet_map:
        ws = wb.create_sheet(sheet_name)
        append_title(ws, f"{cust} 자동조회", "전체코드_입력 시트에서 거래처그룹이 같은 행만 자동 표시합니다.", len(CUSTOMER_HEADERS))
        ws.append([
            make_cell(ws, "조회 거래처", fill=PatternFill("solid", fgColor=LIGHT_BLUE), font=Font(name="맑은 고딕", bold=True), border=True),
            make_cell(ws, cust, fill=PatternFill("solid", fgColor=GRAY), font=Font(name="맑은 고딕", bold=True), border=True),
            *[make_cell(ws, "") for _ in range(len(CUSTOMER_HEADERS) - 2)],
        ])
        ws.append([])
        styled_header(ws, CUSTOMER_HEADERS)
        formula = f'=FILTER(\'전체코드_입력\'!$C${FIRST_DATA_ROW}:$S${last_row},\'전체코드_입력\'!$A${FIRST_DATA_ROW}:$A${last_row}=$B$3,"")'
        ws.append([make_cell(ws, formula)] + [make_cell(ws, "") for _ in range(len(CUSTOMER_HEADERS) - 1)])
        ws.freeze_panes = "E6"
        set_widths(ws, [18, 18, 18, 34, 10, 12, 12, 13, 15, 15, 13, 13, 15, 18, 15, 10, 14])

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    return OUTPUT_PATH, row_count, len(customers)


def verify(path: Path):
    wb = load_workbook(path, read_only=False, data_only=False)
    source = wb["전체코드_입력"]
    first_customer = next(name for name in wb.sheetnames if name not in {"안내", "전체코드_입력"})
    customer = wb[first_customer]
    checks = {
        "sheets": len(wb.sheetnames),
        "source_a5": source["A5"].value,
        "source_c6": source["C6"].value,
        "source_j6_formula": source["J6"].value,
        "source_q6_formula": source["Q6"].value,
        "source_s6_formula": source["S6"].value,
        "customer_sheet": first_customer,
        "customer_a6_formula": customer["A6"].value,
    }
    wb.close()
    return checks


if __name__ == "__main__":
    path, rows, customers = build_workbook()
    checks = verify(path)
    print(json.dumps({"output": str(path), "rows": rows, "customers": customers, "checks": checks}, ensure_ascii=False, indent=2))

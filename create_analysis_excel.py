# -*- coding: utf-8 -*-
"""
AHP 분석 엑셀 생성 - 설문 CSV 데이터와 연동
- 다수 응답자 데이터 입력 시트
- 기하평균 자동 계산
- AHP 가중치, lambda_max, CI, CR 자동 계산
- 차트 시트
- 한계 및 결론 시트
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.chart import BarChart, Reference, PieChart, RadarChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter

# ============================================================
# 데이터 정의
# ============================================================
LEVEL1 = ["자립성", "연계성", "실용성", "포용성", "지속가능성", "확장성"]
LEVEL1_DEF = {
    "자립성": "지역이 외부 의존을 줄이고 자체 혁신역량을 축적하는 능력",
    "연계성": "지역혁신주체 간 네트워크와 협력의 정도",
    "실용성": "지역혁신정책이 기업성과와 산업현장 문제 해결로 연결되는 정도",
    "포용성": "다양한 지역 주체가 혁신과정에 참여하고 성과를 공유하는 정도",
    "지속가능성": "지역혁신체제가 장기적으로 유지·발전될 수 있는 제도적 기반",
    "확장성": "지역 내부 혁신성과가 외부시장, 신산업, 글로벌 네트워크로 확산되는 능력",
}

LEVEL2 = {
    "자립성": ["지역 R&D 역량 강화", "지역기업 기술흡수역량 강화", "지역혁신 인재양성"],
    "연계성": ["산·학·연·관 협력 강화", "기업 간 네트워크 활성화", "중간지원기관 조정 기능 강화"],
    "실용성": ["기술사업화 지원", "기업 기술애로 해결", "매출·고용 창출 연계"],
    "포용성": ["중소기업 정책 접근성 강화", "청년·창업기업 참여 확대", "지역대학 및 지역사회 참여 확대"],
    "지속가능성": ["정책 연속성 확보", "혁신거버넌스 안정화", "혁신인프라 유지·고도화"],
    "확장성": ["외부 지식·기술 연계", "글로벌 시장 진출 지원", "신산업 전환 및 확산"],
}

RI_TABLE = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}

# ============================================================
# 스타일
# ============================================================
TITLE_FONT = Font(name="맑은 고딕", size=16, bold=True, color="1F4E79")
HEADER_FONT = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
SUBHEADER_FONT = Font(name="맑은 고딕", size=11, bold=True, color="1F4E79")
NORMAL_FONT = Font(name="맑은 고딕", size=10)
BOLD_FONT = Font(name="맑은 고딕", size=10, bold=True)
SMALL_FONT = Font(name="맑은 고딕", size=9, color="666666")
RESULT_FONT = Font(name="맑은 고딕", size=11, bold=True, color="27AE60")
WARN_FONT = Font(name="맑은 고딕", size=11, bold=True, color="C0392B")

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
RESULT_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
GEOMEAN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

MAX_RESPONDENTS = 30  # 최대 응답자 수

def style_cell(cell, font=NORMAL_FONT, fill=None, align=CENTER, border=THIN_BORDER, num_fmt=None):
    cell.font = font
    if fill: cell.fill = fill
    cell.alignment = align
    cell.border = border
    if num_fmt: cell.number_format = num_fmt


# ============================================================
# 워크북 생성
# ============================================================
wb = openpyxl.Workbook()

# ============================================================
# Sheet 1: 응답데이터 (Raw Data Input)
# ============================================================
ws_data = wb.active
ws_data.title = "응답데이터"
ws_data.sheet_properties.tabColor = "E67E22"

ws_data.merge_cells("A1:Z1")
style_cell(ws_data["A1"], font=TITLE_FONT, align=Alignment(horizontal="center", vertical="center"))
ws_data["A1"].value = "AHP 설문 응답 데이터 입력"

ws_data.merge_cells("A2:Z2")
ws_data["A2"].value = ("CSV 파일의 데이터를 아래에 붙여넣으세요. 각 행이 한 명의 응답자입니다. "
                       "노란색 셀에 쌍대비교값을 입력합니다. (왼쪽 중요: 1~9, 오른쪽 중요: 1/3~1/9)")
style_cell(ws_data["A2"], font=NORMAL_FONT, align=LEFT_WRAP)
ws_data.row_dimensions[2].height = 35

# 응답자 일반사항 열 정의
INFO_HEADERS = ["응답자", "응답일시", "소속유형", "소속기관", "직위", "관심분야", "경력", "성별", "연락처"]
INFO_COL_COUNT = len(INFO_HEADERS)  # 9
INFO_FILL = PatternFill(start_color="E8DAEF", end_color="E8DAEF", fill_type="solid")

# Build column headers for pairwise comparison values
# Columns 1~9 = 응답자 일반사항, 10~ = 쌍대비교값
pair_headers = []
pair_ids = []  # (section_key, i, j) for mapping

# L1 pairs
for i in range(len(LEVEL1)):
    for j in range(i + 1, len(LEVEL1)):
        pair_headers.append(f"{LEVEL1[i]} vs {LEVEL1[j]}")
        pair_ids.append(("L1", i, j))

# L2 pairs
for key in LEVEL1:
    factors = LEVEL2[key]
    for i in range(len(factors)):
        for j in range(i + 1, len(factors)):
            pair_headers.append(f"[{key}] {factors[i]} vs {factors[j]}")
            pair_ids.append((key, i, j))

total_pairs = len(pair_headers)  # 15 + 18 = 33
PAIR_START_COL = INFO_COL_COUNT + 1  # 쌍대비교 시작 열 (10)

# Header row 3: Section labels
row = 3

# 일반사항 섹션 헤더
ws_data.merge_cells(start_row=row, start_column=1, end_row=row, end_column=INFO_COL_COUNT)
style_cell(ws_data.cell(row=row, column=1, value="응답자 일반사항"),
           font=HEADER_FONT, fill=PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid"))

# L1 section header
l1_start_col = PAIR_START_COL
l1_end_col = l1_start_col + 14  # 15 pairs
ws_data.merge_cells(start_row=row, start_column=l1_start_col, end_row=row, end_column=l1_end_col)
style_cell(ws_data.cell(row=row, column=l1_start_col, value="제1계층: 핵심가치 간 비교 (15쌍)"),
           font=HEADER_FONT, fill=PatternFill(start_color="2E86C1", end_color="2E86C1", fill_type="solid"))

# L2 section headers
col = l1_end_col + 1
for key in LEVEL1:
    n = len(LEVEL2[key])
    num_pairs = n * (n - 1) // 2
    end_c = col + num_pairs - 1
    ws_data.merge_cells(start_row=row, start_column=col, end_row=row, end_column=end_c)
    style_cell(ws_data.cell(row=row, column=col, value=f"{key} ({num_pairs}쌍)"),
               font=HEADER_FONT, fill=PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"))
    col = end_c + 1

# Header row 4: Column labels
row = 4

# 일반사항 헤더
for idx, h in enumerate(INFO_HEADERS):
    style_cell(ws_data.cell(row=row, column=idx + 1, value=h),
               font=HEADER_FONT, fill=PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid"))

# 쌍대비교 헤더
for idx, h in enumerate(pair_headers):
    style_cell(ws_data.cell(row=row, column=PAIR_START_COL + idx, value=h),
               font=Font(name="맑은 고딕", size=8, bold=True, color="FFFFFF"), fill=HEADER_FILL)

# 열 너비 설정
info_widths = [10, 18, 12, 20, 12, 24, 12, 8, 16]
for idx, w in enumerate(info_widths):
    ws_data.column_dimensions[get_column_letter(idx + 1)].width = w
for c in range(PAIR_START_COL, PAIR_START_COL + total_pairs):
    ws_data.column_dimensions[get_column_letter(c)].width = 14

# Data input rows
for resp in range(1, MAX_RESPONDENTS + 1):
    r = row + resp
    # 응답자 번호
    style_cell(ws_data.cell(row=r, column=1, value=f"응답자{resp}"), font=BOLD_FONT, fill=LIGHT_FILL)
    # 일반사항 입력 셀 (보라색 연한 배경)
    for c in range(2, INFO_COL_COUNT + 1):
        style_cell(ws_data.cell(row=r, column=c), fill=INFO_FILL)
    # 쌍대비교 입력 셀 (노란 배경)
    for c in range(PAIR_START_COL, PAIR_START_COL + total_pairs):
        style_cell(ws_data.cell(row=r, column=c), fill=INPUT_FILL, num_fmt='0.0000')

data_start_row = 5
data_end_row = data_start_row + MAX_RESPONDENTS - 1

# Geometric mean row (쌍대비교 열만 기하평균 계산)
geomean_row = data_end_row + 2
style_cell(ws_data.cell(row=geomean_row, column=1, value="기하평균"), font=BOLD_FONT, fill=GEOMEAN_FILL)
for c in range(2, INFO_COL_COUNT + 1):
    ws_data.cell(row=geomean_row, column=c).border = THIN_BORDER
for c in range(PAIR_START_COL, PAIR_START_COL + total_pairs):
    col_letter = get_column_letter(c)
    formula = f'=IF(COUNT({col_letter}{data_start_row}:{col_letter}{data_end_row})=0,1,GEOMEAN({col_letter}{data_start_row}:{col_letter}{data_end_row}))'
    cell = ws_data.cell(row=geomean_row, column=c, value=formula)
    style_cell(cell, font=BOLD_FONT, fill=GEOMEAN_FILL, num_fmt='0.0000')

# Count row
count_row = geomean_row + 1
style_cell(ws_data.cell(row=count_row, column=1, value="응답수"), font=BOLD_FONT, fill=LIGHT_FILL)
for c in range(2, INFO_COL_COUNT + 1):
    ws_data.cell(row=count_row, column=c).border = THIN_BORDER
for c in range(PAIR_START_COL, PAIR_START_COL + total_pairs):
    col_letter = get_column_letter(c)
    formula = f'=COUNT({col_letter}{data_start_row}:{col_letter}{data_end_row})'
    cell = ws_data.cell(row=count_row, column=c, value=formula)
    style_cell(cell, font=BOLD_FONT, fill=LIGHT_FILL)


# ============================================================
# Sheet 2: AHP 분석결과
# ============================================================
ws_calc = wb.create_sheet("AHP분석결과")
ws_calc.sheet_properties.tabColor = "27AE60"

ws_calc.merge_cells("A1:J1")
ws_calc["A1"].value = "AHP 분석 결과 (자동 계산)"
ws_calc["A1"].font = TITLE_FONT

ws_calc.merge_cells("A2:J2")
ws_calc["A2"].value = "응답데이터 시트의 기하평균을 기반으로 자동 계산됩니다."
ws_calc["A2"].font = NORMAL_FONT
ws_calc["A2"].alignment = LEFT_WRAP

# Track which geomean column corresponds to which pair
pair_col_map = {}  # (section_key, i, j) -> column_letter in 응답데이터
for idx, (sec, i, j) in enumerate(pair_ids):
    pair_col_map[(sec, i, j)] = get_column_letter(PAIR_START_COL + idx)


def create_ahp_section(ws, start_row, title, items, section_key):
    """Create full AHP calculation section. Returns (next_row, weight_cell_refs, cr_cell_ref)"""
    n = len(items)
    ri = RI_TABLE.get(n, 1.49)

    # Title
    ws.merge_cells(f"A{start_row}:{get_column_letter(n + 4)}{start_row}")
    style_cell(ws.cell(row=start_row, column=1, value=title), font=SUBHEADER_FONT, fill=LIGHT_FILL)
    start_row += 1

    # === Pairwise Comparison Matrix ===
    ws.cell(row=start_row, column=1, value="[쌍대비교행렬 (기하평균)]").font = BOLD_FONT
    start_row += 1

    mat_header_row = start_row
    style_cell(ws.cell(row=mat_header_row, column=1, value=""), fill=HEADER_FILL)
    for j in range(n):
        style_cell(ws.cell(row=mat_header_row, column=j + 2, value=items[j]),
                   font=HEADER_FONT, fill=HEADER_FILL)

    mat_start = mat_header_row + 1
    for i in range(n):
        style_cell(ws.cell(row=mat_start + i, column=1, value=items[i]), font=BOLD_FONT, fill=LIGHT_FILL, align=LEFT_WRAP)
        for j in range(n):
            cell = ws.cell(row=mat_start + i, column=j + 2)
            cell.border = THIN_BORDER
            cell.alignment = CENTER
            cell.number_format = '0.0000'
            cell.font = NORMAL_FONT

            if i == j:
                cell.value = 1
            elif i < j:
                # Get from geomean row
                col_letter = pair_col_map.get((section_key, i, j))
                if col_letter:
                    cell.value = f"='응답데이터'!{col_letter}{geomean_row}"
            else:
                # Reciprocal: 1/a(j,i)
                ref_cell = f"{get_column_letter(i + 2)}{mat_start + j}"
                cell.value = f"=1/{ref_cell}"

    # Column sum row
    sum_row = mat_start + n
    style_cell(ws.cell(row=sum_row, column=1, value="열 합계"), font=BOLD_FONT, fill=RESULT_FILL)
    for j in range(n):
        cl = get_column_letter(j + 2)
        cell = ws.cell(row=sum_row, column=j + 2)
        cell.value = f"=SUM({cl}{mat_start}:{cl}{mat_start + n - 1})"
        style_cell(cell, font=BOLD_FONT, fill=RESULT_FILL, num_fmt='0.0000')

    # === Normalized Matrix + Weights ===
    norm_title_row = sum_row + 2
    ws.cell(row=norm_title_row, column=1, value="[정규화행렬 및 가중치]").font = BOLD_FONT

    norm_header = norm_title_row + 1
    style_cell(ws.cell(row=norm_header, column=1, value=""), fill=HEADER_FILL)
    for j in range(n):
        style_cell(ws.cell(row=norm_header, column=j + 2, value=items[j]),
                   font=HEADER_FONT, fill=HEADER_FILL)
    style_cell(ws.cell(row=norm_header, column=n + 2, value="가중치(Wi)"),
               font=HEADER_FONT, fill=PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"))

    norm_start = norm_header + 1
    weight_cells = []
    for i in range(n):
        style_cell(ws.cell(row=norm_start + i, column=1, value=items[i]), font=BOLD_FONT, fill=LIGHT_FILL, align=LEFT_WRAP)
        for j in range(n):
            mat_ref = f"{get_column_letter(j + 2)}{mat_start + i}"
            sum_ref = f"{get_column_letter(j + 2)}{sum_row}"
            cell = ws.cell(row=norm_start + i, column=j + 2)
            cell.value = f"={mat_ref}/{sum_ref}"
            style_cell(cell, num_fmt='0.0000')

        # Weight = row average
        first_c = get_column_letter(2)
        last_c = get_column_letter(n + 1)
        w_cell = ws.cell(row=norm_start + i, column=n + 2)
        w_cell.value = f"=AVERAGE({first_c}{norm_start + i}:{last_c}{norm_start + i})"
        style_cell(w_cell, font=RESULT_FONT, fill=RESULT_FILL, num_fmt='0.0000')
        weight_cells.append(f"{get_column_letter(n + 2)}{norm_start + i}")

    # === Consistency Check ===
    cons_title = norm_start + n + 1
    ws.cell(row=cons_title, column=1, value="[일관성 검증]").font = BOLD_FONT

    cons_header = cons_title + 1
    for c, h in enumerate(["항목", "가중치(Wi)", "가중합(Aw)", "Aw/Wi"], 1):
        style_cell(ws.cell(row=cons_header, column=c, value=h), font=HEADER_FONT, fill=HEADER_FILL)

    cons_start = cons_header + 1
    aw_wi_cells = []
    for i in range(n):
        r = cons_start + i
        style_cell(ws.cell(row=r, column=1, value=items[i]), align=LEFT_WRAP)
        ws.cell(row=r, column=1).border = THIN_BORDER

        # Wi
        cell_b = ws.cell(row=r, column=2)
        cell_b.value = f"={weight_cells[i]}"
        style_cell(cell_b, num_fmt='0.0000')

        # Aw = sum(matrix_row * weights)
        terms = []
        for j in range(n):
            mat_ref = f"{get_column_letter(j + 2)}{mat_start + i}"
            terms.append(f"{mat_ref}*{weight_cells[j]}")
        cell_c = ws.cell(row=r, column=3)
        cell_c.value = "=" + "+".join(terms)
        style_cell(cell_c, num_fmt='0.0000')

        # Aw/Wi
        cell_d = ws.cell(row=r, column=4)
        cell_d.value = f"=C{r}/B{r}"
        style_cell(cell_d, num_fmt='0.0000')
        aw_wi_cells.append(f"D{r}")

    # λmax, CI, RI, CR
    res_start = cons_start + n + 1

    labels = ["λmax", "CI", "RI", "CR", "판정결과"]
    for idx, lbl in enumerate(labels):
        style_cell(ws.cell(row=res_start + idx, column=1, value=lbl), font=BOLD_FONT, fill=RESULT_FILL)

    # λmax
    lmax_cell = ws.cell(row=res_start, column=2)
    lmax_cell.value = f"=AVERAGE({','.join(aw_wi_cells)})"
    style_cell(lmax_cell, font=BOLD_FONT, fill=RESULT_FILL, num_fmt='0.0000')

    # CI
    ci_cell = ws.cell(row=res_start + 1, column=2)
    ci_cell.value = f"=(B{res_start}-{n})/({n}-1)" if n > 1 else 0
    style_cell(ci_cell, font=BOLD_FONT, fill=RESULT_FILL, num_fmt='0.0000')

    # RI
    ri_cell = ws.cell(row=res_start + 2, column=2)
    ri_cell.value = ri
    style_cell(ri_cell, font=BOLD_FONT, fill=RESULT_FILL, num_fmt='0.0000')

    # CR
    cr_cell = ws.cell(row=res_start + 3, column=2)
    if ri > 0:
        cr_cell.value = f"=B{res_start + 1}/{ri}"
    else:
        cr_cell.value = 0
    style_cell(cr_cell, font=BOLD_FONT, fill=RESULT_FILL, num_fmt='0.0000')

    # 판정
    judge_cell = ws.cell(row=res_start + 4, column=2)
    if ri > 0:
        judge_cell.value = f'=IF(B{res_start + 3}<=0.1,"PASS (CR<=0.1, 적합)","FAIL (CR>0.1, 부적합 - 재설문 필요)")'
    else:
        judge_cell.value = "PASS (n<=2)"
    style_cell(judge_cell, font=WARN_FONT, fill=RESULT_FILL, align=LEFT_WRAP)
    ws.column_dimensions["B"].width = 40

    next_row = res_start + 6
    cr_ref = f"B{res_start + 3}"
    return next_row, weight_cells, cr_ref


# Build all sections
calc_row = 4
weight_refs = {}
cr_refs = {}

calc_row, w, cr = create_ahp_section(ws_calc, calc_row, "제1계층: 핵심가치 가중치 분석", LEVEL1, "L1")
weight_refs["L1"] = w
cr_refs["L1"] = cr

for key in LEVEL1:
    calc_row, w, cr = create_ahp_section(ws_calc, calc_row, f"제2계층: {key} - 정책요인 가중치 분석", LEVEL2[key], key)
    weight_refs[key] = w
    cr_refs[key] = cr

ws_calc.column_dimensions["A"].width = 30
for c in range(2, 10):
    ws_calc.column_dimensions[get_column_letter(c)].width = 14


# ============================================================
# Sheet 3: 종합순위
# ============================================================
ws_rank = wb.create_sheet("종합순위")
ws_rank.sheet_properties.tabColor = "8E44AD"

ws_rank.merge_cells("A1:G1")
ws_rank["A1"].value = "AHP 종합 가중치 및 우선순위"
ws_rank["A1"].font = TITLE_FONT

# --- Level 1 ---
row = 3
ws_rank.cell(row=row, column=1, value="[제1계층 핵심가치 가중치]").font = BOLD_FONT
row += 1

for c, h in enumerate(["핵심가치", "가중치", "순위", "CR값", "적합성"], 1):
    style_cell(ws_rank.cell(row=row, column=c, value=h), font=HEADER_FONT, fill=HEADER_FILL)
row += 1

l1_start = row
for i, item in enumerate(LEVEL1):
    style_cell(ws_rank.cell(row=row, column=1, value=item), font=BOLD_FONT, align=CENTER)
    w_cell = ws_rank.cell(row=row, column=2)
    w_cell.value = f"='AHP분석결과'!{weight_refs['L1'][i]}"
    style_cell(w_cell, num_fmt='0.0000')

    rank_cell = ws_rank.cell(row=row, column=3)
    rank_cell.value = f"=RANK(B{row},B${l1_start}:B${l1_start + len(LEVEL1) - 1})"
    style_cell(rank_cell)

    if i == 0:
        cr_c = ws_rank.cell(row=row, column=4)
        cr_c.value = f"='AHP분석결과'!{cr_refs['L1']}"
        style_cell(cr_c, num_fmt='0.0000')
        j_c = ws_rank.cell(row=row, column=5)
        j_c.value = f'=IF(D{row}<=0.1,"적합","부적합")'
        style_cell(j_c, font=WARN_FONT)
    else:
        ws_rank.cell(row=row, column=4).border = THIN_BORDER
        ws_rank.cell(row=row, column=5).border = THIN_BORDER
    row += 1
l1_end = row - 1

# --- Global Weights ---
row += 1
ws_rank.cell(row=row, column=1, value="[종합 가중치 (글로벌 가중치)]").font = BOLD_FONT
row += 1

for c, h in enumerate(["핵심가치", "정책요인", "로컬가중치", "핵심가치가중치", "글로벌가중치", "종합순위", "CR값"], 1):
    style_cell(ws_rank.cell(row=row, column=c, value=h), font=HEADER_FONT, fill=HEADER_FILL)
row += 1

global_start = row
all_items = []
for ki, key in enumerate(LEVEL1):
    l1_w = f"B{l1_start + ki}"
    for fi, factor in enumerate(LEVEL2[key]):
        style_cell(ws_rank.cell(row=row, column=1, value=key), align=CENTER)
        style_cell(ws_rank.cell(row=row, column=2, value=factor), align=LEFT_WRAP)

        local_w = ws_rank.cell(row=row, column=3)
        local_w.value = f"='AHP분석결과'!{weight_refs[key][fi]}"
        style_cell(local_w, num_fmt='0.0000')

        l1w = ws_rank.cell(row=row, column=4)
        l1w.value = f"={l1_w}"
        style_cell(l1w, num_fmt='0.0000')

        gw = ws_rank.cell(row=row, column=5)
        gw.value = f"=C{row}*D{row}"
        style_cell(gw, font=RESULT_FONT, fill=RESULT_FILL, num_fmt='0.0000')

        ws_rank.cell(row=row, column=6).border = THIN_BORDER
        ws_rank.cell(row=row, column=6).alignment = CENTER

        if fi == 0:
            cr_c = ws_rank.cell(row=row, column=7)
            cr_c.value = f"='AHP분석결과'!{cr_refs[key]}"
            style_cell(cr_c, num_fmt='0.0000')
        else:
            ws_rank.cell(row=row, column=7).border = THIN_BORDER

        all_items.append(row)
        row += 1

global_end = row - 1

# Fill rank formulas
for r in all_items:
    ws_rank.cell(row=r, column=6).value = f"=RANK(E{r},E${global_start}:E${global_end})"
    style_cell(ws_rank.cell(row=r, column=6))

ws_rank.column_dimensions["A"].width = 14
ws_rank.column_dimensions["B"].width = 30
ws_rank.column_dimensions["C"].width = 14
ws_rank.column_dimensions["D"].width = 16
ws_rank.column_dimensions["E"].width = 14
ws_rank.column_dimensions["F"].width = 10
ws_rank.column_dimensions["G"].width = 10


# ============================================================
# Charts - 각 차트를 별도 시트에 배치
# ============================================================
colors = ["1F4E79", "2E86C1", "27AE60", "E67E22", "8E44AD", "C0392B"]

# 공통 데이터 참조
data1 = Reference(ws_rank, min_col=2, min_row=l1_start - 1, max_row=l1_end)
cats1 = Reference(ws_rank, min_col=1, min_row=l1_start, max_row=l1_end)
data3 = Reference(ws_rank, min_col=5, min_row=global_start - 1, max_row=global_end)
cats3 = Reference(ws_rank, min_col=2, min_row=global_start, max_row=global_end)

# --- Chart 1: 핵심가치 막대차트 ---
ws_c1 = wb.create_sheet("차트1_핵심가치막대")
ws_c1.sheet_properties.tabColor = "E74C3C"
chart1 = BarChart()
chart1.type = "col"
chart1.title = "제1계층 핵심가치 가중치"
chart1.y_axis.title = "가중치"
chart1.style = 10
chart1.width = 28
chart1.height = 18
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
for i, clr in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = clr
    chart1.series[0].data_points.append(pt)
chart1.dataLabels = DataLabelList()
chart1.dataLabels.showVal = True
chart1.dataLabels.numFmt = '0.0000'
ws_c1.add_chart(chart1, "A1")

# --- Chart 2: 핵심가치 파이차트 ---
ws_c2 = wb.create_sheet("차트2_핵심가치파이")
ws_c2.sheet_properties.tabColor = "E74C3C"
pie = PieChart()
pie.title = "핵심가치 비중"
pie.style = 10
pie.width = 28
pie.height = 18
pie.add_data(data1, titles_from_data=True)
pie.set_categories(cats1)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showCatName = True
for i, clr in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = clr
    pie.series[0].data_points.append(pt)
ws_c2.add_chart(pie, "A1")

# --- Chart 3: 핵심가치 레이더차트 ---
ws_c3 = wb.create_sheet("차트3_핵심가치레이더")
ws_c3.sheet_properties.tabColor = "E74C3C"
radar = RadarChart()
radar.type = "marker"
radar.title = "핵심가치 레이더 차트"
radar.style = 26
radar.width = 28
radar.height = 18
radar.add_data(data1, titles_from_data=True)
radar.set_categories(cats1)
ws_c3.add_chart(radar, "A1")

# --- Chart 4: 글로벌 가중치 가로막대 ---
ws_c4 = wb.create_sheet("차트4_글로벌가중치")
ws_c4.sheet_properties.tabColor = "E74C3C"
chart3 = BarChart()
chart3.type = "bar"
chart3.title = "종합 글로벌 가중치 (18개 정책요인)"
chart3.x_axis.title = "글로벌 가중치"
chart3.style = 10
chart3.width = 32
chart3.height = 24
chart3.add_data(data3, titles_from_data=True)
chart3.set_categories(cats3)
color_idx = 0
for ki, key in enumerate(LEVEL1):
    for fi in range(len(LEVEL2[key])):
        pt = DataPoint(idx=color_idx)
        pt.graphicalProperties.solidFill = colors[ki]
        chart3.series[0].data_points.append(pt)
        color_idx += 1
chart3.dataLabels = DataLabelList()
chart3.dataLabels.showVal = True
chart3.dataLabels.numFmt = '0.0000'
ws_c4.add_chart(chart3, "A1")

# --- Charts 5~10: 핵심가치별 정책요인 차트 (각각 별도 시트) ---
item_idx = 0
for ki, key in enumerate(LEVEL1):
    n_factors = len(LEVEL2[key])
    sub_start_r = global_start + item_idx
    sub_end_r = sub_start_r + n_factors - 1

    ws_sub = wb.create_sheet(f"차트{5+ki}_{key}")
    ws_sub.sheet_properties.tabColor = colors[ki]

    sub_chart = BarChart()
    sub_chart.type = "col"
    sub_chart.title = f"{key} - 정책요인 로컬 가중치"
    sub_chart.y_axis.title = "가중치"
    sub_chart.style = 10
    sub_chart.width = 28
    sub_chart.height = 18

    sub_data = Reference(ws_rank, min_col=3, min_row=sub_start_r - 1, max_row=sub_end_r)
    sub_cats = Reference(ws_rank, min_col=2, min_row=sub_start_r, max_row=sub_end_r)
    sub_chart.add_data(sub_data, titles_from_data=True)
    sub_chart.set_categories(sub_cats)
    for j in range(n_factors):
        pt = DataPoint(idx=j)
        pt.graphicalProperties.solidFill = colors[ki]
        sub_chart.series[0].data_points.append(pt)
    sub_chart.dataLabels = DataLabelList()
    sub_chart.dataLabels.showVal = True
    sub_chart.dataLabels.numFmt = '0.0000'

    ws_sub.add_chart(sub_chart, "A1")
    item_idx += n_factors


# ============================================================
# Sheet 5: 한계 및 결론
# ============================================================
ws_conc = wb.create_sheet("한계및결론")
ws_conc.sheet_properties.tabColor = "2C3E50"

ws_conc.merge_cells("A1:H1")
ws_conc["A1"].value = "AHP 분석의 한계 및 결론"
ws_conc["A1"].font = TITLE_FONT

sections = [
    ("1. 연구 목적 및 의의", [
        "본 연구는 충남 지역혁신체제 구축을 위한 핵심가치(자립성, 연계성, 실용성, 포용성, 지속가능성, 확장성)와",
        "18개 정책요인의 상대적 중요도 및 우선순위를 AHP(Analytic Hierarchy Process) 기법을 통해 분석하였다.",
        "이를 통해 지역혁신정책 수립 시 전문가 관점에서의 정책적 우선순위를 실증적으로 도출하고자 하였다.",
    ]),
    ("2. 분석 방법론", [
        "AHP 기법은 Saaty(1980)가 개발한 다기준 의사결정 방법으로, 복잡한 의사결정 문제를 계층화하고",
        "쌍대비교를 통해 각 요인의 상대적 중요도를 도출한다.",
        "본 연구는 2계층 구조(제1계층: 6개 핵심가치, 제2계층: 각 3개 정책요인)로 설계하였으며,",
        "9점 척도를 활용한 쌍대비교를 통해 가중치를 산출하고 일관성 비율(CR<=0.1)을 검증하였다.",
        "",
        "[AHP 계산 절차]",
        "Step 1: 전문가 응답을 기하평균으로 통합하여 그룹 쌍대비교행렬(A) 구성",
        "Step 2: 각 열의 합계 산출",
        "Step 3: 정규화행렬 산출 (각 원소 / 해당 열 합계)",
        "Step 4: 가중치(Wi) = 정규화행렬의 각 행 평균",
        "Step 5: 가중합벡터(Aw) = 원래 행렬 x 가중치 벡터",
        "Step 6: lambda_max = (Aw/Wi)의 평균",
        "Step 7: CI = (lambda_max - n) / (n - 1)",
        "Step 8: CR = CI / RI, CR <= 0.1이면 적합",
        "",
        "[RI(무작위지수) 참조표 - Saaty]",
        "n=3: 0.58 / n=4: 0.90 / n=5: 1.12 / n=6: 1.24 / n=7: 1.32",
    ]),
    ("3. 연구 결과 해석 가이드", [
        "- 제1계층 가중치: 6개 핵심가치 중 어떤 가치가 지역혁신체제에서 가장 중요한지를 나타냄",
        "- 제2계층 로컬 가중치: 각 핵심가치 내에서 정책요인의 상대적 중요도",
        "- 글로벌 가중치: 제1계층 가중치 x 로컬 가중치로 산출, 18개 정책요인의 전체 순위 결정",
        "- CR 값이 0.1을 초과하는 경우 해당 응답의 일관성이 부족하므로 재설문이 필요함",
        "- 다수 응답자의 경우 기하평균(Geometric Mean)으로 통합하여 그룹 판단을 도출함",
    ]),
    ("4. 연구의 한계", [
        "(1) 전문가 패널 구성의 한계",
        "   - AHP 분석은 소수 전문가 패널의 판단에 의존하므로, 패널 구성에 따라 결과가 달라질 수 있음",
        "   - 특정 분야(산업계, 학계, 공공부문) 전문가의 과대/과소 대표 가능성 존재",
        "   - 충남 지역 특수성을 반영한 전문가 선정 기준의 객관성 확보 필요",
        "",
        "(2) AHP 기법의 방법론적 한계",
        "   - 9점 척도의 주관성: 응답자마다 척도 해석이 다를 수 있음",
        "   - 쌍대비교 항목 수 증가 시 응답 피로도 증가 및 일관성 저하 우려",
        "   - 본 연구의 제1계층 6개 항목 비교 시 15개 쌍대비교가 필요하여 응답 부담 존재",
        "   - 정량적 데이터가 아닌 주관적 판단에 기반하므로 결과의 일반화에 주의 필요",
        "",
        "(3) 계층구조 설계의 한계",
        "   - 각 핵심가치별 정책요인을 3개로 한정하여 일부 중요 정책요인이 누락되었을 가능성",
        "   - 정책요인 간 상호 연관성(독립성 가정)이 현실에서는 완전히 성립하지 않을 수 있음",
        "   - ANP(Analytic Network Process) 등 상호 의존성을 고려한 분석 기법 적용 필요성 존재",
        "",
        "(4) 시간적/공간적 한계",
        "   - 특정 시점의 전문가 인식을 반영하므로 정책환경 변화 시 결과 재검증 필요",
        "   - 충남 지역에 특화된 결과로 다른 지역에의 직접 적용에는 한계",
        "",
        "(5) 기하평균 통합의 한계",
        "   - 개별 응답자의 일관성 비율(CR) 검증 후 부적합 응답 제외가 이상적이나,",
        "     본 분석에서는 기하평균 통합 후 그룹 수준의 CR만 산출함",
        "   - 응답자 간 의견 분산(합의도)에 대한 추가 분석 필요",
    ]),
    ("5. 결론 및 정책적 시사점", [
        "본 연구는 AHP 분석을 통해 충남 지역혁신체제 구축을 위한 핵심가치와 정책요인의 우선순위를",
        "체계적으로 도출하였다. 분석 결과를 바탕으로 다음과 같은 정책적 시사점을 제시한다.",
        "",
        "(1) 가중치가 높은 핵심가치를 중심으로 지역혁신정책의 전략적 방향성 설정 필요",
        "(2) 글로벌 가중치 상위 정책요인에 대한 우선적 자원 배분 및 실행계획 수립 필요",
        "(3) 하위 순위 정책요인의 경우에도 장기적 관점에서 기반 조성 정책 병행 필요",
        "(4) 전문가 그룹별(산업계, 학계, 공공부문) 인식 차이 분석을 통한 맞춤형 정책 수립 검토",
        "(5) 정기적 AHP 재분석을 통한 정책환경 변화 모니터링 및 우선순위 갱신 권고",
        "",
        "향후 연구에서는 ANP 기법 적용, 응답자 그룹 간 비교분석, Fuzzy-AHP, DEMATEL-ANP 통합 분석 등을",
        "통해 보다 정교한 정책 우선순위 분석이 이루어질 필요가 있다.",
    ]),
    ("6. 참고문헌", [
        "- Saaty, T. L. (1980). The Analytic Hierarchy Process. McGraw-Hill.",
        "- Saaty, T. L. (2008). Decision making with the analytic hierarchy process.",
        "  International Journal of Services Sciences, 1(1), 83-98.",
        "- Saaty, T. L., & Peniwati, K. (2013). Group Decision Making.",
        "- Forman, E., & Peniwati, K. (1998). Aggregating individual judgments and",
        "  priorities with the analytic hierarchy process. EJOR, 108(1), 165-169.",
        "- 조근태, 조용곤, 강현수 (2003). 앞서가는 리더들의 계층분석적 의사결정. 동현출판사.",
    ]),
]

row = 3
for title, lines in sections:
    ws_conc.merge_cells(f"A{row}:H{row}")
    style_cell(ws_conc.cell(row=row, column=1, value=title), font=SUBHEADER_FONT, fill=LIGHT_FILL, align=LEFT_WRAP)
    row += 1
    for line in lines:
        ws_conc.merge_cells(f"A{row}:H{row}")
        ws_conc.cell(row=row, column=1, value=line).font = NORMAL_FONT
        ws_conc.cell(row=row, column=1).alignment = LEFT_WRAP
        row += 1
    row += 1

ws_conc.column_dimensions["A"].width = 100


# ============================================================
# Sheet 6: 사용법 안내
# ============================================================
ws_guide = wb.create_sheet("사용안내")
ws_guide.sheet_properties.tabColor = "3498DB"
ws_guide.merge_cells("A1:H1")
ws_guide["A1"].value = "엑셀 파일 사용 안내"
ws_guide["A1"].font = TITLE_FONT

guide_text = [
    ("사용 절차", [
        "1. 온라인 설문지(AHP_설문조사.html)를 응답자에게 배포합니다.",
        "2. 응답자가 설문 완료 후 CSV 파일을 다운로드합니다.",
        "3. CSV 파일을 열어 쌍대비교값(9번째 열부터)을 복사합니다.",
        "4. 본 엑셀의 '응답데이터' 시트 노란색 셀에 붙여넣습니다.",
        "5. 'AHP분석결과' 시트에서 자동 계산된 결과를 확인합니다.",
        "6. '종합순위' 시트에서 최종 우선순위를 확인합니다.",
        "7. '차트분석' 시트에서 시각화 결과를 확인합니다.",
    ]),
    ("입력값 설명", [
        "- 왼쪽 항목이 절대적으로 중요: 9",
        "- 왼쪽 항목이 매우 중요: 7",
        "- 왼쪽 항목이 강하게 중요: 5",
        "- 왼쪽 항목이 약간 중요: 3",
        "- 동등하게 중요: 1",
        "- 오른쪽 항목이 약간 중요: 0.3333 (=1/3)",
        "- 오른쪽 항목이 강하게 중요: 0.2 (=1/5)",
        "- 오른쪽 항목이 매우 중요: 0.1429 (=1/7)",
        "- 오른쪽 항목이 절대적으로 중요: 0.1111 (=1/9)",
    ]),
    ("시트 구성", [
        "- 응답데이터: 설문 응답 원시 데이터 입력 (최대 30명)",
        "- AHP분석결과: 쌍대비교행렬, 정규화, 가중치, CI/CR 자동 계산",
        "- 종합순위: 글로벌 가중치 및 18개 정책요인 종합순위",
        "- 차트분석: 막대, 레이더, 파이, 카테고리별 차트 (10개)",
        "- 한계및결론: 논문용 한계점, 결론, 정책적 시사점",
        "- 사용안내: 본 시트 (사용 방법 안내)",
    ]),
    ("주의사항", [
        "- 응답데이터 입력 시 빈 행이 중간에 있으면 기하평균 계산에 영향을 줄 수 있습니다.",
        "- CR > 0.1인 경우 해당 응답의 일관성이 부족하므로 재설문을 권장합니다.",
        "- 개별 응답자의 CR을 확인하려면 각 응답자 데이터를 개별적으로 분석해야 합니다.",
        "- 엑셀 수식이 포함되어 있으므로 셀 서식을 변경하지 마세요.",
    ]),
]

row = 3
for title, lines in guide_text:
    ws_guide.merge_cells(f"A{row}:H{row}")
    style_cell(ws_guide.cell(row=row, column=1, value=title), font=SUBHEADER_FONT, fill=LIGHT_FILL, align=LEFT_WRAP)
    row += 1
    for line in lines:
        ws_guide.merge_cells(f"A{row}:H{row}")
        ws_guide.cell(row=row, column=1, value=line).font = NORMAL_FONT
        ws_guide.cell(row=row, column=1).alignment = LEFT_WRAP
        row += 1
    row += 1

ws_guide.column_dimensions["A"].width = 80

# ============================================================
# 저장
# ============================================================
output = "충남_AHP_분석_결과_v3.xlsx"
wb.save(output)
print(f"파일 생성 완료: {output}")
print("시트:", [ws.title for ws in wb.worksheets])

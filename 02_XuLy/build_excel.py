"""Tạo sổ số liệu đo `BaoCao_SoLieu_Nhom3.xlsx` theo dạng biên bản đo của phòng thí nghiệm hạt nhân.

Số liệu thô lấy từ `nhóm 1.xlsx`; số liệu Monte Carlo và χ² Pearson lấy từ `results.json`
(chạy `python3 02_XuLy/analysis.py` trước). Mọi đại lượng thống kê trong sổ đều là công thức Excel.
"""
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "01_SoLieuGoc" / "nhóm 1.xlsx"
R = json.loads((ROOT / "03_KetQuaTrungGian" / "results.json").read_text())
OUT = ROOT / "04_NopBai" / "BaoCao_SoLieu_Nhom3.xlsx"

FONT = "Times New Roman"
DOC_CODE = "THVLHĐ-GM-N3-01"
SETS = ["S1", "S2", "S3"]
SET_NAME = {"S1": "Bộ 1 – Mức thấp (Phông)", "S2": "Bộ 2 – Mức trung bình (Ra-226)",
            "S3": "Bộ 3 – Mức cao (Ra-226)"}
K_RANGE = {"S1": (0, 16), "S2": (20, 64), "S3": (75, 150)}
MEMBERS = [
    ("24130044", "Phạm Việt Linh"), ("24130048", "Lê Quốc Long"),
    ("24130049", "Ngô Đặng Tuấn Long"), ("24130103", "Dương Quốc Huy"),
    ("24130162", "Trần Trí Dũng"), ("24130147", "Nguyễn Trần Nhiên Đăng"),
    ("24130157", "Lê Anh Đức"), ("24130159", "Trần Thanh Đức"),
    ("24130218", "Nguyễn Vũ Nhật Minh"),
]

thin = Side(style="thin", color="808080")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor="D9E2F3")
SEC_FILL = PatternFill("solid", fgColor="EDEDED")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
RES_FILL = PatternFill("solid", fgColor="E2EFDA")
BLUE = "0000FF"
GREEN = "008000"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def f(bold=False, color="000000", size=12, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def put(ws, ref, value, bold=False, color="000000", fill=None, fmt=None, align=CENTER,
        border=True, size=12, italic=False):
    c = ws[ref]
    c.value = value
    c.font = f(bold, color, size, italic)
    c.alignment = align
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BORDER
    return c


def title_block(ws, title, last_col, subtitle=None):
    """Tiêu đề chuẩn cho mọi trang: đơn vị, tên biểu mẫu, mã hồ sơ."""
    ws.merge_cells(f"A1:{last_col}1")
    put(ws, "A1", "TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN, ĐHQG-HCM  –  KHOA VẬT LÝ – VẬT LÝ KỸ THUẬT",
        bold=True, size=11, border=False)
    ws.merge_cells(f"A2:{last_col}2")
    put(ws, "A2", title, bold=True, size=14, border=False)
    ws.merge_cells(f"A3:{last_col}3")
    put(ws, "A3", subtitle or f"Mã hồ sơ: {DOC_CODE}   |   Bài: Phân bố thống kê trong phân rã phóng xạ   |   Nhóm 3",
        italic=True, size=11, border=False)
    ws.row_dimensions[2].height = 22


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def load_raw():
    import openpyxl
    ws = openpyxl.load_workbook(RAW, data_only=True).active
    cols = {"S1": ("G", 7), "S2": ("C", 6), "S3": ("A", 6)}
    long_cell = {"S1": "H7", "S2": "C108", "S3": "A108"}
    raw = {k: [ws[f"{c}{r0 + i}"].value for i in range(100)] for k, (c, r0) in cols.items()}
    n100 = {k: ws[v].value for k, v in long_cell.items()}
    return raw, n100


# ---------------------------------------------------------------- 1. Thông tin
def sheet_info(wb):
    ws = wb.active
    ws.title = "ThongTin"
    title_block(ws, "BIÊN BẢN GHI NHẬN SỐ LIỆU ĐO BỨC XẠ GAMMA", "F")
    widths(ws, {"A": 6, "B": 34, "C": 26, "D": 18, "E": 18, "F": 24})

    put(ws, "A5", "Chú thích:", bold=True, align=LEFT, border=False)
    put(ws, "B5", "Ô nền vàng", fill=INPUT_FILL, align=LEFT)
    ws.merge_cells("C5:F5")
    put(ws, "C5", "→ thông tin nhóm cần điền/xác nhận. Chữ xanh dương = số liệu nhập tay; "
        "chữ đen = công thức; chữ xanh lá = liên kết từ trang khác.", align=LEFT, border=False, size=11)
    ws.row_dimensions[5].height = 30

    rows = [
        ("I", "THÔNG TIN CHUNG", None, False),
        ("1", "Học phần", "Thực hành Vật lý Hiện đại (THVLHĐ)", True),
        ("2", "Tên bài thực tập", "Phân bố thống kê trong phân rã phóng xạ", False),
        ("3", "Nhóm thực hiện", "Nhóm 3", False),
        ("4", "Phòng thí nghiệm", "", True),
        ("5", "Ngày đo", "", True),
        ("6", "Giảng viên hướng dẫn", "", True),
        ("7", "Mã hồ sơ", DOC_CODE, False),
        ("II", "THIẾT BỊ ĐO", None, False),
        ("1", "Đầu dò", "Ống đếm Geiger–Müller (model/số hiệu: ……)", True),
        ("2", "Bộ đếm (Counter)", "Model/số hiệu: ……", True),
        ("3", "Nguồn cao thế (HV)", "Model/số hiệu: ……", True),
        ("4", "Điện áp làm việc đặt cho GM (V)", "", True),
        ("III", "NGUỒN PHÓNG XẠ", None, False),
        ("1", "Đồng vị", "Ra-226 (dùng cho Bộ 2 và Bộ 3)", False),
        ("2", "Hoạt độ / ngày chuẩn", "", True),
        ("3", "Bộ 1 (mức thấp)", "Phông (“Font” trong số liệu gốc), xác nhận có/không đặt nguồn", True),
        ("IV", "ĐIỀU KIỆN ĐO", None, False),
        ("1", "Thời gian mỗi lần đo ngắn t (s)", 10, False),
        ("2", "Số lần đo ngắn mỗi bộ n", 100, False),
        ("3", "Thời gian phép đo dài T (s)", 100, False),
        ("4", "Khoảng cách nguồn–đầu dò Bộ 1 (cm)", "", True),
        ("5", "Khoảng cách nguồn–đầu dò Bộ 2 (cm)", "", True),
        ("6", "Khoảng cách nguồn–đầu dò Bộ 3 (cm)", "", True),
        ("7", "Nhiệt độ / độ ẩm phòng", "", True),
        ("V", "QUY ƯỚC BÁO CÁO", None, False),
        ("1", "Độ không đảm bảo đo", "Chuẩn (k = 1, ~68%), thống kê đếm Poisson u(N) = √N", False),
        ("2", "Số liệu gốc", "Tệp “nhóm 1.xlsx” (không chỉnh sửa)", False),
        ("3", "Phân tích bổ trợ", "analysis.py (Python) – Monte Carlo, χ² Pearson, đồ thị", False),
    ]
    r = 7
    for no, label, val, need in rows:
        if val is None:
            ws.merge_cells(f"B{r}:F{r}")
            put(ws, f"A{r}", no, bold=True, fill=SEC_FILL)
            put(ws, f"B{r}", label, bold=True, fill=SEC_FILL, align=LEFT)
            for c in "CDEF":
                ws[f"{c}{r}"].border = BORDER
        else:
            ws.merge_cells(f"C{r}:F{r}")
            put(ws, f"A{r}", no)
            put(ws, f"B{r}", label, align=LEFT)
            put(ws, f"C{r}", val, color=BLUE if isinstance(val, (int, float)) else "000000",
                fill=INPUT_FILL if need else None, align=LEFT)
            for c in "DEF":
                ws[f"{c}{r}"].border = BORDER
        r += 1
    labels = [x[1] for x in rows]
    info_t = 7 + labels.index("Thời gian mỗi lần đo ngắn t (s)")
    info_T = 7 + labels.index("Thời gian phép đo dài T (s)")

    r += 1
    ws.merge_cells(f"A{r}:F{r}")
    put(ws, f"A{r}", "VI. DANH SÁCH THÀNH VIÊN NHÓM 3", bold=True, fill=SEC_FILL, align=LEFT)
    r += 1
    for col, h in zip("ABCDEF", ["STT", "Họ và tên", "MSSV", "Vai trò", "Chữ ký", "Ghi chú"]):
        put(ws, f"{col}{r}", h, bold=True, fill=HEAD_FILL)
    for i, (mssv, name) in enumerate(MEMBERS, 1):
        r += 1
        put(ws, f"A{r}", i)
        put(ws, f"B{r}", name, align=LEFT)
        put(ws, f"C{r}", mssv)
        put(ws, f"D{r}", "Thành viên", fill=INPUT_FILL)
        put(ws, f"E{r}", "")
        put(ws, f"F{r}", "")

    r += 2
    for col, h in zip(["A", "C", "E"], ["NGƯỜI ĐO", "NGƯỜI KIỂM TRA", "GIẢNG VIÊN XÁC NHẬN"]):
        end = {"A": "B", "C": "D", "E": "F"}[col]
        ws.merge_cells(f"{col}{r}:{end}{r}")
        put(ws, f"{col}{r}", h, bold=True, border=False)
        ws.merge_cells(f"{col}{r + 1}:{end}{r + 1}")
        put(ws, f"{col}{r + 1}", "(Ký, ghi rõ họ tên)", italic=True, border=False, size=11)
    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False
    return f"ThongTin!$C${info_t}", f"ThongTin!$C${info_T}"


# ---------------------------------------------------------------- 2. Số liệu thô
RAW_R0 = 9  # hàng đầu tiên của số liệu


def sheet_raw(wb, raw, n100, t_ref, T_info):
    ws = wb.create_sheet("SoLieuTho")
    title_block(ws, "BẢNG SỐ LIỆU ĐẾM THÔ – ỐNG ĐẾM GEIGER–MÜLLER", "K")
    widths(ws, {"A": 8, "B": 10, **{L(c): 11 for c in range(3, 12)}})
    ws.merge_cells("A6:K6")
    put(ws, "A6", "N: số đếm ghi nhận trong thời gian t;  u(N) = √N (k = 1);  R = N/t (s⁻¹);  u(R) = √N/t",
        italic=True, size=11, border=False, align=LEFT)
    ws.merge_cells("A7:A8")
    put(ws, "A7", "Lần đo j", bold=True, fill=HEAD_FILL)
    ws.merge_cells("B7:B8")
    put(ws, "B7", "t (s)", bold=True, fill=HEAD_FILL)
    for i, s in enumerate(SETS):
        c0 = 3 + 3 * i
        ws.merge_cells(f"{L(c0)}7:{L(c0 + 2)}7")
        put(ws, f"{L(c0)}7", SET_NAME[s], bold=True, fill=HEAD_FILL, size=11)
        for j, h in enumerate(["N", "u(N)", "R (s⁻¹)"]):
            put(ws, f"{L(c0 + j)}8", h, bold=True, fill=HEAD_FILL)
        ws[f"{L(c0)}7"].border = BORDER
    ws.row_dimensions[7].height = 32
    for j in range(100):
        r = RAW_R0 + j
        put(ws, f"A{r}", j + 1)
        put(ws, f"B{r}", f"={t_ref}", color=GREEN)
        for i, s in enumerate(SETS):
            c0 = 3 + 3 * i
            put(ws, f"{L(c0)}{r}", raw[s][j], color=BLUE)
            put(ws, f"{L(c0 + 1)}{r}", f"=SQRT({L(c0)}{r})", fmt="0.00")
            put(ws, f"{L(c0 + 2)}{r}", f"={L(c0)}{r}/$B{r}", fmt="0.000")
    last = RAW_R0 + 99
    rl = last + 2  # phép đo dài
    ws.merge_cells(f"A{rl}:K{rl}")
    put(ws, f"A{rl}", "PHÉP ĐO DÀI (một lần đo trong thời gian T)", bold=True, fill=SEC_FILL, align=LEFT)
    rl += 1
    put(ws, f"A{rl}", "T")
    put(ws, f"B{rl}", f"={T_info}", color=GREEN)
    for i, s in enumerate(SETS):
        c0 = 3 + 3 * i
        put(ws, f"{L(c0)}{rl}", n100[s], color=BLUE)
        put(ws, f"{L(c0 + 1)}{rl}", f"=SQRT({L(c0)}{rl})", fmt="0.00")
        put(ws, f"{L(c0 + 2)}{rl}", f"={L(c0)}{rl}/$B{rl}", fmt="0.000")
    ws[f"C{RAW_R0}"].comment = Comment("Nguồn: nhóm 1.xlsx, cột G7:G106 (“Font”); phép đo 100 s: ô H7.", "Nhóm 3")
    ws[f"F{RAW_R0}"].comment = Comment("Nguồn: nhóm 1.xlsx, cột C6:C105 (Ra-226); phép đo 100 s: ô C108.", "Nhóm 3")
    ws[f"I{RAW_R0}"].comment = Comment("Nguồn: nhóm 1.xlsx, cột A6:A105 (Ra-226); phép đo 100 s: ô A108.", "Nhóm 3")
    ws.freeze_panes = f"C{RAW_R0}"
    return {s: f"SoLieuTho!${L(3 + 3 * i)}${RAW_R0}:${L(3 + 3 * i)}${last}" for i, s in enumerate(SETS)}, \
        {s: f"SoLieuTho!${L(3 + 3 * i)}${rl}" for i, s in enumerate(SETS)}, f"SoLieuTho!$B${rl}"


# ---------------------------------------------------------------- 3. Tần suất
def sheet_freq(wb, rng, mean_ref):
    """Mỗi bộ một khối: k, O, u(O), E_Poisson, E_Gauss, số hạng χ² (công thức 5)."""
    ws = wb.create_sheet("TanSuat")
    title_block(ws, "BẢNG TẦN SUẤT SỐ ĐẾM VÀ PHÂN BỐ LÝ THUYẾT", "U")
    ws.merge_cells("A4:U4")
    put(ws, "A4", "O: số lần xuất hiện; u(O) = √O; E: tần suất lý thuyết = n·P(k); "
        "χ² theo công thức (5): (O − E)²/O, chỉ tính các bin O > 0", italic=True, size=11, border=False, align=LEFT)
    pos = {}
    heads = ["k", "O", "u(O)", "E Poisson", "E Gauss", "χ² Poisson", "χ² Gauss"]
    for i, s in enumerate(SETS):
        c0 = 1 + 7 * i
        C = [L(c0 + j) for j in range(7)]
        for c, w in zip(C, [6, 6, 7, 10, 10, 10, 10]):
            ws.column_dimensions[c].width = w
        ws.merge_cells(f"{C[0]}6:{C[6]}6")
        put(ws, f"{C[0]}6", SET_NAME[s], bold=True, fill=HEAD_FILL)
        put(ws, f"{C[0]}7", "x̄ =", bold=True)
        ws.merge_cells(f"{C[1]}7:{C[2]}7")
        put(ws, f"{C[1]}7", f"={mean_ref[s]}", color=GREEN, fmt="0.00")
        put(ws, f"{C[3]}7", "n =", bold=True)
        put(ws, f"{C[4]}7", f"=COUNT({rng[s]})", fmt="0")
        for j, h in enumerate(heads):
            put(ws, f"{C[j]}8", h, bold=True, fill=HEAD_FILL, size=11)
        k0, k1 = K_RANGE[s]
        r0 = 9
        for k in range(k0, k1 + 1):
            r = r0 + k - k0
            m, n = f"${C[1]}$7", f"${C[4]}$7"
            put(ws, f"{C[0]}{r}", k, color=BLUE)
            put(ws, f"{C[1]}{r}", f"=COUNTIF({rng[s]},{C[0]}{r})")
            put(ws, f"{C[2]}{r}", f"=SQRT({C[1]}{r})", fmt="0.00")
            put(ws, f"{C[3]}{r}", f"={n}*POISSON({C[0]}{r},{m},FALSE)", fmt="0.00")
            put(ws, f"{C[4]}{r}", f"={n}*NORMDIST({C[0]}{r},{m},SQRT({m}),FALSE)", fmt="0.00")
            put(ws, f"{C[5]}{r}", f"=IF({C[1]}{r}>0,({C[1]}{r}-{C[3]}{r})^2/{C[1]}{r},0)", fmt="0.00")
            put(ws, f"{C[6]}{r}", f"=IF({C[1]}{r}>0,({C[1]}{r}-{C[4]}{r})^2/{C[1]}{r},0)", fmt="0.00")
        r1 = r0 + k1 - k0
        rs = r1 + 1
        put(ws, f"{C[0]}{rs}", "Σ", bold=True, fill=RES_FILL)
        for j in (1, 3, 4, 5, 6):
            put(ws, f"{C[j]}{rs}", f"=SUM({C[j]}{r0}:{C[j]}{r1})", bold=True, fill=RES_FILL,
                fmt="0" if j == 1 else "0.00")
        put(ws, f"{C[2]}{rs}", "", fill=RES_FILL)
        put(ws, f"{C[0]}{rs + 1}", "bin", bold=True, size=10)
        put(ws, f"{C[1]}{rs + 1}", f"=COUNTIF({C[1]}{r0}:{C[1]}{r1},\">0\")", fmt="0")
        ws[f"{C[1]}{rs + 1}"].comment = Comment("Số bin có O > 0 (dùng cho χ² công thức 5).", "Nhóm 3")
        pos[s] = {"k": (C[0], r0, r1), "O": C[1], "uO": C[2], "P": C[3], "G": C[4],
                  "chiP": f"TanSuat!${C[5]}${rs}", "chiG": f"TanSuat!${C[6]}${rs}",
                  "bins": f"TanSuat!${C[1]}${rs + 1}"}
    ws.freeze_panes = "A9"
    return ws, pos


# ---------------------------------------------------------------- 4. Thống kê
def sheet_stats(wb, rng, long_ref, T_ref, t_ref):
    ws = wb.create_sheet("ThongKe")
    title_block(ws, "XỬ LÝ THỐNG KÊ SỐ LIỆU ĐẾM", "F")
    widths(ws, {"A": 6, "B": 50, "C": 16, "D": 16, "E": 16, "F": 30})
    put(ws, "A5", "STT", bold=True, fill=HEAD_FILL)
    put(ws, "B5", "Đại lượng", bold=True, fill=HEAD_FILL)
    for c, s in zip("CDE", SETS):
        put(ws, f"{c}5", SET_NAME[s], bold=True, fill=HEAD_FILL, size=11)
    put(ws, "F5", "Công thức / ghi chú", bold=True, fill=HEAD_FILL)
    ws.row_dimensions[5].height = 45
    rows = {}
    r = 6

    def sec(title):
        nonlocal r
        ws.merge_cells(f"A{r}:F{r}")
        put(ws, f"A{r}", title, bold=True, fill=SEC_FILL, align=LEFT)
        for c in "BCDEF":
            ws[f"{c}{r}"].border = BORDER
        r += 1

    no = [0]

    def row(key, label, tmpl, fmt="0.00", note="", res=False):
        nonlocal r
        no[0] += 1
        put(ws, f"A{r}", no[0])
        put(ws, f"B{r}", label, align=LEFT)
        for c, s in zip("CDE", SETS):
            expr = tmpl(s, c)
            put(ws, f"{c}{r}", expr, fmt=fmt, fill=RES_FILL if res else None,
                bold=res, color=GREEN if isinstance(expr, str) and expr.startswith("=ThongTin") else "000000")
        put(ws, f"F{r}", note, align=LEFT, size=10, italic=True)
        rows[key] = r
        r += 1

    sec("A. ĐẶC TRƯNG THỐNG KÊ CỦA 100 LẦN ĐO 10 s")
    row("n", "Số lần đo n", lambda s, c: f"=COUNT({rng[s]})", "0", "COUNT")
    row("t", "Thời gian mỗi lần đo t (s)", lambda s, c: f"={t_ref}", "0", "Từ trang ThongTin")
    row("sum", "Tổng số đếm ΣN", lambda s, c: f"=SUM({rng[s]})", "0", "SUM")
    row("mean", "Số đếm trung bình x̄", lambda s, c: f"=AVERAGE({rng[s]})", "0.00", "x̄ = ΣN/n", res=True)
    row("s", "Độ lệch chuẩn thực nghiệm s", lambda s, c: f"=STDEV({rng[s]})", "0.00", "STDEV (n − 1)", res=True)
    row("var", "Phương sai s²", lambda s, c: f"=VAR({rng[s]})", "0.00", "VAR (n − 1)")
    row("umean", "Độ không đảm bảo của trung bình u(x̄)", lambda s, c: f"={c}{rows['s']}/SQRT({c}{rows['n']})",
        "0.00", "u(x̄) = s/√n")
    row("sqrt", "Độ lệch chuẩn Poisson lý thuyết σ = √x̄", lambda s, c: f"=SQRT({c}{rows['mean']})", "0.00",
        "σ_Poisson = √µ")
    row("ratio", "Tỉ số s/√x̄", lambda s, c: f"={c}{rows['s']}/{c}{rows['sqrt']}", "0.000", "≈ 1 nếu tuân theo Poisson")
    row("D", "Chỉ số phân tán D = s²/x̄", lambda s, c: f"={c}{rows['var']}/{c}{rows['mean']}", "0.000",
        "Poisson: D = 1")
    row("dchi", "Thống kê phân tán χ² = (n − 1)·D", lambda s, c: f"=({c}{rows['n']}-1)*{c}{rows['D']}", "0.00",
        "Tuân theo χ²(n − 1)")
    row("dp", "p-value hai phía của kiểm định phân tán",
        lambda s, c: f"=2*MIN(CHIDIST({c}{rows['dchi']},{c}{rows['n']}-1),1-CHIDIST({c}{rows['dchi']},{c}{rows['n']}-1))",
        "0.000", "p > 0,05: phương sai phù hợp Poisson")
    row("dconc", "Kết luận về phương sai",
        lambda s, c: f"=IF({c}{rows['dp']}>0.05,\"Phù hợp Poisson\",\"Không phù hợp\")", "@", "", res=True)
    row("rate", "Tốc độ đếm trung bình R = x̄/t (s⁻¹)", lambda s, c: f"={c}{rows['mean']}/{c}{rows['t']}", "0.000",
        "", res=True)
    row("urate", "u(R) (s⁻¹)", lambda s, c: f"={c}{rows['umean']}/{c}{rows['t']}", "0.000", "u(R) = u(x̄)/t")

    sec("B. SO SÁNH VỚI PHÉP ĐO DÀI T = 100 s")
    row("N100", "Số đếm phép đo dài N_T", lambda s, c: f"={long_ref[s]}", "0", "Từ trang SoLieuTho")
    row("N100s", "Quy về 10 s: N_T·t/T", lambda s, c: f"={c}{rows['N100']}*{c}{rows['t']}/{T_ref}", "0.00")
    row("uN100", "u(N_T·t/T) = √N_T·t/T", lambda s, c: f"=SQRT({c}{rows['N100']})*{c}{rows['t']}/{T_ref}", "0.00")
    row("sT", "Độ lệch chuẩn dự đoán cho 10 s từ phép đo dài √(N_T·t/T)",
        lambda s, c: f"=SQRT({c}{rows['N100s']})", "0.00", "So với s ở mục A")
    row("diff", "Hiệu Δ = N_T·t/T − x̄", lambda s, c: f"={c}{rows['N100s']}-{c}{rows['mean']}", "0.00")
    row("udiff", "u(Δ) = √(u(N_T·t/T)² + u(x̄)²)", lambda s, c: f"=SQRT({c}{rows['uN100']}^2+{c}{rows['umean']}^2)",
        "0.00")
    row("z", "Chỉ số z = Δ/u(Δ) (số độ lệch chuẩn)", lambda s, c: f"={c}{rows['diff']}/{c}{rows['udiff']}", "0.00",
        "|z| ≤ 2: phù hợp", res=True)
    row("zconc", "Kết luận", lambda s, c: f"=IF(ABS({c}{rows['z']})<=2,\"Phù hợp (|z| ≤ 2)\",\"Không phù hợp\")",
        "@", "", res=True)

    sec("C. ĐỘ LỆCH LỚN NHẤT SO VỚI TRUNG BÌNH")
    row("max", "Giá trị lớn nhất", lambda s, c: f"=MAX({rng[s]})", "0")
    row("min", "Giá trị nhỏ nhất", lambda s, c: f"=MIN({rng[s]})", "0")
    row("dev", "Độ lệch lớn nhất x − x̄",
        lambda s, c: f"=IF({c}{rows['max']}-{c}{rows['mean']}>={c}{rows['mean']}-{c}{rows['min']},"
                     f"{c}{rows['max']}-{c}{rows['mean']},{c}{rows['min']}-{c}{rows['mean']})", "0.00")
    row("devs", "Tính theo σ = √x̄", lambda s, c: f"=ABS({c}{rows['dev']})/{c}{rows['sqrt']}", "0.00")
    row("p1", "Xác suất Poisson một lần đo lệch ít nhất như vậy (một phía)",
        lambda s, c: f"=IF({c}{rows['dev']}>=0,1-POISSON({c}{rows['max']}-1,{c}{rows['mean']},TRUE),"
                     f"POISSON({c}{rows['min']},{c}{rows['mean']},TRUE))", "0.0000",
        "P(X ≥ max) hoặc P(X ≤ min)")
    row("pn", "Xác suất xuất hiện ít nhất 1 lần trong n lần đo",
        lambda s, c: f"=1-(1-{c}{rows['p1']})^{c}{rows['n']}", "0.000", "1 − (1 − p)ⁿ (n = 100 cơ hội)")
    row("en", "Số lần kỳ vọng trong n lần đo", lambda s, c: f"={c}{rows['n']}*{c}{rows['p1']}", "0.00")
    row("devc", "Kết luận về hệ đếm",
        lambda s, c: f"=IF({c}{rows['pn']}>0.05,\"Hoạt động bình thường\",\"Cần kiểm tra\")", "@",
        "Ngưỡng 5%", res=True)
    return ws, rows


def stats_chi(ws, rows, pos):
    r = max(rows.values()) + 1
    ws.merge_cells(f"A{r}:F{r}")
    put(ws, f"A{r}", "D. KIỂM ĐỊNH χ² (CÔNG THỨC 5): POISSON VÀ GAUSS", bold=True, fill=SEC_FILL, align=LEFT)
    r += 1
    n0 = ws[f"A{r - 2}"].value or 0
    items = [
        ("chiP", "χ² Poisson", lambda s: f"={pos[s]['chiP']}", "0.00", "Σ(O − E)²/O, O > 0"),
        ("chiG", "χ² Gauss", lambda s: f"={pos[s]['chiG']}", "0.00", "Gauss: µ = σ² = x̄"),
        ("bins", "Số bin sử dụng (O > 0)", lambda s: f"={pos[s]['bins']}", "0", ""),
        ("ndf", "Bậc tự do ν = số bin − 1", None, "0", "1 tham số (µ) ước lượng từ dữ liệu"),
        ("rP", "χ²/ν Poisson", None, "0.00", ""),
        ("rG", "χ²/ν Gauss", None, "0.00", ""),
        ("pP", "p-value Poisson", None, "0.000", "CHIDIST(χ², ν)"),
        ("pG", "p-value Gauss", None, "0.000", ""),
        ("best", "Phân bố phù hợp hơn (χ² nhỏ hơn)", None, "@", ""),
    ]
    for key, label, fn, fmt, note in items:
        n0 += 1
        put(ws, f"A{r}", n0)
        put(ws, f"B{r}", label, align=LEFT)
        rows[key] = r
        for c, s in zip("CDE", SETS):
            if fn:
                expr, col = fn(s), GREEN
            else:
                col = "000000"
                expr = {
                    "ndf": f"={c}{rows['bins']}-1",
                    "rP": f"={c}{rows['chiP']}/{c}{rows['ndf']}",
                    "rG": f"={c}{rows['chiG']}/{c}{rows['ndf']}",
                    "pP": f"=CHIDIST({c}{rows['chiP']},{c}{rows['ndf']})",
                    "pG": f"=CHIDIST({c}{rows['chiG']},{c}{rows['ndf']})",
                    "best": f"=IF({c}{rows['chiP']}<={c}{rows['chiG']},\"Poisson\",\"Gauss\")",
                }[key]
            put(ws, f"{c}{r}", expr, fmt=fmt, color=col, fill=RES_FILL if key == "best" else None,
                bold=key == "best")
        put(ws, f"F{r}", note, align=LEFT, size=10, italic=True)
        r += 1
    ws.freeze_panes = "C6"


# ---------------------------------------------------------------- 5. Tốc độ đếm tích lũy
def sheet_running(wb):
    ws = wb.create_sheet("TocDoTichLuy")
    title_block(ws, "TỐC ĐỘ ĐẾM TRUNG BÌNH TÍCH LŨY r_c(j)", "E")
    widths(ws, {"A": 10, "B": 16, "C": 16, "D": 16, "E": 4})
    ws.merge_cells("A4:D4")
    put(ws, "A4", "r_c(j) = Σᵢ₌₁ʲ Nᵢ / Σᵢ₌₁ʲ tᵢ  (s⁻¹) – công thức (3)", italic=True, border=False, align=LEFT)
    put(ws, "A6", "j", bold=True, fill=HEAD_FILL)
    for c, s in zip("BCD", SETS):
        put(ws, f"{c}6", SET_NAME[s], bold=True, fill=HEAD_FILL, size=10)
    ws.row_dimensions[6].height = 45
    raw_cols = {"B": "C", "C": "F", "D": "I"}
    for j in range(100):
        r = 7 + j
        rr = RAW_R0 + j
        put(ws, f"A{r}", j + 1)
        for c in "BCD":
            rc = raw_cols[c]
            put(ws, f"{c}{r}", f"=SUM(SoLieuTho!${rc}${RAW_R0}:{rc}{rr})/SUM(SoLieuTho!$B${RAW_R0}:B{rr})",
                fmt="0.000", color=GREEN)
    for i, c in enumerate("BCD"):
        ch = LineChart()
        ch.title = SET_NAME[SETS[i]]
        ch.y_axis.title = "r_c (s⁻¹)"
        ch.x_axis.title = "Lần đo j"
        ch.add_data(Reference(ws, min_col=i + 2, min_row=6, max_row=106), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=7, max_row=106))
        ch.legend = None
        ch.height, ch.width = 7, 15
        ch.x_axis.tickLblSkip = 10
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        ch.y_axis.scaling.min, ch.y_axis.scaling.max = [(0.3, 1.1), (3.0, 4.6), (10.0, 13.0)][i]
        ch.y_axis.number_format = "0.0"
        ws.add_chart(ch, f"F{6 + 15 * i}")
    ws.freeze_panes = "A7"


# ---------------------------------------------------------------- 6. Đồ thị tần suất
def charts_freq(wb, ws_freq, pos):
    ws = wb.create_sheet("DoThi")
    title_block(ws, "ĐỒ THỊ PHÂN BỐ TẦN SUẤT: THỰC NGHIỆM – POISSON – GAUSS", "R")
    ws.merge_cells("A4:R4")
    put(ws, "A4", "Điểm đen: thực nghiệm (thanh sai số ±√O); đường liền: Poisson; đường đứt: Gauss.",
        italic=True, border=False, align=LEFT)
    for i, s in enumerate(SETS):
        p = pos[s]
        kc, r0, r1 = p["k"]
        col = lambda c: ws_freq[f"{c}{r0}"].column  # noqa: E731
        xref = Reference(ws_freq, min_col=col(kc), min_row=r0, max_row=r1)
        ch = ScatterChart()
        ch.title = SET_NAME[s]
        ch.style = 13
        ch.x_axis.title = "Số đếm trong 10 s"
        ch.y_axis.title = "Số lần xuất hiện"
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        so = Series(Reference(ws_freq, min_col=col(p["O"]), min_row=r0, max_row=r1), xref, title="Thực nghiệm")
        so.marker.symbol = "circle"
        so.marker.size = 5
        so.marker.graphicalProperties.solidFill = "000000"
        so.graphicalProperties.line.noFill = True
        eref = f"TanSuat!${p['uO']}${r0}:${p['uO']}${r1}"
        so.errBars = ErrorBars(errDir="y", errBarType="both", errValType="cust", noEndCap=False,
                               plus=NumDataSource(numRef=NumRef(f=eref)),
                               minus=NumDataSource(numRef=NumRef(f=eref)))
        ch.series.append(so)
        sp = Series(Reference(ws_freq, min_col=col(p["P"]), min_row=r0, max_row=r1), xref, title="Poisson")
        sp.marker.symbol = "none"
        sp.smooth = True
        sp.graphicalProperties.line.solidFill = "127A7A"
        sp.graphicalProperties.line.width = 22000
        ch.series.append(sp)
        if s != "S1":
            sg = Series(Reference(ws_freq, min_col=col(p["G"]), min_row=r0, max_row=r1), xref, title="Gauss")
            sg.marker.symbol = "none"
            sg.smooth = True
            sg.graphicalProperties.line.solidFill = "D35400"
            sg.graphicalProperties.line.dashStyle = "dash"
            sg.graphicalProperties.line.width = 19000
            ch.series.append(sg)
        k0, k1 = K_RANGE[s]
        ch.x_axis.scaling.min, ch.x_axis.scaling.max = k0, k1
        ch.y_axis.scaling.min = 0
        ch.height, ch.width = 9, 17
        ws.add_chart(ch, f"A{6 + 19 * i}")


# ---------------------------------------------------------------- 7. Monte Carlo
def sheet_mc(wb, rows):
    ws = wb.create_sheet("MonteCarlo")
    meta = R["_meta"]
    title_block(ws, f"MÔ PHỎNG MONTE CARLO PHÂN BỐ POISSON ({meta['N_MC']} LẦN THỬ)", "O")
    ws.merge_cells("A4:O4")
    put(ws, "A4", f"Sinh số ngẫu nhiên Poisson bằng numpy (seed = {meta['SEED']}) với µ = x̄ thực nghiệm; "
        "số liệu xuất từ analysis.py → montecarlo/mc_S*.csv (nhập tay, chữ xanh dương).",
        italic=True, size=11, border=False, align=LEFT)
    blocks = {}
    for i, s in enumerate(SETS):
        c0 = 1 + 5 * i
        C = [L(c0 + j) for j in range(4)]
        for c, w in zip(C, [7, 10, 12, 12]):
            ws.column_dimensions[c].width = w
        ws.column_dimensions[L(c0 + 4)].width = 3
        ws.merge_cells(f"{C[0]}6:{C[3]}6")
        put(ws, f"{C[0]}6", SET_NAME[s], bold=True, fill=HEAD_FILL, size=11)
        ws.merge_cells(f"{C[0]}7:{C[1]}7")
        put(ws, f"{C[0]}7", "µ (thực nghiệm)", size=11)
        ws.merge_cells(f"{C[2]}7:{C[3]}7")
        put(ws, f"{C[2]}7", f"=ThongKe!{'CDE'[i]}{rows['mean']}", color=GREEN, fmt="0.00")
        for j, h in enumerate(["k", "Tần suất MC", "E Poisson", "(O−E)²/O"]):
            put(ws, f"{C[j]}8", h, bold=True, fill=HEAD_FILL, size=11)
        mc = R[s]["mc"]
        r0 = 9
        for j, (k, o) in enumerate(zip(mc["k"], mc["obs"])):
            r = r0 + j
            put(ws, f"{C[0]}{r}", k, color=BLUE)
            put(ws, f"{C[1]}{r}", o, color=BLUE)
            put(ws, f"{C[2]}{r}", f"=SUM(${C[1]}${r0}:${C[1]}${r0 + len(mc['k']) - 1})*POISSON({C[0]}{r},${C[2]}$7,FALSE)",
                fmt="0.0")
            put(ws, f"{C[3]}{r}", f"=IF({C[1]}{r}>0,({C[1]}{r}-{C[2]}{r})^2/{C[1]}{r},0)", fmt="0.00")
        r1 = r0 + len(mc["k"]) - 1
        blocks[s] = (C, r0, r1)

    # Bảng so sánh thực nghiệm – mô phỏng
    r = 9 + max(len(R[s]["mc"]["k"]) for s in SETS) + 2
    ws.merge_cells(f"A{r}:O{r}")
    put(ws, f"A{r}", "SO SÁNH THỰC NGHIỆM – MÔ PHỎNG", bold=True, fill=SEC_FILL, align=LEFT)
    r += 1
    ws.merge_cells(f"A{r}:E{r}")
    put(ws, f"A{r}", "Đại lượng", bold=True, fill=HEAD_FILL)
    for j, s in enumerate(SETS):
        cc = L(6 + 3 * j)
        ws.merge_cells(f"{cc}{r}:{L(8 + 3 * j)}{r}")
        put(ws, f"{cc}{r}", SET_NAME[s], bold=True, fill=HEAD_FILL, size=10)
    ws.row_dimensions[r].height = 30
    r += 1

    def line(label, fn, fmt="0.00", note_color="000000"):
        nonlocal r
        ws.merge_cells(f"A{r}:E{r}")
        put(ws, f"A{r}", label, align=LEFT, size=11)
        for c in "BCDE":
            ws[f"{c}{r}"].border = BORDER
        for j, s in enumerate(SETS):
            cc = L(6 + 3 * j)
            ws.merge_cells(f"{cc}{r}:{L(8 + 3 * j)}{r}")
            v = fn(s, j)
            put(ws, f"{cc}{r}", v, fmt=fmt, color=note_color if not (isinstance(v, str) and v.startswith("=")) or
                note_color != BLUE else note_color)
            for c in (L(7 + 3 * j), L(8 + 3 * j)):
                ws[f"{c}{r}"].border = BORDER
        r += 1
        return r - 1

    def rng_of(s, col):
        C, r0, r1 = blocks[s]
        return f"${C[col]}${r0}:${C[col]}${r1}"

    rN = line("Số lần thử MC", lambda s, j: f"=SUM({rng_of(s, 1)})", "0")
    rM = line("Trung bình MC x̄_MC", lambda s, j: f"=SUMPRODUCT({rng_of(s, 0)},{rng_of(s, 1)})/{L(6 + 3 * j)}{rN}")
    rS = line("Độ lệch chuẩn MC s_MC",
              lambda s, j: f"=SQRT(SUMPRODUCT({rng_of(s, 1)},({rng_of(s, 0)}-{L(6 + 3 * j)}{rM})^2)/({L(6 + 3 * j)}{rN}-1))")
    line("u(x̄_MC) = s_MC/√N", lambda s, j: f"={L(6 + 3 * j)}{rS}/SQRT({L(6 + 3 * j)}{rN})")
    rE = line("Trung bình thực nghiệm x̄", lambda s, j: f"=ThongKe!{'CDE'[j]}{rows['mean']}")
    rES = line("Độ lệch chuẩn thực nghiệm s", lambda s, j: f"=ThongKe!{'CDE'[j]}{rows['s']}")
    line("√µ (lý thuyết)", lambda s, j: f"=SQRT({L(6 + 3 * j)}{rE})")
    line("s thực nghiệm / s_MC", lambda s, j: f"={L(6 + 3 * j)}{rES}/{L(6 + 3 * j)}{rS}", "0.000")
    line("χ² (công thức 5) của mẫu MC so với Poisson", lambda s, j: f"=SUM({rng_of(s, 3)})")
    line(f"Lặp {meta['N_REP']} thí nghiệm ảo × 100 lần đo: trung bình của s",
         lambda s, j: R[s]["mc_rep"]["s_mean"], "0.00", BLUE)
    line("… độ lệch chuẩn của s (phân tán thống kê của s)", lambda s, j: R[s]["mc_rep"]["s_std"], "0.00", BLUE)
    line("… tỉ lệ thí nghiệm ảo có s nhỏ hơn s thực nghiệm (%)", lambda s, j: R[s]["mc_rep"]["s_pct"], "0.0", BLUE)
    ws.freeze_panes = "A9"


# ---------------------------------------------------------------- 8. Kết quả
def sheet_result(wb, rows):
    ws = wb.create_sheet("KetQua")
    title_block(ws, "TỔNG HỢP KẾT QUẢ ĐO", "H")
    widths(ws, {"A": 6, "B": 44, "C": 13, "D": 10, "E": 13, "F": 10, "G": 13, "H": 10})
    put(ws, "A5", "STT", bold=True, fill=HEAD_FILL)
    put(ws, "B5", "Đại lượng (k = 1)", bold=True, fill=HEAD_FILL)
    for j, s in enumerate(SETS):
        c = L(3 + 2 * j)
        ws.merge_cells(f"{c}5:{L(4 + 2 * j)}5")
        put(ws, f"{c}5", SET_NAME[s], bold=True, fill=HEAD_FILL, size=10)
        put(ws, f"{c}6", "Giá trị", bold=True, fill=HEAD_FILL, size=10)
        put(ws, f"{L(4 + 2 * j)}6", "u", bold=True, fill=HEAD_FILL, size=10)
    ws.merge_cells("A5:A6")
    ws.merge_cells("B5:B6")
    ws.row_dimensions[5].height = 30
    lines = [
        ("Số đếm trung bình trong 10 s, x̄", "mean", "umean", "0.00"),
        ("Tốc độ đếm R (s⁻¹)", "rate", "urate", "0.000"),
        ("Phép đo dài quy về 10 s", "N100s", "uN100", "0.00"),
        ("Độ lệch chuẩn thực nghiệm s", "s", None, "0.00"),
        ("Độ lệch chuẩn Poisson √x̄", "sqrt", None, "0.00"),
        ("Chỉ số phân tán D = s²/x̄", "D", None, "0.000"),
        ("p-value kiểm định phân tán", "dp", None, "0.000"),
        ("Chỉ số z (10 s so với 100 s)", "z", None, "0.00"),
        ("χ² Poisson (công thức 5) / ν", "chiP", "ndf", "0.00"),
        ("χ² Gauss (công thức 5) / ν", "chiG", "ndf", "0.00"),
        ("Phân bố phù hợp hơn", "best", None, "@"),
        ("Xác suất ≥ 1 lần lệch cực đại trong 100 lần", "pn", None, "0.000"),
    ]
    r = 7
    for i, (label, k, ku, fmt) in enumerate(lines, 1):
        put(ws, f"A{r}", i)
        put(ws, f"B{r}", label, align=LEFT)
        for j in range(3):
            sc = "CDE"[j]
            put(ws, f"{L(3 + 2 * j)}{r}", f"=ThongKe!{sc}{rows[k]}", fmt=fmt, color=GREEN)
            put(ws, f"{L(4 + 2 * j)}{r}", f"=ThongKe!{sc}{rows[ku]}" if ku else "", fmt="0.00" if ku != "ndf" else "0",
                color=GREEN)
        r += 1
    ws[f"D{r - 4}"].comment = Comment("Cột u ở hai dòng χ²: ghi bậc tự do ν.", "Nhóm 3")

    r += 1
    ws.merge_cells(f"A{r}:H{r}")
    put(ws, f"A{r}", "χ² PEARSON (GỘP BIN SAO CHO E ≥ 5) – tính bằng analysis.py, dùng để đối chiếu",
        bold=True, fill=SEC_FILL, align=LEFT)
    r += 1
    for j, h in enumerate(["", "Phân bố", "χ²", "ν", "p", "", "", ""]):
        if h:
            put(ws, f"{L(1 + j)}{r}", h, bold=True, fill=HEAD_FILL)
    r += 1
    for s in SETS:
        for name, key in [("Poisson", "pearson_poi"), ("Gauss", "pearson_gau")]:
            p = R[s][key]
            put(ws, f"B{r}", f"{SET_NAME[s]} – {name}", align=LEFT, size=11)
            put(ws, f"C{r}", p["chi2"], color=BLUE, fmt="0.00")
            put(ws, f"D{r}", p["ndf"], color=BLUE, fmt="0")
            put(ws, f"E{r}", f"=CHIDIST(C{r},D{r})", fmt="0.000")
            r += 1

    r += 1
    ws.merge_cells(f"A{r}:H{r}")
    put(ws, f"A{r}", "NHẬN XÉT", bold=True, fill=SEC_FILL, align=LEFT)
    notes = [
        "1. r_c(j) của cả ba bộ hội tụ về giá trị trung bình và nằm trong dải ±1σ Poisson sau ~20–30 lần đo: "
        "quá trình đếm dừng trong suốt thời gian đo.",
        "2. s ≈ √x̄ (D = 1,10; 1,20; 1,15; p-value > 0,05): phương sai phù hợp phân bố Poisson.",
        "3. Kết quả phép đo 100 s phù hợp với trung bình 100 lần đo 10 s trong khoảng 1,5σ (|z| ≤ 1,43).",
        "4. Bộ mức thấp: độ lệch lớn nhất (x = 12, ≈ 2,8σ) có xác suất xuất hiện ≈ 68% trong 100 lần đo: "
        "hệ đếm hoạt động bình thường.",
        "5. Ở µ ≈ 41 Poisson cho χ² nhỏ hơn Gauss một chút; ở µ ≈ 112 hai phân bố gần như trùng nhau: "
        "xấp xỉ Gauss chỉ tốt khi µ lớn.",
        "6. χ² theo công thức (5) bị phóng đại khi O nhỏ (σ = √O); χ² Pearson gộp bin cho p > 0,1 ở cả ba bộ.",
    ]
    for t in notes:
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        put(ws, f"A{r}", t, align=LEFT, size=11)
        ws.row_dimensions[r].height = 32
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------- 9. Phân công
def sheet_tasks(wb):
    ws = wb.create_sheet("PhanCong")
    title_block(ws, "BẢNG PHÂN CÔNG CÔNG VIỆC – NHÓM 3", "F")
    widths(ws, {"A": 6, "B": 28, "C": 12, "D": 50, "E": 14, "F": 14})
    put(ws, "A4", "Nội dung phân công là đề xuất (nền vàng), nhóm cần điều chỉnh theo thực tế.",
        italic=True, border=False, align=LEFT)
    for c, h in zip("ABCDEF", ["STT", "Họ và tên", "MSSV", "Nhiệm vụ", "Mức độ hoàn thành", "Ghi chú"]):
        put(ws, f"{c}6", h, bold=True, fill=HEAD_FILL)
    tasks = [
        "Lắp đặt hệ đo, đo Bộ 1 (mức thấp)",
        "Tổng hợp số liệu, xử lý thống kê, soạn báo cáo",
        "Đo Bộ 2, ghi điều kiện đo",
        "Đo Bộ 3, phép đo dài 100 s",
        "Vẽ đồ thị phân bố tần suất, thanh sai số",
        "Tính χ², so sánh Poisson – Gauss",
        "Mô phỏng Monte Carlo",
        "Cơ sở lý thuyết, câu hỏi tìm hiểu thêm",
        "Kiểm tra, biên tập báo cáo",
    ]
    for i, ((mssv, name), t) in enumerate(zip(MEMBERS, tasks), 1):
        r = 6 + i
        put(ws, f"A{r}", i)
        put(ws, f"B{r}", name, align=LEFT)
        put(ws, f"C{r}", mssv)
        put(ws, f"D{r}", t, fill=INPUT_FILL, align=LEFT)
        put(ws, f"E{r}", "100%", fill=INPUT_FILL)
        put(ws, f"F{r}", "")
    ws.sheet_view.showGridLines = False


def main():
    raw, n100 = load_raw()
    wb = Workbook()
    t_ref, T_info = sheet_info(wb)
    rng, long_ref, T_ref = sheet_raw(wb, raw, n100, t_ref, T_info)
    ws_stats, rows = sheet_stats(wb, rng, long_ref, T_ref, t_ref)
    mean_ref = {s: f"ThongKe!${c}${rows['mean']}" for c, s in zip("CDE", SETS)}
    ws_freq, pos = sheet_freq(wb, rng, mean_ref)
    stats_chi(ws_stats, rows, pos)
    sheet_running(wb)
    charts_freq(wb, ws_freq, pos)
    sheet_mc(wb, rows)
    sheet_result(wb, rows)
    sheet_tasks(wb)
    # Thứ tự trang: ThongTin, SoLieuTho, ThongKe, TanSuat, TocDoTichLuy, DoThi, MonteCarlo, KetQua, PhanCong
    for ws in wb.worksheets:
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.oddFooter.center.text = f"{DOC_CODE} – &A – Trang &P/&N"
    wb.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()

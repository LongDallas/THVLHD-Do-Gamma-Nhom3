"""Hàm dùng chung cho các notebook hướng dẫn: tìm thư mục bài, đọc số liệu gốc.

Dùng trong notebook:
    from du_lieu import doc_so_lieu
    data, N100 = doc_so_lieu()
"""
from pathlib import Path

import pandas as pd

T_NGAN = 10    # thời gian mỗi lần đo ngắn (s)
T_DAI = 100    # thời gian phép đo dài (s)

TEN_BO = {
    "S1": "Bộ 1 – Mức thấp (Phông)",
    "S2": "Bộ 2 – Mức trung bình (Ra-226)",
    "S3": "Bộ 3 – Mức cao (Ra-226)",
}

# Vị trí trong "nhóm 1.xlsx" (pandas đọc header=None nên chỉ số = số hàng Excel − 1):
#   (cột, hàng đầu, hàng cuối + 1) cho 100 lần đo 10 s, và (hàng, cột) của phép đo 100 s
VI_TRI = {
    "S1": ((6, 6, 106), (6, 7)),      # G7:G106, H7
    "S2": ((2, 5, 105), (107, 2)),    # C6:C105, C108
    "S3": ((0, 5, 105), (107, 0)),    # A6:A105, A108
}


def thu_muc_bai() -> Path:
    """Đi ngược lên từ thư mục hiện tại cho tới khi gặp thư mục chứa 01_SoLieuGoc/."""
    for p in [Path.cwd(), *Path.cwd().parents]:
        if (p / "01_SoLieuGoc").is_dir():
            return p
    raise FileNotFoundError("Không tìm thấy thư mục 01_SoLieuGoc – hãy mở notebook trong thư mục bài 'Đo Gamma'.")


def doc_so_lieu():
    """Trả về (data, N100): data là DataFrame 100 hàng × 3 cột S1..S3 (int), N100 là dict số đếm 100 s."""
    raw = pd.read_excel(thu_muc_bai() / "01_SoLieuGoc" / "nhóm 1.xlsx", header=None)
    data, N100 = {}, {}
    for bo, ((cot, r0, r1), (hang_dai, cot_dai)) in VI_TRI.items():
        data[bo] = raw.iloc[r0:r1, cot].astype(int).to_numpy()
        N100[bo] = int(raw.iloc[hang_dai, cot_dai])
    df = pd.DataFrame(data, index=pd.RangeIndex(1, 101, name="lan_do"))
    return df, N100

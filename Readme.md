# Bài thực tập: Phân bố thống kê trong phân rã phóng xạ – Nhóm 3

Tóm tắt những gì cần làm theo [Poisson.pdf](00_DeBai/Poisson.pdf) và tiến độ hiện tại.

**Mục tiêu:** khảo sát thống kê số đếm của ống Geiger–Müller (GM) khi chiếu tia gamma, so sánh với phân bố **Poisson** và **Gauss**, đối chiếu với mô phỏng **Monte Carlo**.

---

## ⚡ Tình trạng

| Hạng mục | Trạng thái |
|---|---|
| Đo đạc (3 bộ × 100 lần 10 s + 1 lần 100 s) | ✅ Xong ([nhóm 1.xlsx](01_SoLieuGoc/nhóm%201.xlsx)) |
| Phân tích số liệu (mục 2.3) | ✅ Xong ([analysis.py](02_XuLy/analysis.py)) |
| Mô phỏng Monte Carlo (mục 3) | ✅ Xong ([montecarlo/](03_KetQuaTrungGian/montecarlo/)) |
| Sổ số liệu Excel | ✅ [BaoCao_SoLieu_Nhom3.xlsx](04_NopBai/BaoCao_SoLieu_Nhom3.xlsx) |
| Báo cáo Word | ✅ [BaoCao_Nhom3_PhanBoThongKe.docx](04_NopBai/BaoCao_Nhom3_PhanBoThongKe.docx) |
| Notebook hướng dẫn lập trình | ✅ [05_HuongDan/](05_HuongDan/README.md) |
| **Điền các ô nền vàng** (thông tin đo, GVHD, phân công) | ⏳ **Nhóm cần làm** (xem mục 7) |

---

## 0. Số liệu đã có ([nhóm 1.xlsx](01_SoLieuGoc/nhóm%201.xlsx))

| Bộ số liệu | Cột Excel gốc | x̄ ± u(x̄) (10 s) | s | √x̄ | Phép đo 100 s |
|---|---|---|---|---|---|
| Bộ 1 – Mức thấp ("Font" = phông) | G7:G106 | 5,52 ± 0,25 | 2,46 | 2,35 | 45 (ô H7) |
| Bộ 2 – Mức trung bình (Ra-226) | C6:C105 | 41,22 ± 0,70 | 7,04 | 6,42 | 437 (ô C108) |
| Bộ 3 – Mức cao (Ra-226) | A6:A105 | 112,19 ± 1,13 | 11,35 | 10,59 | 1085 (ô A108) |

> ⚠️ Danh sách thành viên trong ảnh là **Nhóm 3**, nhưng tệp số liệu tên `nhóm 1.xlsx`. Tôi dùng "Nhóm 3" trong mọi tài liệu ([ảnh danh sách](01_SoLieuGoc/DanhSach_ThanhVien_Nhom3.png)); nếu sai, đổi lại trong `02_XuLy/build_excel.py` / `02_XuLy/build_report.js`.

---

## 1. Lý thuyết cần nắm

- Poisson: `P(x; µ) = µ^x · e^(−µ) / x!`, với `σ = √µ`.
- Gauss (µ lớn): `p(x; µ) = 1/√(2πµ) · exp(−(x−µ)²/(2µ))`.
- Đi từ nhị thức → Poisson → Gauss và giới hạn áp dụng của từng phân bố (đã trình bày ở Chương 1 và phần Câu hỏi của báo cáo).

## 2. Thí nghiệm

- [x] Lắp hệ đo (nguồn, GM, bộ đếm, HV +500/+900 V, **tăng áp từ từ**), đặt thời gian đo 10 s.
- [x] Đo ba bộ (5–10, 40–50, 100–120 số đếm/10 s): 100 lần × 10 s + 1 lần × 100 s.
- [ ] Ghi **điều kiện đo** (điện áp, khoảng cách, model thiết bị, ngày đo) vào Excel (trang `ThongTin`) và Word (Bảng 2.1).

## 3. Phân tích dữ liệu (mục 2.3)

- [x] **r_c(j)** hội tụ về 0,552 / 4,122 / 11,219 s⁻¹, nên quá trình đếm là dừng (Hình 3.1).
- [x] **x̄, s so với phép đo 100 s:** phù hợp, |z| = 1,43 / 1,12 / 1,06 (< 2σ).
- [x] **Histogram và Poisson** kèm thanh sai số (Hình 3.2–3.4). D = s²/x̄ = 1,10 / 1,20 / 1,15, p > 0,05, nên phương sai phù hợp Poisson.
- [x] **Độ lệch lớn nhất (Bộ 1):** x = 12 (≈ 2,8σ). P(1 lần) = 1,1%, P(≥1 lần trong 100 lần) ≈ 68%, nên hệ đếm hoạt động bình thường.
- [x] **Poisson và Gauss bằng χ² (công thức 5):**
  - µ ≈ 41: Poisson tốt hơn (43,1 so với 44,6).
  - µ ≈ 112: hai phân bố tương đương (34,1 so với 33,4).
  - Bổ sung **χ² Pearson gộp bin** (p > 0,1 ở cả 3 bộ). Lý do: công thức (5) bị phóng đại bởi các bin có O nhỏ (bin k = 8 ở Bộ 1 và k = 44 ở Bộ 2 chiếm khoảng 75% tổng χ²).

## 4. Mô phỏng Monte Carlo (mục 3)

- [x] 1000 lần thử Poisson cho mỗi µ (≈ 5,5 / 41 / 112). File `03_KetQuaTrungGian/montecarlo/mc_S*.csv` có 2 cột (số đếm, tần suất), cùng định dạng với `poissonsim`.
- [x] So sánh với thực nghiệm (Hình 4.1, Bảng 4.1). Lặp thêm 2000 "thí nghiệm ảo" × 100 lần đo: s thực nghiệm nằm ở phân vị 75% / 91% / 85%, tức vẫn trong phạm vi thăng giáng tự nhiên (Hình 4.2).

## 5. Báo cáo (mục 4)

- [x] Bảng tần suất cho từng bộ (Bảng 3.1–3.3, trang `TanSuat` trong Excel).
- [x] Phân tích mục 2.3 và đồ thị so với Poisson; nhận xét về phương sai.
- [x] Gauss so với Poisson cho bộ µ lớn nhất, so sánh χ².
- [x] So sánh thực nghiệm với mô phỏng.
- [ ] Bảng phân công: **hiện là bản đề xuất**, nhóm cần sửa lại theo thực tế.

## 6. Câu hỏi tìm hiểu thêm

- [x] Nguồn gốc Poisson từ phân bố nhị thức và giới hạn áp dụng.
- [x] Xấp xỉ Gauss cho Poisson khi µ lớn (khai triển Stirling, sai số ~1/√µ).

---

## 7. Việc nhóm cần làm trước khi nộp

Các chỗ **tô nền vàng** trong cả Excel và Word là thông tin tôi không có, cần nhóm điền:

1. **Giảng viên hướng dẫn** (bìa Word, trang `ThongTin`).
2. **Model/số hiệu** của ống GM, bộ đếm, nguồn HV; **điện áp làm việc**.
3. **Khoảng cách nguồn–đầu dò** của từng bộ; **ngày đo**, phòng thí nghiệm; hoạt độ nguồn.
4. **Bộ 1:** xác nhận đây là phông (không đặt nguồn) hay nguồn đặt xa. Số liệu gốc chỉ ghi "Font".
5. **Bảng phân công** (Phụ lục B của Word, trang `PhanCong` của Excel): đang là bản đề xuất.
6. **Vai trò** từng thành viên (trang `ThongTin`).
7. Mở file Word và chọn **cập nhật mục lục** (References → Update Table, hoặc nhấp chuột phải vào mục lục → Update Field). Có thể chèn logo trường vào bìa nếu yêu cầu.

---

## 8. Cấu trúc thư mục

Các thư mục được đánh số theo thứ tự công việc: **đề bài, số liệu gốc, xử lý, kết quả trung gian, nộp bài, hướng dẫn**.

```
Đo Gamma/
├── Readme.md                          # file này: kế hoạch, tiến độ, hướng dẫn
│
├── 00_DeBai/                          # tài liệu đề bài (chỉ đọc)
│   └── Poisson.pdf
│
├── 01_SoLieuGoc/                      # số liệu gốc, KHÔNG chỉnh sửa
│   ├── nhóm 1.xlsx                    #   số đếm thô 3 bộ × 100 lần + phép đo 100 s
│   └── DanhSach_ThanhVien_Nhom3.png   #   ảnh danh sách thành viên
│
├── 02_XuLy/                           # mã nguồn xử lý
│   ├── run_all.sh                     #   ★ chạy toàn bộ quy trình
│   ├── analysis.py                    #   bước 1: thống kê, χ², Monte Carlo, đồ thị
│   ├── build_excel.py                 #   bước 2: tạo sổ số liệu Excel
│   ├── build_report.js                #   bước 3: tạo báo cáo Word
│   └── package.json                   #   khai báo thư viện docx (npm)
│
├── 03_KetQuaTrungGian/                # sinh tự động, có thể xóa rồi tạo lại
│   ├── results.json                   #   kết quả số
│   ├── figures/                       #   đồ thị PNG (fig1–fig6) + hình hệ đo (fig_setup_*)
│   └── montecarlo/                    #   số liệu mô phỏng mc_S*.csv
│
├── 04_NopBai/                         # ★ file nộp cho giảng viên
│   ├── BaoCao_Nhom3_PhanBoThongKe.docx
│   └── BaoCao_SoLieu_Nhom3.xlsx
│
└── 05_HuongDan/                       # notebook hướng dẫn kỹ thuật lập trình (xem README bên trong)
    ├── 01_DocSoLieu_ThongKe.ipynb     #   đọc Excel, thống kê, ddof, D, z
    ├── 02_VeBieuDo.ipynb              #   r_c(j), histogram + errorbar, Poisson/Gauss
    ├── 03_KiemDinh_Chi2.ipynb         #   χ² công thức (5), Pearson gộp bin, độ lệch lớn nhất
    ├── 04_MoPhong_MonteCarlo.ipynb    #   poissonsim, hội tụ, chuỗi xung GM, thời gian chết
    └── du_lieu.py                     #   hàm đọc số liệu dùng chung
```

Luồng dữ liệu: `01_SoLieuGoc` → `02_XuLy/analysis.py` → `03_KetQuaTrungGian` → `build_excel.py` / `build_report.js` → `04_NopBai`.

> Hai hình `fig_setup_*.png` được cắt từ `Poisson.pdf` (không do script tạo). **Không xóa** hai file này khi dọn `03_KetQuaTrungGian/`.

### Sổ số liệu Excel: dạng biên bản đo của phòng thí nghiệm hạt nhân

Mỗi trang có tiêu đề đơn vị, mã hồ sơ `THVLHĐ-GM-N3-01`, chân trang đánh số. Độ không đảm bảo báo cáo ở mức k = 1, theo thống kê đếm u(N) = √N. Mọi đại lượng đều là **công thức Excel**, nên sửa số liệu thô thì toàn bộ kết quả tự cập nhật.

| Trang | Nội dung |
|---|---|
| `ThongTin` | Biên bản: học phần, thiết bị, nguồn, điều kiện đo, quy ước, danh sách nhóm, ô ký |
| `SoLieuTho` | N, u(N) = √N, R = N/t cho 100 lần đo × 3 bộ, kèm phép đo 100 s |
| `ThongKe` | x̄, s, u(x̄), √x̄, D, kiểm định phân tán, so sánh 100 s (z), độ lệch lớn nhất, χ² |
| `TanSuat` | Bảng tần suất, E Poisson, E Gauss, các số hạng χ² |
| `TocDoTichLuy` | r_c(j) kèm đồ thị |
| `DoThi` | Đồ thị tần suất: thực nghiệm (thanh sai số), Poisson, Gauss |
| `MonteCarlo` | Tần suất MC, so sánh thực nghiệm–mô phỏng |
| `KetQua` | Tổng hợp kết quả, χ² Pearson, nhận xét |
| `PhanCong` | Bảng phân công |

Quy ước màu: **chữ xanh dương** là số nhập tay, **chữ đen** là công thức, **chữ xanh lá** là liên kết từ trang khác, **nền vàng** là ô cần điền.

### Báo cáo Word: trình bày theo quy chuẩn văn bản Việt Nam

Định dạng: A4, Times New Roman 13, giãn dòng 1,5, lề trên/dưới 2 cm, trái 3 cm, phải 2 cm. Bảng có chú thích ở trên, hình có chú thích ở dưới, đánh số theo chương.

Cấu trúc:

- Bìa (có khung)
- Mục lục, danh mục bảng, danh mục hình (số trang La Mã)
- Mở đầu
- Ch.1 Cơ sở lý thuyết
- Ch.2 Thiết bị và phương pháp
- Ch.3 Kết quả và thảo luận
- Ch.4 Mô phỏng Monte Carlo
- Kết luận
- Câu hỏi tìm hiểu thêm
- Tài liệu tham khảo
- Phụ lục A (số liệu thô), Phụ lục B (phân công)

---

## 9. Chạy lại khi sửa số liệu

```bash
pip3 install openpyxl numpy scipy matplotlib   # một lần
./02_XuLy/run_all.sh                           # tự cài docx (npm) lần đầu, rồi chạy 3 bước
```

Hoặc chạy từng bước:

```bash
python3 02_XuLy/analysis.py      # → 03_KetQuaTrungGian/
python3 02_XuLy/build_excel.py   # → 04_NopBai/BaoCao_SoLieu_Nhom3.xlsx
node 02_XuLy/build_report.js     # → 04_NopBai/BaoCao_Nhom3_PhanBoThongKe.docx
```

> Chạy lại sẽ **ghi đè** các file trong `04_NopBai/`. Lưu ý:
> - **Đóng Word/Excel** trước khi chạy.
> - Nếu đã điền tay các ô vàng, hãy sửa trực tiếp trong file thay vì chạy lại, hoặc sửa thông tin trong script trước khi chạy.

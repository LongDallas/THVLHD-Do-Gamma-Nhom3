# 05_HuongDan: Notebook hướng dẫn lập trình

Bộ 4 notebook hướng dẫn **từng kỹ thuật lập trình** dùng trong bài *Phân bố thống kê trong phân rã phóng xạ*. Tất cả chạy trên số liệu thật của nhóm (`01_SoLieuGoc/nhóm 1.xlsx`). Các notebook **đã được chạy sẵn**, nên mở ra là thấy kết quả và đồ thị.

| # | Notebook | Ứng với đề bài | Kỹ thuật trọng tâm |
|---|---|---|---|
| 01 | [01_DocSoLieu_ThongKe.ipynb](01_DocSoLieu_ThongKe.ipynb) | Mục 2.3, bước 2 | `pd.read_excel` + `iloc`; kiểm tra dữ liệu bằng `assert`; bẫy `ddof`; chỉ số phân tán D; chỉ số z; đóng gói hàm vào module |
| 02 | [02_VeBieuDo.ipynb](02_VeBieuDo.ipynb) | Mục 2.3, bước 1, 3, 5 | `np.cumsum` (vector hóa); `fill_between`; `np.bincount`; `plt.errorbar`; `subplots`; hàm vẽ nhận `ax`; `savefig` |
| 03 | [03_KiemDinh_Chi2.ipynb](03_KiemDinh_Chi2.ipynb) | Mục 2.3, bước 4, 5 | χ² công thức (5); phân tích đóng góp từng bin; χ² Pearson gộp bin bằng `cdf`; `sf` và `cdf`; hiệu ứng "100 cơ hội" |
| 04 | [04_MoPhong_MonteCarlo.ipynb](04_MoPhong_MonteCarlo.ipynb) | Mục 3 | `default_rng(seed)`; viết lại `poissonsim`; luật số lớn; mảng 2 chiều `axis=1`; chuỗi xung GM theo phân bố mũ; thời gian chết; nhị thức → Poisson |

Nên học theo thứ tự 01 → 04. Cuối mỗi notebook có phần **📝 Bài tập** để tự luyện.

## Cách mở

- **VS Code:** cài extension *Jupyter*, mở file `.ipynb`, chọn kernel Python 3.
- **Jupyter Lab:** chạy `jupyter lab` trong thư mục `Đo Gamma/`, rồi mở `05_HuongDan/`.

Thư viện cần có:

```bash
pip3 install numpy pandas scipy matplotlib openpyxl
```

## Tệp trong thư mục

| Tệp | Vai trò |
|---|---|
| `du_lieu.py` | Module dùng chung: `doc_so_lieu()` trả về DataFrame 100 × 3 và số đếm 100 s; `thu_muc_bai()` tự tìm thư mục bài |
| `hinh_tu_ve/` | Hình lưu ra khi chạy notebook 02 (tự sinh, không đưa lên Git) |
| `mo_phong/` | File CSV do `poissonsim` của notebook 04 sinh ra (tự sinh, không đưa lên Git) |

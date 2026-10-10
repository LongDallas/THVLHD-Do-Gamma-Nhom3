// Tạo báo cáo Word `BaoCao_Nhom3_PhanBoThongKe.docx` từ results.json và figures/.
// Chạy từ thư mục gốc: ./02_XuLy/run_all.sh   (cần gói npm `docx`: cd 02_XuLy && npm install)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, PageBreak, Footer, PageNumber,
  NumberFormat, TableOfContents, TabStopType, LevelFormat, VerticalAlign,
  Math: OMath, MathRun, MathFraction, MathSuperScript, MathSubScript, MathRadical, MathSum, MathLimitLower, MathRoundBrackets,
  PageBorderDisplay, PageBorderOffsetFrom, SectionType,
} = require("docx");

const ROOT = path.join(__dirname, "..");
const INTER = path.join(ROOT, "03_KetQuaTrungGian");
const R = JSON.parse(fs.readFileSync(path.join(INTER, "results.json"), "utf8"));
const OUT = path.join(ROOT, "04_NopBai", "BaoCao_Nhom3_PhanBoThongKe.docx");
const SETS = ["S1", "S2", "S3"];
const NAME = { S1: "Bộ 1 (mức thấp)", S2: "Bộ 2 (mức trung bình)", S3: "Bộ 3 (mức cao)" };
const MEMBERS = [
  ["Phạm Việt Linh", "24130044"], ["Lê Quốc Long", "24130048"], ["Ngô Đặng Tuấn Long", "24130049"],
  ["Dương Quốc Huy", "24130103"], ["Trần Trí Dũng", "24130162"], ["Nguyễn Trần Nhiên Đăng", "24130147"],
  ["Lê Anh Đức", "24130157"], ["Trần Thanh Đức", "24130159"], ["Nguyễn Vũ Nhật Minh", "24130218"],
];
const TASKS = [
  "Lắp đặt hệ đo, đo Bộ 1 (mức thấp)", "Tổng hợp số liệu, xử lý thống kê, soạn báo cáo",
  "Đo Bộ 2, ghi điều kiện đo", "Đo Bộ 3, phép đo dài 100 s", "Vẽ đồ thị phân bố tần suất",
  "Tính χ², so sánh Poisson – Gauss", "Mô phỏng Monte Carlo",
  "Cơ sở lý thuyết, câu hỏi tìm hiểu thêm", "Kiểm tra, biên tập báo cáo",
];

const FONT = "Times New Roman";
const SZ = 26;          // 13 pt
const TW = 9070;        // bề rộng vùng chữ A4, lề trái 3 cm, phải 2 cm (DXA)
const cm = (x) => Math.round(x * 567);

// ---------- định dạng số kiểu Việt Nam (dấu phẩy thập phân)
const vn = (x, d = 2) => Number(x).toFixed(d).replace(".", ",");
const pct = (x, d = 1) => vn(x * 100, d) + "%";

// ---------- khối dựng cơ bản
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || SZ, ...o });
const fill = (text) => t(text, { highlight: "yellow" });

function para(children, o = {}) {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({
    children, alignment: o.align || AlignmentType.JUSTIFIED,
    indent: o.noIndent ? undefined : { firstLine: cm(1) },
    spacing: { line: 360, before: o.before || 0, after: o.after ?? 120 },
    keepNext: o.keepNext, ...o.extra,
  });
}
const center = (children, o = {}) => para(children, { ...o, align: AlignmentType.CENTER, noIndent: true });

function bullet(children, level = 0) {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({ children, numbering: { reference: "bullets", level }, alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: 60 } });
}
function numbered(children, ref = "numbers") {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({ children, numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: 60 } });
}

const H1 = (text, first = false) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text, { bold: true, size: 28 })],
  alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240, line: 360 }, pageBreakBefore: !first });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(text, { bold: true })],
  spacing: { before: 200, after: 120, line: 360 }, keepNext: true });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [t(text, { bold: true, italics: true })],
  spacing: { before: 120, after: 80, line: 360 }, keepNext: true });

// Chú thích: bảng ở trên, hình ở dưới
const LISTS = { tables: [], figures: [] };
function tableCaption(no, text) {
  LISTS.tables.push([no, text]);
  return new Paragraph({ children: [t(`Bảng ${no}. `, { bold: true }), t(text, { italics: true })],
    alignment: AlignmentType.CENTER, spacing: { before: 120, after: 80, line: 300 }, keepNext: true });
}
function figure(file, widthPx, no, text, source) {
  const buf = fs.readFileSync(path.join(INTER, "figures", file));
  const { w, h } = pngSize(buf);
  LISTS.figures.push([no, text]);
  const out = [
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: widthPx, height: Math.round(widthPx * h / w) },
        altText: { title: text, description: text, name: file } })] }),
    new Paragraph({ children: [t(`Hình ${no}. `, { bold: true }), t(text, { italics: true })],
      alignment: AlignmentType.CENTER, spacing: { after: source ? 0 : 160, line: 300 } }),
  ];
  if (source) out.push(new Paragraph({ children: [t(source, { italics: true, size: 22 })],
    alignment: AlignmentType.CENTER, spacing: { after: 160 } }));
  return out;
}
function pngSize(buf) { return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; }

// ---------- bảng
const BD = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const BORDERS = { top: BD, bottom: BD, left: BD, right: BD };
function cell(content, width, o = {}) {
  const runs = Array.isArray(content) ? content : [t(String(content), { size: o.size || 24, bold: o.bold })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA }, borders: BORDERS, verticalAlign: VerticalAlign.CENTER,
    shading: o.head ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" } : (o.shade ? { fill: o.shade, type: ShadingType.CLEAR, color: "auto" } : undefined),
    margins: { top: 40, bottom: 40, left: 80, right: 80 }, columnSpan: o.span,
    children: [new Paragraph({ children: runs, alignment: o.align || AlignmentType.CENTER, spacing: { line: 260 } })],
  });
}
function table(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const hdrRows = (Array.isArray(headers[0]) ? headers : [headers]).map((hr) => new TableRow({ tableHeader: true,
    children: hr.map((h, i) => cell(h.text ?? h, h.w ?? widths[i], { head: true, bold: true, span: h.span })) }));
  const body = rows.map((r) => new TableRow({ cantSplit: true,
    children: r.map((v, i) => cell(v, widths[i], { align: o.alignFirstLeft && i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })) }));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [...hdrRows, ...body],
    alignment: AlignmentType.CENTER });
}

// ---------- phương trình (OMML) có đánh số bên phải
const mr = (s) => new MathRun(s);
const sup = (b, e) => new MathSuperScript({ children: typeof b === "string" ? [mr(b)] : b, superScript: typeof e === "string" ? [mr(e)] : e });
const sub = (b, e) => new MathSubScript({ children: typeof b === "string" ? [mr(b)] : b, subScript: typeof e === "string" ? [mr(e)] : e });
const frac = (n, d) => new MathFraction({ numerator: typeof n === "string" ? [mr(n)] : n, denominator: typeof d === "string" ? [mr(d)] : d });
const sqrt = (c) => new MathRadical({ children: typeof c === "string" ? [mr(c)] : c });
const sum = (c, lo, hi) => new MathSum({ children: c, subScript: lo ? [mr(lo)] : undefined, superScript: hi ? [mr(hi)] : undefined });
function eq(no, ...children) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.CENTER, position: Math.round(TW / 2) }, { type: TabStopType.RIGHT, position: TW }],
    spacing: { before: 80, after: 120, line: 360 },
    children: [t("\t"), new OMath({ children }), t(`\t(${no})`)],
  });
}

// ============================================================ TRANG BÌA
function cover() {
  const c = (text, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: o.after ?? 60, before: o.before || 0 },
    children: [t(text, { bold: o.bold, size: o.size || 26, italics: o.it })] });
  const memberRows = MEMBERS.map(([n, id], i) => new TableRow({ children: [
    noBorderCell(`${i + 1}.`, 700, AlignmentType.RIGHT), noBorderCell(n, 4000), noBorderCell(id, 1800)] }));
  return [
    c("ĐẠI HỌC QUỐC GIA THÀNH PHỐ HỒ CHÍ MINH", { size: 26 }),
    c("TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN", { bold: true }),
    c("KHOA VẬT LÝ – VẬT LÝ KỸ THUẬT", { bold: true }),
    c("———————————", { after: 600 }),
    c("BÁO CÁO THỰC HÀNH", { bold: true, size: 36, before: 600, after: 120 }),
    c("HỌC PHẦN: THỰC HÀNH VẬT LÝ HIỆN ĐẠI", { bold: true, size: 28, after: 600 }),
    c("BÀI THỰC TẬP", { size: 28 }),
    c("PHÂN BỐ THỐNG KÊ", { bold: true, size: 40 }),
    c("TRONG PHÂN RÃ PHÓNG XẠ", { bold: true, size: 40, after: 600 }),
    new Paragraph({ alignment: AlignmentType.LEFT, indent: { left: cm(2.5) }, spacing: { after: 60 },
      children: [t("Giảng viên hướng dẫn: ", { bold: true }), fill("……………………………")] }),
    new Paragraph({ alignment: AlignmentType.LEFT, indent: { left: cm(2.5) }, spacing: { after: 60 },
      children: [t("Nhóm thực hiện: ", { bold: true }), t("Nhóm 3")] }),
    new Paragraph({ alignment: AlignmentType.LEFT, indent: { left: cm(2.5) }, spacing: { after: 60 },
      children: [t("Thành viên:", { bold: true })] }),
    new Table({ width: { size: 6500, type: WidthType.DXA }, columnWidths: [700, 4000, 1800], rows: memberRows,
      alignment: AlignmentType.CENTER }),
    c("TP. Hồ Chí Minh, tháng 10 năm 2026", { it: true, before: 900 }),
  ];
}
function noBorderCell(text, w, align = AlignmentType.LEFT) {
  const nb = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: { top: nb, bottom: nb, left: nb, right: nb },
    children: [new Paragraph({ alignment: align, spacing: { after: 20 }, children: [t(text)] })] });
}

// ============================================================ PHẦN ĐẦU
function frontMatter() {
  const listPage = (title, items, prefix) => [
    new Paragraph({ alignment: AlignmentType.CENTER, pageBreakBefore: true, spacing: { after: 240 },
      children: [t(title, { bold: true, size: 28 })] }),
    ...items.map(([no, text]) => new Paragraph({ spacing: { after: 60, line: 320 }, indent: { left: cm(2), hanging: cm(2) },
      children: [t(`${prefix} ${no}. `, { bold: true }), t(text)] })),
  ];
  return { listPage };
}

// ============================================================ NỘI DUNG
function body() {
  const S = (k) => R[k];
  const out = [];
  const push = (...x) => out.push(...x.flat());

  // ---------------------------------------------------------------- MỞ ĐẦU
  push(H1("MỞ ĐẦU", true));
  push(para("Phân rã phóng xạ là ví dụ điển hình của một chuỗi sự kiện ngẫu nhiên và độc lập: thời điểm phân rã của mỗi hạt nhân không phụ thuộc vào các hạt nhân khác. Vì vậy, số đếm ghi nhận được bởi một đầu dò bức xạ trong những khoảng thời gian bằng nhau luôn thăng giáng quanh một giá trị trung bình, và các thăng giáng này tuân theo những quy luật thống kê xác định."));
  push(para("Bài thực tập nhằm các mục tiêu sau:"));
  push(bullet("Khảo sát tính dừng của quá trình đếm thông qua tốc độ đếm trung bình tích lũy."));
  push(bullet("Xác định trị trung bình, độ lệch chuẩn của số đếm và so sánh với dự đoán σ = √µ của phân bố Poisson."));
  push(bullet("So sánh phân bố tần suất thực nghiệm với phân bố Poisson và phân bố Gauss; đánh giá định lượng bằng kiểm định χ²."));
  push(bullet("Mô phỏng Monte Carlo phân bố Poisson và đối chiếu với kết quả thực nghiệm."));
  push(para("Thí nghiệm sử dụng ống đếm Geiger–Müller (GM) và nguồn gamma Ra-226, đo ở ba mức số đếm trung bình khác nhau (khoảng 5, 40 và 110 số đếm trong 10 s); mỗi mức gồm 100 lần đo 10 s và một lần đo 100 s."));

  // ---------------------------------------------------------------- CHƯƠNG 1
  push(H1("CHƯƠNG 1. CƠ SỞ LÝ THUYẾT"));
  push(H2("1.1. Quá trình ngẫu nhiên dừng"));
  push(para("Một quá trình ngẫu nhiên liên tục được gọi là dừng với tốc độ trung bình R nếu số sự kiện x xảy ra trong thời gian T thỏa mãn:"));
  push(eq("1.1", new MathLimitLower({ children: [mr("lim")], limit: [mr("T→∞")] }), new MathRoundBrackets({ children: [frac("x", "T")] }), mr("=R")));
  push(para("Để kiểm tra tính dừng trong thời gian thí nghiệm, ta đo lặp lại số đếm xᵢ trong các khoảng thời gian tᵢ và theo dõi tốc độ đếm trung bình tích lũy:"));
  push(eq("1.2", sub("r", "c"), mr("(j)="), frac([sum([sub("x", "i")], "i=1", "j")], [sum([sub("t", "i")], "i=1", "j")])));
  push(para("Với một quá trình dừng, rc(j) hội tụ về tốc độ đếm thực R khi j tăng; độ phân tán của rc(j) quanh R giảm theo 1/√j."));

  push(H2("1.2. Phân bố nhị thức"));
  push(para("Xét n hạt nhân, mỗi hạt có xác suất p phân rã (và được ghi nhận) trong thời gian T, độc lập với nhau. Xác suất có đúng x sự kiện là phân bố nhị thức:"));
  push(eq("1.3", sub("P", "B"), mr("(x;n,p)="), frac("n!", "x!(n−x)!"), sup("p", "x"), sup("(1−p)", "n−x")));
  push(para("Phân bố nhị thức có trị trung bình µ = np và phương sai σ² = np(1 − p)."));

  push(H2("1.3. Phân bố Poisson"));
  push(para("Khi n → ∞, p → 0 nhưng tích µ = np = RT giữ hữu hạn, phân bố nhị thức tiến tới phân bố Poisson:"));
  push(eq("1.4", mr("P(x;μ)="), frac([sup("μ", "x"), sup("e", "−μ")], "x!")));
  push(para("Trong đó x là số đếm (nguyên, không âm) và µ là số đếm kỳ vọng (không nhất thiết nguyên). Phân bố Poisson chỉ phụ thuộc một tham số và có tính chất đặc trưng: phương sai bằng trị trung bình, σ² = µ, tức σ = √µ. Phân bố bất đối xứng với hệ số bất đối xứng 1/√µ, rõ rệt khi µ nhỏ."));

  push(H2("1.4. Phân bố Gauss"));
  push(para("Khi µ lớn, phân bố Poisson tiệm cận phân bố Gauss (phân bố chuẩn) có trị trung bình µ và phương sai σ² = µ:"));
  push(eq("1.5", mr("p(x;μ)="), frac("1", [sqrt("2πμ")]), sup("e", [mr("−"), frac([sup("(x−μ)", "2")], "2μ")])));
  push(para("Khác với Poisson, phân bố Gauss là phân bố liên tục và đối xứng; nó xác định cả với x âm. Xấp xỉ này càng tốt khi µ càng lớn (thường chấp nhận khi µ ≳ 20)."));

  push(H2("1.5. Các đại lượng thống kê và kiểm định"));
  push(para("Với n giá trị đo x₁, …, xₙ, trị trung bình, độ lệch chuẩn thực nghiệm và độ không đảm bảo của trị trung bình được tính theo:"));
  push(eq("1.6", mr("x̄="), frac("1", "n"), sum([sub("x", "i")], "i=1", "n"), mr(",  s="), sqrt([frac("1", "n−1"), sum([sup([sub("(x", "i"), mr("−x̄)")], "2")], "i=1", "n")]), mr(",  u(x̄)="), frac("s", [sqrt("n")])));
  push(para("Mức độ phù hợp giữa tần suất thực nghiệm Oₖ và tần suất lý thuyết Eₖ = n·P(k) được đánh giá bằng đại lượng χ² theo công thức (5) của tài liệu hướng dẫn, với sai số thực nghiệm σₖ = √Oₖ:"));
  push(eq("1.7", sup("χ", "2"), mr("="), sum([frac([sup([sub("(O", "k"), mr("−"), sub("E", "k"), mr(")")], "2")], [sup([sub("σ", "k")], "2")])], "k", null), mr(",   "), sub("σ", "k"), mr("="), sqrt([sub("O", "k")])));
  push(para("Số bậc tự do ν bằng số bin sử dụng trừ đi số tham số ước lượng từ dữ liệu (ở đây là µ). Khi mô hình đúng, χ²/ν ≈ 1. Do σₖ = √Oₖ không xác định khi Oₖ = 0 và bị đánh giá thấp khi Oₖ nhỏ, báo cáo bổ sung χ² Pearson Σ(Oₖ − Eₖ)²/Eₖ sau khi gộp bin sao cho Eₖ ≥ 5 để đối chiếu."));
  push(para("Ngoài ra, tỉ số phân tán D = s²/x̄ được dùng để kiểm tra trực tiếp tính chất σ² = µ: với số liệu Poisson, đại lượng (n − 1)D tuân theo phân bố χ² với n − 1 bậc tự do."));

  // ---------------------------------------------------------------- CHƯƠNG 2
  push(H1("CHƯƠNG 2. THIẾT BỊ VÀ PHƯƠNG PHÁP THỰC NGHIỆM"));
  push(H2("2.1. Hệ đo"));
  push(para("Hệ đo gồm nguồn phóng xạ gamma đặt đối diện ống đếm Geiger–Müller; ống đếm được cấp điện áp bởi nguồn cao thế và tín hiệu xung được ghi nhận bởi bộ đếm (Hình 2.1, Hình 2.2)."));
  push(figure("fig_setup_diagram.png", 380, "2.1", "Sơ đồ bố trí thí nghiệm", "Nguồn: tài liệu hướng dẫn thực tập [3]."));
  push(figure("fig_setup_photo.png", 560, "2.2", "Hệ đo thực tế tại phòng thí nghiệm", "Nguồn: tài liệu hướng dẫn thực tập [3]."));
  push(tableCaption("2.1", "Thiết bị và điều kiện đo"));
  push(table(["Hạng mục", "Thông số"], [
    ["Đầu dò", [t("Ống đếm Geiger–Müller; model: ", { size: 24 }), fill("……")]],
    ["Bộ đếm", [t("Model: ", { size: 24 }), fill("……")]],
    ["Điện áp làm việc", [fill("……"), t(" V (tăng từ từ đến giá trị làm việc)", { size: 24 })]],
    ["Nguồn phóng xạ", "Ra-226 (Bộ 2 và Bộ 3)"],
    ["Bộ 1 (mức thấp)", [t("Ghi “Font” (phông) trong số liệu gốc; ", { size: 24 }), fill("xác nhận có/không đặt nguồn")]],
    ["Khoảng cách nguồn – đầu dò", [t("Bộ 1: ", { size: 24 }), fill("…"), t(" cm; Bộ 2: ", { size: 24 }), fill("…"), t(" cm; Bộ 3: ", { size: 24 }), fill("…"), t(" cm", { size: 24 })]],
    ["Thời gian đo", "100 lần × 10 s và 1 lần × 100 s cho mỗi bộ"],
    ["Ngày đo / phòng thí nghiệm", [fill("……")]],
  ].map((r) => r.map((v) => (typeof v === "string" ? v : v))), [3000, 6070], { alignFirstLeft: true }));

  push(H2("2.2. Quy trình đo"));
  push(numbered("Bật nguồn cao thế, tăng điện áp từ từ đến điện áp làm việc của ống đếm GM; đặt thời gian đo 10 s.", "steps"));
  push(numbered("Đặt nguồn gamma đối diện đầu dò; điều chỉnh khoảng cách nguồn – đầu dò để đạt ba mức số đếm: 5–10, 40–50 và 100–120 số đếm trong 10 s.", "steps"));
  push(numbered("Với mỗi mức: thực hiện 100 lần đo, mỗi lần 10 s; sau đó thực hiện một lần đo 100 s ở cùng cấu hình.", "steps"));
  push(numbered("Ghi đầy đủ điều kiện đo và số liệu vào bảng (Phụ lục A).", "steps"));

  push(H2("2.3. Phương pháp xử lý số liệu"));
  push(para("Số liệu được xử lý bằng MS Excel (tệp BaoCao_SoLieu_Nhom3.xlsx, toàn bộ đại lượng là công thức) và kiểm tra chéo bằng chương trình Python (analysis.py, sử dụng numpy/scipy/matplotlib). Các bước xử lý theo mục 2.3 của tài liệu hướng dẫn:"));
  push(bullet("Tính và vẽ rc(j) theo công thức (1.2) với j = 1, …, 100."));
  push(bullet("Tính x̄, s, u(x̄), √x̄, D; so sánh với phép đo 100 s quy về 10 s bằng chỉ số z = Δ/u(Δ)."));
  push(bullet("Lập bảng tần suất, tính tần suất Poisson lý thuyết 100·P(k; x̄), vẽ trên cùng đồ thị với số liệu thực nghiệm (kèm thanh sai số ±√Oₖ)."));
  push(bullet("Với bộ có µ thấp nhất, đánh giá độ lệch lớn nhất so với trung bình có tính đến 100 cơ hội xuất hiện."));
  push(bullet("So sánh Poisson và Gauss (µ = σ² = x̄) cho hai bộ µ ≈ 41 và µ ≈ 112 bằng χ² theo công thức (1.7)."));
  push(para("Độ không đảm bảo được báo cáo ở mức chuẩn (k = 1, xác suất bao phủ ≈ 68%). Số liệu được làm tròn khi trình bày; mọi phép tính dùng giá trị đầy đủ."));

  // ---------------------------------------------------------------- CHƯƠNG 3
  push(H1("CHƯƠNG 3. KẾT QUẢ VÀ THẢO LUẬN"));
  push(H2("3.1. Bảng tần suất số đếm"));
  push(para("Bảng 3.1 – 3.3 trình bày tần suất xuất hiện Oₖ của mỗi giá trị số đếm k trong 100 lần đo 10 s, cùng tần suất lý thuyết Poisson Eₖ(P) và Gauss Eₖ(G) với µ = x̄ thực nghiệm. Chỉ liệt kê các giá trị k nằm trong khoảng [min; max] của số liệu; số liệu thô xem ở Phụ lục A."));
  const freqTable = (k, no, groups) => {
    const r = S(k);
    const items = r.k.map((kk, i) => [String(kk), String(r.obs[i]), vn(r.poi[i]), vn(r.gau[i])]);
    const per = Math.ceil(items.length / groups);
    const rows = [];
    for (let i = 0; i < per; i++) {
      const row = [];
      for (let g = 0; g < groups; g++) row.push(...(items[g * per + i] || ["", "", "", ""]));
      rows.push(row);
    }
    const hdr = [];
    for (let g = 0; g < groups; g++) hdr.push("k", "Oₖ", "Eₖ(P)", "Eₖ(G)");
    const w = Math.floor(TW / (groups * 4));
    const widths = Array(groups * 4).fill(0).map((_, i) => (i % 4 < 2 ? Math.floor(w * 0.8) : Math.floor(w * 1.2)));
    return [tableCaption(no, `Tần suất số đếm – ${NAME[k]}, x̄ = ${vn(r.mean)} (n = 100)`), table(hdr, rows, widths)];
  };
  push(freqTable("S1", "3.1", 2));
  push(freqTable("S2", "3.2", 3));
  push(freqTable("S3", "3.3", 3));

  push(H2("3.2. Tốc độ đếm trung bình tích lũy"));
  push(para(`Hình 3.1 biểu diễn rc(j) của ba bộ số liệu. Ở những lần đo đầu, rc(j) dao động mạnh (ví dụ Bộ 1 bắt đầu từ 1,0 s⁻¹, Bộ 2 giảm xuống 3,3 s⁻¹) do số đếm tích lũy còn nhỏ. Sau khoảng 20–30 lần đo, rc(j) của cả ba bộ đi vào dải ±1σ Poisson (√(µj)/(j·t)) quanh giá trị trung bình và hội tụ về ${vn(S("S1").rate_cps, 3)} s⁻¹, ${vn(S("S2").rate_cps, 3)} s⁻¹ và ${vn(S("S3").rate_cps, 3)} s⁻¹. Không quan sát thấy xu hướng trôi theo thời gian; có thể kết luận quá trình đếm là dừng trong suốt thời gian thí nghiệm.`));
  push(figure("fig1_running_mean.png", 520, "3.1", "Tốc độ đếm trung bình tích lũy rc(j) theo số lần đo j; vùng tô: dải ±1σ Poisson"));

  push(H2("3.3. Trị trung bình, độ lệch chuẩn và so sánh với phép đo 100 s"));
  push(tableCaption("3.4", "Các đặc trưng thống kê của 100 lần đo 10 s"));
  push(table(["Đại lượng", NAME.S1, NAME.S2, NAME.S3], [
    ["Tổng số đếm ΣN", ...SETS.map((k) => String(S(k).sum))],
    ["Trung bình x̄ ± u(x̄)", ...SETS.map((k) => `${vn(S(k).mean)} ± ${vn(S(k).u_mean)}`)],
    ["Tốc độ đếm R (s⁻¹)", ...SETS.map((k) => `${vn(S(k).rate_cps, 3)} ± ${vn(S(k).u_rate_cps, 3)}`)],
    ["Độ lệch chuẩn s", ...SETS.map((k) => vn(S(k).s))],
    ["√x̄ (Poisson)", ...SETS.map((k) => vn(S(k).sqrt_mean))],
    ["s/√x̄", ...SETS.map((k) => vn(S(k).s / S(k).sqrt_mean, 3))],
    ["D = s²/x̄", ...SETS.map((k) => vn(S(k).D, 3))],
    ["p-value (kiểm định phân tán)", ...SETS.map((k) => vn(S(k).disp_p, 3))],
    ["Giá trị nhỏ nhất – lớn nhất", ...SETS.map((k) => `${S(k).min} – ${S(k).max}`)],
  ], [2770, 2100, 2100, 2100], { alignFirstLeft: true }));
  push(para(`Độ lệch chuẩn thực nghiệm s lớn hơn √x̄ khoảng 5–10% ở cả ba bộ (s/√x̄ = ${vn(S("S1").s / S("S1").sqrt_mean, 3)}; ${vn(S("S2").s / S("S2").sqrt_mean, 3)}; ${vn(S("S3").s / S("S3").sqrt_mean, 3)}). Tuy nhiên, với n = 100, bản thân s cũng thăng giáng với độ lệch chuẩn tương đối ≈ 1/√(2(n − 1)) ≈ 7%; kiểm định phân tán cho p-value lần lượt ${vn(S("S1").disp_p, 2)}; ${vn(S("S2").disp_p, 2)}; ${vn(S("S3").disp_p, 2)} (đều > 0,05). Vậy phương sai thực nghiệm phù hợp với tính chất σ² = µ của phân bố Poisson.`, { before: 120 }));
  push(tableCaption("3.5", "So sánh trung bình 100 lần đo 10 s với phép đo 100 s (quy về 10 s)"));
  push(table(["Đại lượng", NAME.S1, NAME.S2, NAME.S3], [
    ["Số đếm trong 100 s, N_T", ...SETS.map((k) => String(S(k).N100))],
    ["N_T/10 ± √N_T/10", ...SETS.map((k) => `${vn(S(k).N100_per10)} ± ${vn(S(k).u_N100_per10)}`)],
    ["x̄ ± u(x̄) (100 × 10 s)", ...SETS.map((k) => `${vn(S(k).mean)} ± ${vn(S(k).u_mean)}`)],
    ["Hiệu Δ", ...SETS.map((k) => vn(S(k).diff100))],
    ["z = Δ/u(Δ)", ...SETS.map((k) => vn(S(k).z100))],
    ["√(N_T/10) – σ dự đoán cho 10 s", ...SETS.map((k) => vn(Math.sqrt(S(k).N100_per10)))],
    ["Kết luận", ...SETS.map((k) => (Math.abs(S(k).z100) <= 2 ? "Phù hợp" : "Không phù hợp"))],
  ], [2770, 2100, 2100, 2100], { alignFirstLeft: true }));
  push(para(`Kết quả phép đo 100 s phù hợp với trị trung bình của 100 lần đo 10 s trong phạm vi ${vn(Math.max(...SETS.map((k) => Math.abs(S(k).z100))), 1)} độ lệch chuẩn tổng hợp (|z| < 2 ở cả ba bộ). Lưu ý phép đo 100 s chỉ là một lần đo nên có độ không đảm bảo √N_T lớn hơn nhiều so với u(x̄) của 100 lần đo (tổng thời gian 1000 s). Độ lệch chuẩn dự đoán từ phép đo 100 s, √(N_T/10), cũng gần với s thực nghiệm, đúng như kỳ vọng với phân bố Poisson: chỉ một phép đo đã đủ ước lượng độ thăng giáng của các phép đo khác.`, { before: 120 }));

  push(H2("3.4. Phân bố tần suất thực nghiệm và phân bố Poisson"));
  push(para("Hình 3.2 – 3.4 so sánh tần suất thực nghiệm (điểm, thanh sai số ±√Oₖ) với phân bố Poisson lý thuyết chuẩn hóa theo 100 lần đo (đường liền), và phân bố Gauss (đường đứt) cho hai bộ có µ lớn."));
  push(figure("fig2_hist_S1.png", 500, "3.2", `Phân bố tần suất – ${NAME.S1}, µ = ${vn(S("S1").mean)}`));
  push(figure("fig3_hist_S2.png", 500, "3.3", `Phân bố tần suất – ${NAME.S2}, µ = ${vn(S("S2").mean)}`));
  push(figure("fig4_hist_S3.png", 500, "3.4", `Phân bố tần suất – ${NAME.S3}, µ = ${vn(S("S3").mean)}`));
  push(para("Về tổng thể, hình dạng phân bố thực nghiệm bám theo đường Poisson ở cả ba mức: Bộ 1 thể hiện rõ tính bất đối xứng (đuôi phải dài hơn) của phân bố Poisson khi µ nhỏ; Bộ 2 và Bộ 3 gần như đối xứng. Phần lớn các điểm thực nghiệm nằm trong khoảng ±1–2 thanh sai số so với đường lý thuyết, đúng với mức thăng giáng thống kê mong đợi của mẫu 100 lần đo."));

  push(H2("3.5. Độ lệch lớn nhất ở bộ có tốc độ đếm thấp nhất"));
  const md = S("S1").maxdev;
  push(para(`Với Bộ 1 (x̄ = ${vn(S("S1").mean)}, σ = √x̄ = ${vn(S("S1").sqrt_mean)}), giá trị lệch xa trung bình nhất là x = ${md.value}, độ lệch ${vn(md.dev)} ≈ ${vn(md.nsig, 1)}σ. Theo phân bố Poisson, xác suất một lần đo cho x ≥ ${md.value} là:`));
  push(eq("3.1", mr(`P(x≥${md.value};${vn(S("S1").mean)})=1−`), sum([mr(`P(k;${vn(S("S1").mean)})`)], "k=0", String(md.value - 1)), mr(`≈${vn(md.p1, 4)}`)));
  push(para(`Vì có 100 cơ hội để độ lệch như vậy xảy ra, xác suất có ít nhất một lần đo đạt x ≥ ${md.value} là:`));
  push(eq("3.2", mr("P(≥1 lần)=1−"), sup(`(1−${vn(md.p1, 4)})`, "100"), mr(`≈${vn(md.p100, 2)}`)));
  push(para(`Số lần kỳ vọng là 100 × ${vn(md.p1, 4)} ≈ ${vn(md.expected, 2)}, tức việc quan sát đúng một giá trị như vậy là hoàn toàn bình thường (xác suất ≈ ${pct(md.p100, 0)}). Nếu chỉ xét riêng một lần đo (p ≈ ${pct(md.p1, 1)}) ta có thể lầm tưởng đây là bất thường; khi tính đến số cơ hội, độ lệch quan sát được không cho thấy dấu hiệu hệ đếm hoạt động sai (không có xung nhiễu, phóng điện hay mất số đếm). Hai bộ còn lại cho kết luận tương tự (Bảng 3.6).`));
  push(tableCaption("3.6", "Độ lệch lớn nhất so với trung bình"));
  push(table(["Đại lượng", NAME.S1, NAME.S2, NAME.S3], [
    ["Giá trị lệch xa nhất", ...SETS.map((k) => String(S(k).maxdev.value))],
    ["Độ lệch x − x̄", ...SETS.map((k) => vn(S(k).maxdev.dev))],
    ["Theo σ = √x̄", ...SETS.map((k) => vn(S(k).maxdev.nsig))],
    ["Xác suất 1 lần đo (một phía)", ...SETS.map((k) => vn(S(k).maxdev.p1, 4))],
    ["Xác suất ≥ 1 lần trong 100 lần", ...SETS.map((k) => vn(S(k).maxdev.p100, 3))],
  ], [2770, 2100, 2100, 2100], { alignFirstLeft: true }));

  push(H2("3.6. So sánh phân bố Poisson và Gauss"));
  push(H3("a) So sánh định tính"));
  push(para(`Trên Hình 3.3 (µ = ${vn(S("S2").mean)}), đường Gauss lệch nhẹ sang phải so với Poisson ở vùng đỉnh và thấp hơn ở đuôi phải, phản ánh tính bất đối xứng còn lại của Poisson (hệ số bất đối xứng 1/√µ ≈ ${vn(1 / Math.sqrt(S("S2").mean), 2)}). Ở µ = ${vn(S("S3").mean)} (Hình 3.4), hai đường gần như trùng khít (1/√µ ≈ ${vn(1 / Math.sqrt(S("S3").mean), 2)}). Với Bộ 1 (µ ≈ 5,5) xấp xỉ Gauss không phù hợp: hệ số bất đối xứng ≈ ${vn(1 / Math.sqrt(S("S1").mean), 2)} và Gauss gán xác suất khác 0 cho số đếm âm.`));
  push(H3("b) So sánh định lượng bằng χ²"));
  push(tableCaption("3.7", "Kết quả kiểm định χ² giữa số liệu thực nghiệm và phân bố lý thuyết"));
  push(table([
    [{ text: "Bộ số liệu", w: 1870 }, { text: "χ² theo công thức (1.7)", span: 4, w: 4400 }, { text: "χ² Pearson (gộp bin E ≥ 5)", span: 2, w: 2800 }],
    ["", "χ² Poisson", "χ² Gauss", "ν", "Phù hợp hơn", "Poisson: χ²/ν (p)", "Gauss: χ²/ν (p)"],
  ], SETS.map((k) => [NAME[k], vn(S(k).chi2_poi), vn(S(k).chi2_gau), String(S(k).ndf),
    S(k).chi2_poi <= S(k).chi2_gau ? "Poisson" : "Gauss",
    `${vn(S(k).pearson_poi.chi2, 1)}/${S(k).pearson_poi.ndf} (${vn(S(k).pearson_poi.p, 2)})`,
    `${vn(S(k).pearson_gau.chi2, 1)}/${S(k).pearson_gau.ndf} (${vn(S(k).pearson_gau.p, 2)})`]),
  [1870, 1100, 1100, 800, 1400, 1400, 1400], { alignFirstLeft: true }));
  push(para(`Ở µ ≈ 41, χ² của Poisson (${vn(S("S2").chi2_poi)}) nhỏ hơn của Gauss (${vn(S("S2").chi2_gau)}): phân bố Poisson mô tả số liệu tốt hơn. Ở µ ≈ 112, hai giá trị gần như bằng nhau (${vn(S("S3").chi2_poi)} và ${vn(S("S3").chi2_gau)}, χ²/ν ≈ 0,85), chênh lệch nhỏ hơn nhiều so với thăng giáng thống kê của chính χ² (≈ √(2ν) ≈ 9): hai phân bố phù hợp với số liệu như nhau, xác nhận Gauss là xấp xỉ tốt của Poisson khi µ lớn.`, { before: 120 }));
  push(para(`Lưu ý về công thức (1.7): vì σₖ = √Oₖ, các bin có Oₖ rất nhỏ đóng góp quá mức vào χ². Cụ thể, ở Bộ 1 riêng bin k = 8 (O = 2, E = 8,56) đóng góp 21,5 trên tổng ${vn(S("S1").chi2_poi)}; ở Bộ 2 bin k = 44 (O = 1, E = 5,48) đóng góp 20,0 trên tổng ${vn(S("S2").chi2_poi)}. Đây là lý do χ²/ν của Bộ 1 và Bộ 2 lớn hơn 1. Khi dùng χ² Pearson với bin đã gộp (E ≥ 5), p-value của Poisson là ${vn(S("S1").pearson_poi.p, 2)}; ${vn(S("S2").pearson_poi.p, 2)}; ${vn(S("S3").pearson_poi.p, 2)}, đều lớn hơn 0,05: không có cơ sở bác bỏ giả thuyết số liệu tuân theo phân bố Poisson ở cả ba mức.`));

  // ---------------------------------------------------------------- CHƯƠNG 4
  push(H1("CHƯƠNG 4. MÔ PHỎNG MONTE CARLO"));
  push(H2("4.1. Phương pháp mô phỏng"));
  push(para(`Chương trình Python (analysis.py) sinh các số ngẫu nhiên tuân theo phân bố Poisson bằng hàm numpy.random.Generator.poisson (seed = ${R._meta.SEED}) với µ bằng trị trung bình thực nghiệm của từng bộ. Thực hiện hai mô phỏng:`));
  push(bullet(`Mô phỏng A: ${R._meta.N_MC} lần thử cho mỗi µ (theo yêu cầu mục 3.1 của tài liệu hướng dẫn), xuất ra tệp CSV hai cột (số đếm, tần suất) giống hàm poissonsim.`));
  push(bullet(`Mô phỏng B: lặp lại ${R._meta.N_REP} “thí nghiệm ảo”, mỗi thí nghiệm gồm 100 lần đo như thực tế, để xác định phân bố của độ lệch chuẩn mẫu s và so sánh với s thực nghiệm.`));
  push(H2("4.2. Kết quả và so sánh với thực nghiệm"));
  push(figure("fig5_montecarlo.png", 600, "4.1", `Phân bố mô phỏng Monte Carlo (${R._meta.N_MC} lần thử), phân bố thực nghiệm và Poisson lý thuyết`));
  push(tableCaption("4.1", "So sánh kết quả mô phỏng Monte Carlo và thực nghiệm"));
  push(table(["Đại lượng", NAME.S1, NAME.S2, NAME.S3], [
    ["µ đầu vào (= x̄ thực nghiệm)", ...SETS.map((k) => vn(S(k).mean))],
    [`x̄ MC (${R._meta.N_MC} lần) ± u`, ...SETS.map((k) => `${vn(S(k).mc.mean)} ± ${vn(S(k).mc.u_mean)}`)],
    ["s MC", ...SETS.map((k) => vn(S(k).mc.s))],
    ["√µ", ...SETS.map((k) => vn(S(k).sqrt_mean))],
    ["s thực nghiệm", ...SETS.map((k) => vn(S(k).s))],
    [`Mô phỏng B: s trung bình ± độ lệch chuẩn`, ...SETS.map((k) => `${vn(S(k).mc_rep.s_mean)} ± ${vn(S(k).mc_rep.s_std)}`)],
    ["Tỉ lệ thí nghiệm ảo có s < s thực nghiệm", ...SETS.map((k) => `${vn(S(k).mc_rep.s_pct, 1)}%`)],
  ], [2770, 2100, 2100, 2100], { alignFirstLeft: true }));
  push(para(`Trung bình và độ lệch chuẩn của mẫu ${R._meta.N_MC} lần thử hội tụ về µ và √µ trong phạm vi sai số thống kê (|x̄_MC − µ| < 2u). Hình dạng phân bố mô phỏng trùng với đường Poisson lý thuyết và chồng khớp với histogram thực nghiệm ở cả ba mức (Hình 4.1); phân bố thực nghiệm “gồ ghề” hơn chỉ vì cỡ mẫu nhỏ hơn 10 lần.`, { before: 120 }));
  push(figure("fig6_s_spread.png", 600, "4.2", `Phân bố độ lệch chuẩn mẫu s của ${R._meta.N_REP} thí nghiệm ảo (mỗi thí nghiệm 100 lần đo); đường đỏ: s thực nghiệm`));
  push(para(`Mô phỏng B cho thấy với 100 lần đo, s tự nhiên thăng giáng khoảng ±7% quanh √µ. Giá trị s thực nghiệm nằm ở phân vị ${vn(S("S1").mc_rep.s_pct, 0)}%, ${vn(S("S2").mc_rep.s_pct, 0)}% và ${vn(S("S3").mc_rep.s_pct, 0)}% của phân bố mô phỏng, tức trong khoảng 2σ, phù hợp với kết luận của kiểm định phân tán ở mục 3.3. Việc cả ba giá trị đều nằm ở nửa trên có thể gợi ý một thành phần thăng giáng nhỏ ngoài thống kê Poisson (ví dụ dao động điện áp cao thế hoặc vị trí nguồn), song với cỡ mẫu hiện tại hiệu ứng này không có ý nghĩa thống kê.`));

  // ---------------------------------------------------------------- KẾT LUẬN
  push(H1("KẾT LUẬN"));
  push(para("Từ kết quả đo và phân tích ba bộ số liệu đếm gamma bằng ống đếm Geiger–Müller, nhóm rút ra các kết luận sau:"));
  push(numbered(`Quá trình đếm là quá trình dừng: tốc độ đếm trung bình tích lũy rc(j) hội tụ về ${vn(S("S1").rate_cps, 3)}; ${vn(S("S2").rate_cps, 3)} và ${vn(S("S3").rate_cps, 3)} s⁻¹ mà không có xu hướng trôi.`, "concl"));
  push(numbered(`Trị trung bình 100 lần đo 10 s (${vn(S("S1").mean)} ± ${vn(S("S1").u_mean)}; ${vn(S("S2").mean)} ± ${vn(S("S2").u_mean)}; ${vn(S("S3").mean)} ± ${vn(S("S3").u_mean)}) phù hợp với phép đo 100 s trong phạm vi 1,5σ.`, "concl"));
  push(numbered("Độ lệch chuẩn thực nghiệm phù hợp với √µ (p-value kiểm định phân tán 0,17 – 0,47): phương sai bằng trị trung bình, đặc trưng của phân bố Poisson.", "concl"));
  push(numbered(`Độ lệch lớn nhất ở bộ mức thấp (≈ ${vn(md.nsig, 1)}σ) có xác suất xuất hiện ≈ ${pct(md.p100, 0)} trong 100 lần đo: hệ đếm hoạt động bình thường.`, "concl"));
  push(numbered("Phân bố Poisson mô tả tốt số liệu ở cả ba mức. Ở µ ≈ 41 Poisson phù hợp hơn Gauss; ở µ ≈ 112 hai phân bố cho kết quả tương đương, xác nhận xấp xỉ Gauss chỉ hợp lệ khi µ đủ lớn.", "concl"));
  push(numbered("Mô phỏng Monte Carlo tái tạo đúng phân bố thực nghiệm; độ lệch chuẩn thực nghiệm nằm trong phạm vi thăng giáng tự nhiên của mẫu 100 lần đo.", "concl"));
  push(para("Một số hạn chế: cỡ mẫu 100 lần đo khiến các bin ở đuôi phân bố có rất ít sự kiện, làm χ² theo công thức (1.7) nhạy với các bin có O nhỏ; nên tăng số lần đo hoặc gộp bin khi kiểm định. Thời gian chết của ống GM không được hiệu chỉnh, tuy nhiên ở tốc độ đếm ≤ 12 s⁻¹ ảnh hưởng này dưới 0,2% (với thời gian chết cỡ 100 µs) và không đáng kể.", { before: 120 }));

  // ---------------------------------------------------------------- CÂU HỎI
  push(H1("CÂU HỎI TÌM HIỂU THÊM"));
  push(H2("Câu 1. Nguồn gốc phân bố Poisson từ phân bố nhị thức và giới hạn áp dụng"));
  push(para("Xét mẫu có n hạt nhân phóng xạ. Trong thời gian T, mỗi hạt nhân có xác suất phân rã p = 1 − e^(−λT) và các hạt nhân phân rã độc lập, nên số phân rã x tuân theo phân bố nhị thức (1.3). Với µ = np giữ cố định, viết p = µ/n và cho n → ∞:"));
  push(eq("Q.1", frac("n!", "(n−x)!"), sup([new MathRoundBrackets({ children: [frac("μ", "n")] })], "x"), mr("→"), sup("μ", "x"), mr(",   "), sup([new MathRoundBrackets({ children: [mr("1−"), frac("μ", "n")] })], "n−x"), mr("→"), sup("e", "−μ")));
  push(para("Thay vào (1.3) thu được P(x; µ) = µˣe^(−µ)/x!, chính là phân bố Poisson (1.4). Phương sai np(1 − p) → np = µ khi p → 0."));
  push(para("Điều kiện áp dụng:"));
  push(bullet("Số “phép thử” n rất lớn và xác suất p mỗi phép thử rất nhỏ (p ≪ 1). Với Ra-226 (T½ ≈ 1600 năm), trong 10 s p ≈ λT ≈ 1,4·10⁻¹⁰, điều kiện được thỏa mãn rất tốt; tốc độ đếm cũng không đổi trong thời gian đo."));
  push(bullet("Các sự kiện độc lập: sự xảy ra của sự kiện này không ảnh hưởng đến sự kiện khác. Điều kiện này bị vi phạm với các quá trình có tương quan như phản ứng phân hạch dây chuyền."));
  push(bullet("Hệ đo không làm mất tính ngẫu nhiên: thời gian chết của đầu dò phải không đáng kể so với khoảng thời gian trung bình giữa hai xung. Khi tốc độ đếm cao, các xung bị mất làm phương sai nhỏ hơn trị trung bình (D < 1)."));
  push(bullet("Nếu p không nhỏ (ví dụ đồng vị có chu kỳ bán rã ngắn so với thời gian đo), phải dùng phân bố nhị thức với phương sai np(1 − p) < µ."));

  push(H2("Câu 2. Phép xấp xỉ Gauss đối với phân bố Poisson khi µ lớn"));
  push(para("Đặt x = µ + δ với |δ| ≪ µ và dùng công thức Stirling ln x! ≈ x ln x − x + ½ ln(2πx):"));
  push(eq("Q.2", mr("ln P(x;μ)=x ln μ−μ−ln x!≈−"), frac([sup("δ", "2")], "2μ"), mr("−"), frac("1", "2"), mr("ln(2πμ)+O("), frac([sup("δ", "3")], [sup("μ", "2")]), mr(")")));
  push(para("Bỏ số hạng bậc ba, ta được P(x; µ) ≈ exp(−(x − µ)²/2µ)/√(2πµ), tức phân bố Gauss có trị trung bình µ và phương sai σ² = µ (công thức 1.5). Số hạng bỏ qua tỉ lệ với δ³/µ²; trong vùng |δ| ~ σ = √µ, nó có độ lớn ~1/√µ, cũng chính là hệ số bất đối xứng của Poisson. Do đó:"));
  push(bullet("Khi µ ≳ 20 (sai số tương đối ở đỉnh ≈ 1/(12µ) < 0,5%, bất đối xứng ≲ 0,2), xấp xỉ Gauss chấp nhận được; với µ ≳ 100 hai phân bố gần như trùng nhau."));
  push(bullet("Sai lệch lớn nhất xuất hiện ở các đuôi (|δ| ≫ σ), nơi số hạng bậc ba không còn nhỏ; Gauss đánh giá thấp đuôi phải và đánh giá cao đuôi trái."));
  push(bullet(`Số liệu của nhóm minh họa điều này: µ ≈ 5,5 (bất đối xứng ${vn(1 / Math.sqrt(S("S1").mean), 2)}) Gauss phù hợp kém hơn rõ rệt; µ ≈ 41 Poisson còn phù hợp hơn; µ ≈ 112 hai phân bố tương đương.`));

  // ---------------------------------------------------------------- TÀI LIỆU
  push(H1("TÀI LIỆU THAM KHẢO"));
  const refs = [
    "Bevington, P. R., & Robinson, D. K. (2003). Data Reduction and Error Analysis for the Physical Sciences (3rd ed.). McGraw-Hill, New York.",
    "Knoll, G. F. (2010). Radiation Detection and Measurement (4th ed.). John Wiley & Sons, Hoboken.",
    "Khoa Vật lý – Vật lý Kỹ thuật, Trường ĐH Khoa học Tự nhiên, ĐHQG-HCM. Tài liệu hướng dẫn thực tập: Phân bố thống kê trong phân rã phóng xạ (Poisson.pdf).",
    "Taylor, J. R. (1997). An Introduction to Error Analysis (2nd ed.). University Science Books, Sausalito.",
    "JCGM 100:2008. Evaluation of measurement data – Guide to the expression of uncertainty in measurement (GUM).",
  ];
  refs.forEach((r, i) => push(new Paragraph({ spacing: { after: 100, line: 340 }, indent: { left: cm(1), hanging: cm(1) },
    alignment: AlignmentType.JUSTIFIED, children: [t(`[${i + 1}]\t${r}`)], tabStops: [{ type: TabStopType.LEFT, position: cm(1) }] })));

  // ---------------------------------------------------------------- PHỤ LỤC
  push(H1("PHỤ LỤC"));
  push(H2("Phụ lục A. Số liệu đếm thô"));
  push(para("Số đếm ghi nhận trong mỗi lần đo 10 s; lần đo thứ j = 10·(hàng − 1) + cột. Số liệu gốc: tệp “nhóm 1.xlsx”."));
  SETS.forEach((k, i) => {
    const x = S(k).x;
    const rows = [];
    for (let r = 0; r < 10; r++) rows.push([String(r + 1), ...x.slice(r * 10, r * 10 + 10).map(String)]);
    push(tableCaption(`A.${i + 1}`, `Số đếm 10 s – ${NAME[k]}; phép đo 100 s: N = ${S(k).N100}`));
    push(table(["", ...Array.from({ length: 10 }, (_, c) => String(c + 1))], rows, [970, ...Array(10).fill(810)]));
  });
  push(H2("Phụ lục B. Bảng phân công công việc"));
  push(para([t("Nội dung phân công dưới đây là "), fill("đề xuất, nhóm cần điều chỉnh theo thực tế"), t(".")]));
  push(tableCaption("B.1", "Phân công công việc các thành viên Nhóm 3"));
  push(table(["STT", "Họ và tên", "MSSV", "Nhiệm vụ", "Hoàn thành"],
    MEMBERS.map(([n, id], i) => [String(i + 1), n, id, TASKS[i], "100%"]), [650, 2600, 1300, 3420, 1100]));
  return out;
}

// ============================================================ LẮP RÁP TÀI LIỆU
const content = body();
const { listPage } = frontMatter();
const pageNumFooter = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 24 })] })] });
const pageProps = (extra = {}) => ({
  page: { size: { width: 11906, height: 16838 }, margin: { top: cm(2), bottom: cm(2), left: cm(3), right: cm(2), footer: cm(1) }, ...extra },
});

const doc = new Document({
  creator: "Nhóm 3", title: "Báo cáo thực hành – Phân bố thống kê trong phân rã phóng xạ",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: SZ } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true }, paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: SZ, bold: true, italics: true }, paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: cm(1.25), hanging: cm(0.5) } } } }] },
    ...["numbers", "steps", "concl"].map((ref) => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: cm(1.25), hanging: cm(0.6) } } } }] })),
  ] },
  sections: [
    { // Bìa – khung viền đôi, không đánh số trang
      properties: { ...pageProps({ borders: {
        pageBorders: { display: PageBorderDisplay.ALL_PAGES, offsetFrom: PageBorderOffsetFrom.PAGE },
        pageBorderTop: { style: BorderStyle.DOUBLE, size: 12, color: "1F3864", space: 24 },
        pageBorderBottom: { style: BorderStyle.DOUBLE, size: 12, color: "1F3864", space: 24 },
        pageBorderLeft: { style: BorderStyle.DOUBLE, size: 12, color: "1F3864", space: 24 },
        pageBorderRight: { style: BorderStyle.DOUBLE, size: 12, color: "1F3864", space: 24 } } }) },
      children: cover(),
    },
    { // Mục lục, danh mục – số trang La Mã
      properties: { ...pageProps({ pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } }), type: SectionType.NEXT_PAGE },
      footers: { default: pageNumFooter },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [t("MỤC LỤC", { bold: true, size: 28 })] }),
        new TableOfContents("Mục lục", { hyperlink: true, headingStyleRange: "1-2" }),
        ...listPage("DANH MỤC BẢNG", LISTS.tables, "Bảng"),
        ...listPage("DANH MỤC HÌNH", LISTS.figures, "Hình"),
      ],
    },
    { // Nội dung – số trang Ả Rập
      properties: { ...pageProps({ pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } }), type: SectionType.NEXT_PAGE },
      footers: { default: pageNumFooter },
      children: content,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("saved", OUT); });

<p align="center">
  <img src="assets/logo.svg" width="88" alt="humanizer-doan-vi logo" />
</p>

<h1 align="center">Humanizer Báo Cáo & Đồ Án Tiếng Việt (KMA Edition)</h1>

<p align="center">
  <strong>Agent hỗ trợ điền nội dung vào file Word khung có sẵn, chuẩn hóa định dạng Đồ án/BTL chuẩn KMA và khử giọng văn AI Tiếng Việt thực nghiệm.</strong><br/>
  Phiên bản plugin: <code>2.7.0</code> &bull; Phiên bản skill: <code>5.8.0-vi-kma-ooxml-safe</code>
</p>

---

## Agent này làm được gì?

Khi viết **Báo cáo Bài tập lớn (BTL)** hoặc **Đồ án tốt nghiệp (ĐATN)** bằng AI, sinh viên thường gặp 4 vấn đề lớn:
1. AI tự bịa bố cục khác hẳn file khung của môn học/trường hoặc tự chế số liệu không có trong tài liệu tham khảo.
2. Giọng văn đọc lên lộ rõ "mùi AI" và văn dịch tiếng Anh (mở đoạn rào đón, câu nào cũng đều tăm tắp 4–5 câu, lặp cụm *"Không chỉ... mà còn..."*, *"đóng vai trò then chốt..."*, lạm dụng bullet in đậm, nhồi từ khóa ở kết luận chương).
3. Sao chép từ AI vào Word bị mất hết Style chuẩn, sai khổ giấy/lề/số trang La Mã–Ả Rập, còn sót ký tự Markdown (`` ` ``, `- `, `**...**`), thậm chí sửa `.docx` bằng script thông thường làm **Microsoft Word báo lỗi hỏng file hoặc ép Save As**.
4. AI hay tự vẽ sơ đồ ASCII/Mermaid xấu hoặc lệch kích thước ảnh thực tế.

**`humanizer-doan-vi`** giải quyết trọn vẹn quy trình đó trong một agent duy nhất:
- **Điền đúng vào file khung `.docx` bạn có sẵn:** Giữ nguyên 100% trang bìa, mục lục tự động (`TOC`), `Section Break` và cây đề mục gốc; chỉ tổng hợp nội dung từ tài liệu đính kèm và báo rõ nếu có mục bị thiếu tài liệu cần bổ sung.
- **Thể hiện góc nhìn sinh viên nghiên cứu:** Phân biệt rõ mô hình/dịch vụ quy mô lớn tham khảo từ tài liệu với công cụ mã nguồn mở thực tế nhóm chọn cho bài thực nghiệm; kết luận chương viết thành nhận định đúc kết chứ không nhồi danh sách từ khóa.
- **Hòa nhập tài liệu tự nhiên:** Cấm viết các câu lộ nguồn rập khuôn (*"Theo tài liệu đính kèm cho biết..."*); mặc định không chèn số trích dẫn `[1]`, `[2]` làm ngắt mạch giữa đoạn văn (trừ khi bạn yêu cầu rõ), tập trung nguồn ở cuối báo cáo.
- **Chuẩn hóa Format Word ĐATN KMA cho cả BTL lẫn ĐATN:** Tự động đưa về khổ `A4`, lề `Trên 2.0 - Dưới 2.0 - Trái 3.0 - Phải 1.5 cm`, `Header/Footer 1.25 cm`, phông `Times New Roman 14pt` (`Code` dùng `Courier New 12pt`), đánh số trang căn giữa (`Bìa` không số, phần đầu `i, ii...`, thân bài `1, 2...`), bảo toàn tuyệt đối cấu trúc **OOXML** để mở và bấm `Ctrl+S` lưu trực tiếp trong Word không lỗi.
- **Xử lý hình ảnh thông minh:** Tuyệt đối không vẽ sơ đồ ASCII/Mermaid. Chưa có ảnh thì chèn khung giữ chỗ sạch (`[KHUNG HÌNH X.Y - CHÈN ẢNH TẠI ĐÂY]`); khi bạn đã dán ảnh chụp thật vào Word, agent xem ảnh đối chiếu tên hình và chạy căn giữa + co giãn tự động theo style `Hnhv` / `Tnhnhv`.
- **Bộ lọc khử giọng AI & văn dịch sâu Tiếng Việt:** Dựa trên dữ liệu thực nghiệm `ViDetect`, `VietAIDetector`, `PhoBERT`, `human`, `viet-style` và `viet-nhu-nguoi` để kiểm nhịp câu (`CV`), phá khuôn 4 bước, bắt văn dịch (*"bạn/của bạn"*, động từ rỗng, từ thừa nghĩa, chuỗi câu cụt, lạm dụng dấu `:`), chuẩn dấu câu (`V01`–`V04`, `H12`) và đối chiếu chống sai/mất số liệu nguồn.

---

## Hai chế độ viết báo cáo

Cả hai chế độ đều dùng **chung chuẩn trình bày Word ĐATN KMA** (`A4`, lề, phông chữ, đánh số trang, bộ Styles), chỉ khác về **giọng văn và yêu cầu nội dung học thuật**:

| Tiêu chí | Bài tập lớn / Báo cáo môn học (`@che-do-btl`) | Đồ án tốt nghiệp / Khóa luận (`@che-do-do-an`) |
|---|---|---|
| **Format trình bày Word** | Chuẩn ĐATN (`A4`, lề `2.0-2.0-3.0-1.5 cm`, TNR `14pt`, Footer `i, ii / 1, 2`) | Chuẩn ĐATN (`A4`, lề `2.0-2.0-3.0-1.5 cm`, TNR `14pt`, Footer `i, ii / 1, 2`) |
| **Chiều sâu & Phạm vi** | Sát kiến thức bài giảng/tài liệu môn học, rõ ràng, gãy gọn | Văn phong học thuật chuẩn mực, phân tích chuyên sâu |
| **Code / Kịch bản thực hành** | Dưới lệnh/đoạn mã quan trọng giải thích rõ cờ, tham số, hàm chính | Mã nguồn dài đưa xuống `Phụ lục`, thân bài tập trung thuật toán & kiến trúc |
| **Ngôi xưng** | Thân bài khách quan / *"ta, chúng ta"*; Mở đầu & Kết luận nhóm được dùng *"nhóm chúng em / nhóm thực hiện"* | Thân bài *"ta / chúng ta / trong phạm vi đồ án này"*; chỉ xưng *"tôi"* ở `Lời cảm ơn` |
| **Yêu cầu riêng** | Có thể trình bày từng bước thực nghiệm + khung kết quả lab | Cuối mỗi chương bắt buộc có `Kết luận chương <số>`; cấm câu cảm tính ở `Kết luận` chung |

---

## Luồng hoạt động trực quan

```text
  [File khung .docx] + [Tài liệu lý thuyết / PDF / Mã nguồn]
                         │
                         ▼
   1. inspect          Đọc cây đề mục có sẵn, kiểm tra format & ảnh hiện có
                         │
                         ▼
   2. Soạn & Biên tập  Bám sát nguồn, khử giọng AI & văn dịch, giữ đúng số liệu
                         │
                         ▼
   3. write-md         Ghi thẳng vào .docx + chuẩn hóa Format ĐATN & OOXML sạch
      (--apply-format)   (chưa có ảnh thật -> đặt khung giữ chỗ Hnhv/Tnhnhv)
                         │
                         ▼
   4. align-images     (Sau khi bạn dán ảnh thật) Căn giữa Hnhv, co giãn <= 15.5x11cm
                         │
                         ▼
   5. check &          Chấm điểm bề mặt (0-100), nhịp câu CV, dấu câu, đối chiếu --source
      validate-xml     và kiểm định toàn vẹn XML để mở/Ctrl+S trong Word an toàn
```

---

## Cài đặt nhanh

### 1. Yêu cầu môi trường
- **Python 3.9+** (toàn bộ chức năng đọc/ghi `.docx`, chuẩn hóa format, căn ảnh, kiểm tra văn phong AI và đối chiếu dữ kiện chỉ dùng **thư viện chuẩn Python**, không bắt buộc cài thêm gói ngoài).
- *(Tùy chọn)* Cài `pymupdf` nếu muốn công cụ đọc trực tiếp file tài liệu nguồn `.pdf`:
  ```powershell
  pip install pymupdf
  ```
- *(Tùy chọn)* Nếu muốn chạy chấm điểm neural cục bộ qua PhoBERT (`vietnamese-ai-detection-v2`), cài thêm `torch` và `transformers` rồi đặt thư mục model cạnh thư mục plugin.

### 2. Cài vào Antigravity / AI Agent
Sao chép toàn bộ thư mục `humanizer-doan-vi` vào thư mục plugin cá nhân của bạn:
- **Windows:** `C:\Users\<Ten_User>\.gemini\config\plugins\humanizer-doan-vi`
- **macOS / Linux:** `~/.gemini/config/plugins/humanizer-doan-vi`

Hoặc clone trực tiếp từ GitHub vào đúng vị trí đó (ví dụ trên Windows):
```powershell
git clone <LINK_REPO_GITHUB> "$env:USERPROFILE\.gemini\config\plugins\humanizer-doan-vi"
```

---

## Hướng dẫn sử dụng trong khung chat

Khi mở workspace chứa file báo cáo và tài liệu của bạn, chỉ cần gõ `@` để gọi rule chế độ kèm tên file là agent tự chạy theo đúng quy trình.

### Tình huống 1: Điền nội dung Bài tập lớn từ file khung & tài liệu đính kèm
```text
@che-do-btl @File_khung_BTL.docx @Slide_ly_thuyet.pdf
Hãy kiểm tra khung đề mục trong File_khung_BTL.docx rồi điền nội dung các mục từ tài liệu đính kèm, giữ chuẩn format và đặt khung giữ chỗ hình ảnh phù hợp.
```

### Tình huống 2: Viết hoặc hoàn thiện từng chương Đồ án tốt nghiệp
```text
@che-do-do-an @Do_an_cua_toi.docx @Tai_lieu_tham_khao.pdf
Hãy điền và hoàn thiện Chương 1 & Chương 2 vào đúng đề mục có sẵn trong file .docx, đảm bảo chuẩn KMA và kiểm tra văn phong AI trước khi hoàn tất.
```

### Tình huống 3: Đã tự chèn ảnh chụp màn hình vào Word, nhờ căn chỉnh và kiểm tra lại tên hình
```text
Tôi vừa dán toàn bộ ảnh chụp thực nghiệm vào file Bao_cao.docx. Hãy xem từng ảnh trong file để kiểm tra tên hình/câu dẫn cho khớp thực tế và chạy căn chỉnh chuẩn Hnhv + kiểm tra lỗi OOXML giúp tôi.
```

### Tình huống 4: Đã có file `.docx` viết sẵn, chỉ muốn chuẩn hóa lại Format ĐATN hoặc khử giọng AI
```text
@che-do-btl Kiểm tra giúp tôi văn phong AI, lỗi dấu câu và chuẩn hóa lại toàn bộ format trang/lề/chân trang ĐATN cho file Bao_cao_nhom.docx mà không làm thay đổi nội dung bài viết.
```

---

## Sử dụng công cụ dòng lệnh (`scripts/da_docx_helper.py`)

Nếu bạn muốn tự chạy kiểm tra hoặc chuẩn hóa thủ công trong terminal (chạy từ bên trong thư mục `humanizer-doan-vi`):

```powershell
# 1. Xem cây đề mục, khổ giấy, căn lề, chân trang, ảnh và thống kê nhịp câu của file .docx
python scripts/da_docx_helper.py inspect "duong_dan/bao_cao.docx"

# 2. Tạo một file .docx mới từ bản mẫu KMA (giữ bìa/mục lục, làm sạch phần nội dung mẫu RSA)
python scripts/da_docx_helper.py clone-template "duong_dan/bao_cao_moi.docx" --clean-body

# 3. Điền nội dung từ file Markdown (.md) vào đúng đề mục có sẵn trong .docx và áp dụng Format ĐATN
python scripts/da_docx_helper.py write-md "duong_dan/bao_cao.docx" "duong_dan/noi_dung.md" --apply-format

# 4. Chuẩn hóa riêng Format ĐATN (khổ A4, lề, chân trang i/ii & 1/2, Times New Roman 14pt, bộ Styles)
python scripts/da_docx_helper.py apply-format "duong_dan/bao_cao.docx"

# 5. Tự động căn giữa style Hnhv và co giãn tỷ lệ toàn bộ ảnh thật trong thân bài (<= 15.5 x 11.0 cm)
python scripts/da_docx_helper.py align-images "duong_dan/bao_cao.docx"

# 6. Chèn lẻ một khung giữ chỗ hình ảnh đóng khung nét đứt (Hnhv + Tnhnhv) dưới một đề mục chỉ định
python scripts/da_docx_helper.py add-placeholder "duong_dan/bao_cao.docx" --fig "2.1" --title "Kiến trúc tổng quan hệ thống" --hint "Mô tả các thành phần cần thể hiện trong hình" --after-heading "2.1"

# 7. Rà soát toàn diện: định dạng ĐATN, OOXML, đề mục, văn dịch sâu, nhịp câu CV & chấm điểm bề mặt 0-100
python scripts/da_docx_helper.py check "duong_dan/bao_cao.docx" --mode btl
python scripts/da_docx_helper.py check "duong_dan/bao_cao.docx" --mode do-an --source "duong_dan/tai_lieu_goc.pdf"

# 8. Kiểm tra nhanh tính toàn vẹn OOXML để chắc chắn Word mở & Ctrl+S lưu trực tiếp không lỗi
python scripts/da_docx_helper.py validate-xml "duong_dan/bao_cao.docx"

# 9. Phân tích chuẩn hóa VietAIDetector, thống kê ViDetect/CV và chấm thử PhoBERT cục bộ (nếu có)
python scripts/da_docx_helper.py score-local "duong_dan/bao_cao.docx"
```

---

## Cấu trúc thư mục

```text
humanizer-doan-vi/
├── plugin.json                        # Thông tin plugin & câu gợi ý khởi chạy nhanh
├── README.md                          # Tài liệu hướng dẫn tổng quan và cách sử dụng
├── SKILL.md                           # Bộ quy tắc cốt lõi cho AI Agent (bản gốc)
├── assets/
│   └── logo.svg                       # Biểu tượng plugin
├── agents/
│   └── viet-luan-da.md                # Cấu hình subagent chuyên biên tập & kiểm định báo cáo
├── rules/
│   ├── chuan-do-an-kma.md             # Rule nền tảng (P1-P6, Format ĐATN, OOXML, Humanizer)
│   ├── che-do-btl.md                  # Rule kích hoạt nhanh bằng @che-do-btl
│   └── che-do-do-an.md                # Rule kích hoạt nhanh bằng @che-do-do-an
├── skills/
│   └── humanizer-doan-vi/
│       └── SKILL.md                   # Skill chính được hệ thống agent nạp khi kích hoạt
├── scripts/
│   └── da_docx_helper.py              # Công cụ Python xử lý .docx/OOXML, hình ảnh & bộ lọc AI
└── templates/
    └── Mau_trinh_bay_DA.docx          # File mẫu trình bày Đồ án tốt nghiệp chuẩn Phụ lục 8 KMA
```

---

## Lưu ý khi sử dụng

- **Đóng file Word trước khi ghi:** Nếu file `.docx` đang mở trong Microsoft Word (có file khóa `~$...`), công cụ sẽ báo dừng an toàn để tránh xung đột ghi file. Hãy đóng Word rồi chạy lại.
- **Ý nghĩa điểm dấu hiệu bề mặt (`0–100`):** Thang điểm chỉ phản ánh mật độ các mẫu diễn đạt khuôn sáo và độ đều nhịp câu (`CV`) trên văn bản để hỗ trợ bạn biên tập tự nhiên hơn, không dùng để kết luận máy móc nguồn gốc văn bản.

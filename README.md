# fetch_pdf.py — đọc PDF tiếng Việt (kể cả font cũ và bản scan)

Script Python một file, trích chữ từ PDF ra Markdown, xử lý các lỗi hay gặp với PDF tiếng Việt:

- **Font cũ** TCVN3/ABC (`.VnTime`…), VNI (`VNI-Times`…), Vn*2 (`VnPalatino2`…): tự nhận bảng mã của từng font và chuyển sang Unicode. Ví dụ `ViÖt Nam`, `Vieät Nam`, `Viïåt Nam` đều ra `Việt Nam`.
- **Lỗi khi xuất PDF từ Office**: `ƣ` thành `ư`, `Ð` thành `Đ`, lỗi UTF-8 kiểu `Viá»‡t` thành `Việt`.
- **Bản scan**: OCR bằng Tesseract (model tiếng Việt), không cần cài Tesseract riêng.
- **Mất khoảng trắng** sau chữ có dấu (`dạdày` thành `dạ dày`), rất hay gặp khi OCR.

## Cài đặt

Cần Python 3.10 trở lên.

**Windows (PowerShell):**

```powershell
curl.exe -L -O https://raw.githubusercontent.com/CT333-boop/vn-pdf-ocr/main/fetch_pdf.py
py -m pip install pymupdf
New-Item -ItemType Directory -Force "$env:LOCALAPPDATA\tessdata" | Out-Null
curl.exe -L -o "$env:LOCALAPPDATA\tessdata\vie.traineddata" https://github.com/tesseract-ocr/tessdata_fast/raw/main/vie.traineddata
```

**macOS / Linux:**

```bash
curl -L -O https://raw.githubusercontent.com/CT333-boop/vn-pdf-ocr/main/fetch_pdf.py
python3 -m pip install pymupdf
mkdir -p ~/.local/share/tessdata
curl -L -o ~/.local/share/tessdata/vie.traineddata https://github.com/tesseract-ocr/tessdata_fast/raw/main/vie.traineddata
```

Dòng đầu tải script về thư mục hiện tại (hoặc `git clone https://github.com/CT333-boop/vn-pdf-ocr`). File `vie.traineddata` (0,5 MB) chỉ cần cho OCR. Nếu chỉ đọc PDF có sẵn chữ thì bỏ bước này. Có thể đặt file ở thư mục khác và trỏ biến môi trường `TESSDATA_PREFIX` tới thư mục đó.

## Dùng

```bash
py fetch_pdf.py "duong-dan/tai-lieu.pdf" ket-qua
py fetch_pdf.py "duong-dan/tai-lieu.pdf" ket-qua --ocr          # OCR các trang scan / trang lỗi font
py fetch_pdf.py "duong-dan/tai-lieu.pdf" ket-qua --ocr-all      # OCR mọi trang
py fetch_pdf.py https://example.com/file.pdf ket-qua --pages 1-20,35
```

(macOS/Linux: thay `py` bằng `python3`.)

Kết quả nằm trong thư mục `ket-qua/`:

| File | Nội dung |
|---|---|
| `text.md` | Toàn bộ chữ, mỗi trang mở đầu bằng `<!-- p. N -->`, trang OCR ghi `<!-- p. N (OCR) -->` |
| `summary.json` | Tóm tắt: số trang, các lỗi đã sửa, trang scan, trang nghi lỗi font, thời gian OCR |
| `meta.json` | Metadata, mục lục, các font cũ phát hiện được |
| `text.raw.md` | Chữ gốc trước khi sửa (chỉ có khi có sửa) |

Cách chạy hợp lý: chạy lần đầu không có `--ocr`. Nếu `summary.json` báo `image_only_pages` (trang scan) hoặc `suspect_pages` (chữ còn lỗi font) thì chạy lại với `--ocr`.

## Độ chính xác và tốc độ

- **OCR**: khoảng 1 giây/trang, sách 145 trang mất khoảng 2,5 phút. Sai khoảng 0,5% ký tự, chủ yếu ở dấu chồng trên nguyên âm mũ (`ể`/`ề`, `ấ`/`ắ`). Câu nào cần trích nguyên văn thì nên so lại với bản gốc.
- **Ngôn ngữ OCR**: mặc định chỉ dùng `vie`, vẫn đọc được chữ tiếng Anh xen kẽ. Không nên dùng `--ocr-lang vie+eng`, vì model tiếng Anh làm mất dấu tiếng Việt.
- **Font cũ**: đã thử trên sách 452 trang dùng font Vn*2, 99,94% từ có dấu ra âm tiết tiếng Việt hợp lệ. Bảng VNI mới thử bằng chuỗi mẫu, chưa có PDF VNI tiếng Việt thật.

## Giấy phép

- `fetch_pdf.py`: MIT (xem [LICENSE](LICENSE)).
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF): AGPL-3.0, hoặc giấy phép thương mại của Artifex.
- [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast): Apache-2.0.

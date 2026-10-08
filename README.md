# Báo cáo chất lượng công việc Pháp chế - Phòng Kiểm soát nội bộ

Ứng dụng Streamlit 1 trang, đọc dữ liệu từ Google Drive (chỉ đọc, không sửa file nguồn).

## Chạy trên máy

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cấu trúc

| File | Vai trò |
|---|---|
| `app.py` | Toàn bộ ứng dụng |
| `requirements.txt` | Thư viện cần cài |
| `.streamlit/config.toml` | Màu giao diện |
| `.streamlit/secrets.toml` | (Không đưa lên GitHub) Khai báo `SHEET_ID` khi chạy local |

## Cấu hình nguồn dữ liệu

Mặc định app dùng ID file được ghi trong `app.py`. Có thể ghi đè bằng secrets (khuyên dùng khi đưa lên GitHub):

```toml
# .streamlit/secrets.toml (local)  hoặc  App settings > Secrets (Streamlit Cloud)
SHEET_ID = "ID_FILE_GOOGLE_DRIVE"
```

File nguồn phải để chế độ **Anyone with the link – Viewer** để app tải được. Nếu không tải được, app hiện ô upload file `.xlsx`.

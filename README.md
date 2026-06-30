# Khao sat du lich cong ty Meiwa

## Chay local

```powershell
cd "D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\khao sat di lich cong ty meiwa"
pip install -r requirements.txt
streamlit run app.py
```

## Cach dung

- Nhap thong tin nhan vien
- Bam `LUU KET QUA KHAO SAT`
- App tu dong:
  - luu vao danh sach trong app
  - cap nhat file tong Excel
  - neu da cau hinh Google Drive thi day len Drive ngay
- Ai can lay file thi bam `TAI FILE TONG`

## Vi tri file tong local

- `exports\Form khao sat Du lich Cong ty meiwa nam 2026.xlsx`

## Cau hinh Google Drive

1. Tao Service Account trong Google Cloud
2. Bat Google Drive API
3. Chia se thu muc Google Drive dich cho email Service Account voi quyen Editor
4. Tao file `.streamlit\secrets.toml` dua theo mau `.streamlit\secrets.toml.example`
5. Dien `folder_id` cua thu muc Google Drive muon luu file tong
6. Neu muon hien link mo file tren app, dien them `shared_url`

## Dua app len Render

App da co san file `render.yaml`, nen co the deploy tren [Render Dashboard](https://dashboard.render.com/) theo cach nay:

1. Dua thu muc app len GitHub
2. Vao Render va chon `New +`
3. Chon `Blueprint`
4. Chon repo GitHub chua app
5. Render se doc file `render.yaml` va tao web service
6. O phan Environment variables, dien neu can:
   - `GOOGLE_DRIVE_FOLDER_ID`
   - `GOOGLE_DRIVE_SHARED_URL`
   - `GOOGLE_SERVICE_ACCOUNT_JSON`
7. Deploy

Co the copy nhanh mau bien moi truong tu file:

- `render-env.example.txt`

Luu y:

- `DATA_DIR=/var/data` da duoc cau hinh san de giu `1 file tong duy nhat`
- Render can `persistent disk` de file tong khong bi mat sau khi restart
- Neu da cau hinh Google Drive, moi lan luu khao sat app se dong bo them len Drive
- `GOOGLE_SERVICE_ACCOUNT_JSON` tren Render phai dan thanh 1 dong JSON

Sau khi len online:

- Dien thoai va may tinh deu mo duoc bang link web
- Bam `LUU KET QUA KHAO SAT` se cap nhat file tong
- Bam `TAI FILE TONG` se tai file ve
- Neu da cau hinh Drive, file tong se dong bo len Google Drive

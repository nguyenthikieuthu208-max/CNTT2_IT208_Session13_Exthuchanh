### Chức năng 1: Terminal

**Một dòng lệnh tạo cây thư mục** (tạo thư mục con thì thư mục gốc tự được tạo):

powershell

```powershell
New-Item -ItemType Directory -Path smartfarm-system/docs, smartfarm-system/scripts, smartfarm-system/data, smartfarm-system/backup
```

**Tạo `.env` bằng đường dẫn tương đối rồi ẩn file:**

powershell

```powershell
cd smartfarm-system/data
Set-Content -Path ..\.env -Value "DB_PASSWORD=demo_123456"
attrib +h ..\.env
```

`..` nghĩa là lên một cấp, tức thư mục gốc. Kiểm tra bằng `Get-ChildItem .. -Force`: file `.env` phải hiện với thuộc tính `h`.

### Chức năng 2: `scripts/auto_report.ps1`

Giả sử `log-he-thong.log` được đặt trong `data/`. Nếu bạn để chỗ khác thì sửa biến `$log`.

powershell

```powershell
# AUDIT: xác nhận đã thẩm định script
$log = "..\data\log-he-thong.log"
$out = "..\backup\bao-cao-su-co.txt"

Write-Output "BAO CAO SU CO SMARTFARM" > $out
Write-Output "Ngay gio: $(Get-Date)" >> $out

Get-Content $log | Select-String -Pattern "CRITICAL|ERROR" >> $out

$dem = (Get-Content $log | Select-String -Pattern "CRITICAL|ERROR" | Measure-Object).Count
Write-Output "TONG SO SU CO: $dem" >> $out
```

Chạy từ thư mục `scripts/`:

powershell

```powershell
.\auto_report.ps1
# nếu bị chặn:
powershell -ExecutionPolicy Bypass -File .\auto_report.ps1
```

Script không có lệnh xóa nào nên qua phần audit.

### Chức năng 3: Excel (sheet `trang-thai-tram`)

Giả sử cột A là `ma_tram`, B là `ma_linh_kien`, C là `dung_luong_ram`, D là `ngay_lap_dat`.

-   **E1** đặt tên `ma_linh_kien_sach`, **E2**: `=TRIM(B2)`, rồi kéo xuống hết dữ liệu.
-   **F1** đặt tên `ten_thiet_bi`, **F2**: `=VLOOKUP(E2,'danh-muc-linh-kien'!$A$2:$D$20,2,FALSE)`, rồi kéo xuống.

Lưu ý:

-   Thay `$D$20` bằng hàng cuối thật của sheet danh mục. Cột A của sheet đó là `ma_linh_kien`, và số `2` là cột `ten_thiet_bi`.
-   Tra theo cột đã TRIM (E) thì mới hết `#N/A`. Nếu vẫn còn `#N/A`, kiểm tra xem mã có sai chính tả hay còn ký tự lạ không.
-   Bật AutoFilter bằng `Ctrl+Shift+L`. Không Merge Cell và không để hàng trống giữa dữ liệu.

### Chức năng 4: `docs/README.md`

Phần bảng phần cứng bên dưới là ví dụ, bạn thay bằng 3 thiết bị thật trong sheet `danh-muc-linh-kien`.

markdown

```markdown
# SmartFarm Monitor

Hệ thống giám sát trạm cảm biến nông trại.

## Hướng dẫn chạy script

```powershell
cd scripts
.\auto_report.ps1
```

## Checklist kiểm tra thiết bị

- [x] Kiểm tra nguồn điện trạm sensor
- [ ] Kiểm tra kết nối mạng
- [ ] Kiểm tra dung lượng lưu trữ

## Bảng thông tin phần cứng

| Thiết bị được đặt tên | Đơn giá | Loại lưu trữ | |--------------|--------------|--------------| | Thiết bị A | 100000 | SSD | | Thiết bị B | 200000 | lái xe chăm chỉ | | Thiết bị C | 150000 | Thẻ SD |





## Phân tích hệ thống

** IPO của chức năng lọc **

- ** Đầu vào: ** file ` log-he-thong.log ` - ** Xử lý: ** ` Get-Content ` | ` Select -String ` filter line TIÊU CHUẨN/LỖI, rồi đếm số dòng - ** Đầu ra : ** file ` bao-cao-su-co.txt ` bao gồm tiêu đề, dòng lỗi và tổng số sự cố 
 


** Phân loại DIKW **

| Phát biểu | Tầng | Lý do | |----------|------|-------| | (b) ` 2026-10-12 08:15:01 [LỖI] cảm biến-03 ... ` | Dữ liệu | Dữ liệu thô, chưa được xử lý | | (c) "Hôm nay ghi nhận 8 sự cố, 4 từ cảm biến-03" | Thông tin | Dữ liệu đã được tổng hợp, có ngữ cảnh | | (d) “sensor-03 hay mất kết nối buổi chiều làm sóng yếu…” | Kiến thức | Rút ra quy luật, nguyên nhân từ thông tin | | (a) "Quyết định thay đổi modem cảm biến-03 trước thứ sáu" | Trí tuệ | Quyết định hành động dựa trên tri thức |

 




## B- phần cứng

File log 500 MB tải qua mạng 100 Mbps:

- Quy đổi: 100 Mbps ` 8 = 12,5 MB/s - Thời gian: 500 MB ` 12,5 MB/s = ** 40 giây **

`                                                                                                                                                                                                                ` `

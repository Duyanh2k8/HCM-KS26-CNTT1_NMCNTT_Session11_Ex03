# Bài 3 : Chuẩn hóa quy trình tra cứu dữ liệu kho hàng
## Phần 1. Phân tích
Hàm VLOOKUP dùng để tìm mã linh kiện trong bảng Danh_Muc, sau đó trả về tên linh kiện hoặc đơn giá tương ứng trong bảng Don_Hang.

Hàm có cú pháp:
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```
Ý nghĩa 4 tham số:
| Tham số |	Ý nghĩa trong bài toán|
|---|---|
| lookup_value |	Mã linh kiện cần tìm, ví dụ ô A2 chứa LK01 trong Sheet Don_Hang.|
| table_array |	Vùng dữ liệu tra cứu trong Sheet Danh_Muc, gồm cột Mã linh kiện, Tên linh kiện và Đơn giá.|
| col_index_num	| Số thứ tự cột cần lấy kết quả trong vùng tra cứu: 2 để lấy Tên linh kiện, 3 để lấy Đơn giá.|
| FALSE |	Yêu cầu tìm mã linh kiện chính xác, không tìm giá trị gần đúng.|
## Phần 2. Triển khai
Giả định cấu trúc dữ liệu:
- Sheet Don_Hang: Ô A2 chứa Mã linh kiện cần tra cứu.
- Sheet Danh_Muc: Bảng tham chiếu nằm trong vùng từ A2 đến C100 (Cột A: Mã linh kiện, Cột B: Tên linh kiện, Cột C: Đơn giá).
### 1.	Công thức lấy Tên linh kiện (Điền tại ô B2 - Sheet Don_Hang)
```excel
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$100, 2, FALSE), "Mã không tồn tại")
```
### 2.	Công thức lấy Đơn giá (Điền tại ô C2 - Sheet Don_Hang)
```excel
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$100, 3, FALSE), 0)
```

# Ứng Dụng Chấm Điểm Trắc Nghiệm Tự Động (OMR Web App)

Ứng dụng web tự động căn chỉnh và chấm điểm phiếu trắc nghiệm dạng bubble từ file PDF nhiều trang.

### Tính năng chính:
- Tự động nắn chỉnh trang scan bị nghiêng dựa trên 4 điểm neo góc.
- Nhận diện Số thứ tự (2 chữ số) và Mã đề.
- Chọn nhanh mốc đề từ 5 đến 50; mỗi mốc có một bộ đáp án 80 câu trong `answers.json`.
- Tự động lấy 80 câu từ bộ đáp án của mốc đã chọn, không cần nhập thủ công.
- Xuất file Excel đính kèm ảnh cắt phần tên/lớp viết tay để đối chiếu.

### Cấu hình đáp án:

Chỉnh sửa `answers.json` để khai báo đáp án A/B/C/D cho từng mốc. Mỗi mốc phải có
đủ 80 câu liên tiếp từ 1 đến 80. Sau khi thay đổi file trong
khi ứng dụng đang chạy, chọn **Clear cache** hoặc khởi động lại Streamlit để nạp lại dữ liệu.

Có thể khai báo hướng dẫn riêng cho từng mốc trong mục `_instructions`; hướng dẫn
sẽ tự động xuất hiện trên sidebar khi giáo viên chọn mốc tương ứng.

Hiện tại đã có đáp án chính thức cho các mốc 5, 10 và 15. Các mốc chưa được cấu hình sẽ bị
khóa chấm và hiển thị cảnh báo trên giao diện.

### Triển khai cục bộ:
```bash
pip install -r requirements.txt
streamlit run app.py

# Ứng Dụng Chấm Điểm Trắc Nghiệm Tự Động (OMR Web App)

Ứng dụng web tự động căn chỉnh và chấm điểm phiếu trắc nghiệm dạng bubble từ file PDF nhiều trang.

### Tính năng chính:
- Tự động nắn chỉnh trang scan bị nghiêng dựa trên 4 điểm neo góc.
- Nhận diện Số thứ tự (2 chữ số) và Mã đề.
- Chấm điểm theo file đáp án (Excel/CSV).
- Xuất file Excel đính kèm ảnh cắt phần tên/lớp viết tay để đối chiếu.

### Triển khai cục bộ:
```bash
pip install -r requirements.txt
streamlit run app.py
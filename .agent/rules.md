# Quy tắc dành cho Agent (Agent Rules)

1. **Xác nhận trước khi xóa (Confirm before deletion)**:
   - Bất cứ khi nào Agent muốn thực hiện thao tác xóa bất kỳ tệp tin, thư mục, dữ liệu lịch sử hoặc tài nguyên nào (bao gồm việc sử dụng lệnh `Remove-Item`, `os.remove()`, `shutil.rmtree()`, v.v.), Agent **bắt buộc phải mô tả rõ mục đích và hỏi ý kiến xác nhận của người dùng**.
   - Agent chỉ được phép thực hiện hành động xóa sau khi người dùng đã phản hồi đồng ý rõ ràng trong cuộc hội thoại.

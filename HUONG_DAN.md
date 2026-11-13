# Phòng làm bài trực tuyến · Vật lí 12 · Chương 2 · Chất khí

Cùng động cơ và cách dùng với phòng khối 10, 11 và phòng Vật lí nhiệt: học sinh vào bằng link, nhập họ tên là làm được, không cần tài khoản. Giáo viên vào `/teacher` bằng mật khẩu để mở câu, đóng câu, xem % đúng và danh sách đúng/sai, tải kết quả `.csv`.

**Hiện phòng đang trống**: chưa có câu hỏi nào, chờ giáo viên đưa bài lên. Phòng này là một dự án riêng, không liên quan và không ảnh hưởng tới các phòng khác đang chạy.

| Trang | Link |
|---|---|
| Phòng làm bài (học sinh) | `/` |
| Giáo viên | `/teacher` |

Khi chưa có câu hỏi, học sinh thấy "Chờ giáo viên mở câu 1…" và giáo viên thấy thông báo chưa có câu hỏi.

## Đưa thêm bài lên
1. Đưa nguồn (trang HTML luyện tập, file Word ngân hàng câu hỏi, hoặc dạng khác) vào thư mục `nguon-goc`.
2. Nhờ Claude chuyển sang `data/cfg.json`, `data/key.json` và `public/img`. **Nội dung giữ nguyên 100%**, chuyển bằng chương trình, không viết lại. Có thể thêm nguồn nhiều lần, mỗi lần một bài hoặc nhiều bài.
3. Kiểm tra: `node tools/kiem_tra.js` (đúng định dạng, đáp án gốc được chấm đúng, đủ ảnh).
4. Thử trong máy rồi mới đưa lên GitHub. Xem chi tiết các bước và định dạng câu hỏi trong `QUY_TRINH_TAO_PHONG_MOI.md`.

Công thức viết dạng `$...$` (ví dụ `$\dfrac{p_1V_1}{T_1}$`) hiển thị bằng KaTeX, hỗ trợ `\dfrac`, `\mathrm`, `\overline`... Nếu thư viện không tải được thì dùng bộ vẽ đơn giản có sẵn.

## Cách 1. Chạy trên laptop của giáo viên (cả lớp cùng Wi-Fi)
1. Cài Node.js (bản LTS) từ nodejs.org, nếu máy chưa có.
2. Bấm đôi `chay-tren-may.bat`, nhập mật khẩu giáo viên (gõ không dấu).
3. Cửa sổ hiện dòng **Cung Wi-Fi: http://192.168.x.x:3000/**. Học sinh mở địa chỉ đó. Giáo viên mở thêm `/teacher`.

## Cách 2. Đưa lên mạng bằng Render (học sinh vào từ bất kì đâu)
1. Tạo kho GitHub riêng cho phòng này (Public, không tick README), ví dụ `Phong-chuong-2-chat-khi`.
2. Trên Render: **Configure account** để cấp quyền cho kho mới, rồi **New, Blueprint**, chọn kho. Render đọc `render.yaml` (dịch vụ `phong-chuong-2-chat-khi`).
3. Ở ô `TEACHER_PASS`, **tự nhập** mật khẩu giáo viên (không ghi vào kho GitHub). Bấm **Apply**.
4. Địa chỉ thường là `https://phong-chuong-2-chat-khi.onrender.com`. Học sinh: địa chỉ đó. Giáo viên: thêm `/teacher`.

### Cần biết khi dùng gói miễn phí của Render
- Máy chủ tự ngủ sau khoảng 15 phút không ai dùng. Mở `/teacher` trước giờ học 5 phút.
- Bài làm lưu trong bộ nhớ máy chủ và mất khi máy chủ ngủ, khởi động lại hoặc cập nhật mã. Cuối buổi bấm **Tải kết quả (.csv)**. Không đẩy bản mới khi học sinh đang làm.
- Tên học sinh được lưu trên máy chủ của Render. Chỉ dùng họ tên và lớp.

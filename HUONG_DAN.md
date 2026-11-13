# Phòng làm bài trực tuyến · Vật lí 12 · Chương 2 · Chất khí

Cùng động cơ và cách dùng với phòng khối 10, 11 và phòng Vật lí nhiệt: học sinh vào bằng link, nhập họ tên là làm được, không cần tài khoản. Giáo viên vào `/teacher` bằng mật khẩu để mở câu, đóng câu, xem % đúng và danh sách đúng/sai, tải kết quả `.csv`.

**Nội dung hiện có**: 265 câu (174 trắc nghiệm, 36 đúng sai, 55 trả lời ngắn) chuyển nguyên văn từ 2 tệp Word nguồn (có bản sao trong `nguon-goc/`): Buổi 1–6 (167 câu, theo thứ tự bộ ôn theo buổi) và 5 nhóm "Ngân hàng · Bài 8–12" (các câu còn lại). Hai câu nguồn ghi "không có đáp án" (B9-TH-I-19, B12-VD-I-03) chưa được đưa vào; xem `tools/bao_cao_chuyen_doi.txt`. Tạo lại dữ liệu: `python tools/tao_du_lieu_docx.py` rồi `node tools/kiem_tra.js`. Phòng này là một dự án riêng, không ảnh hưởng các phòng khác.

| Trang | Link |
|---|---|
| Phòng làm bài (học sinh) | `/` |
| Giáo viên | `/teacher` |

Nhiều câu trong nguồn không có lời giải riêng (chỉ có đáp án); khi đó phần giải thích để trống.

**Mở câu theo dạng hoặc theo bộ**: ở trang giáo viên, dòng "Mở câu theo" chọn *Dạng câu* (3 tab: Trắc nghiệm, Đúng sai, Trả lời ngắn) hoặc *Bộ đề* (11 bộ: Buổi 1–6 và Ngân hàng Bài 8–12, mỗi bộ đủ các dạng câu). Ở chế độ bộ, "Mở câu" mở câu kế tiếp của bộ đang chọn, "Mở hết bộ này" mở cả bộ. Học sinh thấy tab theo đúng chế độ giáo viên đang chọn.

**Mô phỏng**: 39 câu có tình huống cụ thể (bơm bóng, nén pit-tông, nung nóng đẳng áp/đẳng tích, khí thoát qua van, bọt khí nổi lên, chuyển động Brown) có thêm khung mô phỏng chạy được, đặt ngay dưới hình gốc. Danh sách và nhãn ở `tools/mo_phong.json`; mô phỏng chỉ diễn lại quá trình đề đã cho, không theo tỉ lệ, không hiện đại lượng cần tìm. Các câu hỏi khái niệm mà mô phỏng sẽ để lộ đáp án thì không gắn.

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

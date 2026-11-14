# Phòng làm bài trực tuyến · Vật lí 12 · Chương 2 · Chất khí

Cùng động cơ và cách dùng với phòng khối 10, 11 và phòng Vật lí nhiệt: học sinh vào bằng link, nhập họ tên là làm được, không cần tài khoản. Giáo viên vào `/teacher` bằng mật khẩu để mở câu, đóng câu, xem % đúng và danh sách đúng/sai, tải kết quả `.csv`.

**Nội dung hiện có**: 521 câu (329 trắc nghiệm, 81 đúng sai, 111 trả lời ngắn) chuyển nguyên văn từ tệp Word `nguon-goc/CHUONG2_KHI_LI_TUONG_tu_de_thi_tong_hop_DA_GOP_523_cau.docx`, chia thành **26 buổi** đúng như file "CHUONG2_KHI_LI_TUONG_Bo_on_theo_buoi_DU_523_cau.docx": Bài 8 buổi 1–4, Bài 9 buổi 5–9, Bài 10 buổi 10–13, Bài 11 buổi 14–22, Bài 12 buổi 23–26. Mỗi buổi tối đa 18 câu trắc nghiệm, 4 câu đúng sai, 6 câu trả lời ngắn; câu chùm đi cùng nhau. Hai câu nguồn ghi "không có đáp án" (B9-TH-I-34, B12-VD-I-02) chưa được đưa vào; xem `tools/bao_cao_chuyen_doi_523.txt`. Tạo lại dữ liệu: `python -I -X utf8 tools/tao_du_lieu_523.py` rồi `node tools/kiem_tra.js`. (Bộ 265 câu trước đây tạo bằng `tools/tao_du_lieu_docx.py`, vẫn giữ để dùng lại khi cần.) Phòng này là một dự án riêng, không ảnh hưởng các phòng khác.

| Trang | Link |
|---|---|
| Phòng làm bài (học sinh) | `/` |
| Giáo viên | `/teacher` |

Nhiều câu trong nguồn không có lời giải riêng (chỉ có đáp án); khi đó phần giải thích để trống.

**Ba kiểu mở câu** (trang giáo viên, dòng "Mở câu theo"):
1. **Nhóm (buổi) · 26 nhóm**: mỗi buổi là một tab đủ các dạng câu. "Mở câu" mở câu kế tiếp của buổi đang chọn, "Mở hết bộ này" mở cả buổi.
2. **Đề ôn trộn ngẫu nhiên · 18 TN – 4 ĐS – 6 TLN**: bấm "Soạn và mở đề ôn" để bốc ngẫu nhiên một đề từ các bài trong toàn chương, chỉ lấy câu chưa mở (đề sau không trùng đề trước), câu chùm đi đủ cặp. Đề mở cả một lần, học sinh làm theo 3 tab. Phải "Đóng tất cả & hiện đáp án" rồi mới soạn đề tiếp. Muốn dùng lại cả ngân hàng: "Làm lại từ đầu" (xóa kết quả, nhớ tải .csv trước). Ma trận đề ghi ở `matrix` trong `data/cfg.json` (sinh từ `tools/tao_du_lieu_523.py`).
3. **Dạng câu · lộn xộn toàn chương**: 3 tab Trắc nghiệm / Đúng sai / Trả lời ngắn. "Mở 1 câu ngẫu nhiên" và "Mở 5 câu ngẫu nhiên" rút ngẫu nhiên trong toàn chương các câu của tab đang chọn (câu chùm mở cả cặp).

Muốn kết hợp tùy ý (ví dụ chỉ câu đúng sai của Buổi 5): "Cài đặt thêm" → "Mở câu theo nhóm và số lượng". Học sinh thấy tab theo kiểu giáo viên đang chọn.

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

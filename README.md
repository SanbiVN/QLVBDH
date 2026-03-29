# QLVBDH
Tự động hóa gửi Văn bản đi trong QLVBĐH thuộc hệ thống iOffice VNPT cho đơn vị hành chính công

Ứng dụng tự động hóa việc đăng nhập hệ thống iOffice VNPT, thu thập thông tin nhập liệu để nhập nhanh thông tin. Và gửi Đăng ký Văn bản đi lên hệ thống iOffice VNPT cho đơn vị hành chính công. Giúp giảm bớt thời gian tự nhập thao tác tay quá nhiều. 
Ứng dụng sử dụng công nghệ web nhân chromium để duyệt web. Tự động hoàn toàn việc cập nhật Driver cho trình duyệt.
Việc bạn cần làm là kiểm soát quá trình gửi lên, xem xét sai soát thông tin đầu vào dừng chương trình để sử lại thông tin. 


## TẢI XUỐNG
<!-- items that need to be updated release to release -->
[ptUserAddin]: https://github.com/SanbiVN/QLVBDH/releases/download/v1.0/QLVBDH_v1.0.zip
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/QLVBDH/total.svg)](https://github.com/SanbiVN/QLVBDH/releases/) 
 
|  Thông tin   | Tải xuống |
|--------------|-----------|
| QLVBDH v1.0 | [QLVBDH_v1.0.zip][ptUserAddin] | 

<!-- 
[Nhấn tải TaxCodeVN Add-in](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
[![Lượt tải](https://img.shields.io/github/downloads/SanbiVN/TaxCodeVN/total.svg)](https://github.com/SanbiVN/TaxCodeVN/releases/download/tax_code/TaxCodeVN_v4.2.rar) 
 -->
 Ứng dụng chỉ hoạt động trên HĐH Windows 10 trở lên (Không bao gồm máy sử dụng chip ARM). \
(Add-in được khóa mật khẩu VBA, vì lí do bảo mật khi ứng dụng hoạt động ở máy tính Công ty)

## HƯỚNG DẪN TÓM TẮT
 - Dành cho Excel Add-in:
    - **Tải tệp về** > **Giải nén vào thư mục phù hợp** > **Bỏ block tệp (nếu có)**
    - **Nhấn vào tệp QLVBDH.xlsm để mở với Excel** > **Nhập đường dẫn trực thuộc khu vực** > **Nhấn đăng nhập** > **Nhập mật khẩu**
    - **LoadData** (Để tải thông tin danh mục nhập nhanh)
    - Nhập các mục có dấu * là bắt buộc

(hình ảnh cho khung công tác nhập liệu đăng ký văn bản đi)
<img width="960" height="863" alt="qlvbdh1" src="https://github.com/user-attachments/assets/a085864f-35e9-4f70-b2b3-6dfc575c131d" />


## ​HƯỚNG DẪN SỬ DỤNG

<img width="943" height="139" alt="1774782184311" src="https://github.com/user-attachments/assets/7a05d9db-9cd0-4a43-ae89-ba3bb3acd987" />

#### Ứng dụng bao gồm các mục:

- Nhập liên kết URL thuộc vnptioffice.vn và tài khoản VNPT iOffice.
- Đăng nhập: Đăng nhập vào trang để tải lên dữ liệu. *Trình điều khiển tự động kết nối lại với phiên đăng nhập.
- Tải lên: Thực hiện tải lên dữ liệu đến trang web tự động.
- LoadData: Tải thông tin và danh mục từ Web về để nhập liệu nhanh. Sau khi load, các dữ liệu sẽ được lưu vào trang tính, chỉ cần chọn vào dòng nhập, sẽ tự động hiện nhanh danh sách chọn.
- Độ trễ nghỉ giữa mỗi lần gửi lên
- Tạo danh mục nhập nhanh

#### Chỉ thị trong trang tính:

- Dấu * ở tiêu đề cột là yêu cầu nhập.
- Ô có (+) là ô có thể nhập thêm, tách bởi xuống dòng. (Có thể xóa đi nếu muốn chỉ nhập 1 thông tin)


Phiên bản đầu tiên có thể có lỗi nhất định. Nếu gặp lỗi hãy đăng bài bên dưới góp ý, báo lỗi.


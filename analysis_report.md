# BÁO CÁO PHÂN TÍCH DỮ LIỆU CHỈ SỐ DOW JONES (DJI)
**Phương pháp áp dụng: Tiền xử lý dữ liệu, Phân tích thành phần chính (PCA) & Hồi quy tuyến tính (OLS)**

---

## 1. Giới thiệu bài toán
Chỉ số Dow Jones Industrial Average (DJI) là một trong những chỉ số chứng khoán quan trọng nhất toàn cầu, phản ánh sức khỏe của nền kinh tế Mỹ và thị trường tài chính quốc tế. Bài toán đặt ra là phân tích sự tác động của các chỉ số tài chính, kỹ thuật, tỷ giá và các cổ phiếu công nghệ lớn lên biến động của chỉ số DJI (giá đóng cửa - cột `Close`). 

Do số lượng biến độc lập rất lớn (đa cộng tuyến cao), nghiên cứu áp dụng phương pháp **Phân tích thành phần chính (PCA)** để giảm chiều dữ liệu và loại bỏ đa cộng tuyến, sau đó sử dụng các thành phần chính để xây dựng mô hình **Hồi quy tuyến tính OLS** nhằm giải thích sự biến động của chỉ số DJI.

---

## 2. Mô tả dữ liệu ban đầu
Dữ liệu được cung cấp từ tệp `Processed_DJI.csv` bao gồm:
- **Số lượng quan sát**: 1.984 dòng.
- **Số lượng biến ban đầu**: 84 cột.
- **Biến phụ thuộc**: Cột `Close` (Giá đóng cửa của chỉ số DJI).
- **Các nhóm biến độc lập**:
  - Các chỉ số kỹ thuật: `Volume`, `mom` (Momentum), `ROC` (Rate of Change), `EMA` (Exponential Moving Average).
  - Lãi suất và trái phiếu: `DTB3` (3-Month Treasury Bill), `DGS10` (10-Year Treasury Constant Maturity), v.v.
  - Giá hàng hóa: `Oil` (Dầu thô), `Gold` (Vàng).
  - Tỷ giá hối đoái: `GBP`, `JPY`, `CAD`, `CNY`, v.v.
  - Giá cổ phiếu của các công ty công nghệ lớn: `AAPL`, `AMZN`, `MSFT`, v.v.
  - Các chỉ số chứng khoán quốc tế khác: `HSI` (Hang Seng), `FTSE` (Anh), `SSEC` (Thượng Hải), `NYSE`.
  - Các biến hợp đồng tương lai (Futures) kết thúc bằng ký tự `F` (ví dụ: `gold-F`, `copper-F`).

---

## 3. Tiền xử lý dữ liệu
Quy trình tiền xử lý được thực hiện nghiêm ngặt qua các bước:
1. **Loại bỏ các cột kết thúc bằng chữ 'F'**: Theo yêu cầu của đề bài, các biến phái sinh kết thúc bằng 'F' bị xóa bỏ trước khi thực hiện phân tích sâu hơn. Tổng số **17 cột** đã bị loại bỏ (như `gold-F`, `copper-F`, `DJI-F`...). Số cột còn lại sau khi loại bỏ là **67 cột** (bao gồm cả cột `Date` và `Close`).
2. **Chuyển đổi dữ liệu số**: Tất cả các cột trừ cột `Date` được chuyển đổi sang kiểu dữ liệu số thực. Các giá trị lỗi không chuyển đổi được sẽ tự động chuyển thành giá trị khuyết (`NaN`).
3. **Xử lý dữ liệu thiếu bằng Moving Average bậc 5 (MA5)**:
   - Các giá trị khuyết được điền bằng giá trị trung bình trượt của 5 ngày trước đó (`rolling window=5`, `min_periods=1`). Phương pháp này giúp giữ lại xu hướng chuỗi thời gian của dữ liệu tốt hơn việc điền bằng trung bình toàn bộ cột.
   - Đối với các vị trí nằm ở đầu chuỗi (không thể tính toán MA5), giá trị khuyết còn lại được điền bằng giá trị trung bình (`mean`) của chính cột đó.
   - **Kết quả**: Dữ liệu sau xử lý hoàn toàn sạch, số lượng giá trị thiếu giảm về **0**.

---

## 4. Giảm chiều bằng PCA
Để trích xuất thông tin hữu ích và loại bỏ hiện tượng đa cộng tuyến giữa 65 biến độc lập (sau khi đã loại bỏ cột `Date` và biến mục tiêu `Close`), dữ liệu được chuẩn hóa bằng phương pháp `StandardScaler` (đưa về phân phối chuẩn có trung bình bằng 0 và độ lệch chuẩn bằng 1) trước khi thực hiện PCA.

### Bảng phương sai giải thích của các PC đầu tiên:

| Thành phần chính (PC) | Phương sai giải thích riêng lẻ (Explained Variance) | Phương sai tích lũy (Cumulative Variance) |
| :--- | :---: | :---: |
| **PC1** | 18.82% | 18.82% |
| **PC2** | 18.58% | 37.40% |
| **PC3** | 7.87% | 45.27% |
| **PC4** | 5.33% | 50.60% |
| **PC5** | 4.91% | 55.51% |
| **PC6** | 4.39% | 59.90% |
| **PC7** | 3.22% | 63.12% |
| **PC8** | 2.69% | 65.81% |
| **PC9** | 2.34% | 68.15% |
| **PC10** | 2.01% | 70.16% |
| **...** | ... | ... |
| **PC24** | **1.21%** | **90.02%** |

### Chọn số lượng thành phần chính:
Dựa trên tiêu chí phương sai tích lũy giải thích được tối thiểu **90%** tổng phương sai của dữ liệu gốc, chúng ta chọn **24 thành phần chính đầu tiên (PC1 đến PC24)**. Tổng phương sai giải thích tích lũy của 24 PC này đạt **90.02%**.

Biểu đồ Scree dưới đây biểu diễn tỷ lệ phương sai giải thích và phương sai tích lũy của các PC:

![Biểu đồ phương sai giải thích](file:///d:/Hoc. thuc. hanh`/Ths-Phân tích dữ liệu/baitapnhom-phantichdulieu/pca_variance_plot.png)

---

## 5. Xây dựng mô hình hồi quy tuyến tính OLS
Mô hình hồi quy tuyến tính OLS được xây dựng với biến phụ thuộc là chỉ số `Close` và biến độc lập là 24 thành phần chính đã được chuẩn hóa ở trên cùng hệ số tự do (constant).

### Kết quả tóm tắt từ mô hình OLS:
- **Số lượng quan sát**: 1.984
- **Số lượng biến độc lập**: 24 PC + hằng số (`const`)
- **Hệ số xác định $R^2$ (R-squared)**: **0.959 (95.9%)**
- **Hệ số xác định điều chỉnh $Adj. R^2$**: **0.958 (95.8%)**
- **Kiểm định F (F-statistic)**: **1907**
- **Prob (F-statistic)**: **0.00**

### Phân tích và diễn giải kết quả:

1. **Độ phù hợp của mô hình ($R^2$ và $Adj. R^2$)**:
   - Giá trị $R^2 = 0.959$ cho biết **95.9%** biến động của giá đóng cửa DJI được giải thích bởi 24 thành phần chính được trích xuất từ tập dữ liệu gốc. Đây là một con số rất cao, chứng tỏ mô hình có khả năng giải thích và độ phù hợp tuyệt vời đối với dữ liệu thực tế.
   
2. **Ý nghĩa thống kê tổng thể (Prob (F-statistic))**:
   - Giá trị p-value của kiểm định F bằng **0.00** (< 0.05), chứng tỏ mô hình hồi quy có ý nghĩa thống kê cực kỳ cao ở mọi mức ý nghĩa thông thường (1%, 5%). Ta bác bỏ giả thuyết $H_0$ (tất cả các hệ số hồi quy đều bằng 0).

3. **Đánh giá các hệ số hồi quy (Coefficients & P-values)**:
   - **Hằng số tự do (`const`)**: Có giá trị là $15450$ (với $p < 0.001$), đại diện cho giá trị trung bình của chỉ số DJI khi các thành phần chính bằng 0.
   - **Các PC có tác động cùng chiều mạnh nhất**:
     - **PC3**: Hệ số $\beta = 538.07$, $p < 0.001$. Tác động cùng chiều mạnh nhất lên DJI.
     - **PC14**: Hệ số $\beta = 243.04$, $p < 0.001$.
     - **PC13**: Hệ số $\beta = 211.63$, $p < 0.001$.
   - **Các PC có tác động ngược chiều mạnh nhất**:
     - **PC2**: Hệ số $\beta = -797.41$, $p < 0.001$. Tác động ngược chiều mạnh nhất lên chỉ số DJI.
     - **PC9**: Hệ số $\beta = -694.42$, $p < 0.001$.
     - **PC20**: Hệ số $\beta = -495.34$, $p < 0.001$.
     - **PC10**: Hệ số $\beta = -432.36$, $p < 0.001$.
   - **Ý nghĩa thống kê của từng PC**:
     - **22/24** thành phần chính có ý nghĩa thống kê cực kỳ cao ($p$-value < 0.05).
     - Chỉ có **PC8** ($p = 0.184 > 0.05$) và **PC21** ($p = 0.590 > 0.05$) là không có ý nghĩa thống kê ở mức ý nghĩa 5%. Việc phần lớn các PC đều có ý nghĩa thống kê cho thấy các thông tin kỹ thuật, tài chính được nén qua PCA giữ vai trò rất quan trọng trong giải thích biến động của chỉ số Dow Jones.

---

## 6. Kết luận
- **Tiền xử lý thành công**: Dữ liệu thiếu được giải quyết triệt để bằng Moving Average bậc 5 kết hợp trung bình cột, đảm bảo không làm mất các mẫu chuỗi thời gian quan trọng.
- **Hiệu quả của PCA**: Từ 65 biến độc lập có độ tương quan chéo cao (gây ra hiện tượng đa cộng tuyến nghiêm trọng nếu đưa trực tiếp vào hồi quy OLS), PCA đã giảm xuống còn 24 biến độc lập (PC) đại diện mà vẫn giữ lại được tới **90.02%** lượng thông tin của dữ liệu gốc.
- **Sức mạnh giải thích của mô hình**: Mô hình hồi quy OLS sử dụng các PC cho kết quả cực kỳ ấn tượng với $R^2 = 95.9\%$, các kiểm định thống kê đều chứng minh mô hình hoàn toàn tin cậy và có ý nghĩa thực tiễn cao trong việc phân tích cấu trúc biến động của chỉ số Dow Jones.

*Khuyến nghị phát triển*: Trong tương lai, nhóm có thể áp dụng thêm các mô hình dự báo chuỗi thời gian phi tuyến tính như LSTM (Deep Learning) hoặc các mô hình lai (Hybrid models) kết hợp PCA để tăng khả năng dự báo giá đóng cửa của DJI theo thời gian thực.

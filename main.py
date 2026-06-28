import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import statsmodels.api as sm
import matplotlib.pyplot as plt
import os

def run_analysis():
    # 1. Đọc dữ liệu
    csv_path = "Processed_DJI.csv"
    print(f"--- BƯỚC 5: Đọc dữ liệu từ {csv_path} ---")
    df = pd.read_csv(csv_path)
    
    # Ghi thông tin ban đầu ra file
    with open("initial_check.txt", "w", encoding="utf-8") as f:
        f.write("=== THÔNG TIN DỮ LIỆU BAN ĐẦU ===\n")
        f.write(f"Số dòng: {df.shape[0]}, Số cột: {df.shape[1]}\n\n")
        f.write("--- 5 dòng đầu tiên ---\n")
        f.write(df.head().to_string() + "\n\n")
        f.write("--- Thông tin các cột ---\n")
        # Capture df.info()
        import io
        buffer = io.StringIO()
        df.info(buf=buffer)
        f.write(buffer.getvalue() + "\n\n")
        f.write("--- Số lượng giá trị bị thiếu (top 20 cột thiếu nhiều nhất) ---\n")
        f.write(df.isnull().sum().sort_values(ascending=False).head(20).to_string() + "\n")
    
    print("Đã ghi thông tin ban đầu vào file 'initial_check.txt'.")
    print(f"Số dòng ban đầu: {df.shape[0]}, Số cột: {df.shape[1]}")
    
    # 2. Xóa các cột kết thúc bằng chữ F (Futures)
    print("\n--- BƯỚC 6: Xóa các cột kết thúc bằng chữ 'F' ---")
    cols_before = df.shape[1]
    df = df[[col for col in df.columns if not col.endswith("F")]]
    cols_after = df.shape[1]
    print(f"Số cột đã xóa: {cols_before - cols_after}. Số cột còn lại: {cols_after}")
    
    # 3. Xử lý dữ liệu thiếu bằng Moving Average bậc 5
    print("\n--- BƯỚC 7: Xử lý dữ liệu thiếu bằng Moving Average bậc 5 ---")
    # Tách cột Date ra trước khi tính toán
    date_col = None
    if "Date" in df.columns:
        date_col = df["Date"]
        df_numeric = df.drop(columns=["Date"])
    else:
        df_numeric = df.copy()
        
    # Chuyển đổi sang số (coerce lỗi thành NaN)
    for col in df_numeric.columns:
        df_numeric[col] = pd.to_numeric(df_numeric[col], errors="coerce")
        
    # Áp dụng Moving Average bậc 5 (rolling window=5, min_periods=1)
    df_numeric = df_numeric.fillna(df_numeric.rolling(window=5, min_periods=1).mean())
    # Nếu còn thiếu (ở các dòng đầu tiên hoặc cột hoàn toàn rỗng), điền bằng giá trị trung bình cột
    df_numeric = df_numeric.fillna(df_numeric.mean(numeric_only=True))
    # Dự phòng cuối cùng: điền bằng 0
    df_numeric = df_numeric.fillna(0)
    
    # Ghép lại cột Date
    if date_col is not None:
        df = pd.concat([date_col, df_numeric], axis=1)
    else:
        df = df_numeric
        
    print(f"Số lượng giá trị thiếu sau khi xử lý: {df.isnull().sum().sum()}")
    
    # 4. Chọn biến mục tiêu
    # Biến mục tiêu là Close (đại diện cho DJI)
    print("\n--- BƯỚC 8 & 9: Tách biến mục tiêu và ma trận thuộc tính ---")
    if "Close" not in df.columns:
        raise ValueError("Không tìm thấy cột 'Close' trong tập dữ liệu!")
        
    y = df["Close"]
    X = df.drop(columns=["Close"])
    
    if "Date" in X.columns:
        X = X.drop(columns=["Date"])
        
    print(f"Biến mục tiêu y: 'Close' (giá đóng cửa DJI)")
    print(f"Ma trận thuộc tính X có kích thước: {X.shape}")
    
    # 5. Chuẩn hóa dữ liệu
    print("\n--- BƯỚC 10: Chuẩn hóa dữ liệu (StandardScaler) ---")
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # 6. Chạy PCA
    print("\n--- BƯỚC 11: Chạy PCA ---")
    pca = PCA()
    X_pca = pca.fit_transform(X_scaled)
    
    explained_var = pca.explained_variance_ratio_
    cum_var = explained_var.cumsum()
    
    pca_table = pd.DataFrame({
        "PC": [f"PC{i+1}" for i in range(len(explained_var))],
        "Explained Variance": explained_var,
        "Cumulative Variance": cum_var
    })
    
    # Ghi bảng kết quả PCA ra file
    pca_table.to_csv("pca_variance_table.csv", index=False)
    with open("pca_variance_table.txt", "w", encoding="utf-8") as f:
        f.write("=== BẢNG PHƯƠNG SAI GIẢI THÍCH CỦA CÁC THÀNH PHẦN CHÍNH (PCA) ===\n")
        f.write(pca_table.to_string(index=False) + "\n")
        
    print("Đã ghi kết quả PCA vào file 'pca_variance_table.txt' và 'pca_variance_table.csv'.")
    print("Top 15 Principal Components:")
    print(pca_table.head(15))
    
    # 7. Chọn số lượng thành phần chính giải thích ít nhất 90% phương sai tích lũy
    print("\n--- BƯỚC 12: Chọn số lượng thành phần chính ---")
    # Số thành phần chính có phương sai tích lũy < 90%
    n_components = (pca_table["Cumulative Variance"] < 0.90).sum() + 1
    cum_explained = cum_var[n_components - 1]
    print(f"Số thành phần chính giải thích được ít nhất 90% phương sai tích lũy: {n_components}")
    print(f"Tổng phương sai tích lũy được giải thích bởi {n_components} PC đầu tiên: {cum_explained:.4f} ({cum_explained*100:.2f}%)")
    
    # Vẽ biểu đồ Scree và lưu
    plt.figure(figsize=(10, 6))
    plt.plot(range(1, len(explained_var) + 1), cum_var, marker='o', color='b', label='Cumulative Variance')
    plt.bar(range(1, len(explained_var) + 1), explained_var, alpha=0.5, color='g', align='center', label='Individual Variance')
    plt.axhline(y=0.90, color='r', linestyle='--', label='Ngưỡng 90%')
    plt.axvline(x=n_components, color='m', linestyle='--', label=f'Số PC đã chọn: {n_components}')
    plt.title('Biểu đồ phương sai giải thích và phương sai tích lũy của các PC', fontsize=14)
    plt.xlabel('Thành phần chính (PC)', fontsize=12)
    plt.ylabel('Tỷ lệ phương sai giải thích', fontsize=12)
    plt.xlim(0.5, min(len(explained_var) + 0.5, 30))  # Giới hạn hiển thị 30 PC đầu tiên để biểu đồ rõ ràng hơn
    plt.ylim(0, 1.05)
    plt.grid(True, linestyle=':', alpha=0.6)
    plt.legend(loc='best')
    plt.tight_layout()
    plt.savefig("pca_variance_plot.png", dpi=300)
    plt.close()
    print("Đã vẽ và lưu biểu đồ phương sai giải thích: 'pca_variance_plot.png'")
    
    # 8. Chạy hồi quy OLS
    print("\n--- BƯỚC 13: Chạy hồi quy OLS ---")
    X_model = X_pca[:, :n_components]
    X_model = sm.add_constant(X_model)  # Thêm hằng số tự do beta_0
    
    # Tạo danh sách tên biến cho mô hình hồi quy (Hằng số + PC1, PC2, ...)
    column_names = ["const"] + [f"PC{i+1}" for i in range(n_components)]
    
    model = sm.OLS(y, X_model).fit()
    
    # In ra terminal và lưu tóm tắt mô hình
    summary_str = model.summary(xname=column_names).as_text()
    print(summary_str)
    
    with open("ols_regression_results.txt", "w", encoding="utf-8") as f:
        f.write("=== KẾT QUẢ HỒI QUY TUYẾN TÍNH OLS TRÊN CÁC THÀNH PHẦN CHÍNH ===\n")
        f.write(f"Số lượng PC được sử dụng: {n_components}\n")
        f.write(summary_str + "\n")
        
    print("\nĐã lưu kết quả hồi quy OLS vào file 'ols_regression_results.txt'.")
    print("Toàn bộ phân tích hoàn tất thành công!")

if __name__ == "__main__":
    run_analysis()

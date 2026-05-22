import os
import sys
import pandas as pd
import math

# Configure stdout to use UTF-8 to prevent encoding errors on Windows console
sys.stdout.reconfigure(encoding='utf-8')

def split_files():
    # Thư mục chứa dữ liệu đầu vào và đầu ra
    base_dir = r"d:\Downloads\takeout-20260430T043432Z-3-001\data du lieu"
    output_dir = os.path.join(base_dir, "splitted_150")
    
    # Tạo thư mục đầu ra nếu chưa có
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 Đã tạo thư mục đầu ra: {output_dir}")
        
    chunk_size = 150

    # 1. Xử lý chia nhỏ tệp Excel KoCham_Company_List.xlsx
    xlsx_path = os.path.join(base_dir, "KoCham_Company_List.xlsx")
    if os.path.exists(xlsx_path):
        print(f"\n📖 Đang đọc tệp Excel: {xlsx_path} ...")
        df_xlsx = pd.read_excel(xlsx_path)
        total_rows = len(df_xlsx)
        num_parts = math.ceil(total_rows / chunk_size)
        print(f"📊 Tìm thấy {total_rows} công ty. Sẽ chia làm {num_parts} phần (mỗi phần tối đa {chunk_size} dòng).")
        
        for i in range(num_parts):
            start_idx = i * chunk_size
            end_idx = min(start_idx + chunk_size, total_rows)
            chunk_df = df_xlsx.iloc[start_idx:end_idx]
            
            part_num = i + 1
            part_filename = f"KoCham_Company_List_Part_{part_num}.xlsx"
            part_path = os.path.join(output_dir, part_filename)
            
            chunk_df.to_excel(part_path, index=False)
            print(f"   💾 Đã lưu: {part_filename} (Dòng {start_idx + 1} -> {end_idx})")
    else:
        print(f"❌ Không tìm thấy tệp Excel tại: {xlsx_path}")

    # 2. Xử lý chia nhỏ tệp CSV Book1.csv (đọc với bảng mã CP1258)
    csv_path = os.path.join(base_dir, "Book1.csv")
    if os.path.exists(csv_path):
        print(f"\n📖 Đang đọc tệp CSV: {csv_path} ...")
        df_csv = pd.read_csv(csv_path, encoding='cp1258')
        total_rows = len(df_csv)
        num_parts = math.ceil(total_rows / chunk_size)
        print(f"📊 Tìm thấy {total_rows} công ty trong CSV. Sẽ chia làm {num_parts} phần.")
        
        for i in range(num_parts):
            start_idx = i * chunk_size
            end_idx = min(start_idx + chunk_size, total_rows)
            chunk_df = df_csv.iloc[start_idx:end_idx]
            
            part_num = i + 1
            part_filename = f"Book1_Part_{part_num}.csv"
            part_path = os.path.join(output_dir, part_filename)
            
            # Lưu dạng CSV với bảng mã CP1258 để giữ nguyên định dạng tiếng Việt của bạn
            chunk_df.to_csv(part_path, index=False, encoding='cp1258')
            print(f"   💾 Đã lưu: {part_filename} (Dòng {start_idx + 1} -> {end_idx})")
    else:
        print(f"❌ Không tìm thấy tệp CSV tại: {csv_path}")

    print("\n🎉 Hoàn thành việc phân tách tất cả các tệp thành công!")

if __name__ == "__main__":
    split_files()

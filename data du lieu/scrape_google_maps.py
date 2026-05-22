import os
import sys
import re
import time
import random
import urllib.parse
import pandas as pd
from playwright.sync_api import sync_playwright

# Đảm bảo in tiếng Việt ra console không bị lỗi mã hóa (UTF-8)
sys.stdout.reconfigure(encoding='utf-8')

# Đường dẫn file Excel
EXCEL_PATH = r"d:\Downloads\takeout-20260430T043432Z-3-001\data du lieu\KoCham_Company_List.xlsx"
SAVE_INTERVAL = 5  # Lưu định kỳ sau mỗi 5 doanh nghiệp thành công
HEADLESS = True    # Để False nếu bạn muốn mở trình duyệt lên xem thực tế chạy

def log_msg(msg):
    """In log có định dạng thời gian trực quan"""
    current_time = time.strftime("%H:%M:%S")
    print(f"[{current_time}] {msg}")

def safe_save(df, path):
    """Lưu file Excel an toàn bằng cách ghi ra file tạm rồi ghi đè để tránh bị hỏng file nếu dừng đột ngột"""
    temp_path = path + ".tmp"
    try:
        df.to_excel(temp_path, index=False)
        if os.path.exists(temp_path):
            if os.path.exists(path):
                os.remove(path)
            os.rename(temp_path, path)
            log_msg("💾 Đã lưu dữ liệu định kỳ vào file Excel thành công!")
    except Exception as e:
        log_msg(f"❌ Lỗi khi ghi file Excel: {e}")
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def clean_rating_and_reviews(text):
    """
    Trích xuất số sao và lượt đánh giá từ chuỗi văn bản của Google Maps.
    Ví dụ:
    - "4,5(12)" -> rating=4.5, reviews=12
    - "4.8 (1,234)" -> rating=4.8, reviews=1234
    - "4.2(1.150 đánh giá)" -> rating=4.2, reviews=1150
    """
    if not text:
        return None, None
    
    try:
        parts = text.split("(")
        rating_str = parts[0].strip().replace(",", ".")
        rating = float(rating_str)
        
        reviews_num = 0
        if len(parts) > 1:
            reviews_str = parts[1].split(")")[0]
            # Loại bỏ toàn bộ ký tự không phải số (ví dụ dấu chấm phân cách hàng nghìn hoặc chữ "đánh giá")
            clean_reviews = re.sub(r"\D", "", reviews_str)
            if clean_reviews:
                reviews_num = int(clean_reviews)
                
        return rating, reviews_num
    except Exception as e:
        log_msg(f"⚠️ Không thể parse text '{text}': {e}")
        return None, None

def scrape_google_maps():
    # 1. Đọc file Excel bằng Pandas
    if not os.path.exists(EXCEL_PATH):
        log_msg(f"❌ Không tìm thấy file Excel tại đường dẫn: {EXCEL_PATH}")
        return
    
    log_msg(f"📖 Đang đọc file Excel: {EXCEL_PATH} ...")
    df = pd.read_excel(EXCEL_PATH)
    
    # Chuẩn hóa tên cột (loại bỏ khoảng trắng thừa nếu có)
    df.columns = [col.strip() for col in df.columns]
    
    # Đảm bảo các cột kết quả tồn tại trong DataFrame
    required_cols = ["Google Maps Link", "Rating Google Maps", "Số Lượt Đánh Giá"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
            
    total_rows = len(df)
    log_msg(f"📊 Tổng số doanh nghiệp trong danh sách: {total_rows}")
    
    # Lọc danh sách cần xử lý (những dòng CHƯA CÓ cả Link và Rating để hỗ trợ Resume)
    pending_rows = df[df["Rating Google Maps"].isna() | df["Google Maps Link"].isna()]
    log_msg(f"🔄 Số doanh nghiệp cần cào (chưa có kết quả): {len(pending_rows)}")
    
    if len(pending_rows) == 0:
        log_msg("🎉 Tất cả doanh nghiệp đã được cào dữ liệu hoàn tất!")
        return

    # 2. Khởi chạy Playwright
    with sync_playwright() as p:
        log_msg("🚀 Đang khởi động trình duyệt...")
        browser = p.chromium.launch(
            headless=HEADLESS, 
            args=["--no-sandbox", "--disable-setuid-sandbox", "--lang=vi-VN"]
        )
        
        # Tạo context với ngôn ngữ tiếng Việt và User Agent phổ biến để định vị chuẩn hơn
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            locale="vi-VN",
            viewport={"width": 1280, "height": 800}
        )
        
        page = context.new_page()
        success_count = 0
        
        # Lặp qua các dòng cần xử lý
        for idx, row in pending_rows.iterrows():
            company_name = str(row["Tên Công Ty"]).strip()
            city = str(row["Thành Phố"]).strip() if pd.notna(row["Thành Phố"]) else ""
            address = str(row["Địa Chỉ"]).strip() if pd.notna(row["Địa Chỉ"]) else ""
            
            if not company_name or company_name.lower() == "nan":
                log_msg(f"⚠️ Bỏ qua dòng {idx + 1}: Tên công ty bị rỗng.")
                continue
                
            # Tạo chuỗi truy vấn tìm kiếm tối ưu
            # Kết hợp: Tên công ty + Thành phố + Vietnam để Google Maps khoanh vùng chuẩn xác
            search_query = f"{company_name} {city} Vietnam"
            encoded_query = urllib.parse.quote_plus(search_query)
            search_url = f"https://www.google.com/maps/search/{encoded_query}"
            
            log_msg(f"🔎 [{idx + 1}/{total_rows}] Đang tìm kiếm: '{company_name}' tại {city}...")
            
            try:
                # Điều hướng tới URL tìm kiếm Google Maps
                page.goto(search_url, timeout=30000)
                
                # Chờ xem kết quả tải ra:
                # 1. Nếu mở trực tiếp chi tiết: div.F7nice (rating) hoặc h1.DUwDvf (tiêu đề địa điểm)
                # 2. Nếu trả về danh sách kết quả: a.hfpxzc hoặc a[href*="/maps/place/"]
                # Chờ tối đa 8 giây cho bất kỳ dấu hiệu nào xuất hiện
                try:
                    page.wait_for_selector(
                        'div.F7nice, h1.DUwDvf, a.hfpxzc, a[href*="/maps/place/"]', 
                        timeout=8000
                    )
                except Exception:
                    # Timeout - Không tìm thấy bất kỳ dấu hiệu nào (Không có kết quả)
                    log_msg(f"❌ Không tìm thấy kết quả nào cho: '{company_name}'")
                    df.at[idx, "Rating Google Maps"] = "Không tìm thấy"
                    df.at[idx, "Số Lượt Đánh Giá"] = 0
                    df.at[idx, "Google Maps Link"] = "Không tìm thấy"
                    success_count += 1
                    continue
                
                # Kiểm tra xem có phải danh sách kết quả hay không
                list_selector = 'a.hfpxzc, a[href*="/maps/place/"]'
                list_elements = page.locator(list_selector)
                
                if list_elements.count() > 0:
                    # Nếu là danh sách, click vào phần tử đầu tiên để xem chi tiết
                    log_msg("ℹ️ Tìm thấy danh sách kết quả. Đang chọn kết quả đầu tiên...")
                    list_elements.first.click()
                    # Chờ trang chi tiết tải ra
                    try:
                        page.wait_for_selector('h1.DUwDvf, div.F7nice', timeout=5000)
                    except:
                        pass
                
                # Đợi ổn định trang 1 giây
                page.wait_for_timeout(1000)
                
                # Trích xuất dữ liệu từ trang chi tiết
                # 1. Điểm đánh giá và Số lượt đánh giá
                rating_val = None
                reviews_val = 0
                
                rating_locator = page.locator('div.F7nice')
                if rating_locator.count() > 0:
                    rating_text = rating_locator.first.inner_text().strip()
                    if rating_text:
                        rating_val, reviews_val = clean_rating_and_reviews(rating_text)
                
                # 2. URL Google Maps chính xác của địa điểm
                final_url = page.url
                
                if rating_val is not None:
                    log_msg(f"✅ Đã tìm thấy: {rating_val} ⭐ | {reviews_val} lượt đánh giá")
                    df.at[idx, "Rating Google Maps"] = rating_val
                    df.at[idx, "Số Lượt Đánh Giá"] = reviews_val
                else:
                    log_msg("⚠️ Địa điểm tồn tại nhưng chưa có lượt đánh giá nào (0 ⭐).")
                    df.at[idx, "Rating Google Maps"] = 0.0
                    df.at[idx, "Số Lượt Đánh Giá"] = 0
                
                df.at[idx, "Google Maps Link"] = final_url
                success_count += 1
                
            except Exception as e:
                log_msg(f"❌ Lỗi khi xử lý dòng {idx + 1} ({company_name}): {e}")
                
            # Nghỉ ngẫu nhiên từ 1 đến 2.5 giây để tránh bị Google chặn (Polite Crawling)
            sleep_time = random.uniform(1.0, 2.5)
            time.sleep(sleep_time)
            
            # Lưu định kỳ (Periodic Save)
            if success_count % SAVE_INTERVAL == 0 and success_count > 0:
                safe_save(df, EXCEL_PATH)
                
        # Lưu lần cuối khi hoàn thành vòng lặp
        safe_save(df, EXCEL_PATH)
        log_msg("🎉 Đã quét xong toàn bộ danh sách!")
        browser.close()

if __name__ == "__main__":
    scrape_google_maps()

import os
import sys
import re
import time
import random
import urllib.parse
import pandas as pd
from playwright.sync_api import sync_playwright

# Configure stdout to use UTF-8 to prevent encoding errors on Windows console
sys.stdout.reconfigure(encoding='utf-8')

SAVE_INTERVAL = 1  # Save after EACH company to see immediate progress
HEADLESS = True    # Running in headless mode

def log_msg(msg):
    """Print message with timestamp"""
    current_time = time.strftime("%H:%M:%S")
    print(f"[{current_time}] {msg}")
    sys.stdout.flush()  # Flush output immediately

def safe_save(df, path):
    """Save Excel file safely by writing to a temp file first"""
    temp_path = path + ".tmp.xlsx"
    try:
        df.to_excel(temp_path, index=False)
        if os.path.exists(temp_path):
            if os.path.exists(path):
                os.remove(path)
            os.rename(temp_path, path)
            log_msg(f"💾 Đã cập nhật file Excel thành công!")
    except Exception as e:
        log_msg(f"❌ Lỗi khi ghi file Excel: {e}")
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

def clean_rating_and_reviews(text):
    """Extract rating and review count from Google Maps text"""
    if not text:
        return None, None
    try:
        parts = text.split("(")
        rating_str = parts[0].strip().replace(",", ".")
        rating = float(rating_str)
        
        reviews_num = 0
        if len(parts) > 1:
            reviews_str = parts[1].split(")")[0]
            clean_reviews = re.sub(r"\D", "", reviews_str)
            if clean_reviews:
                reviews_num = int(clean_reviews)
        return rating, reviews_num
    except Exception as e:
        log_msg(f"⚠️ Không thể parse text '{text}': {e}")
        return None, None

def scrape_part(part_num):
    base_dir = r"d:\Downloads\takeout-20260430T043432Z-3-001\data du lieu"
    part_filename = f"KoCham_Company_List_Part_{part_num}.xlsx"
    excel_path = os.path.join(base_dir, "splitted_150", part_filename)
    
    if not os.path.exists(excel_path):
        log_msg(f"❌ Không tìm thấy file Excel tại đường dẫn: {excel_path}")
        return
    
    log_msg(f"📖 Đang đọc tệp: {part_filename} ...")
    df = pd.read_excel(excel_path)
    
    # Strip column names
    df.columns = [col.strip() for col in df.columns]
    
    required_cols = ["Google Maps Link", "Rating Google Maps", "Số Lượt Đánh Giá"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
            
    total_rows = len(df)
    log_msg(f"📊 Tổng số doanh nghiệp trong phần {part_num}: {total_rows}")
    
    # Select rows that don't have ratings or Google Maps Link yet
    pending_rows = df[df["Rating Google Maps"].isna() | df["Google Maps Link"].isna() | (df["Google Maps Link"] == "")]
    log_msg(f"🔄 Số doanh nghiệp cần cào: {len(pending_rows)}")
    
    if len(pending_rows) == 0:
        log_msg(f"🎉 Tất cả doanh nghiệp trong phần {part_num} đã được cào hoàn tất!")
        return

    with sync_playwright() as p:
        log_msg("🚀 Đang khởi động trình duyệt Chromium...")
        browser = p.chromium.launch(
            headless=HEADLESS, 
            args=["--no-sandbox", "--disable-setuid-sandbox", "--lang=vi-VN"]
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            locale="vi-VN",
            viewport={"width": 1280, "height": 800}
        )
        
        page = context.new_page()
        success_count = 0
        
        for idx, row in pending_rows.iterrows():
            company_name = str(row["Tên Công Ty"]).strip()
            city = str(row["Thành Phố"]).strip() if pd.notna(row["Thành Phố"]) else ""
            
            if not company_name or company_name.lower() == "nan":
                log_msg(f"⚠️ Bỏ qua dòng {idx + 1}: Tên công ty bị rỗng.")
                continue
                
            search_query = f"{company_name} {city} Vietnam"
            encoded_query = urllib.parse.quote_plus(search_query)
            search_url = f"https://www.google.com/maps/search/{encoded_query}"
            
            log_msg(f"🔎 [{idx + 1}/{total_rows}] Đang tìm kiếm: '{company_name}' tại {city}...")
            
            try:
                page.goto(search_url, timeout=30000)
                
                try:
                    page.wait_for_selector(
                        'div.F7nice, h1.DUwDvf, a.hfpxzc, a[href*="/maps/place/"]', 
                        timeout=8000
                    )
                except Exception:
                    log_msg(f"❌ Không tìm thấy kết quả cho: '{company_name}'")
                    df.at[idx, "Rating Google Maps"] = "Không tìm thấy"
                    df.at[idx, "Số Lượt Đánh Giá"] = 0
                    df.at[idx, "Google Maps Link"] = "Không tìm thấy"
                    success_count += 1
                    safe_save(df, excel_path)
                    continue
                
                list_selector = 'a.hfpxzc, a[href*="/maps/place/"]'
                list_elements = page.locator(list_selector)
                
                if list_elements.count() > 0:
                    log_msg("ℹ️ Tìm thấy danh sách kết quả. Chọn kết quả đầu tiên...")
                    list_elements.first.click()
                    try:
                        page.wait_for_selector('h1.DUwDvf, div.F7nice', timeout=5000)
                    except:
                        pass
                
                page.wait_for_timeout(1000)
                
                rating_val = None
                reviews_val = 0
                
                rating_locator = page.locator('div.F7nice')
                if rating_locator.count() > 0:
                    rating_text = rating_locator.first.inner_text().strip()
                    if rating_text:
                        rating_val, reviews_val = clean_rating_and_reviews(rating_text)
                
                final_url = page.url
                
                if rating_val is not None:
                    log_msg(f"✅ Đã tìm thấy: {rating_val} ⭐ | {reviews_val} lượt đánh giá")
                    df.at[idx, "Rating Google Maps"] = rating_val
                    df.at[idx, "Số Lượt Đánh Giá"] = reviews_val
                else:
                    log_msg("⚠️ Địa điểm chưa có lượt đánh giá (0 ⭐).")
                    df.at[idx, "Rating Google Maps"] = 0.0
                    df.at[idx, "Số Lượt Đánh Giá"] = 0
                
                df.at[idx, "Google Maps Link"] = final_url
                success_count += 1
                
                # Save immediately
                safe_save(df, excel_path)
                
            except Exception as e:
                log_msg(f"❌ Lỗi khi xử lý dòng {idx + 1} ({company_name}): {e}")
                
            sleep_time = random.uniform(1.0, 2.0)
            time.sleep(sleep_time)
            
        log_msg(f"🎉 Đã quét xong toàn bộ phần {part_num}!")
        browser.close()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            p_num = int(sys.argv[1])
            scrape_part(p_num)
        except Exception as e:
            print(f"Lỗi tham số: {e}")
    else:
        print("Vui lòng nhập số phần cần cào. Ví dụ: python scrape_split_part.py 1")

from playwright.sync_api import sync_playwright
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Starting Playwright test...")
try:
    with sync_playwright() as p:
        print("Launching browser...")
        browser = p.chromium.launch(headless=True)
        print("Browser launched successfully!")
        page = browser.new_page()
        print("Navigating to Google Maps...")
        page.goto("https://www.google.com/maps/search/ADD+Group+JSC+Hanoi+Vietnam", timeout=20000)
        print("Page loaded. Title:", page.title())
        browser.close()
    print("SUCCESS: Playwright works!")
except Exception as e:
    print(f"FAILED: {e}")

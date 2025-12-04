from playwright.sync_api import sync_playwright, expect
import os

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("http://localhost:5000")

        # Verify initial state
        expect(page.get_by_text("Content Source")).to_be_visible()
        expect(page.get_by_label("Web Search")).to_be_checked()

        # Click Upload Files radio
        page.get_by_label("Upload Files").click()

        # Verify file inputs appear
        expect(page.get_by_text("Upload Content Files (TXT, CSV, Excel)")).to_be_visible()
        expect(page.locator("#contentFile")).to_be_visible()

        # Using exact text match or locator since label might be ambiguous
        expect(page.locator("#fileTopic")).to_be_visible()

        # Verify Chart Data input exists
        expect(page.get_by_text("Chart Data (Optional)")).to_be_visible()

        # Take screenshot
        page.screenshot(path="frontend_verification/upload_ui.png")
        print("Screenshot saved to frontend_verification/upload_ui.png")

        browser.close()

if __name__ == "__main__":
    run()

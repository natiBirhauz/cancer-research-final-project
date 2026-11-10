"""
take_gui_screenshots.py
-----------------------
Automated screenshot generator for the Surprisal Analysis (SA) – Minimal ES Discovery Platform.
Uses Playwright to capture high-resolution screenshots of:
1. Step 1 (English): Overview, Illustration, Guide, Templates, Start Button
2. Step 1 (Hebrew): Overview in Hebrew, Methodology Illustration, RTL Layout
3. Step 2 (English): Pre-Flight Dataset Inspection & Execution Parameters with (i) Tooltips
4. Step 2 (Hebrew): Hebrew Pre-Flight Check & Parameters
5. Step 2 (English): Execution Results - Summary KPI Cards, ES Bar Chart, Conflict-Free Heatmap, Deliverables
6. Step 2 (English): Interactive Output Tables (Optimal_ES_SetCover, Per_Dataset_Coverage_and_R2)
"""

import os
import glob
import time
from playwright.sync_api import sync_playwright
from web_gui import create_web_application

def capture_screenshots():
    out_dir = os.path.join("Docs", "screenshots")
    os.makedirs(out_dir, exist_ok=True)

    print("1. Launching Gradio Web GUI server on port 7865...")
    app = create_web_application()
    app.launch(
        server_name="127.0.0.1",
        server_port=7865,
        prevent_thread_lock=True,
        inbrowser=False
    )
    time.sleep(2)

    eval_files = glob.glob(os.path.abspath('Data for algorithm evaluation/*/data.xlsx'))
    if not eval_files:
        eval_files = [os.path.abspath('Sample_Input_Template.xlsx')]

    print("2. Launching Playwright Chrome browser...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            headless=True
        )
        context = browser.new_context(
            viewport={'width': 1280, 'height': 960},
            device_scale_factor=1.5
        )
        page = context.new_page()
        page.goto("http://127.0.0.1:7865/?__theme=light", wait_until="domcontentloaded")
        page.wait_for_selector(".sa-header", timeout=15000)
        time.sleep(2)

        # 1. Step 1 English Screenshot
        print("   Capturing: 01_step1_overview_en.png...")
        path_1 = os.path.join(out_dir, "01_step1_overview_en.png")
        page.screenshot(path=path_1, full_page=True)

        # 2. Step 1 Hebrew Screenshot
        print("   Switching language to Hebrew...")
        page.locator('label').filter(has_text="עברית (Hebrew)").click()
        time.sleep(2)
        print("   Capturing: 02_step1_overview_he.png...")
        path_2 = os.path.join(out_dir, "02_step1_overview_he.png")
        page.screenshot(path=path_2, full_page=True)

        # 3. Switch back to English and proceed to Step 2
        print("   Switching back to English and proceeding to Step 2...")
        page.locator('label').filter(has_text="English (Default)").click()
        time.sleep(1)
        page.locator('button').filter(has_text="Start Analysis (Proceed to Step 2)").click()
        time.sleep(1)

        # 4. Upload dataset files
        print(f"   Uploading {len(eval_files)} dataset files for Pre-Flight Check...")
        file_input = page.locator('input[type="file"]')
        file_input.set_input_files(eval_files)
        
        # Wait for precheck to compute
        print("   Waiting for Pre-Flight Dataset Inspection to complete...")
        page.wait_for_selector('text=Pre-Flight Dataset Inspection Complete', timeout=20000)
        time.sleep(1)

        print("   Capturing: 03_step2_precheck_en.png...")
        path_3 = os.path.join(out_dir, "03_step2_precheck_en.png")
        page.screenshot(path=path_3, full_page=False)

        # 5. Capture Pre-Check in Hebrew
        print("   Capturing: 04_step2_precheck_he.png...")
        page.locator('label').filter(has_text="עברית (Hebrew)").click()
        time.sleep(1.5)
        path_4 = os.path.join(out_dir, "04_step2_precheck_he.png")
        page.screenshot(path=path_4, full_page=False)

        # 6. Switch back to English & run Greedy Analysis
        page.locator('label').filter(has_text="English (Default)").click()
        time.sleep(1)
        print("   Running Greedy Set Cover Analysis...")
        page.locator('button').filter(has_text="Run Greedy Set Cover Analysis").click()

        # Wait for completion
        page.wait_for_selector('text=Execution Complete', timeout=30000)
        time.sleep(2)

        # 7. Step 2 Results - Visual Charts & Heatmap
        print("   Capturing: 05_step2_results_visual_charts_en.png...")
        path_5 = os.path.join(out_dir, "05_step2_results_visual_charts_en.png")
        page.screenshot(path=path_5, full_page=True)

        # 8. Step 2 Results - Tables Tab
        print("   Clicking on Output Tables tab...")
        page.locator('button[role="tab"]').filter(has_text="Interactive Output Tables").click()
        time.sleep(1.5)
        print("   Capturing: 06_step2_results_tables_en.png...")
        path_6 = os.path.join(out_dir, "06_step2_results_tables_en.png")
        page.screenshot(path=path_6, full_page=True)

        browser.close()

    app.close()
    print("\nAll screenshots generated successfully in Docs/screenshots/!")
    for f in sorted(os.listdir(out_dir)):
        p = os.path.join(out_dir, f)
        print(f"  - {f} ({os.path.getsize(p):,} bytes)")

if __name__ == '__main__':
    capture_screenshots()


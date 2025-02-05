import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import base64

def capture_full_page_screenshot_base64(region) -> bytes:
    """
    使用 Selenium + CDP 截图，固定网页宽度为 1100。
    高度会自适应整页内容，从而可以完整截取所有内容。
    """
    if region != "sg":
        html_file_path = "/home/app/DeCard_Reports/DeCard_Report_Script_output.html"
    else:
        html_file_path = "/home/app/DeCard_Reports/SG_DeCard_Report_Script_output.html"

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-web-security')
    chrome_options.add_argument('--allow-file-access-from-files')

    driver = webdriver.Chrome(options=chrome_options)

    try:
        abs_path = os.path.abspath(html_file_path)
        file_url = 'file:///' + abs_path.replace('\\', '/')

        driver.get(file_url)
        time.sleep(2)  

  
        driver.set_window_size(1100, 1000)
        time.sleep(1)  

        total_height = driver.execute_script(
            "return Math.max("
            "document.body.scrollHeight, "
            "document.documentElement.scrollHeight)"
        )

        print(f"After setting width=1100, total_height is {total_height}")

        driver.set_window_size(1100, total_height)
        time.sleep(1)


        screenshot = driver.execute_cdp_cmd("Page.captureScreenshot", {
            "fromSurface": True,
            "captureBeyondViewport": True,
            "clip": {
                "x": 0,
                "y": 0,
                "width": 1100,     
                "height": total_height,
                "scale": 1
            }
        })

        image_bytes = base64.b64decode(screenshot['data'])
        return image_bytes

    except Exception as e:
        print("An error occurred:", e)
        return b""  
    finally:
        driver.quit()

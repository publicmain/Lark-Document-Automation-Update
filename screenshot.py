# screenshot.py
import base64
import io
import time
import logging
from html2image import Html2Image
from PIL import Image

logger = logging.getLogger("screenshot")
logger.setLevel(logging.DEBUG)

def capture_full_page_screenshot_base64(region) -> bytes:
    # logger.debug("进入 capture_full_page_screenshot_base64 函数, region=%s", region)
    # if region != "sg":
    #     html_file_path = "/home/app/DeCard_Reports/DeCard_Report_Script_output.html"
    #     logger.debug("使用 Global 模板, html_file_path=%s", html_file_path)
    # else:
    #     html_file_path = "/home/app/DeCard_Reports/SG_DeCard_Report_Script_output.html"
    #     logger.debug("使用 SG 模板, html_file_path=%s", html_file_path)
    # html_file_path = 'DeCard_Report_Script_output.html'
    # logger.debug("重置 html_file_path 为 %s", html_file_path)
    html_file_path = 'DeCard_Report_Script_output.html'
    temp_image_path = 'temp_screenshot.png'
    hti = Html2Image()
    hti.browser.flags = [
        '--window-size=1400,9000', 
        '--disable-gpu',
        '--no-sandbox',
    ]
    logger.debug("浏览器启动参数设置为: %s", hti.browser.flags)
    try:
        hti.screenshot(html_file=html_file_path, save_as=temp_image_path)
        logger.debug("截图已保存至 %s", temp_image_path)
    except Exception as e:
        logger.error("截图失败: %s", e, exc_info=True)
        raise

    try:
        with Image.open(temp_image_path) as img:
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
            logger.debug("图片成功转换为 base64 格式")
    except Exception as e:
        logger.error("图片转换为 base64 失败: %s", e, exc_info=True)
        raise

    if ',' in img_base64:
        img_base64 = img_base64.split(',')[1]
        logger.debug("移除 base64 前缀后的字符串长度: %d", len(img_base64))
    image_data = base64.b64decode(img_base64)
    logger.debug("base64 数据已解码为二进制数据, 长度: %d", len(image_data))
    return image_data

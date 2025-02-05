# docx_utils.py
import time
import json
import logging
from datetime import datetime, timedelta
from typing import Optional
import pytz
import requests

from screenshot import capture_full_page_screenshot_base64
from larkAPI import get_tenant_access_token

logger = logging.getLogger("docx_utils")
logger.setLevel(logging.DEBUG)

class LarkAPIError(Exception):
    """自定义异常类用于 Lark API 错误"""
    def __init__(self, status_code: int, error_code: int, message: str):
        super().__init__(f"HTTP {status_code} - Error {error_code}: {message}")
        self.status_code = status_code
        self.error_code = error_code
        self.message = message

def get_yesterday_beijing_str() -> str:
    tz_beijing = pytz.timezone('Asia/Shanghai')
    now_beijing = datetime.now(tz_beijing)
    yesterday_beijing = now_beijing - timedelta(days=1)
    date_str = yesterday_beijing.strftime("%Y.%m.%d")
    logger.debug("昨天的北京时间字符串为: %s", date_str)
    return date_str

def get_child_blocks(
    document_id: str,
    block_id: str,
    access_token: str,
    page_token: Optional[str] = None,
    page_size: int = 500,
    max_retries: int = 5,
    backoff_factor: float = 0.5
):
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{block_id}/children"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    params = {
        "document_revision_id": -1,
        "page_size": page_size
    }
    if page_token:
        params["page_token"] = page_token
    logger.debug("请求 get_child_blocks, URL: %s, 参数: %s", url, params)

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.get(url, headers=headers, params=params)
            logger.debug("第 %d 次请求 get_child_blocks, 状态码: %d", attempt + 1, resp.status_code)
            if resp.status_code == 200:
                data = resp.json()
                if data["code"] == 0:
                    items = data["data"].get("items", [])
                    has_more = data["data"].get("has_more", False)
                    next_token = data["data"].get("page_token")
                    logger.debug("get_child_blocks 成功获取 %d 个子块, has_more=%s", len(items), has_more)
                    return items, next_token, has_more
                else:
                    raise LarkAPIError(resp.status_code, data["code"], data["msg"])
            elif resp.status_code in [400, 429, 503]:
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                error_code = data.get("code", resp.status_code)
                message = data.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    logger.debug("get_child_blocks 遇到限频或服务不可用, 等待 %.2f 秒后重试（第 %d 次）", wait_time, attempt + 1)
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                try:
                    data = resp.json()
                    error_code = data.get("code", resp.status_code)
                    message = data.get("msg", "Unknown error")
                except Exception:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            logger.debug("get_child_blocks 网络错误: %s, 等待 %.2f 秒后重试（第 %d 次）", e, wait_time, attempt + 1)
            time.sleep(wait_time)
            attempt += 1
    raise Exception("超过最大重试次数，get_child_blocks 操作失败。")

def delete_child_blocks_batch(
    document_id: str,
    block_id: str,
    start_index: int,
    end_index: int,
    access_token: str,
    max_retries: int = 5,
    backoff_factor: float = 0.5
):
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{block_id}/children/batch_delete"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    params = {
        "document_revision_id": -1
    }
    payload = {
        "start_index": start_index,
        "end_index": end_index
    }
    logger.debug("请求 delete_child_blocks_batch, URL: %s, payload: %s", url, payload)

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.delete(url, headers=headers, params=params, json=payload)
            logger.debug("第 %d 次请求 delete_child_blocks_batch, 状态码: %d", attempt + 1, resp.status_code)
            if resp.status_code == 200:
                data = resp.json()
                if data["code"] == 0:
                    logger.debug("delete_child_blocks_batch 成功删除子块")
                    return data["data"]
                else:
                    raise LarkAPIError(resp.status_code, data["code"], data["msg"])
            elif resp.status_code in [400, 429, 503]:
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                error_code = data.get("code", resp.status_code)
                message = data.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    logger.debug("delete_child_blocks_batch 遇到限频或服务不可用, 等待 %.2f 秒后重试（第 %d 次）", wait_time, attempt + 1)
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                try:
                    data = resp.json()
                    error_code = data.get("code", resp.status_code)
                    message = data.get("msg", "Unknown error")
                except Exception:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            logger.debug("delete_child_blocks_batch 网络错误: %s, 等待 %.2f 秒后重试（第 %d 次）", e, wait_time, attempt + 1)
            time.sleep(wait_time)
            attempt += 1
    raise Exception("超过最大重试次数，delete_child_blocks_batch 操作失败。")

def create_text_block(
    document_id: str,
    parent_block_id: str,
    access_token: str,
    index: int = 0,
    max_retries: int = 5,
    backoff_factor: float = 0.5
) -> str:
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    date_str = get_yesterday_beijing_str()
    logger.debug("当前使用的日期字符串: %s", date_str)

    part1 = {"text_run": {"content": "本文对"}}
    part2 = {
        "text_run": {
            "content": "DeCard",
            "text_element_style": {
                "underline": True,
                "text_color": 3,
            }
        }
    }
    part3 = {
        "text_run": {
            "content": (
                "产品用户进行基础客户画像和行为分析（初步），为客户运营提供方案，"
                "以提升客户体验以及推动产品使用和销售的重要策略。"
                "卡产品包含：MV3789 - Master新币卡，UV3701 - 银联新币卡，数据截止 "
            )
        }
    }
    part4 = {
        "text_run": {
            "content": date_str,
            "text_element_style": {"underline": True}
        }
    }
    part5 = {"text_run": {"content": "。"}}

    payload = {
        "index": index,
        "children": [
            {
                "block_type": 2,
                "text": {
                    "elements": [part1, part2, part3, part4, part5]
                }
            }
        ]
    }
    logger.debug("请求 create_text_block, URL: %s, payload: %s", url, payload)
    resp = requests.post(url, headers=headers, json=payload)
    try:
        resp.raise_for_status()
    except Exception as e:
        logger.error("create_text_block 请求失败: %s", e, exc_info=True)
        raise
    resp_json = resp.json()
    if resp_json.get("code") == 0:
        children_data = resp_json["data"]["children"]
        if children_data:
            new_text_block_id = children_data[0]["block_id"]
            logger.debug("创建文本块成功, new_text_block_id=%s", new_text_block_id)
            return new_text_block_id
        else:
            raise Exception("No child returned in response while creating text block.")
    else:
        raise Exception(f"Create text block failed: {resp_json.get('msg')}")

def clear_document_blocks(
    document_id: str,
    access_token: str,
    max_retries: int = 5,
    backoff_factor: float = 0.5
):
    block_id = document_id  # 文档根节点 block_id
    while True:
        items, next_token, has_more = get_child_blocks(
            document_id=document_id,
            block_id=block_id,
            access_token=access_token,
            max_retries=max_retries,
            backoff_factor=backoff_factor
        )
        child_count = len(items)
        if child_count == 0:
            logger.debug("当前已无任何子块, 文档已清空。")
            break

        delete_child_blocks_batch(
            document_id=document_id,
            block_id=block_id,
            start_index=0,
            end_index=child_count,
            access_token=access_token,
            max_retries=max_retries,
            backoff_factor=backoff_factor
        )
        logger.debug("已删除 %d 个子块。", child_count)

        # 如需确认是否清空则继续循环
        if not has_more and child_count < 500:
            continue

    logger.debug("文档根节点清空完毕。")

def create_image_block(
    document_id: str,
    parent_block_id: str,
    access_token: str,
    index: int = 5,
    max_retries: int = 5,
    backoff_factor: float = 0.5
) -> str:
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = {
        "index": index,
        "children": [
            {
                "block_type": 27,
                "image": {"token": ""}
            }
        ]
    }
    logger.debug("请求 create_image_block, URL: %s, payload: %s", url, payload)
    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.post(url, headers=headers, json=payload)
            logger.debug("第 %d 次请求 create_image_block, 状态码: %d", attempt + 1, resp.status_code)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                    children_data = resp_json["data"]["children"]
                    if children_data:
                        image_block_id = children_data[0]["block_id"]
                        logger.debug("创建图片块成功, image_block_id=%s", image_block_id)
                        return image_block_id
                    else:
                        raise LarkAPIError(200, -1, "No child returned in response.")
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    logger.debug("create_image_block 遇到限频或服务不可用, 等待 %.2f 秒后重试（第 %d 次）", wait_time, attempt + 1)
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            logger.debug("create_image_block 网络错误: %s, 等待 %.2f 秒后重试（第 %d 次）", e, wait_time, attempt + 1)
            time.sleep(wait_time)
            attempt += 1
    raise Exception("超过最大重试次数，create_image_block 操作失败。")

def upload_image_data(
    image_block_id: str,
    image_data: bytes,
    file_name: str,
    access_token: str,
    max_retries: int = 5,
    backoff_factor: float = 0.5
) -> str:
    url = "https://open.larksuite.com/open-apis/drive/v1/medias/upload_all"
    headers = {"Authorization": f"Bearer {access_token}"}
    files = {"file": (file_name, image_data, "application/octet-stream")}
    data = {
        "file_name": file_name,
        "parent_type": "docx_image",
        "parent_node": image_block_id,
        "size": str(len(image_data)),
        "extra": '{"drive_route_token":"V0Hxd7J8PoYwYNxA1z4l7dWrg3f"}'
    }
    logger.debug("请求 upload_image_data, URL: %s, 文件名: %s, 数据大小: %d", url, file_name, len(image_data))
    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.post(url, headers=headers, files=files, data=data)
            logger.debug("第 %d 次请求 upload_image_data, 状态码: %d", attempt + 1, resp.status_code)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                    file_token = resp_json["data"]["file_token"]
                    logger.debug("上传图片成功, file_token=%s", file_token)
                    return file_token
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                try:
                    resp_json = resp.json()
                except Exception:
                    resp_json = {}
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    logger.debug("upload_image_data 遇到限频或服务不可用, 等待 %.2f 秒后重试（第 %d 次）", wait_time, attempt + 1)
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                try:
                    resp_json = resp.json()
                    error_code = resp_json.get("code", resp.status_code)
                    message = resp_json.get("msg", "Unknown error")
                except Exception:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            logger.debug("upload_image_data 网络错误: %s, 等待 %.2f 秒后重试（第 %d 次）", e, wait_time, attempt + 1)
            time.sleep(wait_time)
            attempt += 1
    raise Exception("超过最大重试次数，upload_image_data 操作失败。")

def replace_image_in_block(
    document_id: str,
    block_id: str,
    file_token: str,
    access_token: str,
    max_retries: int = 5,
    backoff_factor: float = 0.5
) -> None:
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{block_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = {"replace_image": {"token": file_token}}
    logger.debug("请求 replace_image_in_block, URL: %s, payload: %s", url, payload)
    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.patch(url, headers=headers, json=payload)
            logger.debug("第 %d 次请求 replace_image_in_block, 状态码: %d", attempt + 1, resp.status_code)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                    logger.debug("replace_image_in_block 图片替换成功！")
                    return
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    logger.debug("replace_image_in_block 遇到限频或服务不可用, 等待 %.2f 秒后重试（第 %d 次）", wait_time, attempt + 1)
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            logger.debug("replace_image_in_block 网络错误: %s, 等待 %.2f 秒后重试（第 %d 次）", e, wait_time, attempt + 1)
            time.sleep(wait_time)
            attempt += 1
    raise Exception("超过最大重试次数，replace_image_in_block 操作失败。")

def insert_image_example_in_memory(DOCUMENT_ID, REGION):
    """
    1) 使用 Selenium 截图得到二进制图像数据  
    2) 清空文档  
    3) 创建图片块  
    4) 直接上传二进制数据  
    5) 替换图片块  
    """
    PARENT_BLOCK_ID = DOCUMENT_ID
    IMAGE_NAME = "screenshot_in_memory.png"
    ACCESS_TOKEN = get_tenant_access_token()
    logger.debug("获取 ACCESS_TOKEN 成功: %s", ACCESS_TOKEN)

    # 1) 截图
    image_data = capture_full_page_screenshot_base64(REGION)
    logger.debug("截图得到的 image_data 长度: %d", len(image_data))

    # 2) 清空文档
    clear_document_blocks(document_id=DOCUMENT_ID, access_token=ACCESS_TOKEN)
    # 可选：创建文本块
    create_text_block(
        document_id=DOCUMENT_ID,
        parent_block_id=PARENT_BLOCK_ID,
        access_token=ACCESS_TOKEN
    )
    # 3) 创建图片块
    image_block_id = create_image_block(
        document_id=DOCUMENT_ID,
        parent_block_id=PARENT_BLOCK_ID,
        access_token=ACCESS_TOKEN,
        index=1
    )
    logger.debug("图片块创建成功, block_id=%s", image_block_id)

    # 4) 上传图片数据
    file_token = upload_image_data(
        image_block_id=image_block_id,
        image_data=image_data,
        file_name=IMAGE_NAME,
        access_token=ACCESS_TOKEN
    )
    logger.debug("上传图片成功, file_token=%s", file_token)

    # 5) 替换图片块
    replace_image_in_block(
        document_id=DOCUMENT_ID,
        block_id=image_block_id,
        file_token=file_token,
        access_token=ACCESS_TOKEN
    )
    logger.debug("图片块替换完成")

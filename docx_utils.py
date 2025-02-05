import requests
import time
import json
from typing import Optional, Dict
import os
from datetime import datetime, timedelta
import pytz
from larkAPI import get_tenant_access_token
class LarkAPIError(Exception):
    """自定义异常类用于Lark API错误"""
    def __init__(self, status_code: int, error_code: int, message: str):
        super().__init__(f"HTTP {status_code} - Error {error_code}: {message}")
        self.status_code = status_code
        self.error_code = error_code
        self.message = message

def get_yesterday_beijing_str() -> str:
    """
    获取北京时间的昨天日期，格式为 yyyy.MM.dd
    """
    # 1) 指定时区为 "Asia/Shanghai"（中国大陆使用 UTC+8）
    tz_beijing = pytz.timezone('Asia/Shanghai')

    # 2) 获取当前北京时间
    now_beijing = datetime.now(tz_beijing)

    # 3) 往前减去一天
    yesterday_beijing = now_beijing - timedelta(days=1)

    # 4) 格式化为 "YYYY.MM.DD"
    return yesterday_beijing.strftime("%Y.%m.%d")

def get_child_blocks(
    document_id: str,
    block_id: str,
    access_token: str,
    page_token: Optional[str] = None,
    page_size: int = 500,
    max_retries: int = 5,
    backoff_factor: float = 0.5
):
    """
    使用 'GET /docx/v1/documents/:document_id/blocks/:block_id/children'
    获取指定 block 下的子块列表（分页）。
    :return: (child_items, next_page_token, has_more)
    """
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{block_id}/children"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    params = {
        "document_revision_id": -1,  # -1 表示最新版本
        "page_size": page_size
    }
    if page_token:
        params["page_token"] = page_token

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.get(url, headers=headers, params=params)
            if resp.status_code == 200:
                data = resp.json()
                if data["code"] == 0:
                    items = data["data"].get("items", [])
                    has_more = data["data"].get("has_more", False)
                    next_token = data["data"].get("page_token")
                    return items, next_token, has_more
                else:
                    raise LarkAPIError(
                        status_code=resp.status_code,
                        error_code=data["code"],
                        message=data["msg"]
                    )
            elif resp.status_code in [400, 429, 503]:
                # 频率限制或服务器暂时不可用，重试
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                error_code = data.get("code", resp.status_code)
                message = data.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    print(f"[get_child_blocks] 频率限制或服务不可用，等待{wait_time}秒后重试...（第{attempt+1}次）")
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                # 其他 HTTP 错误
                try:
                    data = resp.json()
                    error_code = data.get("code", resp.status_code)
                    message = data.get("msg", "Unknown error")
                except:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            print(f"[get_child_blocks] 网络错误: {e}，等待{wait_time}秒后重试...（第{attempt+1}次）")
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
    """
    调用 'DELETE /docx/v1/documents/:document_id/blocks/:block_id/children/batch_delete'
    删除指定范围 [start_index, end_index) 的子块。
    """
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

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.delete(url, headers=headers, params=params, json=payload)
            if resp.status_code == 200:
                data = resp.json()
                if data["code"] == 0:
                    return data["data"]
                else:
                    raise LarkAPIError(
                        status_code=resp.status_code,
                        error_code=data["code"],
                        message=data["msg"]
                    )
            elif resp.status_code in [400, 429, 503]:
                # 频率限制或服务器暂时不可用，重试
                try:
                    data = resp.json()
                except Exception:
                    data = {}
                error_code = data.get("code", resp.status_code)
                message = data.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    print(f"[delete_child_blocks_batch] 频率限制或服务不可用，等待{wait_time}秒后重试...（第{attempt+1}次）")
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                # 其他 HTTP 错误
                try:
                    data = resp.json()
                    error_code = data.get("code", resp.status_code)
                    message = data.get("msg", "Unknown error")
                except:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            print(f"[delete_child_blocks_batch] 网络错误: {e}，等待{wait_time}秒后重试...（第{attempt+1}次）")
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
    """
    在文档 (或父块) 下创建一个文本块 (block_type=2)，并插入指定文字。
    返回新建文本块的block_id，方便后续操作（如果需要）。
    """
    import requests
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{parent_block_id}/children"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    date_str = get_yesterday_beijing_str()  # 形如 "2025.01.20"

   
    part1 = {
        "text_run": {
            "content": "本文对"
        }
    }

    part2 = {
        "text_run": {
            "content": "DeCard",  
            "text_element_style": {
                "underline":True,        # 下划线
                "text_color": 3,  # 黄色背景
            }
        }
    }
    part3 = {
        "text_run": {
            "content":
              "产品用户进行基础客户画像和行为分析（初步），为客户运营提供方案，"
              "以提升客户体验以及推动产品使用和销售的重要策略。"
              "卡产品包含：MV3789 - Master新币卡，UV3701 - 银联新币卡，数据截止 "
        }
    }
    part4 = {
        "text_run": {
            "content": date_str,  
            "text_element_style": {
                "underline":True
            }
        }
    }
    # 5) 句号
    part5 = {
        "text_run": {
            "content": "。"
        }
    }

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

    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    resp_json = resp.json()
    if resp_json.get("code") == 0:
        children_data = resp_json["data"]["children"]
        if len(children_data) > 0:
            new_text_block_id = children_data[0]["block_id"]
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
    """
    清空文档根节点下所有子块。
    可能需要多次分页获取并循环删除，直到没有子块为止。
    """
    block_id = document_id  #
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
            print("[clear_document_blocks] 当前已无任何子块，文档已清空。")
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
        print(f"[clear_document_blocks] 已删除 {child_count} 个子块。")

        # 如果还有更多页，理论上可以继续处理，但因为我们一次就把当前页的所有子块都删除了，
        # 所以下一个 page_token 也就失效了。循环重新获取可确认是否真的清空。
        # 若文档子块非常多，需要多次循环，直到没有子块返回为止。
        if not has_more and child_count < 500:
            # 如果这页 < 500 并且 has_more=False，基本可以判断没有子块了
            # 但还要再进一次循环确认
            continue

    print("[clear_document_blocks] 文档根节点清空完毕。")


def create_image_block(
    document_id: str,
    parent_block_id: str,
    access_token: str,
    index: int = 5,
    max_retries: int = 5,
    backoff_factor: float = 0.5
) -> str:
    """
    第一步：创建图片 Block

    :param document_id: 文档 ID
    :param parent_block_id: 父块的 block_id。如果需要插入到文档根节点，可将 document_id 直接作为 block_id。
    :param access_token: 鉴权用的 token
    :param index: 在父块 children 中的插入位置，默认插在开头
    :param max_retries: 发生限频或网络错误时的最大重试次数
    :param backoff_factor: 指数退避的基数，默认0.5秒
    :return: 创建成功后返回的 Image Block 的 ID
    """
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
                "image": {
                    "token": ""    
                }
            }
        ]
    }

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.post(url, headers=headers, json=payload)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                 
                    children_data = resp_json["data"]["children"]
                    if len(children_data) > 0:
                        return children_data[0]["block_id"]
                    else:
                        raise LarkAPIError(200, -1, "No child returned in response.")
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                # 频率限制或服务器暂时不可用，进行重试
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    print(f"[create_image_block] 频率限制或服务不可用，等待{wait_time}秒后重试...（第{attempt + 1}次）")
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                # 其他未预料的HTTP错误
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            # 网络错误，重试
            wait_time = backoff_factor * (2 ** attempt)
            print(f"[create_image_block] 网络错误: {e}，等待{wait_time}秒后重试...（第{attempt + 1}次）")
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
    """
    上传图片二进制数据到飞书云空间，对应一个已有的图片块 (image_block_id)。
    不需要在本地生成文件。
    """

    url = "https://open.larksuite.com/open-apis/drive/v1/medias/upload_all"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    # 构造 multipart/form-data
    files = {
        "file": (file_name, image_data, "application/octet-stream")
    }
    data = {
        "file_name": file_name,
        "parent_type": "docx_image",
        "parent_node": image_block_id,
        "size": str(len(image_data)),
        "extra": '{"drive_route_token":"V0Hxd7J8PoYwYNxA1z4l7dWrg3f"}'
    }

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.post(url, headers=headers, files=files, data=data)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                    return resp_json["data"]["file_token"]
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                # 频率限制或服务器暂时不可用，进行重试
                try:
                    resp_json = resp.json()
                except Exception:
                    resp_json = {}
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    print(f"[upload_image_data] 频率限制或服务不可用，等待{wait_time}秒后重试...（第{attempt+1}次）")
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                # 其他未预料的HTTP错误
                try:
                    resp_json = resp.json()
                    error_code = resp_json.get("code", resp.status_code)
                    message = resp_json.get("msg", "Unknown error")
                except:
                    error_code = resp.status_code
                    message = resp.text
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            wait_time = backoff_factor * (2 ** attempt)
            print(f"[upload_image_data] 网络错误: {e}，等待{wait_time}秒后重试...（第{attempt+1}次）")
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
    """
    第三步：将图片素材 token 替换到指定的图片 Block

    :param document_id: 文档 ID
    :param block_id: 图片 Block 的 block_id
    :param file_token: 第二步上传素材后返回的 file_token
    :param access_token: 鉴权用的 token
    :param max_retries: 发生限频或网络错误时的最大重试次数
    :param backoff_factor: 指数退避的基数，默认0.5秒
    """
    url = f"https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks/{block_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = {
        "replace_image": {
            "token": file_token
        }
    }

    attempt = 0
    while attempt <= max_retries:
        try:
            resp = requests.patch(url, headers=headers, json=payload)
            if resp.status_code == 200:
                resp_json = resp.json()
                if resp_json.get("code") == 0:
                    print("[replace_image_in_block] 图片替换成功！")
                    return
                else:
                    raise LarkAPIError(resp.status_code, resp_json.get("code"), resp_json.get("msg"))
            elif resp.status_code in [400, 429, 503]:
                # 频率限制或服务器暂时不可用，进行重试
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                if resp.status_code == 429 or error_code == 99991400:
                    wait_time = backoff_factor * (2 ** attempt)
                    print(f"[replace_image_in_block] 频率限制或服务不可用，等待{wait_time}秒后重试...（第{attempt + 1}次）")
                    time.sleep(wait_time)
                    attempt += 1
                else:
                    raise LarkAPIError(resp.status_code, error_code, message)
            else:
                # 其他未预料的HTTP错误
                resp_json = resp.json()
                error_code = resp_json.get("code", resp.status_code)
                message = resp_json.get("msg", "Unknown error")
                raise LarkAPIError(resp.status_code, error_code, message)
        except requests.RequestException as e:
            # 网络错误，重试
            wait_time = backoff_factor * (2 ** attempt)
            print(f"[replace_image_in_block] 网络错误: {e}，等待{wait_time}秒后重试...（第{attempt + 1}次）")
            time.sleep(wait_time)
            attempt += 1
    
    raise Exception("超过最大重试次数，replace_image_in_block 操作失败。")


def insert_image_example_in_memory(DOCUMENT_ID,REGION):
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

    from screenshot import capture_full_page_screenshot_base64
    image_data = capture_full_page_screenshot_base64(REGION)
    
    # 2) 清空文档

    clear_document_blocks(document_id=DOCUMENT_ID, access_token=ACCESS_TOKEN)

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
 

    # 4) 直接上传二进制数据
   
    file_token = upload_image_data(
        image_block_id=image_block_id,
        image_data=image_data,
        file_name=IMAGE_NAME,
        access_token=ACCESS_TOKEN
    )
   

    # 5) 替换图片块
   
    replace_image_in_block(
        document_id=DOCUMENT_ID,
        block_id=image_block_id,
        file_token=file_token,
        access_token=ACCESS_TOKEN
    )
   


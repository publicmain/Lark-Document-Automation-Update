import requests
import time
import base64
import json
import os



API_URL_TEMPLATE = 'https://open.larksuite.com/open-apis/docx/v1/documents/{document_id}/blocks'

MAX_RETRIES = 5
INITIAL_BACKOFF = 1  

def get_tenant_access_token():
    app_id = "cli_a6ec0559c4be502f"
    app_secret = "JlWPUAUbiuZNlYcI6DkVjbgs3holBWhI"
    url = 'https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal/'
    headers = {
        'Content-Type': 'application/json; charset=utf-8'
    }
    payload = {
        "app_id": app_id,
        "app_secret": app_secret
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        data = response.json()
        if data.get('code') == 0:
            return data.get('tenant_access_token')
        else:
            raise Exception(f"Error getting access token: {data.get('msg')}")
    else:
        raise Exception(f"HTTP error: {response.status_code}")

def get_document_blocks(document_id, access_token, page_size=500, page_token=None, user_id_type='open_id', document_revision_id=-1):
    url = API_URL_TEMPLATE.format(document_id=document_id)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json; charset=utf-8'
    }
    params = {
        'page_size': page_size,
        'user_id_type': user_id_type,
        'document_revision_id': document_revision_id
    }
    if page_token:
        params['page_token'] = page_token

    response = requests.get(url, headers=headers, params=params)
    return response

def handle_rate_limiting(response):
    """
    根据 API 文档，当接口返回 HTTP 状态码 400 及错误码 99991400 时，表示请求被限频。
    """
    if response.status_code == 400:
        try:
            data = response.json()
            if data.get('code') == 99991400:
                return True
        except:
            pass
    return False

def fetch_all_blocks(document_id, access_token):
    all_blocks = []
    page_token = None
    retries = 0
    backoff = INITIAL_BACKOFF

    while True:
        response = get_document_blocks(document_id, access_token, page_size=500, page_token=page_token)
        
        if response.status_code == 200:
            data = response.json()
            if data.get('code') != 0:
                raise Exception(f"API Error: {data.get('msg')}")
            blocks = data.get('data', {}).get('items', [])
            all_blocks.extend(blocks)
            has_more = data.get('data', {}).get('has_more', False)
            page_token = data.get('data', {}).get('page_token', None)
            print(f"Fetched {len(blocks)} blocks. Has more: {has_more}")
            if not has_more:
                break
            # 重置重试计数
            retries = 0
            backoff = INITIAL_BACKOFF
        elif handle_rate_limiting(response):
            if retries < MAX_RETRIES:
                print(f"Rate limited. Retrying after {backoff} seconds...")
                time.sleep(backoff)
                retries += 1
                backoff *= 2  # 指数退避
            else:
                raise Exception("Max retries exceeded due to rate limiting.")
        else:
            raise Exception(f"HTTP Error: {response.status_code} - {response.text}")
    
    return all_blocks

def main():

    DOCUMENT_ID = 'V0Hxd7J8PoYwYNxA1z4l7dWrg3f'

    try:
        # 获取 Access Token
        access_token = get_tenant_access_token()
        print(f"Access Token: {access_token}")

        # 获取所有块
        blocks = fetch_all_blocks(DOCUMENT_ID, access_token)
        print(f"Total blocks fetched: {len(blocks)}")

        # 处理或保存块信息
        with open('blocks.json', 'w', encoding='utf-8') as f:
            json.dump(blocks, f, ensure_ascii=False, indent=4)
        print("Blocks have been saved to blocks.json")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()

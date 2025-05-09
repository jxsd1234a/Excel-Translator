# -*- coding: utf-8 -*-
import openpyxl
import requests
import random
import hashlib
from concurrent.futures import ThreadPoolExecutor
from threading import Semaphore

# 并发控制信号量，限制同时运行的翻译请求数量
semaphore = Semaphore(2)  # 根据API频率限制调整数值

def make_md5(s):
    """生成MD5签名"""
    return hashlib.md5(s.encode('utf-8')).hexdigest()

def translate_text(appid, appkey, text, from_lang='en', to_lang='zh'):
    """调用百度翻译API进行文本翻译
    Args:
        appid: API账户ID
        appkey: API密钥
        text: 待翻译文本
        from_lang: 源语言(默认英语)
        to_lang: 目标语言(默认简体中文)
    Returns:
        翻译后的文本或原始文本(翻译失败时)
    """
    # 生成随机盐值
    salt = random.randint(32768, 65536)
    # 构建签名原始字符串
    sign_str = appid + text + str(salt) + appkey
    sign = make_md5(sign_str)  # 生成32位小写MD5签名
    
    # 配置API请求参数
    url = "http://api.fanyi.baidu.com/api/trans/vip/translate"
    payload = {
        'appid': appid,        # API账户ID
        'q': text,             # 待翻译内容(未编码)
        'from': from_lang,     # 源语言
        'to': to_lang,         # 目标语言
        'salt': salt,          # 随机盐值
        'sign': sign,          # 加密签名
        'needIntervene': 1     # 启用术语库(需在控制台配置)
    }
    
    try:
        response = requests.post(url, params=payload, timeout=10)
        result = response.json()
        if 'trans_result' in result:
            return result['trans_result'][0]['dst']
        print(f"翻译失败：{text}，响应：{result}")
        return text
    except Exception as e:
        print(f"请求异常：{str(e)}")
        return text

def translate_excel(input_path, output_path, appid, appkey):
    """处理Excel文件翻译任务
    Args:
        input_path: 输入文件路径
        output_path: 输出文件路径 
        appid: API账户ID
        appkey: API密钥
    """
    # 加载Excel工作簿
    wb = openpyxl.load_workbook(input_path)
    
    # 遍历所有工作表
    for sheet in wb.worksheets:
        print(f"处理工作表：{sheet.title}")
        # 收集需要翻译的单元格(第二列且非空)
        cells = [cell for row in sheet.iter_rows(min_col=2, max_col=2) 
                for cell in row if cell.value and cell.value.strip()]
        
        # 创建线程池执行翻译任务
        with ThreadPoolExecutor(max_workers=3) as executor:  
            # 映射任务到线程池
            executor.map(process_cell, cells, [appid]*len(cells), [appkey]*len(cells))
    
    # 保存翻译结果
    wb.save(output_path)
    print(f"翻译完成！保存至：{output_path}")

def process_cell(cell, appid, appkey):
    """处理单个单元格翻译任务"""
    with semaphore:  # 通过信号量控制并发
        original = str(cell.value).strip()  # 原始文本处理
        translated = translate_text(appid, appkey, original)  # 调用翻译API
        cell.value = translated  # 回写翻译结果
        print(f"已翻译：{original} → {translated}")

if __name__ == "__main__":
    # 百度API凭证
    APP_ID = '你的APP_ID'
    APP_KEY = '你的APP_KEY'
    
    # 文件路径
    INPUT_FILE = '名字.xlsx'
    # 文件导出名称
    OUTPUT_FILE = '名字-翻译.xlsx'
    
    translate_excel(INPUT_FILE, OUTPUT_FILE, APP_ID, APP_KEY)

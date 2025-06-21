
import sys
import os
import pandas as pd
from typing import List
import urllib.request
import time
import json
import logging
from datetime import datetime
import schedule
import threading
from flask import Flask, jsonify

from alibabacloud_dingtalk.oauth2_1_0.client import Client as DingTalkOauth2Client
from alibabacloud_dingtalk.storage_1_0.client import Client as DingTalkStorageClient
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_dingtalk.oauth2_1_0 import models as oauth2_models
from alibabacloud_dingtalk.storage_1_0 import models as storage_models
from alibabacloud_tea_util import models as util_models
from alibabacloud_tea_util.client import Client as UtilClient

# 配置信息
APP_KEY = ''  # APP_KEY
APP_SECRET = ''  # APP_SECRET
SPACE_ID = ''  # 钉盘空间ID
DENTRY_ID = ''  # 文件ID
UNION_ID = ''  # 用户unionId

# 创建日志目录
logs_dir = 'logs'
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)

# 配置日志
log_file = os.path.join(logs_dir, f'bl_data_{datetime.now().strftime("%Y%m%d")}.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('bl_data_collector')

# 获取访问令牌
def get_access_token():
    logger.info("开始获取访问令牌")
    config = open_api_models.Config()
    config.protocol = 'https'
    config.region_id = 'central'
    client = DingTalkOauth2Client(config)

    request = oauth2_models.GetAccessTokenRequest(
        app_key=APP_KEY,
        app_secret=APP_SECRET
    )

    try:
        response = client.get_access_token(request)
        logger.info("访问令牌获取成功")
        return response.body.access_token
    except Exception as e:
        logger.error(f"获取access_token失败: {e}")
        return None

# 获取文件下载信息
def get_file_download_info(access_token):
    logger.info(f"开始获取文件下载信息: 空间ID={SPACE_ID}, 文件ID={DENTRY_ID}")
    config = open_api_models.Config()
    config.protocol = 'https'
    config.region_id = 'central'
    client = DingTalkStorageClient(config)

    headers = storage_models.GetFileDownloadInfoHeaders()
    headers.x_acs_dingtalk_access_token = access_token

    request = storage_models.GetFileDownloadInfoRequest(
        union_id=UNION_ID
    )

    try:
        response = client.get_file_download_info_with_options(
            SPACE_ID, DENTRY_ID, request, headers, util_models.RuntimeOptions()
        )

        # 返回下载所需的信息
        resource_url = response.body.header_signature_info.resource_urls[0]
        headers = {
            'x-oss-date': response.body.header_signature_info.headers['x-oss-date'],
            'Authorization': response.body.header_signature_info.headers['Authorization']
        }

        logger.info("文件下载信息获取成功")
        return resource_url, headers
    except Exception as e:
        logger.error(f"获取文件下载信息失败: {e}")
        return None, None

# 下载并解析Excel文件
def download_and_parse_excel():
    logger.info("开始下载并解析Excel文件")
    try:
        # 获取访问令牌
        access_token = get_access_token()
        if not access_token:
            raise Exception("获取访问令牌失败")

        # 获取文件下载信息
        resource_url, headers = get_file_download_info(access_token)
        if not resource_url or not headers:
            raise Exception("获取文件下载信息失败")

        # 下载文件，使用固定文件名
        excel_file = "bl_data.xlsx"  # 固定文件名

        logger.info(f"开始下载文件到: {excel_file}")
        req = urllib.request.Request(resource_url, headers=headers)
        with urllib.request.urlopen(req) as response, open(excel_file, 'wb') as out_file:
            data = response.read()
            out_file.write(data)
            file_size = len(data)
            logger.info(f"文件下载完成，大小: {file_size} 字节")
            if file_size == 0:
                logger.error("下载文件为空")
                return []

        logger.info("文件下载完成，开始解析")
        # 解析Excel
        df = pd.read_excel(excel_file)
        logger.info(f"Excel文件包含 {len(df)} 行数据，列名: {list(df.columns)}")  # 新增日志

        # 检查BL号码列
        if "BL号码" not in df.columns:
            logger.warning("Excel文件缺少'BL号码'列，返回空数据")
            return []

        # 提取BL号码数据
        bl_data = []
        for index, row in df.iterrows():
            bl_number = row["BL号码"]
            if pd.notna(bl_number) and str(bl_number).strip():
                bl_data.append({
                    "bl_number": str(bl_number).strip(),
                    "submit_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "status": "pending"
                })

        logger.info(f"Excel文件解析完成，共提取 {len(bl_data)} 条BL号码")
        return bl_data
    except Exception as e:
        logger.error(f"下载和解析Excel时出错: {e}")
        return []

# 保存数据到本地文件
def save_data(data):
    logger.info("开始保存数据到本地文件")
    try:
        # 明确列名，确保空数据也有正确格式
        df = pd.DataFrame(data, columns=["bl_number", "submit_time", "status"])  # 修改：明确列名
        df.to_csv("bl_data.csv", index=False)
        logger.info(f"数据保存成功，总共 {len(df)} 条记录（已覆盖原文件）")
        return "bl_data.csv"
    except Exception as e:
        logger.error(f"保存数据时出错: {e}")
        return None

# Flask应用 - API服务器
app = Flask(__name__)

@app.route('/api/bl_numbers', methods=['GET'])
def get_bl_numbers():
    logger.info("接收到获取BL号码请求")
    try:
        df = pd.read_csv("bl_data.csv")
        all_data = df.to_dict('records')
        logger.info(f"返回 {len(all_data)} 条数据")
        return jsonify({"success": True, "data": all_data})
    except FileNotFoundError:
        logger.warning("未找到数据文件，返回空列表")
        return jsonify({"success": True, "data": []})
    except Exception as e:
        logger.error(f"获取BL号码时出错: {e}")
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/update_status/<bl_number>/<status>', methods=['POST'])
def update_status(bl_number, status):
    logger.info(f"接收到更新状态请求: BL号码={bl_number}, 状态={status}")
    try:
        df = pd.read_csv("bl_data.csv")
        if bl_number in df['bl_number'].values:
            df.loc[df['bl_number'] == bl_number, 'status'] = status
            df.to_csv("bl_data.csv", index=False)
            logger.info(f"BL号码 {bl_number} 状态更新为 {status}")
            return jsonify({"success": True})
        else:
            logger.warning(f"未找到BL号码: {bl_number}")
            return jsonify({"success": False, "error": "BL号码不存在"})
    except Exception as e:
        logger.error(f"更新状态时出错: {e}")
        return jsonify({"success": False, "error": str(e)})

# 定时任务
def scheduled_job():
    logger.info("开始执行定时任务")
    bl_data = download_and_parse_excel()
    filename = save_data(bl_data)  # 修改：始终保存数据，即使为空
    if filename:
        logger.info(f"数据已保存到 {filename}，共 {len(bl_data)} 条记录")
    else:
        logger.error("保存数据失败")

def run_scheduler():
    logger.info("启动定时任务调度器")
    # 立即执行一次
    scheduled_job()

    # 设置定时任务
    schedule.every(8).hours.do(scheduled_job)
    logger.info("已设置每8小时执行一次定时任务")

    while True:
        schedule.run_pending()
        time.sleep(59)

# 主函数
def main():
    logger.info("程序启动")
    try:
        # 启动定时任务线程
        scheduler_thread = threading.Thread(target=run_scheduler)
        scheduler_thread.daemon = True
        scheduler_thread.start()
        logger.info("定时任务线程已启动")

        # 启动Flask服务
        logger.info("启动API服务器...")
        app.run(host='0.0.0.0', port=6030)

    except Exception as e:
        logger.error(f"程序运行出错: {e}")

if __name__ == '__main__':
    main()

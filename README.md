# DingTalk-Scheduled-Table-Fetch
钉钉定时表格获取
# Excel 数据获取与处理程序

这是一个用于从钉钉云盘下载 Excel 文件、解析其中的数据并提供 API 接口来访问这些数据的 Python 程序。该程序支持定时任务，可以定期从钉钉云盘更新数据，并通过 Flask 提供 RESTful API 来查询和更新数据状态。

## 功能概述

- **从钉钉云盘下载 Excel 文件**：使用钉钉开放平台的 API 获取文件下载链接。
- **解析 Excel 文件**：读取 Excel 文件中的“BL号码”列数据。
- **保存数据到 CSV 文件**：将解析后的数据保存为本地 CSV 文件。
- **Flask API 服务**：
  - `/api/bl_numbers`：返回所有 BL 号码数据。
  - `/api/update_status/<bl_number>/<status>`：更新某个 BL 号码的状态。
- **定时任务**：每 8 小时自动从钉钉云盘下载并更新数据。

## 技术栈

- **Python 3.x**
- **第三方库**：
  - `pandas`: 用于数据解析和处理。
  - `urllib.request`: 用于下载文件。
  - `schedule`: 定时任务调度。
  - `flask`: 提供 HTTP API。
  - `alibabacloud-dingtalk`: 钉钉开放平台 SDK。
  - `logging`: 日志记录。
  - `datetime`, `time`, `os`, `sys`: 标准库支持。

## 配置说明

在 `excel_get.py` 文件中需要配置以下信息：

| 参数名        | 描述                           |
|---------------|--------------------------------|
| `APP_KEY`     | 钉钉应用的 App Key             |
| `APP_SECRET`  | 钉钉应用的 App Secret          |
| `SPACE_ID`    | 钉盘空间 ID                    |
| `DENTRY_ID`   | 文件 ID                        |
| `UNION_ID`    | 用户 Union ID                  |

## 运行环境要求

- Python 3.x

- - 程序会启动一个定时任务，每 8 小时从钉钉云盘下载一次 Excel 文件。
- 同时启动 Flask 服务，默认监听 `0.0.0.0:6030`。

## API 接口文档

### 获取所有 BL 号码

- **GET** `/api/bl_numbers`

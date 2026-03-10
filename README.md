# SharePoint 文件下载工具

这是一个用于从 SharePoint 下载文件的 Python 工具。该工具通过 Microsoft Graph API 获取 SharePoint 文件并下载到本地。

## 功能特点

- 解析 SharePoint 共享链接
- 自动获取站点 ID 和驱动器 ID
- 通过 Microsoft Graph API 下载文件
- 支持流式下载，适用于大文件
- 错误处理和日志记录

## 系统要求

- Python 3.6+
- requests 库
- python-dotenv 库

## 安装依赖

```bash
pip install requests python-dotenv
```

## 配置说明

1. 创建一个 `.env` 文件在项目根目录
2. 在 `.env` 文件中添加你的 Microsoft Graph API 访问令牌：

```
ACCESS_TOKEN=你的Microsoft Graph API访问令牌
```

## 使用方法

### 命令行使用

直接运行脚本，将需要下载的 SharePoint 共享链接作为参数：

```bash
python Sharepoint_download.py
```

脚本中已包含一个示例链接，你可以直接运行测试。

### 作为模块使用

你也可以将此脚本作为模块导入，使用 `download_from_sharepoint` 函数：

```python
from Sharepoint_download import download_from_sharepoint

# SharePoint 共享链接
shared_url = 'https://your-domain.sharepoint.com/:u:/r/sites/YourSite/Shared%20Documents/path/to/file.ext?param=value'

# 下载文件到指定目录
download_from_sharepoint(shared_url, save_dir='./downloads')
```

## 函数说明

### `split_url(shared_url)`

解析 SharePoint 共享链接，提取域名、站点路径和文件路径。

参数:
- `shared_url`: SharePoint 共享链接字符串

返回:
- 元组 (domain, site_path, file_path)

### `get_site_id(domain, site_path)`

获取 SharePoint 站点 ID。

参数:
- `domain`: SharePoint 域名
- `site_path`: 站点路径

返回:
- 站点 ID 字符串或 None（如果失败）

### `get_drive_id(site_id)`

获取 SharePoint 文档库驱动器 ID。

参数:
- `site_id`: SharePoint 站点 ID

返回:
- 驱动器 ID 字符串或 None（如果失败）

### `get_file_info(site_id, drive_id, file_path)`

获取文件信息，包括下载链接。

参数:
- `site_id`: SharePoint 站点 ID
- `drive_id`: 驱动器 ID
- `file_path`: 文件路径

返回:
- 包含文件信息的字典或 None（如果失败）

### `download_file(download_url, local_path)`

从 URL 下载文件到本地路径。

参数:
- `download_url`: 文件下载 URL
- `local_path`: 本地保存路径

### `download_from_sharepoint(shared_url, save_dir='.')`

主函数，从 SharePoint 下载文件。

参数:
- `shared_url`: SharePoint 共享链接
- `save_dir`: 本地保存目录（默认为当前目录）

## 注意事项

1. 确保你的访问令牌有足够的权限访问目标 SharePoint 文件
2. 如果文件名包含特殊字符，工具会自动处理
3. 下载的文件会以 "downloadUrl_" 为前缀保存，以避免文件名冲突

## 示例

```python
from Sharepoint_download import download_from_sharepoint

# 示例 SharePoint 链接
shared_url = 'https://riedelcommunications.sharepoint.com/:u:/r/sites/SimplyLiveInternal/Shared%20Documents/R%26D/VideoEngine/TcTableAnalyzer/11.26.4.5/TcTableAnalyzer11.26.4.5.zip?csf=1&web=1&e=MnZxu8'

# 下载文件到当前目录
download_from_sharepoint(shared_url)
```

## 故障排除

- 如果遇到权限错误，请检查你的访问令牌是否有效
- 如果 URL 解析失败，请检查共享链接格式是否正确
- 如果下载失败，请检查网络连接和文件是否仍然存在

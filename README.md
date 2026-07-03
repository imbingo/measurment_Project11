# Metrology Data Platform

量测数据采集配置平台，用于在内网环境管理生产编号、量测项、指标配置、CSV/Excel/Image OCR 数据源采集、结果追溯、Dashboard 监控、审计日志和导出。

## 当前推荐版本

当前推荐入口文件：

```powershell
python .\metrology_data_platform_v2_7.py
```

版本信息：

```text
APP_VERSION = V2.7
APP_TITLE = 量测数据采集配置平台 V2.7 - 稳定化增强版
```

历史入口文件会继续保留，便于追溯和回退；新部署和试运行请使用 V2.7。

## 快速启动

PowerShell：

```powershell
.\start_metrology_v2_7.ps1
```

CMD：

```bat
start_metrology_v2_7.bat
```

默认访问地址：

```text
http://127.0.0.1:8023
```

可通过环境变量调整监听地址和端口：

```powershell
$env:MDCP_HOST="0.0.0.0"
$env:MDCP_PORT="8023"
$env:MDCP_DISPLAY_IP="192.168.1.20"
$env:MDCP_ADMIN_PASSWORD="请改成强密码"
python .\metrology_data_platform_v2_7.py
```

## 默认账号

默认测试账号：

```text
admin / admin123
```

`admin/admin123` 仅限本地测试。正式部署必须设置 `MDCP_ADMIN_PASSWORD`，也可设置 `MDCP_ADMIN_USERNAME`。

V2.7 已落地最小角色权限：

- `viewer`：查看 Dashboard、采集结果、采集日志、采集任务、模板和配置检查；不能新增、编辑、作废或手动采集。
- `engineer`：可维护生产编号、量测项、指标和模板，可测试读取和手动采集；不能管理用户，不能作废结果。
- `admin`：全部权限，包括用户管理、结果作废、日志清理和系统级操作。

## 局域网访问

服务器电脑启动时设置：

```powershell
$env:MDCP_HOST="0.0.0.0"
$env:MDCP_DISPLAY_IP="服务器内网IP"
.\start_metrology_v2_7.ps1
```

局域网其他电脑访问：

```text
http://服务器内网IP:8023
```

公司电脑如果无法访问，需要由 IT 按公司安全策略放行服务器电脑 TCP 8023 入站访问。建议仅放行内网网段或指定电脑 IP。

## 主要功能

- 生产编号管理
- 量测项和指标配置
- CSV、Excel、Image OCR 数据源采集
- 手动采集和定时采集
- `collect_job` 采集任务生命周期追踪
- 多 Sheet 模板解析和模板库
- 新增生产编号时批量套用模板
- 模板套用 snapshot 追溯
- 采集结果查询、导出 Excel
- 采集结果作废，不物理删除
- OCR 原文、ROI、图片路径和配置快照追溯
- Dashboard 统计、采集任务状态、采集失败 Top、24 小时无数据提示、OCR 追溯风险提示
- 配置完整性检查页面和 Dashboard Top 问题
- 最小用户管理和角色权限控制
- 操作审计日志

## V2.7 P1 稳定化增强

### collect_job 采集任务追踪

每次手动采集、定时采集和测试读取都会写入 `collect_job`：

- `pending`：任务已创建
- `running`：采集正在执行
- `success`：采集完成
- `failed`：采集失败
- `timeout`：读取超时
- `skipped`：同一量测项已有任务运行，跳过本次

页面入口：`采集任务`。Dashboard 会显示最近 24 小时任务成功、失败、超时和运行中数量。

### 模板套用 snapshot

每次套用模板会在 `template_apply_log.snapshot_json` 保存完整快照，包括模板名称/版本、数据源类型、生产编号字段、工序字段、Excel Sheet、表头行、指标列表和实际生成的量测项配置。后续模板修改不会改变历史快照。

### 配置完整性检查

页面入口：`配置检查`。当前检查：

- 无工序配置的量测项
- 无启用指标的量测项
- 数据源路径为空
- Excel 数据源但 Sheet 未填写
- Image OCR 配置为空或 JSON 解析失败
- 模板没有指标
- 模板缺少生产编号字段

## 数据源支持范围

### CSV

适合设备导出 `.csv` 文件，按配置的生产编号字段匹配数据行。支持常见中文编码自动回退。

### Excel

支持 `.xlsx/.xlsm`。可配置 Sheet 名称和表头所在行，适合设备或人工维护的多列量测表。

### Image OCR

支持 `.png/.jpg/.jpeg/.bmp/.tif/.tiff`。使用本地离线 OCR：`Pillow`、`opencv-python-headless`、`pytesseract` 和本机 Tesseract-OCR。

图片路径可配置为单个文件、目录或 glob。正式自动抓取建议使用“共享图片目录/inbox”：

- `数据源路径` 填设备导出图片的共享目录，例如 `\\192.168.1.100\metrology_images`。
- `collect_mode` 设为 `all_stable`，每轮扫描目录内多张已经写完的图片。
- 从文件名解析生产编号和工序，例如 `PROD_A_STEP1.png`，用 `production_code_from_filename_regex` 和 `process_from_filename_regex`。
- 如果文件名没有生产编号，也可以在图片上框选生产编号 ROI，配置 `fields.production_code` 的 OCR 正则。
- 不需要每个生产编号配置一个单独目录，也不要求目录里只有一张图片；重复图片会按来源行 hash 跳过。

示例：

```json
{
  "file_pattern": "*.png",
  "collect_mode": "all_stable",
  "max_files_per_collect": 50,
  "route_by_image_production_code": true,
  "production_code_from_filename_regex": "(?P<production_code>.+)_[^_.]+\\.[^.]+$",
  "process_from_filename_regex": "_(?P<process_step>[^_.]+)\\.",
  "ocr": {"lang": "eng", "psm": 6, "scale": 2.0, "threshold": true},
  "metrics": {
    "Rx": {"roi": [0.05, 0.10, 0.30, 0.12], "regex": "Rx\\s*[:=]?\\s*([-+]?\\d+(?:\\.\\d+)?)"},
    "Ry": {"roi": [0.05, 0.24, 0.30, 0.12], "regex": "Ry\\s*[:=]?\\s*([-+]?\\d+(?:\\.\\d+)?)"},
    "Z": {"roi": [0.05, 0.38, 0.30, 0.12], "regex": "Z\\s*[:=]?\\s*([-+]?\\d+(?:\\.\\d+)?)"}
  }
}
```

OCR 数据源必须保留追溯信息。V2.7 会尽量写入：

- `source_file_path`
- `source_file_mtime`
- `ocr_raw_text`
- `ocr_roi_json`
- `ocr_confidence`
- `ocr_config_json`

当前 OCR 置信度尚未接入逐 ROI 结果，`ocr_confidence` 可能为空。OCR 结果不能当作与 CSV/Excel 同等可信数据，产线试运行时需要结合原图和 OCR 原文复核。

## 重要注意事项

- 默认 `admin/admin123` 仅限本地测试。
- 正式部署必须设置 `MDCP_ADMIN_PASSWORD`。
- `*.db`、`*.db-wal`、`*.db-shm` 不提交仓库。
- `template_upload_cache/`、`ocr_debug/`、`exports/` 不提交仓库。
- OCR 图片识别依赖设备版式稳定，ROI 和正则必须经过现场样图验证。
- 运行本程序的电脑或服务器账号必须能读取配置的 UNC 共享路径。
- 采集结果只允许作废，不直接物理删除，以便审计和追溯。

## OCR 离线依赖

源码方式运行 OCR 前安装：

```powershell
pip install -r .\requirements_ocr.txt
```

Windows 还需要安装 Tesseract-OCR。若安装在非标准路径，可设置：

```powershell
$env:MDCP_TESSERACT_CMD="C:\Program Files\Tesseract-OCR\tesseract.exe"
```

无互联网产线电脑建议使用离线 EXE 包和公司 IT 分发流程。

## 文档入口和后续计划

- `README_EXE_OFFLINE.md`：产线离线 EXE 包说明
- `requirements_ocr.txt`：OCR Python 依赖
- `fastapi_blueprint/`：后续迁移到 FastAPI/PostgreSQL/Worker 的参考蓝图

后续产品化建议：

- 接入公司统一账号体系
- 将采集 worker 与 Web 进程拆分
- 迁移数据库到 PostgreSQL
- 增加 OCR ROI 配置可视化工具


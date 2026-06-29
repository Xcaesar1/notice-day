# notice-day

亚马逊 Seller Central 账号状况异常采集与钉钉通知工具。

当前目标是单主机生产运行：

- 采集美国站店铺账号状况异常商品
- 提取 ASIN / SKU / 问题类型 / 日期 / 当前处理
- 用 SQLite 做去重，只通知新增或核心内容变化的异常
- 按固定 Markdown 模板发送到钉钉群

## 当前状态

- 旧链路：`zclaw`
  - 兼容保留
  - 适合现有 NotePC / 旧验证路径
- 新链路：`webdriver`
  - 已接入官方紫鸟 WebDriver HTTP 接口
  - 通过 `startBrowser -> debuggingPort -> CDP -> account health API` 采集
  - 这是 Office-PC 生产推荐链路

## 目录

- 代码：`Q:\notice-day`
- 本地状态：`Q:\notice-day\.local-state\account-health-notifier`
- 示例配置：[config.example.json](Q:\notice-day\config.example.json)
- 单主机 runbook：[docs/single-primary-pc-runbook.md](Q:\notice-day\docs\single-primary-pc-runbook.md)

## 快速开始

```powershell
Set-Location Q:\notice-day
python -m pip install -r requirements.txt
python account_health_notifier.py init-config --config .local-state\account-health-notifier\config.json
python account_health_notifier.py doctor --json
python account_health_notifier.py validate-config --json
```

然后编辑本地配置：

- `collector.backend`
- `dingtalk.webhook_url`
- `dingtalk.secret`
- `dingtalk.send_enabled`
- `runtime.primary_host`

## Office-PC 推荐配置

推荐把 `collector.backend` 改成 `webdriver`，并设置：

```json
{
  "collector": {
    "backend": "webdriver"
  },
  "ziniao_webdriver": {
    "client_path": "D:\\ziniao\\ziniao.exe",
    "port": 9515,
    "company": "你的企业名",
    "username": "你的紫鸟账号",
    "password": "",
    "password_env": "ZINIAO_WEBDRIVER_PASSWORD"
  }
}
```

推荐把紫鸟密码放进环境变量，而不是写进配置文件：

```powershell
$env:ZINIAO_WEBDRIVER_PASSWORD = "你的密码"
```

如果要长期生效，请在 Office-PC 的系统环境变量里设置 `ZINIAO_WEBDRIVER_PASSWORD`。

## 常用命令

```powershell
# 只测钉钉发送链路
python account_health_notifier.py send-test --send --json

# 只做本地去重自测
python account_health_notifier.py self-test --json

# 旧链路：ZClaw 采集
python account_health_notifier.py zclaw-collect-stores --json

# 新链路：官方 WebDriver + CDP 采集
python account_health_notifier.py webdriver-collect-stores --json
python account_health_notifier.py webdriver-collect-stores --json --stores BYF --limit 1

# 生产链路：采集 -> 去重 -> 钉钉
python account_health_notifier.py production-run --dry-run --json
python account_health_notifier.py production-run --send --json

# 强制指定 backend
python account_health_notifier.py production-run --backend webdriver --dry-run --json

# 配置预检
python account_health_notifier.py validate-config --json
python account_health_notifier.py validate-config --require-send-ready --json

# 安装 6 小时任务计划
python account_health_notifier.py install-schedule --dry-run --json
python account_health_notifier.py install-schedule --json
```

## WebDriver 生产链路

当 `collector.backend=webdriver` 时，`production-run` 会执行：

1. 启动或复用 `ziniao.exe --run_type=web_driver --ipc_type=http --port=9515`
2. 调用 `updateCore`
3. 调用 `getBrowserList`
4. 过滤美国站亚马逊店铺
5. 对每个店铺调用 `startBrowser(browserOauth)`
6. 使用返回的 `debuggingPort` 连接 CDP
7. 导航到 Seller Central / 账户状况页
8. 读取当月 1 日到当天的未解决异常
9. 写入本地结果与 SQLite 去重状态
10. 按 Markdown 模板发钉钉

## 默认日期范围

默认总是“当月 1 日到今天”：

- 2026-06-29 运行时：`2026-06-01` 到 `2026-06-29`
- 2026-07-25 运行时：`2026-07-01` 到 `2026-07-25`

只有补数或排障时才显式传：

```powershell
python account_health_notifier.py webdriver-collect-stores --start-date 2026-06-01 --end-date 2026-06-29 --json
```

## 单主机约束

- 真实发送只允许 `runtime.primary_host` 指定的那台机器执行
- 不要让 NotePC 和 Office-PC 同时开真实定时任务
- 如果要迁移到 Office-PC，需要一并迁移：
  - `.local-state\account-health-notifier\state.sqlite`
  - 本地 config

否则历史异常会被当成新异常再次推送。

## 验证

本地代码验证命令：

```powershell
python -m py_compile account_health_notifier.py ziniao_webdriver.py cdp_account_health.py ziniao_cdp.py zclaw_account_health.py
python -m unittest discover -s tests -p "test_*.py" -v
```

当前仓库这套测试已通过。

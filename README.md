# notice-day

亚马逊 Seller Central 账号状况异常商品通知器。

当前版本用于 Office-PC 本地定时运行:

- 通过紫鸟浏览器 CDP 直连能力检查 Seller Central 页面是否可读。
- 从账号状况结果 Excel 读取未解决异常明细。
- 提取受影响商品的 ASIN 和 SKU。
- 用 SQLite 记录已通知项, 新增或核心内容变化才再次通知。
- 通过 `Q:\Dingcli\dws-call.cmd` 调用钉钉 DWS 应用机器人发送 Markdown 到指定群。

## 目录约定

- 代码目录: `Q:\notice-day`
- 默认状态目录: `Q:\notice-day\.local-state\account-health-notifier`
- 默认 DWS 入口: `Q:\Dingcli\dws-call.cmd`
- 默认 Excel 结果目录: `C:\Users\god\Desktop\RPA下载结果\账户状况异常明细`

`.local-state/` 不提交到 Git, 用于保存本地配置、SQLite 状态库、DWS 参数临时文件和运行结果。

## 初始化

```powershell
Set-Location Q:\notice-day
python account_health_notifier.py init-config --config .local-state\account-health-notifier\config.json
python account_health_notifier.py doctor --json
python account_health_notifier.py validate-config --json
python account_health_notifier.py validate-config --require-send-ready --json
```

初始化后编辑 `.local-state\account-health-notifier\config.json`, 填入:

- `dingtalk.robot_code`
- `dingtalk.group_open_conversation_id`
- 确认 `dingtalk.send_enabled` 是否为 `true`

不要把真实配置复制进仓库文件。

## 常用命令

```powershell
# 本地自测: 验证去重、内容变化重新通知和 dry-run 链路
python account_health_notifier.py self-test --json

# 从最新 Excel dry-run, 不真实发钉钉
python account_health_notifier.py run --dry-run --json

# 配置预检, 第二条会要求机器人和群配置满足真实定时通知
python account_health_notifier.py validate-config --json
python account_health_notifier.py validate-config --require-send-ready --json

# 紫鸟 CDP 直连 smoke: 不读 Cookie/token, 只读当前 Seller Central 页面的 title/url/body 摘要
python account_health_notifier.py cdp-smoke --json
python ziniao_cdp.py probe --json --port 9222

# CDP 能力边界诊断和心跳观察
python account_health_notifier.py cdp-doctor --json
python account_health_notifier.py cdp-watch --duration-seconds 20 --interval-seconds 2 --json
python account_health_notifier.py cdp-lifecycle-test --json

# CDP single-store account-health collection from the current Ziniao Seller Central window.
# Default date range is current month, e.g. 2026-06-01 to 2026-06-25 on 2026-06-25.
python account_health_notifier.py cdp-collect-current --json --port 9222
python account_health_notifier.py cdp-collect-current --json --port 9222 --categories safe --start-date 2026-06-01 --end-date 2026-06-25

# CDP multi-open-window collection. It scans currently open Ziniao Seller Central targets,
# collects current-month issues, writes Excel/JSON, and marks unopened US stores as missing.
python account_health_notifier.py cdp-collect-open --json
python account_health_notifier.py cdp-collect-open --json --require-all-open
python account_health_notifier.py cdp-collect-open --json --categories safe,restricted --output-dir .local-state\account-health-notifier\runs\cdp-open-smoke

# 安装紫鸟 CDP daemon 登录自启, 先 dry-run 查看 schtasks 命令
python account_health_notifier.py install-cdp-daemon --dry-run --json
python account_health_notifier.py install-cdp-daemon --json

# sample 自测不代表全店铺覆盖, 需要显式跳过覆盖校验
python account_health_notifier.py run --source-type sample --dry-run --skip-store-coverage --json

# 全店铺解析, 只生成报告, 不触发钉钉和去重状态
python account_health_notifier.py parse --json
python account_health_notifier.py parse --source-dir "C:\Users\god\Desktop\RPA下载结果\账户状况异常明细" --json
python account_health_notifier.py parse --source-dir "C:\Users\god\Desktop\RPA下载结果\账户状况异常明细" --require-all-stores --json

# 指定 Excel dry-run
python account_health_notifier.py run --source-excel "C:\path\账户状况监控_20260624_120000.xlsx" --dry-run --json

# 发送一条机器人测试消息
python account_health_notifier.py send-test --send --json

# 安装 6 小时一次的 Windows 任务计划, 先 dry-run 查看命令
python account_health_notifier.py install-schedule --dry-run --json
python account_health_notifier.py install-schedule --json
```

启用真实定时通知前, 建议先运行 `parse --require-all-stores`: 只有 `coverage_ok=true` 时才代表本次结果覆盖店铺清单里的全部美国站店铺。
正式 `run` 默认启用 `require_all_stores_before_send`; 如果店铺清单无法读取或美国站店铺缺失, 会返回 `coverage_failed` 并且不会进入钉钉发送。
`install-schedule` 默认会执行 `--require-send-ready` 级别预检; 缺少机器人, 群 ID 或 `send_enabled=true` 时会返回 `preflight_failed`, 不会安装任务计划。

## 安全边界

- CDP 能读取浏览器页面和登录态上下文, 排障命令默认不读取 Cookie/token, 日志和提交中也不能写入 Cookie/token。
- 紫鸟 CDP 守护进程必须先于店铺浏览器启动; 已打开且未带 `--remote-debugging-port` 的旧窗口需要关闭后重开。
- `cdp-lifecycle-test` 默认不关闭窗口; 只有显式加 `--close-target` 才会关闭当前 Seller Central target。
- daemon 日志只记录 PID, 端口, target 标题和 URL, 不记录 Cookie/token/localStorage/sessionStorage。
- 脚本不创建群, 不添加机器人, 不修改群成员。
- 缺少 `robot_code` 或 `group_open_conversation_id` 时不能真实发送。
- `send_enabled=false` 时, `run` 默认 dry-run; 需要真实发送时使用 `--send` 或将配置改为 `true`。
- 单次发送失败不会标记为已通知, 下一轮会重试。
- DWS 返回非 0 或入口缺失会记录为 `send_failed`, 不会被误判为 dry-run 成功。

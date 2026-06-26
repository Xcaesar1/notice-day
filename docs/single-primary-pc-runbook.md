# 单主机生产运行手册

## 结论

当前项目固定采用单主机生产模式:

- NotePC: 开发, 调试, 稳定性验证.
- Office-PC: 唯一生产运行机器, 唯一钉钉真实发送机器.

不要让两台 PC 同时启用真实定时发送. 多机同时运行会导致状态库不一致, 进而重复通知同一批异常.

## 生产运行策略

生产任务在 Office-PC 上每 6 小时运行一次:

```powershell
Set-Location Q:\notice-day
python account_health_notifier.py run --send --json
```

正式运行前必须通过:

```powershell
python account_health_notifier.py doctor --json
python account_health_notifier.py validate-config --require-send-ready --json
python account_health_notifier.py self-test --json
```

安装 Windows 任务计划前先 dry-run:

```powershell
python account_health_notifier.py install-schedule --dry-run --json
python account_health_notifier.py install-schedule --json
```

## 采集范围

默认日期范围是当月 1 日到当天:

- `2026-06-25` 运行时采集 `2026-06-01` 到 `2026-06-25`.
- `2026-07-25` 运行时采集 `2026-07-01` 到 `2026-07-25`.

不要在定时任务里写死月份日期. 只有临时补数或排障时才手动传 `--start-date` 和 `--end-date`.

## 耗时预估

已验证的 14 个美国站店铺串行采集耗时:

- 快速轮: 约 3.65 分钟.
- 慢速轮: 约 5.75 分钟.

生产设计按 10-20 分钟容错. 当前 `run_lock_ttl_minutes` 为 240, 足够避免重复任务重叠.

## 固定流程

1. 读取店铺清单, 只处理美国站.
2. 单店串行打开和采集, 不启用多店铺并发启动.
3. 日期范围使用当月 1 日到当天.
4. 解析未解决账号状况异常里的 ASIN/SKU.
5. 使用 SQLite 状态库去重.
6. 只推送新增或核心内容变化的异常.
7. 使用 `field-block-v1` 钉钉 Markdown 模板.
8. 无新增异常时不发群, 只写本地结果和日志.

## 禁止项

- 不要同时在 NotePC 和 Office-PC 启用 `send_enabled=true`.
- 不要同时安装两台机器的 6 小时任务计划.
- 不要用 CDP daemon 作为正式采集主链路. 实测它会干扰紫鸟 `store open`.
- 不要并发调用多个 `ziniao-cli store open`. 实测会让 ZClaw Bridge 短暂不可用.
- 不要提交 `.local-state`, webhook, secret, SQLite 状态库或运行结果.

## 迁移到 Office-PC

在 Office-PC 上执行:

1. 拉取仓库到 `Q:\notice-day`.
2. 安装 Python 依赖: `pip install -r requirements.txt`.
3. 安装并配置 `ziniao-cli`.
4. 登录紫鸟客户端, 确认能打开 14 个美国站店铺.
5. 初始化本机配置:

```powershell
python account_health_notifier.py init-config --config .local-state\account-health-notifier\config.json
```

6. 填写 `.local-state\account-health-notifier\config.json`:

- `dingtalk.webhook_url`
- `dingtalk.secret`
- `dingtalk.send_enabled=true`
- `source.store_list_path`

7. 从 NotePC 迁移 SQLite 去重库到 Office-PC:

```text
.local-state\account-health-notifier\state.sqlite
```

如果不迁移该文件, Office-PC 第一次运行会把历史异常当成新增重新推送.

8. 在 Office-PC 上做 dry-run 和真实测试:

```powershell
python account_health_notifier.py run --dry-run --json
python account_health_notifier.py send-test --send --json
```

9. 确认无误后, 只在 Office-PC 安装任务计划.

10. NotePC 保留 `send_enabled=false`, 不安装生产任务计划.

## 验收标准

- Office-PC 每 6 小时运行一次.
- 每轮能覆盖 14 个美国站店铺.
- 当前月无新增异常时不发群.
- 有新增或变化时, 钉钉只收到最新异常.
- 消息格式符合 `docs/dingtalk-markdown-template.md`.
- 失败店铺写入本地结果, 不阻断后续店铺.

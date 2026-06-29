# 单主机生产运行手册

## 结论

生产只允许一台机器负责真实采集和真实钉钉发送。

当前推荐：

- 开发与排障：NotePC
- 正式生产：Office-PC

真实生产链路推荐使用 `collector.backend=webdriver`。

## 为什么必须单主机

如果两台机器同时运行真实任务，会出现：

- 同一批异常重复推送
- SQLite 去重状态不一致
- 定时任务互相覆盖定位
- 真实发送责任边界不清楚

所以必须在配置里固定：

```json
{
  "runtime": {
    "primary_host": "OFFICE-PC",
    "enforce_primary_host_for_send": true
  }
}
```

## Office-PC 正式配置

### 1. 准备目录

建议：

- 项目目录：`E:\notice-day`
- 本地状态目录：`E:\notice-day\.local-state\account-health-notifier`
- 店铺清单：`E:\notice-day\data\store-list.xlsx`

### 2. 安装依赖

```powershell
Set-Location E:\notice-day
python -m pip install -r requirements.txt
npm install -g @ziniao-open/cli
ziniao-cli --version
```

### 3. 初始化配置

```powershell
python account_health_notifier.py init-config --config .local-state\account-health-notifier\config.json
```

### 4. 配置 collector=webdriver

本地 `config.json` 至少填写：

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
  },
  "dingtalk": {
    "method": "webhook",
    "webhook_url": "你的 webhook",
    "secret": "你的 secret",
    "send_enabled": true
  },
  "runtime": {
    "primary_host": "OFFICE-PC",
    "enforce_primary_host_for_send": true
  }
}
```

推荐把紫鸟密码写进环境变量：

```powershell
$env:ZINIAO_WEBDRIVER_PASSWORD = "你的密码"
```

## 迁移旧去重状态

如果 NotePC 上已经跑过真实通知，要把旧状态库迁移到 Office-PC：

```text
.local-state\account-health-notifier\state.sqlite
```

否则 Office-PC 第一次会把历史未解决异常重新当作新增推送。

## 上线前检查

先跑：

```powershell
python account_health_notifier.py doctor --json
python account_health_notifier.py validate-config --json
python account_health_notifier.py validate-config --require-send-ready --json
python account_health_notifier.py self-test --json
```

然后做一次真实采集 dry-run：

```powershell
python account_health_notifier.py production-run --backend webdriver --dry-run --json
```

如果只想先做小验证：

```powershell
python account_health_notifier.py webdriver-collect-stores --stores BYF --limit 1 --json
```

## 安装 6 小时任务计划

```powershell
python account_health_notifier.py install-schedule --dry-run --json
python account_health_notifier.py install-schedule --json
```

任务计划最终执行的是：

```powershell
python account_health_notifier.py production-run --config .local-state\account-health-notifier\config.json
```

## 运行边界

- 默认只采集“当月 1 日到今天”
- 只处理美国站
- 失败店铺会写入本地结果，但本轮不会误报“全部正常”
- 没有新增异常时不发群

## 故障定位

### validate-config 报 webdriver 缺项

检查：

- `collector.backend=webdriver`
- `ziniao_webdriver.client_path`
- `ziniao_webdriver.company`
- `ziniao_webdriver.username`
- `ziniao_webdriver.password` 或 `password_env`

### WebDriver 端口起不来

先确认：

- `D:\ziniao\ziniao.exe` 路径正确
- 9515 端口未被占用
- 没有残留的普通紫鸟主进程干扰

### startBrowser 成功但页面是 about:blank

这是已知场景。当前实现会自动继续通过 CDP 导航到 Seller Central，再进入账户状况采集。

### Amazon 出现登录 / 切换账户 / 二次验证

当前实现会把这类页面识别成“业务未就绪”，在结果里返回可读错误。正式无人值守前，先确保该批店铺在 Office-PC 上已经具备稳定登录态。

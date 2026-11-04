# PLEX 每日商机工具

销售顾问打开就能看"今天该干什么"的本地仪表盘。**不依赖豆包、不花一分钱**。

## 这个工具做什么

- 每天早上自动抓**巨潮资讯网免费公告接口**，找"汽车零部件企业 + 海外建厂（泰国/墨西哥/匈牙利/塞尔维亚/摩洛哥）+ 增资/设立/投产"相关新公告；
- 对每条公告跑冻结规则：MES 开发外派＝自建＝否决；MES 实施选型＝可打；注册资本小但增资大＝壳公司警告；绿地新厂＝IATF 需求窗口；
- 把此前两轮已验证的 12 家合格商机 + 6 家否决样本作为**种子数据**内置，仪表盘永远不是空的；
- 每家 case 一张应对卡，分两半：
  - ✅ **系统已自动完成**：分数、决策人、杠杆机构、园区落点、证据链接；
  - 🖐️ **你需要手动做**：具体到动作（约九慧信息转介、上门泰中罗勇园区、首通话问 IATF USI……）。

**全免费数据源**：巨潮资讯网（cninfo）公开公告接口。
**不爬的东西**（已在仪表盘"人工步骤区"标明）：猎聘、BOSS 直聘、脉脉、领英、iatf-customerportal——这些需登录或有反爬，ToS 风险，工具不碰，你自己登录看。

## 一分钟开始用（本地）

1. 双击 `run_daily.bat`。
2. 跑完会自动打开 `dashboard.html`。
3. 手机想随时看：把同一个 `dashboard.html` 发到微信文件传输助手/邮件附件，它是单文件自包含网页。

## 让它每天自动跑（二选一）

### 方式 A：Windows 任务计划程序（本机自动跑）

一条命令搞定（在 PowerShell 里跑一次，不需要管理员）：

```powershell
schtasks /Create /TN "PLEX每日商机" /TR "\"C:\Users\Lenovo\Doubao\chats\2026-11-04\new-chat\PLEX每日商机工具\run_daily.bat\"" /SC DAILY /ST 09:00 /RL LIMITED /F
```

以后每天 9:00 自动跑，跑完不弹窗（脚本静默生成文件）。想看结果直接双击 `dashboard.html`。

取消：`schtasks /Delete /TN "PLEX每日商机" /F`

### 方式 B：GitHub Actions（云端免费跑，手机随时看）

适合你不想让电脑开机的场景。Actions 每天 UTC 01:00（北京时间 09:00）自动抓、自动把 dashboard.html 部署到 GitHub Pages，**手机浏览器直接打开网址就能看**。

一次性 10 分钟设置：

1. 注册免费 GitHub 账号（已有跳过）。
2. 新建一个 Private 或 Public repo，名字随便（例如 `plex-daily`）。
3. 把整个 `PLEX每日商机工具\` 文件夹里的所有内容 push 上去（GitHub 网页端 "uploading an existing file" 拖拽即可，包含 `.github/workflows/update.yml`）。
4. 进入 repo → **Settings → Pages** → Source 选 "GitHub Actions"。
5. 进入 repo → **Actions** 标签 → 选 "daily-update" → 点 "Run workflow" 手动跑一次，确认绿色对勾。
6. 跑完后 Pages 地址是 `https://<你的用户名>.github.io/<repo名>/dashboard.html`，加到手机书签。

之后不用管，每天 9 点自动更新。

## 可选：免费推送到微信/手机

如果配了环境变量，脚本跑完会发一句话摘要到你手机；没配就静默跳过，不报错。

**Server酱（推送到微信）**：
1. 微信扫码登录 https://sct.ftqq.com/ （用微信一键登录即可，免费版每天 5 条）。
2. 复制你的 SENDKEY。
3. 本地：在 PowerShell 里 `setx SERVERCHAN_KEY "你的key"`，重开终端生效。
   GitHub：repo → Settings → Secrets and variables → Actions → New repository secret，名字 `SERVERCHAN_KEY`。

**Bark（推送到 iPhone）**：
1. App Store 装 Bark，打开复制你的 Key（形如 `https://api.day.app/xxxxxxx`）。
2. 同上把 `BARK_KEY` 配成本地环境变量或 GitHub Secret。

**SITE_URL**（可选）：把 GitHub Pages 地址（如 `https://you.github.io/plex-daily/dashboard.html`）配成 Secret，推送消息里会直接带链接。

## 目录结构

```
PLEX每日商机工具\
├── dashboard.html          ← 双击即开的仪表盘（脚本每次运行覆盖）
├── leads.xlsx              ← 每 case 一行的台账
├── run_daily.bat           ← 本地一键运行
├── scripts\
│   └── run_daily.py        ← 唯一脚本：抓取 + 规则 + 生成 HTML/XLSX + 推送
├── data\                   ← 11 个种子 JSON（两轮已验证研究）
├── logs\                   ← 每次运行日志
└── .github\workflows\update.yml   ← GitHub Actions 配置
```

## 诚实标注

仪表盘每条数据都标了来源：
- 🟦 **种子：已验证研究**——来自前两轮逐家实测，证据链接可点；
- 🟩 **今日抓取**——来自巨潮资讯网，标题级弱推断，行业是否汽车零部件需人工判断；
- 🟧 **人工**——工具不代做（招聘平台/领英/IATF 门户）。

抓不到数据时仪表盘会写"今日 0 条（源不可达或无新公告）"，**不会编造**。

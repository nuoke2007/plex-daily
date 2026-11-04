# -*- coding: utf-8 -*-
"""
PLEX 每日商机工具 · 本地每日运行脚本
- 全免费、本地运行：唯一外部源是巨潮资讯网免费公告查询接口（cninfo），失败优雅降级。
- 不爬任何登录墙/反爬平台（猎聘/BOSS/脉脉/领英/IATF customerportal 一律不碰，标人工）。
- 产物：dashboard.html（双击即开）+ leads.xlsx + logs/run_YYYYMMDD.log
"""
import json, os, re, sys, datetime, urllib.request, urllib.parse
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
LOGS = ROOT / "logs"
LOGS.mkdir(exist_ok=True)
TODAY = datetime.date.today().isoformat()
NOW = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
log_lines = []

def log(msg):
    line = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}"
    log_lines.append(line)
    print(line)

# ============================================================
# 1. 内置种子 case（来自此前已完成的两轮验证；不要让仪表盘是空的）
# ============================================================
# feas: 高/中/低；prio: 高优先/中/低
CASES = [
 dict(id="M-02", name="苏州瑞玛精密", code="002976.SZ", base="墨西哥 瓜纳华托 Apaseo El Grande",
      score=8, grade="C", chief="董事长/总经理 陈晓敏（已查证）", feas="高",
      leverage="园区运营方 JLB Parques（瑞玛地块出让方）+ 现有客户劳士领 Silao 厂（同州 50 公里）",
      park="Colinas de Apaseo / JLB Parques",
      system_done="已查证决策人；已确认园区运营方；已确认劳士领 Silao 厂存在且与瑞玛同州；ICP=8（C 级）",
      manual=["联系劳士领 Silao 厂（公开厂长 Jorge Navarrete），请其向瑞玛 Apaseo 厂引荐",
              "上门 JLB Parques（瑞玛 2025-09 完成 52,297 平米购地交割，出让方握租户名单）",
              "首通电话直接问：海外新厂 IATF USI 是多少、计划由哪家 CB 首审"],
      evidence=["https://www.cheersson.com/article/category.php?id=7"], source="seed"),
 dict(id="M-03", name="浙江银轮股份", code="002126.SZ", base="墨西哥 蒙特雷（新设子公司）",
      score=10, grade="C", chief="董事长 徐小敏（已查证）", feas="高",
      leverage="九慧信息（同时做银轮 2015 S/4HANA + 保隆 2021 SAP，有商业协同）+ 与中原内配互背书 + AMS 上海",
      park="蒙特雷（园区未披露）",
      system_done="已查证决策人；已确认九慧为其 SAP 实施商；已确认与中原内配共享康明斯/戴姆勒/约翰迪尔；已确认 2025-12 增资 2.69 亿元",
      manual=["约九慧信息的 SAP 顾问转介（它做完 SAP 主干，客户下一步就要上 MES，与 PLEX 互补）",
              "决定先签银轮还是先签中原内配（两家共享三家主机厂客户，先签一家才能背书另一家）",
              "Automechanika Shanghai 2026（12/2–5）直奔银轮展位"],
      evidence=["https://www.yinlun.cn/Products-and-Service"], source="seed"),
 dict(id="T-01", name="飞龙股份", code="002536.SZ", base="泰国 罗勇 Ban Khai Nong Lalok（EEC）",
      score=11, grade="B", chief="副董事长/总经理 孙耀忠（已查证）", feas="高",
      leverage="泰中罗勇工业园运营方（华立+Amata 合资，中国公司）",
      park="泰中罗勇工业园 TCR（官网直认飞龙为二区租户）",
      system_done="已查证决策人；已确认园区运营方；已确认 2025Q3 投产爬坡；ICP=11（B 级）",
      manual=["上门泰中罗勇工业园开发有限公司——一次上门可同时触达飞龙与中原内配两家",
              "问飞龙：泰国厂 MES 是否已选型、是否会复制国内版本",
              "新厂 site 级 IATF 证书通常滞后投产 6–18 个月，正是 MES 立项窗口"],
      evidence=["https://www.tcrindustrial.com/"], source="seed"),
 dict(id="T-02", name="中原内配", code="002448.SZ", base="泰国 罗勇 Ban Khai Hwy 3143",
      score=11, grade="B", chief="董事长 薛亚辉（已查证）", feas="高",
      leverage="泰中罗勇工业园 + 与银轮互背书（共享康明斯/戴姆勒/约翰迪尔）",
      park="泰中罗勇工业园 TCR（2024-01 与运营方签土地购买协议）",
      system_done="已查证决策人；已确认园区落点；已确认已批量供货通用/康明斯/戴姆勒/约翰迪尔",
      manual=["同飞龙：上门泰中罗勇园区运营方，一次见两家",
              "与银轮的背书顺序由你定（建议先签银轮）",
              "用通用/康明斯/戴姆勒/约翰迪尔二方审核倒推 MES 立项时间"],
      evidence=["http://paper.cnstock.com/html/2026-08/27/content_2261032.htm"], source="seed"),
 dict(id="E-03", name="保隆科技", code="603197.SH", base="匈牙利 Érd",
      score=13, grade="B", chief="集团 CIO（姓名待验证）", feas="高",
      leverage="BOSS 直聘在招 MES 实施岗（唯一）+ 九慧信息 + HIPA + AMS 上海",
      park="匈牙利（HIPA 2023-09 公开报道揭幕）",
      system_done="已确认 BOSS 在招“系统实施高级顾问（MES 方向）（驻匈牙利）”（已关闭投递）；已确认九慧做过其 SAP；年报有 SAP 升级原文",
      manual=["联系 BOSS 岗位 HR 吴女士，问匈牙利项目 MES 是否仍在选型",
              "约九慧信息（同时是银轮实施商，一次见两家）",
              "Automechanika Shanghai 2026（12/2–5）直奔保隆展位",
              "通过 HIPA 引荐匈牙利本地运营团队"],
      evidence=["https://hipa.hu/"], source="seed"),

 dict(id="M-04", name="祥鑫科技", code="002965.SZ", base="墨西哥 蒙特雷",
      score=10, grade="C", chief="董事长 谢祥娃（已查证）", feas="中",
      leverage="国金证券（同时服务瑞玛/祥鑫/时代新材）+ Cumulus 蒙特雷本地 SI",
      park="蒙特雷（园区未披露）",
      system_done="已查证决策人；已确认国金为其再融资保荐；已确认 2025-06 增资 1,500 万美元",
      manual=["问祥鑫董办：蒙特雷落在哪个园区、是否与银轮同园",
              "通过国金项目组尝试引荐（接受关系衰减，不强求）"],
      evidence=["https://www.jinge.com.cn/"], source="seed"),
 dict(id="T-03", name="科达利", code="002850.SZ", base="泰国 罗勇（注册地曼谷）",
      score=13, grade="B", chief="董事长 励建立（已查证）", feas="中",
      leverage="中金公司（2017 IPO 服务至今，同时覆盖飞龙/继峰/均胜共 4 家）",
      park="泰国地址“以最终确定为准”，园区未定",
      system_done="已查证决策人；已确认集团 SAP；已确认 2026-01 仍在建；已识别“500 万泰铢注册资本=壳”警告",
      manual=["向科达利董办或中金项目组确认泰国厂最终落点园区",
              "注意：泰国注册时注册资本仅 500 万泰铢，真实投资是后续增资到 5 亿泰铢——不要按注册资本低估"],
      evidence=["https://www.kedali.com/"], source="seed"),
 dict(id="T-04", name="斯菱智驱", code="301550.SZ", base="泰国 大城府 洛加纳工业园",
      score=9, grade="C", chief="董事长 姜岭（已查证）", feas="中",
      leverage="洛加纳工业园运营方 + Stanley 泰国厂（Pathumthani，30–50 公里）",
      park="洛加纳工业园 Rojana（官网+深交所公告直认）",
      system_done="已查证决策人；已确认园区落点；已确认 Stanley 泰国厂距大城府 30–50 公里",
      manual=["通过 Stanley 中国（天津/武汉/广州）或泰国厂联系人，向斯菱大城府团队引荐",
              "洛加纳是泰资大园区、非中资专属，引荐意愿需当面验证"],
      evidence=["https://www.rojana.com/"], source="seed"),
 dict(id="E-01", name="继峰股份", code="603997.SH", base="德国 Ursensollen",
      score=14, grade="A", chief="集团 CIO（姓名待验证）", feas="中",
      leverage="海通证券（2019 格拉默并购独立财务顾问）+ 容诚（审计）",
      park="德国存量基地，非绿地园区",
      system_done="已确认集团模具车间已用益模 MES（局部，非整椅线）；已确认海通为格拉默并购财顾；ICP=14（A 级）",
      manual=["找海通并购项目组问格拉默交易后续 IT 整合情况",
              "首通话直接问：整椅新厂是否会继承格拉默欧洲老厂 MES 模板",
              "对德国本地收件人不要发冷邮件（UWG §7），改走 IAA Mobility 2027 + 领英 InMail"],
      evidence=["https://www.nb-jf.com/"], source="seed"),
 dict(id="E-04", name="均胜电子", code="600699.SH", base="摩洛哥",
      score=14, grade="A", chief="集团 CIO（姓名待验证）", feas="中",
      leverage="DQS（摩洛哥 site IATF 发证机构）+ 中金公司",
      park="摩洛哥（AMDIE，270 家汽车供应商）",
      system_done="已确认 AUMOVIO Systems Morocco 由 DQS 发 IATF 证书（2026-09 可查）；已确认中金为 2022 定增保荐；ICP=14（A 级）",
      manual=["通过 DQS 摩洛哥审核组建立联系（审核员最了解现场管理水平）",
              "登录 Joyson Safety 全球门户查摩洛哥法语区 IT 岗（需人工登录，工具不代爬）"],
      evidence=["https://www.amdie.gov.ma/"], source="seed"),

 dict(id="M-01", name="上海永茂泰", code="605208.SH", base="墨西哥（州未披露）",
      score=10, grade="C", chief="董事长 徐宏（已查证）", feas="低",
      leverage="天健会计（2020 至今年审）+ 华泰联合（2021 IPO 保荐）",
      park="墨西哥州与园区均未披露，园区通道断",
      system_done="已查证决策人；五个通道四个未命中；年报“深化 MES 应用”不写品牌，V1 无法判定",
      manual=["打听厂房租赁的物业方，拿到基地所在州",
              "直接发董办邮件（ymtauto@ymtauto.com）+ 30 分钟电话确认国内是否已有 MES 品牌——关系法对这家无效，回到邮件"],
      evidence=["https://www.ymtauto.com/"], source="seed"),
 dict(id="E-02", name="时代新材", code="600458.SH", base="塞尔维亚 Niš",
      score=10, grade="C", chief="集团 CIO（姓名待验证）", feas="低",
      leverage="塞尔维亚 RAS + Niš 工业区（星宇车灯已在那）",
      park="Niš 工业区（RAS 公开报道 BOGE 4 万平米绿地）",
      system_done="已确认 RAS 为真实招商机构；已确认 Niš 已有中资汽车零部件集聚",
      manual=["通过星宇车灯（同在 Niš）的联系人引荐——比 RAS 更短",
              "远离所有现有客户，背书最弱，优先度排最后"],
      evidence=["https://ras.gov.rs/"], source="seed"),
]

REJECTED = [
 dict(id="MX-01", name="新泉股份", code="603179.SH", why="V1 命中：可转债问询回复披露 MES 为西门子（400 万元投资）、ERP 金蝶。已有重型 MES，饱和客户。"),
 dict(id="MX-02", name="爱柯迪", code="600933.SH", why="V1 命中：在招“MES 开发工程师”要求长期外派墨西哥＝自建 MES 团队。"),
 dict(id="TX-01", name="浦林成山", code="1809.HK", why="轮胎属流程型，与 PLEX 离散工单模型不适配；泰国基地复产，无新增立项窗口。"),
 dict(id="TX-02", name="奇精机械", code="603677.SH", why="汽车零部件仅占营收 16%，主业洗衣机离合器（家电 77%），泰国扩产服务家电出口。"),
 dict(id="EX-01", name="宁波华翔", code="002048.SZ", why="2025-02 以 1 欧元向德国 Mutares 出售欧洲 6 家子公司＝反向剥离。"),
 dict(id="EX-02", name="森麒麟", code="002984.SZ", why="轮胎流程型；智能制造经验已从青岛复制到泰国/摩洛哥。"),
]

# ============================================================
# 2. 每日自动发现（best-effort；失败优雅降级）
# ============================================================
OVERSEAS_KW = ["泰国", "墨西哥", "匈牙利", "塞尔维亚", "摩洛哥", "东欧", "海外", "境外", "出海"]
ACTION_KW = ["增资", "设立", "对外投资", "购地", "新建", "扩建", "投产", "子公司", "生产基地", "工厂"]
# 单关键词检索（多词 AND 太严会漏），动作关键词在客户端过滤
AUTO_KEYWORDS = ["泰国", "墨西哥", "匈牙利", "塞尔维亚", "摩洛哥"]

def fetch_cninfo(searchkey, column="szse", days=14, page_size=15):
    url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
    se_date = f"{(datetime.date.today()-datetime.timedelta(days=days)).isoformat()}~{datetime.date.today().isoformat()}"
    payload = urllib.parse.urlencode({
        "pageNum": 1, "pageSize": page_size, "column": column, "tabName": "fulltext",
        "searchkey": searchkey, "seDate": se_date, "isHLtitle": "true",
    }).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": "http://www.cninfo.com.cn/new/fulltextSearch?notautosubmit=&keyWord=" + urllib.parse.quote(searchkey),
    })
    with urllib.request.urlopen(req, timeout=12) as r:
        return json.loads(r.read().decode("utf-8"))

def auto_rule(title):
    flags = []
    flags.append("行业是否为汽车零部件离散制造：需人工判断（工具不猜）")
    if ("MES" in title or "制造执行" in title) and any(k in title for k in ["开发", "二次开发", "外派", "驻外", "驻海外"]):
        flags.append("V1 命中（自建 MES 团队，不打）")
    if ("MES" in title or "制造执行" in title) and any(k in title for k in ["运维", "实施", "选型", "上线"]):
        flags.append("MES 采购信号（可打）")
    if "注册资本" in title and ("增资" in title or "总投资" in title):
        flags.append("壳公司警告（注册资本小≠实际投资，查后续增资公告）")
    if any(k in title for k in ["设立", "新建", "绿地", "生产基地"]):
        flags.append("新厂 IATF site 证书通常滞后投产 6–18 个月＝MES 需求窗口")
    return flags

new_findings = []
source_status = "ok"
source_note = ""
seen_titles = set()

for kw in AUTO_KEYWORDS:
    for col in ("szse", "sse"):
        try:
            data = fetch_cninfo(kw, column=col)
            anns = data.get("announcements") or []
            hit_this = 0
            for a in anns:
                title = re.sub(r"<[^>]+>", "", a.get("announcementTitle", "") or "")
                if not title or title in seen_titles:
                    continue
                seen_titles.add(title)
                # 必须同时含动作关键词；行业是否为汽车零部件交给人工
                if not any(k in title for k in ACTION_KW):
                    continue
                if any(b in title for b in ["回购", "减持", "增持", "分红", "股东会", "法律意见书", "评级", "股东大会", "摘要"]):
                    continue
                ann_id = a.get("announcementId")
                link = f"http://www.cninfo.com.cn/new/disclosure/detail?stockCode={a.get('secCode')}&announcementId={ann_id}"
                new_findings.append(dict(
                    sec=a.get("secCode"), name=a.get("secName"), title=title,
                    date=a.get("announcementTime"), url=link,
                    rule=auto_rule(title),
                ))
                hit_this += 1
            log(f"关键词“{kw}”({col})：拉到 {len(anns)} 条，新增命中 {hit_this} 条")
        except Exception as e:
            source_status = "degraded"
            source_note = f"抓取失败：{type(e).__name__}: {e}"
            log(f"关键词“{kw}”({col})失败：{e}")

log(f"今日自动发现：{len(new_findings)} 条（源状态：{source_status}）")

# ============================================================
# 3. 生成 dashboard.html
# ============================================================
def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# 手动项总数
manual_total = sum(len(c["manual"]) for c in CASES)

CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;background:#fff;color:#1a1a1a;font-size:13.5px;line-height:1.6}
.wrap{max-width:1280px;margin:0 auto;padding:32px 36px 60px}
.top{display:flex;align-items:baseline;gap:16px;flex-wrap:wrap;margin-bottom:6px}
h1{font-size:24px;color:#003DA5;letter-spacing:-.3px}
.sub{color:#5a6b7d;font-size:12.5px;margin-bottom:20px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:26px}
.kpi{border:1px solid #e3ebf3;border-radius:8px;padding:12px 16px;background:#fafcfe}
.kpi .n{font-size:22px;font-weight:700;color:#003DA5}
.kpi .l{font-size:11.5px;color:#5a6b7d;margin-top:2px}
.sec{margin-bottom:30px}
.sec-h{font-size:16px;color:#003DA5;border-bottom:2px solid #003DA5;padding-bottom:7px;margin-bottom:14px;font-weight:700}
.tag{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600;margin-right:4px}
.tag-seed{background:#e8f0fa;color:#0059b3}
.tag-auto{background:#e8f6ee;color:#0a7f3f}
.tag-manual{background:#fdf3e3;color:#b26a00}
.card{border:1px solid #e3ebf3;border-radius:10px;margin-bottom:12px;overflow:hidden}
.card-h{display:flex;align-items:center;gap:10px;padding:11px 16px;background:#f7fafd;cursor:pointer}
.card-h .cid{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:#00B5E2;font-weight:700;min-width:38px}
.card-h .nm{font-size:14px;font-weight:700;color:#003DA5}
.card-h .code{font-size:11px;color:#8494a4}
.card-h .base{font-size:12px;color:#5a6b7d;margin-left:4px}
.card-h .right{margin-left:auto;display:flex;gap:8px;align-items:center}
.feas{font-size:11.5px;font-weight:700;padding:2px 9px;border-radius:5px}
.feas-高{background:#e8f6ee;color:#0a7f3f}
.feas-中{background:#fdf3e3;color:#b26a00}
.feas-低{background:#fdeaea;color:#b02a2a}
.grade{font-size:11px;color:#8494a4}
.card-b{padding:14px 16px;display:none}
.card.open .card-b{display:block}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.col{border-radius:8px;padding:12px 14px}
.col-auto{background:#f4faf6;border-left:3px solid #0a7f3f}
.col-manual{background:#fff9ec;border-left:3px solid #b26a00}
.col h4{font-size:12.5px;margin-bottom:8px}
.col-auto h4{color:#0a7f3f}
.col-manual h4{color:#b26a00}
.col ul{margin-left:18px;font-size:12.5px}
.col li{margin-bottom:5px}
.row{padding:4px 0;font-size:12.5px;color:#33475b}
.row b{color:#1f2d3d}
table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:10px}
th{background:#003DA5;color:#fff;padding:7px 9px;text-align:left;font-weight:600}
td{padding:7px 9px;border-bottom:1px solid #e8eef5;vertical-align:top}
tr:nth-child(even) td{background:#fafcfe}
.empty{background:#f5f7fa;border:1px dashed #cdd9e5;border-radius:8px;padding:16px;color:#5a6b7d;font-size:12.5px}
a{color:#0059b3;text-decoration:none}
.foot{font-size:11.5px;color:#8494a4;border-top:1px solid #e8eef5;padding-top:12px;margin-top:20px}
"""

def card(c):
    manual = "".join(f"<li>{esc(m)}</li>" for m in c["manual"])
    return f"""
<div class="card{' open' if c['feas']=='高' else ''}">
  <div class="card-h" onclick="this.parentNode.classList.toggle('open')">
    <span class="cid">{c['id']}</span><span class="nm">{esc(c['name'])}</span><span class="code">{c['code']}</span>
    <span class="base">{esc(c['base'])}</span>
    <span class="right"><span class="grade">ICP {c['score']}/{c['grade']}</span><span class="feas feas-{c['feas']}">可行性 {c['feas']}</span></span>
  </div>
  <div class="card-b">
    <div class="row"><b>决策人：</b>{esc(c['chief'])}</div>
    <div class="row"><b>杠杆点：</b>{esc(c['leverage'])}</div>
    <div class="row"><b>园区落点：</b>{esc(c['park'])}</div>
    <div class="cols" style="margin-top:10px">
      <div class="col col-auto"><h4>✅ 系统已自动完成</h4><div style="font-size:12.5px">{esc(c['system_done'])}</div>
        <div style="margin-top:8px;font-size:11.5px"><span class="tag tag-seed">种子：已验证研究</span><a href="{esc(c['evidence'][0])}" target="_blank">原文链接</a></div>
      </div>
      <div class="col col-manual"><h4>🖐️ 你需要手动做</h4><ul>{manual}</ul></div>
    </div>
  </div>
</div>"""

high = [c for c in CASES if c["feas"] == "高"]
mid = [c for c in CASES if c["feas"] == "中"]
low = [c for c in CASES if c["feas"] == "低"]

if new_findings:
    auto_html = "".join(
        f'<tr><td>{esc(f["sec"])} {esc(f["name"])}</td><td>{esc(f["title"])}</td>'
        f'<td>{"<br>".join(esc(x) for x in f["rule"])}</td>'
        f'<td><a href="{esc(f["url"])}" target="_blank">公告原文</a></td></tr>'
        for f in new_findings)
    auto_block = f'<table><tr><th>代码/公司</th><th>公告标题</th><th>自动判定</th><th>链接</th></tr>{auto_html}</table>'
else:
    note = f"今日自动发现 0 条。" + (f"（{esc(source_note)}）" if source_note else "（近 14 天内未命中“海外建厂+汽车零部件”新公告，或关键词需更新）")
    auto_block = f'<div class="empty">{note}<br><span class="tag tag-manual">人工</span> 可登录巨潮资讯网手工调整关键词：http://www.cninfo.com.cn</div>'

reject_html = "".join(f"<tr><td>{r['id']}</td><td>{r['name']} {r['code']}</td><td>{esc(r['why'])}</td></tr>" for r in REJECTED)

HTML = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PLEX 每日商机</title><style>{CSS}</style></head><body>
<div class="wrap">
<div class="top"><h1>PLEX 每日商机 · 今天该干什么</h1></div>
<div class="sub">全免费本地工具｜数据源：巨潮资讯网免费公告接口 + 已验证研究种子｜猎聘/BOSS/脉脉/领英/IATF 一律不爬，标为人工步骤</div>

<div class="kpis">
  <div class="kpi"><div class="n">{NOW}</div><div class="l">上次更新时间</div></div>
  <div class="kpi"><div class="n">{len(new_findings)}</div><div class="l">今日新增线索（自动抓取）</div></div>
  <div class="kpi"><div class="n">{manual_total}</div><div class="l">待你手动跟进项</div></div>
  <div class="kpi"><div class="n">{len(high)}</div><div class="l">高优先 case（今天先看）</div></div>
</div>

<div class="sec">
  <div class="sec-h">今日自动发现</div>
  {auto_block}
</div>

<div class="sec">
  <div class="sec-h">高优先 · 今天先动这 {len(high)} 家</div>
  {''.join(card(c) for c in high)}
</div>

<div class="sec">
  <div class="sec-h">中优先</div>
  {''.join(card(c) for c in mid)}
</div>

<div class="sec">
  <div class="sec-h">低优先（改邮件更快）</div>
  {''.join(card(c) for c in low)}
</div>

<div class="sec">
  <div class="sec-h">否决池（不要再花时间）</div>
  <table><tr><th>ID</th><th>公司</th><th>否决原因</th></tr>{reject_html}</table>
</div>

<div class="sec">
  <div class="sec-h">人工步骤区（工具不自动做，需你登录）</div>
  <table>
    <tr><th>平台</th><th>为什么必须人工</th><th>动作</th></tr>
    <tr><td>猎聘 / BOSS 直聘</td><td>需登录，且 ToS 禁止批量抓取</td><td>手动查保隆 MES 岗 HR 是否仍在选型；查目标公司是否新挂“海外 IT”岗</td></tr>
    <tr><td>脉脉 / 领英</td><td>需登录，无公开 API</td><td>手动搜 12 家集团 CIO / 海外厂长实名；德国本地收件人走领英 InMail（UWG §7）</td></tr>
    <tr><td>iatf-customerportal.org</td><td>本轮实测被安全策略拦截</td><td>手动输入新厂 site 英文名查 USI 与发证 CB；首通话直接问客户</td></tr>
    <tr><td>AMS Shanghai 2026（12/2–5）</td><td>线下活动</td><td>提前在官网 exhibitor search 拉展位图，直奔银轮 M-03、保隆 E-03 展位</td></tr>
  </table>
</div>

<div class="foot">
  <span class="tag tag-seed">种子：已验证研究</span> 来自两轮实测（12 家合格 + 6 家否决）。
  <span class="tag tag-auto">自动：今日抓取</span> 来自巨潮资讯网免费公告接口，失败已降级。
  <span class="tag tag-manual">人工</span> 工具不代做，需你自己登录。<br>
  日志：logs/run_{datetime.date.today().strftime('%Y%m%d')}.log
</div>
</div></body></html>"""

out_html = ROOT / "dashboard.html"
out_html.write_text(HTML, encoding="utf-8")
log(f"dashboard.html 已生成：{out_html}")

# ============================================================
# 4. 生成 leads.xlsx
# ============================================================
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "case 总表"
    cols = ["ID","公司","代码","海外基地","分数","等级","可行性","决策人","杠杆机构/园区","系统已自动完成","你需要手动做","证据链接"]
    ws.append(cols)
    for c in CASES:
        ws.append([c["id"], c["name"], c["code"], c["base"], c["score"], c["grade"], c["feas"],
                   c["chief"], c["leverage"] + "｜园区：" + c["park"],
                   c["system_done"], "；".join(c["manual"]), c["evidence"][0]])
    for i, w in enumerate([6,14,11,26,6,6,7,22,40,40,50,30], start=1):
        ws.column_dimensions[chr(64+i) if i<=26 else "A"+chr(64+i-26)].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="003DA5")
    ws2 = wb.create_sheet("今日自动发现")
    ws2.append(["代码","公司","公告标题","自动判定","链接","日期"])
    for f in new_findings:
        ws2.append([f["sec"], f["name"], f["title"], "；".join(f["rule"]), f["url"], f["date"]])
    wb.save(ROOT / "leads.xlsx")
    log("leads.xlsx 已生成")
except Exception as e:
    log(f"leads.xlsx 生成失败（不影响 dashboard）：{e}")

# ============================================================
# 6. 可选推送钩子（Server酱 / Bark；未配环境变量则静默跳过）
# ============================================================
def notify():
    sc = os.environ.get("SERVERCHAN_KEY", "").strip()
    bk = os.environ.get("BARK_KEY", "").strip()
    site = os.environ.get("SITE_URL", "").strip()
    if not (sc or bk):
        log("未配置 SERVERCHAN_KEY / BARK_KEY，跳过推送")
        return
    title = f"PLEX 每日商机 {TODAY}"
    body = f"今日新增 {len(new_findings)} 条线索；待手动跟进 {manual_total} 项；高优先 {len(high)} 家。"
    if site:
        body += f"\n{site}"
    if sc:
        try:
            data = urllib.parse.urlencode({"title": title, "desp": body}).encode("utf-8")
            req = urllib.request.Request(f"https://sctapi.ftqq.com/{sc}.send", data=data, timeout=8)
            urllib.request.urlopen(req, timeout=8).read()
            log("Server酱推送成功")
        except Exception as e:
            log(f"Server酱推送失败（已忽略）：{e}")
    if bk:
        try:
            url = f"https://api.day.app/{bk}/{urllib.parse.quote(title)}/{urllib.parse.quote(body)}"
            urllib.request.urlopen(url, timeout=8).read()
            log("Bark 推送成功")
        except Exception as e:
            log(f"Bark 推送失败（已忽略）：{e}")

notify()

# ============================================================
# 7. 写日志
# ============================================================
log_path = LOGS / f"run_{datetime.date.today().strftime('%Y%m%d')}.log"
log_path.write_text("\n".join(log_lines), encoding="utf-8")
print(f"\n完成。仪表盘：{out_html}")

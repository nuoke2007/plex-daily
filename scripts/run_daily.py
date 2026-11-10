# -*- coding: utf-8 -*-
"""
PLEX 每日商机工具 · 本地每日运行脚本
- 全免费、本地运行：唯一外部源是巨潮资讯网免费公告查询接口（cninfo），失败优雅降级。
- 不爬任何登录墙/反爬平台（猎聘/BOSS/脉脉/领英/IATF customerportal 一律不碰，标人工）。
- 产物：dashboard.html（双击即开）+ leads.xlsx + logs/run_YYYYMMDD.log
"""
import json, os, re, sys, time, datetime, hashlib, urllib.request, urllib.parse
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


# 每家公司的“一键动作”按钮（id -> [{label,href,kind}]）
ACTIONS = {
 "M-02": [
   {"label":"🔍 查 JLB Parques 园区（握租户名单）","href":"https://www.google.com/search?q=JLB+Parques+Parque+Industrial+Apaseo","kind":"link"},
   {"label":"🔍 查劳士领 Silao 厂联系人","href":"https://www.google.com/search?q=Rochling+Silao+Guanajuato+plant","kind":"link"}],
 "M-03": [
   {"label":"🔍 查九慧信息（SAP 实施商，求转介）","href":"https://www.google.com/search?q=%E4%B9%9D%E6%85%A7%E4%BF%A1%E6%81%AF+SAP","kind":"link"},
   {"label":"📅 Automechanika 上海 12/2-5","href":"https://www.google.com/search?q=Automechanika+Shanghai+2026","kind":"link"}],
 "T-01": [
   {"label":"🏭 打开泰中罗勇工业园官网","href":"https://www.tcrindustrial.com/","kind":"link"}],
 "T-02": [
   {"label":"🏭 打开泰中罗勇工业园官网","href":"https://www.tcrindustrial.com/","kind":"link"}],
 "E-03": [
   {"label":"🏛️ 打开 HIPA 匈牙利招商","href":"https://hipa.hu/","kind":"link"},
   {"label":"🔍 BOSS 直聘查 MES 岗","href":"https://www.zhipin.com/","kind":"link"}],
 "M-04": [
   {"label":"🔍 查祥鑫科技投资者关系邮箱","href":"https://www.google.com/search?q=%E7%A5%A5%E9%91%AB%E7%A7%91%E6%8A%80+%E6%8A%95%E8%B5%84%E8%80%85%E5%85%B3%E7%B3%BB","kind":"link"}],
 "T-03": [
   {"label":"🏢 打开科达利官网","href":"https://www.kedali.com/","kind":"link"}],
 "T-04": [
   {"label":"🏭 打开洛加纳工业园官网","href":"https://www.rojana.com/","kind":"link"}],
 "E-01": [
   {"label":"🔍 领英搜 Grammer/继峰 CIO","href":"https://www.linkedin.com/search/results/all/?keywords=Grammer%20CIO","kind":"link"},
   {"label":"📅 IAA Mobility 2027","href":"https://www.google.com/search?q=IAA+Mobility+2027","kind":"link"}],
 "E-04": [
   {"label":"🏛️ 打开 AMDIE 摩洛哥招商","href":"https://www.amdie.gov.ma/","kind":"link"}],
 "M-01": [
   {"label":"📧 一键发邮件给永茂泰董办","href":"mailto:ymtauto@ymtauto.com?subject=%E5%85%B3%E4%BA%8E%E5%A2%A8%E8%A5%BF%E5%93%A5%E7%94%9F%E4%BA%A7%E5%9F%BA%E5%9C%B0%20MES%20/%20MOM%20%E9%80%89%E5%9E%8B%E7%9A%84%E6%B2%9F%E9%80%9A&body=%E6%82%A8%E5%A5%BD%EF%BC%9A%E6%88%91%E4%BB%AC%E4%B8%93%E6%B3%A8%E6%B1%BD%E8%BD%A6%E9%9B%B6%E9%83%A8%E4%BB%B6%E7%A6%BB%E6%95%A3%E5%88%B6%E9%80%A0%E6%B5%B7%E5%A4%96%E5%B7%A5%E5%8E%82%E7%9A%84%20MES/MOM%20%E6%96%B9%E6%A1%88%E3%80%82%E6%B3%A8%E6%84%8F%E5%88%B0%E8%B4%B5%E5%8F%B8%E5%9C%A8%E5%A2%A8%E8%A5%BF%E5%93%A5%E5%B8%83%E5%B1%80%EF%BC%8C%E6%83%B3%E8%AF%B7%E6%95%99%EF%BC%9A1%EF%BC%89%E5%A2%A8%E8%A5%BF%E5%93%A5%E5%9F%BA%E5%9C%B0%20MES%20%E6%98%AF%E5%90%A6%E5%B7%B2%E9%80%89%E5%9E%8B%EF%BC%9B2%EF%BC%89%E5%9B%BD%E5%86%85%E7%8E%B0%E6%9C%89%20MES%20%E5%93%81%E7%89%8C%E6%98%AF%E5%93%AA%E5%AE%B6%E3%80%82%E7%9B%BC%E5%A4%8D%EF%BC%8C%E8%B0%A2%E8%B0%A2%E3%80%82","kind":"mail"}],
 "E-02": [
   {"label":"🏛️ 打开 RAS 塞尔维亚招商","href":"https://ras.gov.rs/","kind":"link"}],
}

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
NOISE_PLEX = ["回购","减持","增持","分红","股东会","法律意见书","评级","股东大会","摘要","季度报告","年度报告","半年度报告","董事会","监事会","独董","声明","承诺","议案","决议","问询","关注函","担保","质押","会计","减值","审计","更正","补充","批复","处罚","提示","召开","通知","换届","选举","聘任","辞职","薪酬","制度","修订","停牌","复牌","龙虎榜","可转换","债券","票据","授信","贷款"]
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

INDUSTRY_KW = [
    (["电池","锂电","储能","正极","负极","电解液","隔膜","光伏","新能源","材料"], "新能源/电池材料"),
    (["轮胎","橡胶"], "轮胎橡胶"),
    (["汽车","汽配","零部件","车架","底盘","座椅","线束","车灯","制动","传动"], "汽车零部件"),
    (["PCB","电路板","半导体","元件","电子","精密制造","元器件"], "电子/精密制造"),
    (["家电","空调","冰箱","电机","电器","压缩机","照明","灯具","光电","电气"], "家电/照明"),
    (["制药","医药","生物","医疗"], "医药"),
    (["化工","涂料","颜料","有色","金属","钢材","铝","铜","贵金属"], "材料/流程工业"),
    (["食品","饮料","乳","酒"], "食品饮料"),
]

def infer_industry(text):
    for kws, lab in INDUSTRY_KW:
        if any(k in text for k in kws):
            return lab
    return "制造业（综合）"

def auto_rule(title, name=""):
    flags = []
    text = f"{name} {title}"
    ind = infer_industry(text)
    overseas = any(k in text for k in ["泰国","墨西哥","匈牙利","塞尔维亚","摩洛哥","越南","印尼","印度","海外","境外","新加坡"])
    acts = any(k in title for k in ["增资","设立","对外投资","购地","新建","扩建","投产","子公司","生产基地","工厂"])
    if overseas and acts:
        flags.append(f"🏭 行业判断：{ind}；海外绿地建厂＝PLEX 标准画像（当地无 MES 包袱、SaaS 几周上线）→ 建议主动接触海外工厂 IT/制造负责人")
    else:
        flags.append(f"🏭 行业判断：{ind}；按公告属产能扩张，是否汽车安全/功能件需首通电话确认")
    if ("MES" in title or "制造执行" in title) and any(k in title for k in ["开发", "二次开发", "外派", "驻外", "驻海外"]):
        flags.append("⚠️ 对方在自建 MES 团队，短期不打，转为长期观察")
    if ("MES" in title or "制造执行" in title) and any(k in title for k in ["运维", "实施", "选型", "上线"]):
        flags.append("🎯 直接 MES 采购信号（可立即约）")
    if "注册资本" in title and ("增资" in title or "总投资" in title):
        flags.append("📌 注册资本小≠实际投资，留意后续大额增资公告确认规模")
    if any(k in title for k in ["设立", "新建", "绿地", "生产基地"]):
        flags.append("💡 新厂 IATF/客户审核通常滞后投产 6–18 个月，正是 MES 选型窗口")
    return flags

new_findings = []
source_status = "ok"
source_note = ""
seen_titles = set()

for kw in AUTO_KEYWORDS:
    for col in ("szse", "sse"):
        try:
            data = fetch_cninfo(kw, column=col, days=365, page_size=30)
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
                if any(b in title for b in NOISE_PLEX):
                    continue
                ann_id = a.get("announcementId")
                link = f"http://www.cninfo.com.cn/new/disclosure/detail?stockCode={a.get('secCode')}&announcementId={ann_id}"
                ts = a.get("announcementTime")
                # 按公告距今月份分级：海外绿地 MES 选型通常在公告后 3-9 个月
                rt = auto_rule(title, a.get("secName") or "")
                phase = "window"; agelab = ""
                if ts:
                    try:
                        d = datetime.datetime.fromtimestamp(ts/1000)
                        days = (datetime.datetime.now() - d).days
                        agelab = f"{d.strftime('%Y-%m-%d')} · 距今{round(days/30,1)}个月"
                        if days <= 30:
                            phase = "fresh"; rt.append(f"⏳ {agelab}——刚公告，海外厂还在立项，先观察 2-3 个月")
                        elif days <= 270:
                            phase = "window"; rt.append(f"🎯 {agelab}——正好卡在海外厂 MES 选型窗口，现在就该接触")
                        else:
                            phase = "old"; rt.append(f"🗄 {agelab}——超过 9 个月，大概率已选型，仅作存量向上销售")
                    except Exception:
                        pass
                new_findings.append(dict(
                    sec=a.get("secCode"), name=a.get("secName"), title=title,
                    date=ts, url=link, rule=rt, phase=phase,
                ))
                hit_this += 1
            log(f"关键词“{kw}”({col})：拉到 {len(anns)} 条，新增命中 {hit_this} 条")
        except Exception as e:
            source_status = "degraded"
            source_note = f"抓取失败：{type(e).__name__}: {e}"
            log(f"关键词“{kw}”({col})失败：{e}")

# 按选型窗口排序：window 在前，fresh 次之，old 最后
_ord = {"window": 0, "fresh": 1, "old": 2}
new_findings.sort(key=lambda f: _ord.get(f.get("phase"), 9))
new_findings = new_findings[:15]

log(f"今日自动发现：{len(new_findings)} 条（源状态：{source_status}）")

# ---- 多产品信号雷达（FTPC / ThingWorx / Windchill / DataMosaix）----
PRODUCTS = [
 dict(key="FTPC", emoji="🏭", name="FTPC 企业级 MES/MOM",
   pos="新基地投产 + 旧基地升级；离散·流程·混合全行业，对标西门子 Opcenter 的企业级 MOM 平台",
   who="多基地集团要统一标准 / 新基地投产 / 旧产线升级改造；头部制造与多工厂案例密集（整车·电池·轮胎·电子·流程工业）",
   three=["集团几个基地的 MES 是一套标准吗？新厂能否直接复制？",
          "ERP 与车间执行层是自动集成，还是人工对账？",
          "新基地 MES 定了没？旧基地现在的系统是要升级还是替换？"],
   hook="孚能科技赣州+广州 60GWh、巨湾技研南沙、格力钛 4 厂、奇瑞芜湖/大连/青岛/鄂尔多斯、长安多工厂、东风日产 7 厂统一平台；出海有玲珑塞尔维亚、长城泰国。实测手工输入-70%、人为记录错误-90%、偏差调查时间-50%。",
   manual=["查罗克韦尔存量 PLC/DCS 客户（天然 FTPC 客户，新厂/升级都对口）",
           "通过服务大厂的 SI 引荐",
           "搜“募投项目/新基地/电池·储能产线/智能工厂验收”"],
   scan_kw=["电池","储能","轮胎","新基地","智能工厂"]),
 dict(key="TWX", emoji="📡", name="PTC ThingWorx (IIoT)",
   pos="设备互联、OEE 透明化、预测性维护，打通 OT→IT 最后一公里",
   who="已上 MES/FTPC 的集团要做 OEE/预测维护/能耗；设备品牌杂、OEE 靠人工、停机原因不准",
   three=["OEE 是每厂看 HMI，还是集团有统一实时看板？",
          "非计划停机损失现在有准确统计吗？",
          "不同品牌（西门子/三菱/发那科）设备的数据怎么抓？"],
   hook="长安鱼嘴（冲焊涂总、非计划宕机↓）、中海油（预测性维护模型准确率>90%）、科勒佛山（卫浴行业首个 IIoT，90% 设备接入 ThingWorx）；海外 Falcon Group OEE +160%、单机利用率 8%→57%。",
   manual=["优先从 FTPC/PLC 存量客户向上卖（长安、奇瑞都是 MES+IIoT 一起上）",
           "搜“工业互联网/设备数采/OT 网络工程师”在招岗",
           "展会/论坛找抱怨“设备孤岛”的企业"],
   scan_kw=["工业互联网","预测性维护","设备联网","OEE","数字化车间","技改","机器视觉","能耗"]),
 dict(key="WIND", emoji="📐", name="PTC Windchill (PLM)",
   pos="多研发中心协同、设计 BOM 一致、工程变更闭环（数字主线）",
   who="变型/衍生设计多的汽配、家电、精密模具、工程机械；≥2 个研发中心或多厂各自研发",
   three=["多地研发中心的产品数据有没有统一源头？",
          "设计 BOM 到工艺 BOM 是自动传递还是人工整理？",
          "工程变更闭环率如何？在制品库存有没有因变更损失过？"],
   hook="Nidec 家电（巴西/墨西哥/中国/欧洲数据孤岛、1.7 万个 BOM 统一）、Vaillant（工程变更周期 -25%、首件返工率从 50% 降下来）、Polaris（Creo→eBOM→mBOM→sBOM 数字主线）。",
   manual=["优先走 PTC Creo 代理商（CAD 卖进去，PLM 是自然延伸）",
           "搜“PLM 工程师 / BOM 工程师 / 研发 IT 经理”岗",
           "领英/脉脉关注研发总监，评论“研发协同”帖"],
   scan_kw=["研发中心","PLM","产品数据","协同研发","技术中心","创新中心","实验室","研发投入"]),
 dict(key="DMX", emoji="📊", name="FactoryTalk DataMosaix / FTA",
   pos="工业 DataOps：把 OT/IT/ET 数据上下文化，工艺工程师自己做跨站点分析",
   who="流程工业（采矿/化工/食饮）或设备密集型离散；已有 Historian/PI/SCADA/MES 但数据分散、单点 POC 推不开",
   three=["工艺工程师做一次跨站点根因分析，找数据要多久？",
          "有没有哪个分析项目 POC 跑通了、但全集团推不开？",
          "能源附加费或废品率，现在有实时监控吗？"],
   hook="官方实测：维护成本 -20~30%、人工生产率 +50~60%、资产停机 -5~8%。代表客户 Hindalco 铝业（矿业）、Fonterra 恒天然（食饮）。不是建数据湖，而是让工艺工程师自己做跨站点分析。",
   manual=["从 FactoryTalk Historian / PI 存量客户向上销售（存数据→用数据）",
           "找 CDO / 数字化转型负责人，不找传统 IT 运维",
           "参加数据规模 / 数字孪生主题闭门会"],
   scan_kw=["数据中台","数字孪生","工业大数据","数据治理","数字化转型","智能制造","工业互联网平台","数据平台"]),
]

product_findings = {}
for p in PRODUCTS:
    hits = []
    _loc = set()
    for kw in p["scan_kw"]:
        for col in ("szse", "sse"):
            try:
                data = fetch_cninfo(kw, column=col, days=365, page_size=30)
                anns = data.get("announcements") or []
                for a in anns:
                    title = re.sub(r"<[^>]+>", "", a.get("announcementTitle", "") or "")
                    if not title or title in seen_titles or title in _loc:
                        continue
                    if any(b in title for b in ["回购","减持","增持","分红","股东会","法律意见书","评级","股东大会","摘要","反馈","回复","批复","处罚"]):
                        continue
                    _loc.add(title)
                    seen_titles.add(title)
                    link = f"http://www.cninfo.com.cn/new/disclosure/detail?stockCode={a.get('secCode')}&announcementId={a.get('announcementId')}"
                    ts = a.get("announcementTime")
                    nm = re.sub(r"<[^>]+>", "", a.get("secName") or "")
                    hits.append(dict(sec=a.get("secCode"), name=nm, title=title, url=link, kw=kw, ts=ts))
            except Exception as e:
                log(f"{p['key']} 关键词 {kw}({col}) 失败：{e}")
    product_findings[p["key"]] = hits
    log(f"{p['name']} 信号命中 {len(hits)} 条")

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
.legend{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:22px}
.lg{border-radius:8px;padding:10px 12px;font-size:12px;line-height:1.55}
.lg-h{background:#e8f6ee;border-left:3px solid #0a7f3f}
.lg-m{background:#fdf3e3;border-left:3px solid #b26a00}
.lg-l{background:#fdeaea;border-left:3px solid #b02a2a}
.acts{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 4px}
.btn{display:inline-block;padding:9px 14px;border-radius:7px;font-size:12.5px;font-weight:600;text-decoration:none;border:1px solid transparent}
.btn-mail{background:#003DA5;color:#fff}
.btn-link{background:#eef4fb;color:#003DA5;border-color:#c5d8ee}
.step{margin:5px 0;font-size:12.5px;line-height:1.55}
.step b{color:#003DA5;margin-right:4px}
.ev{margin-top:6px;font-size:11.5px;color:#6b7d90;line-height:1.5}
details{margin-top:10px}summary{cursor:pointer;font-size:11.5px;color:#8494a4}
.navchips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px}
.chip{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:18px;background:#003DA5;color:#fff;font-size:12.5px;font-weight:600;text-decoration:none;border:1px solid #003DA5}
.chip.alt{background:#fff;color:#003DA5;border:1px solid #c5d8ee}
.navlogo{height:18px;width:auto;display:block;background:#fff;border-radius:3px;padding:1px 3px}
.chip .navlogo{background:#fff}
.hdrlogo{height:34px;width:auto;vertical-align:middle;background:#fff;border-radius:4px;padding:2px 6px}
.prod{border:1px solid #e3ebf3;border-radius:12px;margin-bottom:18px;overflow:hidden}
.prod-h{padding:12px 16px;background:#003DA5;color:#fff;font-size:15px;font-weight:700}
.prod-h small{font-weight:400;font-size:12px;opacity:.85;margin-left:8px}
.prod-b{padding:14px 16px;display:grid;grid-template-columns:1.2fr 1fr;gap:16px}
@media(max-width:760px){.prod-b{grid-template-columns:1fr}}
.prod-sop .row{margin-bottom:7px;font-size:12.5px}
.prod-sop .row b{color:#003DA5}
.q{background:#f7fafd;border-radius:6px;padding:7px 10px;margin:4px 0;font-size:12px;color:#33475b}
.hook{background:#fff9ec;border-left:3px solid #b26a00;border-radius:6px;padding:8px 11px;font-size:12.5px;margin-top:6px}
.prod-signal{background:#fafcfe;border-radius:8px;padding:10px 12px;border:1px solid #e8eef5}
.prod-signal h5{font-size:12.5px;color:#0a7f3f;margin-bottom:6px}
.sighit{font-size:12px;padding:5px 0;border-bottom:1px dashed #e8eef5}
.sighit a{font-weight:600}
.sighit .nm{color:#8494a4;font-size:11px}
.lead{border:1px solid #e3ebf3;border-radius:9px;padding:11px 13px;margin-bottom:9px;background:#fff}
.lead-top{display:flex;align-items:center;gap:8px;margin-bottom:4px}
.lead-top .lc{color:#8494a4;font-size:11px}
.lead-top .indtag{margin-left:auto;background:#eef4fb;color:#003DA5;font-size:11px;padding:2px 8px;border-radius:10px;white-space:nowrap}
.badge{font-size:11px;font-weight:700;padding:2px 9px;border-radius:10px;color:#fff}
.badge.gh{background:#0a7f3f}.badge.gm{background:#b26a00}.badge.gl{background:#9aa7b4}
.lead-t{display:block;font-size:13px;font-weight:600;color:#003DA5;margin:3px 0}
.why{font-size:12px;color:#33475b;margin:2px 0}
.nxt{font-size:12px;color:#b26a00;margin:3px 0 7px}
"""

def card(c):
    acts = ACTIONS.get(c["id"], [])
    btn = "".join(
        f'<a class="btn {"btn-mail" if a.get("kind")=="mail" else "btn-link"}" href="{a["href"]}" target="_blank">{esc(a["label"])}</a>'
        for a in acts)
    steps = "".join(f'<div class="step"><b>{i+1}.</b>{esc(m)}</div>' for i,m in enumerate(c["manual"]))
    return f"""
<div class="card open">
  <div class="card-h" onclick="this.parentNode.classList.toggle('open')">
    <span class="cid">{c['id']}</span><span class="nm">{esc(c['name'])}</span><span class="code">{c['code']}</span>
    <span class="base">{esc(c['base'])}</span>
    <span class="right"><span class="grade">ICP {c['score']}/{c['grade']}</span><span class="feas feas-{c['feas']}">可行性 {c['feas']}</span></span>
  </div>
  <div class="card-b">
    <div class="row"><b>决策人：</b>{esc(c['chief'])}</div>
    <div class="row"><b>杠杆点：</b>{esc(c['leverage'])}</div>
    <div class="row"><b>园区落点：</b>{esc(c['park'])}</div>
    <div class="acts">{btn or '<span style="font-size:12px;color:#8494a4">无外部一键动作，按下方步骤联系</span>'}</div>
    <div style="margin-top:8px"><b style="font-size:12.5px;color:#b26a00">📋 你要做（按顺序）：</b>{steps}</div>
    <details><summary>系统已查证 / 证据（点开看）</summary>
      <div class="ev">{esc(c['system_done'])}</div>
      <div class="ev"><a href="{esc(c['evidence'][0])}" target="_blank">原文链接 ↗</a></div>
    </details>
  </div>
</div>"""

NOISE=["季度报告","年度报告","半年度报告","报告","董事会","监事会","股东大会","独董","独立董事","候选人","声明","承诺","议案","决议","摘要","法律意见","评级","问询","关注函","担保","质押","回购","减持","增持","分红","利润分配","会计","减值","审计","更正","补充","答复","批复","处罚","提示","召开","通知","换届","选举","聘任","辞职","薪酬","制度","修订","持股","停牌","复牌","龙虎榜","股东会","可转换","债券","票据","授信","银行","贷款","担保"]

def industry_of(nm,t):
    s=nm+t
    if any(w in s for w in ["电池","锂电","储能","星源","恩捷","天能","孚能","巨湾","电解液","隔膜","正极","负极","电芯"]): return "电池/储能"
    if any(w in s for w in ["轮胎","橡胶"]): return "轮胎"
    if any(w in s for w in ["汽车","整车","奇瑞","长安","长城","比亚迪","零跑","理想","哪吒","乘用车","商用车","汽车零部件"]): return "汽车/整车"
    if any(w in s for w in ["医药","生物","制药","医疗","药业"]): return "制药"
    if any(w in s for w in ["电子","半导体","电路","光电","精密","PCB"]): return "电子"
    return "制造"

LEV={
 "电池/储能":"对标孚能60GWh、巨湾南沙、格力钛4厂；找制造副总/数字化总监——电池新厂 MES 选型窗口，问设备进场节奏",
 "轮胎":"对标玲珑塞尔维亚、万力；找制造总监——多厂标准统一+出海复制",
 "汽车/整车":"对标奇瑞4基地、长安多工厂、东风日产7厂；找集团IT/制造总监——多厂 MES 标准化",
 "制药":"用 FTPC 合规优势（偏差调查-50%、审核时间-80%）；找QA/IT——电子批记录",
 "电子":"对标伟创力全球13厂、KLA；找IT——多厂快速复制",
 "制造":"找集团数字化/IT负责人，用行业头部案例切入",
}

def age_info(ts):
    """返回 (距今天数, 距今年份月份标签, 阶段)。阶段: fresh(<1月,观察) / window(1-9月,选型窗口) / old(>9月,可能已选型)"""
    if not ts:
        return (None, "", "window")
    try:
        d = datetime.datetime.fromtimestamp(ts/1000)
        days = (datetime.datetime.now() - d).days
        ms = round(days/30, 1)
        lab = d.strftime("%Y-%m")
        if days <= 30:
            phase = "fresh"
        elif days <= 270:
            phase = "window"
        else:
            phase = "old"
        return (days, f"{d.strftime('%Y-%m-%d')} · 距今{ms}个月", phase)
    except Exception:
        return (None, "", "window")

def analyze(key, nm, title, ts=None):
    t=title
    if any(w in t for w in NOISE):
        return None
    ind=industry_of(nm,t)
    lev=LEV.get(ind,LEV["制造"])
    days,agelab,phase=age_info(ts)
    # 信号本身的强弱
    sig=None
    if key=="FTPC":
        if any(w in t for w in ["收购","交割","并购","股权收购","受让"]):
            sig=("并购整合", f"{ind}企业并购/交割，新厂要并入集团体系——多厂 MES 统一或替换窗口")
        elif any(w in t for w in ["新建","投产","生产基地","募投","扩产","奠基","开工","试生产","项目备案","扩建","产线","环境影响","环评","项目投资"]):
            sig=("新产能建设", f"{ind}企业在建新产能/新基地，按惯例 MES/MOM 在设备进场前后选型")
        elif any(w in t for w in ["智能工厂","数字化车间","技改","改造升级","灯塔"]):
            sig=("技改升级", f"{ind}企业做智能工厂/技改，旧 MES 可能换代")
    elif key=="TWX":
        if any(w in t for w in ["智能工厂","数字化车间","灯塔","预测性维护","OEE","设备联网","设备数据","能耗","技改"]):
            sig=("设备智能化", f"{ind}企业做设备/OEE/能耗智能化，ThingWorx 对口窗口")
    elif key=="WIND":
        if any(w in t for w in ["研发中心","技术中心","创新中心","研究院","实验室"]):
            sig=("研发中心建设", f"{ind}企业新建研发中心，PLM/BOM 统一窗口")
    elif key=="DMX":
        if any(w in t for w in ["数据中台","工业大数据","数字孪生","数据治理","工业互联网平台","数据湖","数据平台","数字化转型","智能制造","SCADA","Historian","数据采集","云平台"]):
            sig=("数据平台/数字化建设", f"{ind}企业建数据平台/推进数字化转型，DataMosaix 跨站点数据上下文化窗口")
    if not sig:
        return None
    stag,sdesc=sig
    # 按公告时点重分级：1-9个月是选型黄金窗口
    if phase=="window":
        grade="高"; why=f"{sdesc}。📅 公告于 {agelab}——正好卡在选型/设备进场期，现在就该接触。"
    elif phase=="fresh":
        grade="中"; why=f"{sdesc}。📅 公告于 {agelab}——刚起步，先盯 1-2 个月，等设计/设备阶段再切入。"
    else:
        grade="低"; why=f"{sdesc}。📅 公告于 {agelab}——超过 9 个月，大概率已选型，仅作存量向上销售线索。"
    return (grade, why, f"杠杆：{lev}", agelab, ind)

GRADE_CLS={"高":"gh","中":"gm","低":"gl"}
PRODUCT_WHO_ROLE={
 "FTPC":("集团制造/生产副总 或 信息化(IT)负责人",
         "多工厂MES/MOM标准统一是制造与IT一把手联合拍板，找车间主任或纯IT运维都够不到预算"),
 "TWX":("设备/自动化(OT)总监 或 工业互联网负责人",
        "设备数采/OEE/预测维护是设备口的项目预算，找纯IT部门不触发"),
 "WIND":("研发IT/PLM负责人 或 研发中心(技术)总监",
         "BOM与工程变更归研发体系，找IT运维或车间都不对口"),
 "DMX":("CDO/数字化转型负责人 或 数据团队负责人",
        "跨站点数据上下文化是高管工程，找传统IT运维推不动"),
 "PLEX":("集团制造/运营副总 或 生产计划(PMC/IT)负责人",
         "PLEX 是制造运营与排产口径，拍板在制造一把手+计划/IT，销售或采购不决策"),
}

def qk_panel(key, who, three, hook, rhint):
    """即查模式：任意公司名 → 立刻生成 6 条精准查询路径 + 该产品打法 + 一键复制AI深查指令"""
    qs="".join(f'<div style="margin:2px 0">· {esc(q)}</div>' for q in three)
    return f'''
<div class="qk" id="qk-{key}">
  <div style="font-size:13px;font-weight:700;color:#003DA5">🔎 查一家不在雷达里的公司（即查模式）</div>
  <div style="font-size:12px;color:#5a6b7b;margin:3px 0 8px">输入公司名或股票代码，立刻给出这家公司该查什么、查哪几个源、找到人后怎么递话。真正查出具名人+邮箱那一步交给 AI（一键复制粘贴即可）。</div>
  <div style="display:flex;gap:8px;flex-wrap:wrap">
    <input class="qk-name" type="text" placeholder="公司名 / 股票代码，如 拓普集团、601689" autocomplete="off"
      style="flex:1;min-width:210px;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:8px 11px;font-size:13px;outline:none">
    <button class="qk-go" style="cursor:pointer;border:none;border-radius:8px;padding:8px 16px;font-size:13px;background:#003DA5;color:#fff;font-weight:600">生成查询路径</button>
  </div>
  <div class="qk-out" style="display:none;margin-top:10px;padding:11px;background:#f7f9fc;border:1px solid #dbe4ef;border-radius:8px">
    <div class="qk-title" style="font-size:12.5px;font-weight:700;color:#003DA5;margin-bottom:6px"></div>
    <div style="font-size:11.8px;color:#5a6b7b;margin-bottom:4px">① 先查这 6 个源（点开即搜，无需自己编关键词）：</div>
    <div class="qk-links" style="margin:4px 0 10px"></div>
    <div style="font-size:11.8px;color:#5a6b7b">② 拿着结果对照这款产品的打法：</div>
    <div style="font-size:12px;line-height:1.75;color:#243b53;margin-top:3px">
      <div>👉 <b>该找谁：</b>{esc(rhint)}</div>
      <div>🎯 <b>打谁（ICP）：</b>{esc(who)}</div>
      <div>❓ <b>三问验证真需求：</b>{qs}</div>
    </div>
    <div style="font-size:11.8px;color:#5a6b7b;margin-top:8px">③ 查到具名人后，按下面这条链递话：</div>
    <div style="font-size:11.8px;line-height:1.75;color:#33414f;margin-top:3px">董秘/证券部公开问询转介 → 年报查证券部直线 → 行业论坛换名片 → 内部CRM找现有AM warm intro</div>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      <button class="qk-queue" style="cursor:pointer;border:none;background:#7a5af8;color:#fff;border-radius:8px;padding:7px 13px;font-size:12.5px;font-weight:600">🚀 交给 AI 自动查（排队）</button>
      <button class="qk-copy" style="cursor:pointer;border:1px solid #003DA5;background:#fff;color:#003DA5;border-radius:8px;padding:7px 13px;font-size:12.5px;font-weight:600">📋 复制指令（急用·分钟级）</button>
      <span class="qk-tip" style="font-size:11.5px;color:#8494a4;align-self:center">排队＝每天 09:00 / 20:30 自动查完入档；急用就复制指令发我，当场查</span>
    </div>
  </div>
  <details style="margin-top:8px"><summary style="cursor:pointer;font-size:11.8px;color:#8494a4">钩子话术（查到人后用）</summary>
    <div style="font-size:11.8px;line-height:1.7;color:#33414f;margin-top:4px">{esc(hook)}</div></details>
</div>
<style>.qk{{margin:10px 0;padding:12px;background:#fff;border:1px solid #dbe4ef;border-radius:10px;box-shadow:0 1px 3px rgba(0,61,165,.06)}}</style>
<script>
(function(){{
 var box=document.getElementById('qk-{key}');if(!box)return;
 var inp=box.querySelector('.qk-name'),out=box.querySelector('.qk-out'),
     ttl=box.querySelector('.qk-title'),lk=box.querySelector('.qk-links'),
     cp=box.querySelector('.qk-copy'),tip=box.querySelector('.qk-tip'),cur='';
 var enc=function(s){{return encodeURIComponent(s);}};
 function build(){{
  var n=(inp.value||'').trim();
  if(!n){{inp.focus();return;}}
  cur=n;
  var L=[
   ['① 巨潮公告（查新基地/并购/技改/募投）','http://www.cninfo.com.cn/new/fulltextSearch?notautosubmit=&keyWord='+enc(n)],
   ['② 官网/简介（查几大基地、海外布局）','https://www.baidu.com/s?wd='+enc(n+' 官网 公司简介 生产基地')],
   ['③ 年报·董秘电话（第2步打证券部直线用）','https://www.baidu.com/s?wd='+enc(n+' 年报 董秘 联系电话')],
   ['④ 出海/新厂/并购信号（近12个月）','https://www.baidu.com/s?wd='+enc(n+' 海外 建厂 投资 新基地 2025 2026')],
   ['⑤ 在招岗（在招=正在选型，验证窗口）','https://www.zhipin.com/web/geek/job?query='+enc(n)],
   ['⑥ 信息化/数字化负责人（找具名人）','https://www.baidu.com/s?wd='+enc(n+' 信息中心 数字化 负责人 MES 智能制造')]
  ];
  ttl.textContent='【'+n+'】的信号源与动作（产品：{key}）';
  lk.innerHTML=L.map(function(x){{return '<a href="'+x[1]+'" target="_blank" style="display:inline-block;margin:0 8px 6px 0;padding:5px 11px;border:1px solid #c8d4e0;border-radius:14px;font-size:11.8px;color:#003DA5;text-decoration:none;background:#fff">'+x[0]+' ↗</a>';}}).join('');
  out.style.display='block';
 }}
 box.querySelector('.qk-go').onclick=build;
 inp.onkeydown=function(e){{if(e.key==='Enter')build();}};
 var qb=box.querySelector('.qk-queue');
 if(qb)qb.onclick=function(){{
  var n=(inp.value||'').trim();
  if(!n){{inp.focus();return;}}
  var fn='queue/'+enc(n)+'-{key}.txt';
  var val=enc(n+' | {key}');
  var url='https://github.com/nuoke2007/plex-daily/new/main?filename='+fn+'&value='+val;
  window.open(url,'_blank');
  tip.textContent='已打开 GitHub：点绿色「Commit changes」即加入队列';tip.style.color='#7a5af8';
 }};
 cp.onclick=function(){{
  var t='帮我深查【'+cur+'】这家公司，产品线是 {key}：\\n'
   +'1) 近12个月有没有新基地投产/海外建厂/并购/技改/研发中心/数据平台类信号，判断它接下来会不会有 {key} 的选型需求，给明确判断（不要让我自己去判断）；\\n'
   +'2) 总部所在省市→大区（用于区域筛选）；\\n'
   +'3) 具名对接人：制造/信息化/研发/数字化负责人，写真实姓名+出处链接；查不到就标“未具名”，绝不编造；\\n'
   +'4) 若查到具名人，推断其企业邮箱（标注“格式推断、首封先测退信”，附备选格式）并加“📧 发邮件”按钮；\\n'
   +'5) 按该产品打法给出四步搭线动作链（董秘转介→证券部直线→行业论坛→内部CRM/AM转介），写明具体转介目标；\\n'
   +'6) 把结果写进 scripts/run_daily.py 的 PRODUCT_DOSSIERS 并部署到 GitHub Pages。';
  var done=function(){{tip.textContent='已复制 ✓ 粘到对话里发我即可';tip.style.color='#137333';}};
  if(navigator.clipboard&&navigator.clipboard.writeText){{navigator.clipboard.writeText(t).then(done,function(){{window.prompt('复制下面这段发给我：',t);}});}}
  else{{window.prompt('复制下面这段发给我：',t);}}
 }};
}})();
</script>'''


def render_product(p):
    raw = product_findings.get(p["key"], [])
    _seen=set(); _raw=[]
    for h in raw:
        k=h.get("sec")  # 同一家公司只保留一条最强信号，避免刷屏
        if k in _seen: continue
        _seen.add(k); _raw.append(h)
    raw=_raw
    leads=[]
    dropped=0
    for h in raw:
        nm=(h["name"] or h["sec"] or "").strip()
        r=analyze(p["key"], nm, h["title"], h.get("ts"))
        if r is None:
            dropped+=1; continue
        g,judge,lev,agelab,ind = r
        leads.append((g,judge,lev,agelab,ind,nm,h))
    order={"高":0,"中":1,"低":2}
    leads.sort(key=lambda x: order[x[0]])
    total=len(leads)
    leads=leads[:15]
    # 各产品对董秘的搭线话术
    PRODUCT_ASK={
      "FTPC":"关注到贵司新基地投产/多厂扩张，希望15分钟对接制造/IT负责人交流多工厂MES/MOM标准统一",
      "TWX":"关注到贵司设备联网/OEE/智能化项目，希望15分钟对接设备/自动化负责人交流OT数据采集与预测维护",
      "WIND":"关注到贵司研发中心/产品数据建设，希望15分钟对接研发IT/PLM负责人交流BOM与工程变更管理",
      "DMX":"关注到贵司数据平台/工业大数据项目，希望15分钟对接CDO/数字化负责人交流跨站点数据上下文化分析",
    }
    ask=PRODUCT_ASK.get(p["key"],"希望对接信息化负责人交流")
    # 按总部区域分组
    from collections import OrderedDict
    by_region=OrderedDict()
    for g,judge,lev,agelab,ind,nm,h in leads:
        reg=region_of(h.get("sec"))
        by_region.setdefault(reg,[]).append((g,judge,lev,agelab,ind,nm,h))
    cards=""
    rg_order=[]
    for reg, items in by_region.items():
        rg=reg.split("·")[0]
        if rg not in rg_order: rg_order.append(rg)
        cards+=f'<div class="rg-head" data-rg="{esc(rg)}" style="margin:12px 0 4px;font-size:13px;font-weight:700;color:#003DA5;border-left:4px solid #003DA5;padding-left:8px">🗺️ 总部区域：{esc(reg)}（{len(items)}家）</div>'
        for g,judge,lev,agelab,ind,nm,h in items:
            code=h.get("sec") or ""
            _rw=PRODUCT_WHO_ROLE.get(p["key"],("信息化/数字化负责人","找对部门比找对公司更重要"))
            _d=PRODUCT_DOSSIERS.get(code)
            if _d:
                _role,_why=_d["role"],_d["why"]
                _chief=''
                if _d.get("auto"):
                    _demand=_d.get("demand") or ""
                    _dc="#137333" if _demand=="是" else ("#b3541e" if _demand=="待验证" else "#8494a4")
                    _chief=(f'<div style="font-size:12px;line-height:1.6;margin-top:2px">'
                            f'<span style="color:#7a5af8;font-weight:700">🤖 AI 自动查证</span>'
                            f' <span style="color:{_dc};font-weight:700">选型判断：{esc(_demand)}</span>'
                            f'<span style="color:#5a6b7b"> —— {esc(_d.get("demand_reason") or "")}</span></div>')
                    _chief += f'<div style="font-size:12px;color:#33414f;line-height:1.6">🔎 信号：{esc(_d.get("signal") or "")}</div>'
                _chief += f'<div style="font-size:12px;color:#0b6bcb;line-height:1.6;margin-top:2px">👤 {esc(_d["chief"])}</div>'
                if _d.get("src"):
                    _chief += f'<div style="font-size:11.5px;margin-top:1px"><a class="ev" style="color:#0b6bcb" href="{esc(_d["src"])}" target="_blank">🔗 点开核实人名出处</a></div>'
                _ins=_d.get("insider")
                if _ins:
                    _it=esc(_ins.get("title") or "内部线索人"); _inm=esc(_ins.get("name","")); _iu=esc(_ins.get("src",""))
                    _chief += ('<div style="font-size:12px;color:#7a4b00;line-height:1.6;margin-top:2px">🧩 内部线索人：'
                               + _it + ' <b>' + _inm + '</b>'
                               + ' <span style="color:#8494a4">（可作引荐桥梁：同一条项目/技术线上的人，比冷邮件近一层）</span>'
                               + ' <a class="ev" style="color:#0b6bcb" href="' + _iu + '" target="_blank">🔗 出处</a></div>')
                if _d.get("email"):
                    import urllib.parse as _u2
                    _e=_d["email"]
                    _sub=_u2.quote(f"【{nm}】多基地制造/数字化标准统一交流——罗克韦尔自动化")
                    _body=_u2.quote(
                        f"您好：\n"
                        f"我是罗克韦尔自动化的售前顾问。关注到{nm}近期的产能/海外布局动向，"
                        f"我们在多工厂制造执行与数据统一上有行业头部落地案例，想约15分钟电话交流，"
                        f"看是否对贵司当前规划有借鉴。\n"
                        f"如方向不对，烦请转介对口同事，感谢！\n罗克韦尔自动化")
                    _chief+=(f'<div style="font-size:12px;color:#0b6bcb;line-height:1.6;margin-top:3px">📮 <b>推断邮箱：</b><code style="background:#eef4fb;padding:1px 5px;border-radius:4px">{esc(_e["to"])}</code>'
                             f' <span style="color:#8494a4">（备选：{esc(_e["alt"])}；{esc(_e["note"])}）</span>'
                             f' <a class="ev" href="mailto:{_e["to"]}?subject={_sub}&body={_body}">📧 直接写邮件</a></div>')
            else:
                _role,_why=_rw[0],_rw[1]; _chief=''
            cards+=(f'<div class="lead" data-rg="{esc(rg)}">'
              f'<div class="lead-top"><span class="badge {GRADE_CLS[g]}">{g}</span><b>{esc(nm)}</b><span class="lc">{esc(code)}</span>'
              f'<span class="indtag">🏭 {esc(ind)}</span><span class="indtag">🗺️ {esc(rg)}</span></div>'
              f'<a class="lead-t" href="{h["url"]}" target="_blank">{esc(h["title"])}</a>'
              f'<div class="why"><b>🔍 关键信息：</b>{esc(judge)}</div>'
              f'<div class="nxt"><b>🎯 下一步怎么做：</b>{esc(lev).replace("杠杆：","")}</div>'
              f'{_chief}'
              f'{contact_links(code,nm,ask,role=_role,why=_why,grade=g)}'
              f'<div style="margin-top:6px"><a class="btn btn-link" href="{h["url"]}" target="_blank">📄 看公告原文</a></div>'
              '</div>')
    # 区域筛选条
    chips='<button class="rg-chip active" data-rg="全部">全部</button>'+"".join(
        f'<button class="rg-chip" data-rg="{esc(rg)}">{esc(rg)}</button>' for rg in rg_order)
    rg_filter_js="""
<style>.rg-chip{cursor:pointer;border:1px solid #c8d4e0;background:#fff;color:#33475b;border-radius:14px;padding:4px 12px;font-size:12px;margin:0 6px 6px 0}.rg-chip.active{background:#003DA5;color:#fff;border-color:#003DA5}.rg-bar{position:sticky;top:0;background:#fff;padding:8px 0;z-index:9;border-bottom:1px solid #eef2f7}.srch{width:100%;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:7px 11px;font-size:12.5px;color:#33475b;margin:2px 0 8px 0;outline:none}.srch:focus{border-color:#003DA5;box-shadow:0 0 0 3px rgba(0,61,165,.08)}.srch-empty{display:none;padding:10px;font-size:12px;color:#8494a4;text-align:center}</style>
<script>
(function(){var bar=document.getElementById('rgbar');if(!bar)return;
var inp=bar.querySelector('.srch'), rg='全部';
function apply(){
 var q=(inp&&inp.value||'').trim().toLowerCase(), n=0;
 document.querySelectorAll('.lead,.rg-head').forEach(function(el){
   var okr=(rg==='全部'||el.getAttribute('data-rg')===rg);
   var oks=true;
   if(q){ oks=(el.textContent||'').toLowerCase().indexOf(q)>=0 && el.classList.contains('lead'); }
   var show=okr&&oks;
   if(show&&el.classList.contains('lead')) n++;
   el.style.display=show?'':'none';});
 var em=bar.parentNode.querySelector('.srch-empty');
 if(em){ em.style.display=(q&&n===0)?'block':'none'; }
}
bar.querySelectorAll('.rg-chip').forEach(function(c){c.onclick=function(){
 bar.querySelectorAll('.rg-chip').forEach(function(x){x.classList.remove('active')});c.classList.add('active');
 rg=c.getAttribute('data-rg'); apply();};});
if(inp){ inp.oninput=apply; }
})();
</script>"""
    if not cards:
        cards='<div class="empty" style="padding:14px;font-size:12.5px;color:#8494a4">今日无合格线索（已过滤财报/治理类公告，只留有新产能·并购·技改信号的；雷达每天自动扫）</div>'
        chips=""; rg_filter_js=""
    qs="".join(f'<div class="q">问{i+1}：{esc(q)}</div>' for i,q in enumerate(p["three"]))
    return f"""
<div class="prod" id="{p['key']}">
  <div class="prod-h">{p['emoji']} {esc(p['name'])}<small>{esc(p['pos'])}</small></div>
  <div class="prod-b" style="grid-template-columns:1fr">
    <div>
      <div style="font-size:12.5px;color:#33475b;margin:2px 0 10px">🆕 今日自动线索 <b style="color:#003DA5">{total}</b> 条 · 已按优先级排序（高在前）</div>
      <div class="rg-bar" id="rgbar">📍 按总部区域筛选：{chips}<input class="srch" type="search" placeholder="🔍 搜公司名 / 股票代码 / 行业（如 安琪、600298）" autocomplete="off"></div>
      <div class="srch-empty">没有匹配的线索 —— 用下面的「即查模式」查这家公司</div>
      {cards}
      {qk_panel(p["key"], p["who"], p["three"], p["hook"], PRODUCT_WHO_ROLE.get(p["key"],("信息化/数字化负责人",""))[0])}
      {rg_filter_js}
      <details style="margin-top:12px"><summary style="font-size:12px;color:#8494a4">📖 这款产品的打法参考（打谁/三问/钩子，平时不用看）</summary>
        <div class="prod-sop" style="margin-top:8px">
          <div class="row"><b>打谁：</b>{esc(p['who'])}</div>
          {qs}
          <div class="hook" style="margin-top:6px"><b>钩子：</b>{esc(p['hook'])}</div>
        </div>
      </details>
    </div>
  </div>
</div>"""

FILEMAP={"PLEX":"dashboard.html","FTPC":"ftpc.html","TWX":"thingworx.html","WIND":"windchill.html","DMX":"datamosaix.html"}
LOGO={"PLEX":"logo-plex.png","FTPC":"logo-ft.png","TWX":"logo-twx.png","WIND":"logo-wind.png","DMX":"logo-ft.png"}
SHORT={"PLEX":"PLEX","FTPC":"FTPC","TWX":"ThingWorx","WIND":"Windchill","DMX":"DataMosaix"}
NAVITEMS=[(FILEMAP[k], k) for k in ["PLEX","FTPC","TWX","WIND","DMX"]]
def nav_html(active):
    chips=""
    for f,k in NAVITEMS:
        cls="chip" if k==active else "chip alt"
        chips+=(f'<a class="{cls}" href="{f}">'
                f'<img class="navlogo" src="{LOGO[k]}" alt="">'
                f'<span>{SHORT[k]}</span></a>')
    return f'<div class="navchips">{chips}</div>'

high=[c for c in CASES if c["feas"]=="高"]
mid=[c for c in CASES if c["feas"]=="中"]
low=[c for c in CASES if c["feas"]=="低"]

if new_findings:
    auto_html="".join(
        f'<tr><td>{esc(f["sec"])} {esc(f["name"])}</td><td>{esc(f["title"])}</td>'
        f'<td>{"<br>".join(esc(x) for x in f["rule"])}</td>'
        f'<td><a href="{esc(f["url"])}" target="_blank">公告原文</a></td></tr>'
        for f in new_findings)
    auto_block=f'<table><tr><th>代码/公司</th><th>公告标题</th><th>自动判定</th><th>链接</th></tr>{auto_html}</table>'
else:
    note=f"今日自动发现 0 条。"+(f"（{esc(source_note)}）" if source_note else "（近 12 个月内未命中海外建厂新公告）")
    auto_block=f'<div class="empty">{note}</div>'

# 每日滚动线索：按公告时点分高(窗口)/中(刚公告)/低(超9月)三档
import urllib.parse as _up2
PHASE_GRADE={"window":"高","fresh":"中","old":"低"}

# ===== AI 逐家深挖档案（真查后写入，按股票代码）=====
# 每家：loc / chief / lev / steps / hook(可空用默认) / note
DOSSIERS={
 "603166":dict(loc="墨西哥（筹建中，2026-03公告设全资子公司）",
   chief="桂林总部：数字化牵头人王长顺（公开谈过 MES/5G 工业互联网）；墨西哥子公司经理将由董事聘任（尚未到岗=切入窗口）",
   lev="国内已贯通 PLM/ERP/MES/WMS——是成熟 MES 用户；但墨西哥是绿地新厂、当地无 IT。桂林总部 MES 团队最关心“海外厂怎么快速复制、不养本地运维”",
   hook="“王总，国内 MES 已经跑顺，墨西哥新厂若再买一套传统 MES，当地服务器+专职 IT 至少两年才回本；PLEX 是 SaaS，和国内体系并行、几周上线，墨西哥厂厂长不用养 IT。”",
   steps=["先约桂林总部数字化/IT 负责人（王长顺线），不谈买新 MES，谈“墨西哥厂怎么零运维快速上线”",
          "墨西哥子公司经理尚未聘任——通过董秘办/证券部问筹建组，赶在团队到位前把 IT 方案锁定",
          "对标：国内 MES 已投，海外厂用 PLEX 补位，避免重复买重型 MES"]),
 "001380":dict(loc="墨西哥克雷塔罗州（与万安科技合资 65/35，2028 投产）",
   chief="董事长金雷（公开谈墨西哥本地化供应链闭环）；合资公司由华纬金晟实业控股，万安科技（香港）北美持股35%",
   lev="克雷塔罗=福特/宝马汽车带；这是与万安科技的合资厂——MES 选型要过两家股东。董事长金雷是决策拍板人",
   hook="“金总，克雷塔罗厂贴着福特/宝马，客户审厂要看追溯；PLEX SaaS 几周上线，合资厂两边股东都不用为 IT 运维扯皮。”",
   steps=["直接攻董事长金雷（他亲自讲墨西哥战略），定位是“2028 投产前把 IT 架构定掉”",
          "同步接触合资方万安科技（35%股东），它自己也有出海 IT 诉求，可捆绑",
          "克雷塔罗中资企业集中，托当地中资园区/中企商会引荐已落地厂长"]),
 "002454":dict(loc="墨西哥蒙特雷（2026 年中投产，现处设备调试阶段）",
   chief="松芝国际控股（全资母公司）统筹；墨西哥松芝尚未正式经营、无专职 IT=真空窗口",
   lev="工厂正处设备调试、未投产——正是 MES 选型最后窗口，投产再上就晚；3.6 亿增资说明投真金",
   hook="“松芝蒙特雷厂马上投产，现在上 PLEX，投产第一天就有排产/追溯；等投产后再补系统，要停产折腾。”",
   steps=["抢在 2026 年中投产前接触松芝国际控股 IT——设备调试期是最后选型窗",
          "通过董秘/互动易松芝证券部转海外项目负责人",
          "对标汽车热管理客户（北美主机厂审厂），主打投产即合规"]),
 "300852":dict(loc="泰国（子公司一品电路，2024H2 已量产，现 1.5 亿营收爬坡）",
   chief="一品电路（泰国）管理层 + 总部 IT；泰国厂正建本地化人才梯队、优化品控",
   lev="已过绿地期、在爬坡扩产（9.3 亿定增加码）；痛点是“中国+泰国双基地标准不统一、品控要对齐”——更偏 FTPC 统一标准，PLEX 补泰国厂追溯",
   hook="“泰国一品电路在爬坡，国内 PCB 厂 MES 标准怎么一键复制到泰国厂、品控数据回传总部？PLEX 几周上线先把泰国厂追溯跑起来。”",
   steps=["定位泰国一品电路爬坡期的品控/IT 负责人，问双基地标准怎么对齐",
          "已量产=不是绿地冷启动，主打“快速补追溯+双基地看板”，不推翻现有系统",
          "注意：此案更偏 FTPC 多厂统一，PLEX 只切泰国厂快速上线这一块"]),
 "301039":dict(loc="墨西哥（集团海外布局；泰国 DS 厂已自动化灯塔）",
   chief="集团数字化流程中心（一办三中心之一）；灯塔项目组李晓甫博士线",
   lev="中集车辆是灯塔工厂老手、数字化能力极强——不是 PLEX 冷启动对象。真正机会是 FTPC/ThingWorx 向墨西哥新厂复制灯塔方法论",
   hook="（此案 PLEX 非首选）“集团灯塔工厂经验要复制到墨西哥新厂，要不要用 ThingWorx/FTPC 把泰国 DS 那套自动化产线标准直接搬过去？”",
   steps=["不推 PLEX，转推 FTPC/ThingWorx——找集团数字化流程中心",
          "借泰国 DS 厂灯塔案例做背书",
          "标为“存量灯塔客户出海复制”，不按绿地 MES 打"]),
 "600261":dict(loc="泰国春武里府 Hermes Industrial Estate（爱马仕园区，107.3亩，刚购地）",
   chief="浙江阳光照明总部：海外项目负责人/IT；泰国阳光为全资子公司，管理层待配",
   lev="已签 Hermes 园区土地协议、付定金，纯绿地刚拿地=IT 方案还没定；园区运营方握同园区租户名单",
   hook="“春武里基地刚拿地，正是把 IT 架构一次定掉的窗口；PLEX SaaS 几周上线，泰国厂不用配专职 IT。”",
   steps=["找 Hermes 园区运营方，同园区中资厂一次摸清哪些已上 MES",
          "通过浙江总部董秘办转海外项目负责人，赶在厂房设计阶段锁定 MES",
          "照明行业多品种小批量，主打快速追溯+北美客户验厂"]),
 "002543":dict(loc="泰国（一/二期已量产，三期2026-05奠基，250万台）+ 埃及基地",
   chief="万和电气(泰国)制造负责人；集团多基地数字化牵头",
   lev="一、二期已投产，三期在建——不是绿地冷启动，痛点是“中国+泰国+埃及多基地标准统一”，偏 FTPC；PLEX 只切三期新产线快速上线",
   hook="（偏 FTPC）“三期新厂要不要和泰国一期、埃及厂用同一套标准？FTPC 多基地统一，避免每个厂各搞一套。”",
   steps=["不主推 PLEX，转 FTPC 多基地统一标准",
          "找集团制造/数字化，问一/二期 MES 是怎么上的、三期要不要复制",
          "埃及+泰国双海外，主打集团级标准拉通"]),
 "002815":dict(loc="泰国巴真府 洛加纳工业园（一期2026-07封顶，年底设备调试）",
   chief="集团董事朱雪花、副总裁余忠（公开出席封顶）；泰国厂厂长/IT 待配",
   lev="2026 年底设备安装调试=MES 选型最后窗口；高端 PCB(服务器/汽车)对追溯严苛",
   hook="“崇达泰国厂年底装设备，投产前把 PLEX 配上，投第一天就有排产追溯，别等投产后停产补系统。”",
   steps=["赶 2026 年底设备调试期接触，这是最后选型窗",
          "通过副总裁余忠线转泰国厂筹建组",
          "PCB 行业主打 SPC/追溯，对标四会富仕泰国厂"]),
 "001359":dict(loc="泰国北柳府（通过平安新加坡设平安泰国，云母绝缘材料基地，~3100万美元）",
   chief="湖北平安电工总部 IT/海外项目组；实施主体平安泰国待组建",
   lev="2026-07 刚立项、定增13亿投泰国=纯绿地、IT 完全空白；材料行业批次追溯要求高",
   hook="“泰国云母材料基地刚立项，PLEX SaaS 几周上线，不用等国内 MES 团队飞过去。”",
   steps=["找总部海外项目/IT，立项期就切入 IT 架构",
          "定增募资落地=预算在，直接谈方案",
          "材料批次追溯切入"]),
 "605277":dict(loc="泰国（注册资本3亿泰铢）；已布局墨西哥/匈牙利/日本多基地",
   chief="新亚电子 IT/数字化；旗下中德电缆自带墨西哥/匈牙利厂",
   lev="多海外基地(泰国/墨西哥/匈牙利/日本)+AI 高频铜缆扩产——是多厂标准统一问题，偏 FTPC，不是单厂 PLEX",
   hook="（偏 FTPC/ThingWorx）“泰国+墨西哥+匈牙利多厂，MES 标准怎么统一、AI 铜缆产线数据怎么汇？FTPC 多基地拉通。”",
   steps=["转 FTPC 多厂统一，不推单厂 PLEX",
          "通过中德电缆现有海外厂找数字化牵头人",
          "AI 算力线缆扩产=产线数据采集可加 ThingWorx"]),
 "002843":dict(loc="泰国（香港泰嘉100%控股，预计2027-10投产，已设海外事业部）",
   chief="泰嘉海外事业部（2025新组建）；泰国厂厂长待配",
   lev="制造业单项冠军，长沙/连云港/德国/印度已有基地，泰国是第五个=多基地复制问题，偏 FTPC；海外事业部刚组建正是切入时",
   hook="（偏 FTPC）“泰嘉刚组建海外事业部、泰国厂2027投产，正好把长沙/连云港那套 MES 标准一键复制到泰国。”",
   steps=["找新组建的海外事业部负责人，时机正好",
          "多基地复制打 FTPC，不打单厂 PLEX",
          "德国/印度已有基地，问现有 MES 能否统一"]),
 "300791":dict(loc="泰国春武里产业园（2027投产，在建，总6万㎡）",
   chief="助理副总裁黄舒丽（公开牵头泰国项目）；CDMO 质量负责人",
   lev="保健品 CDMO、GMP/FDA 合规要求极高；2027投产=现在选型；多基地(全球7大工厂)标准统一",
   hook="“仙乐泰国产业园2027投产，GMP 批次追溯是硬门槛；PLEX/FTPC 几周上线，直接满足 FDA 审计。”",
   steps=["找助理副总裁黄舒丽线（公开牵头项目）",
          "CDMO 行业主打 GMP 电子批记录/批次追溯，合规切入",
          "全球7大基地，问标准统一需求"]),
 "002429":dict(loc="墨西哥（经新加坡兆驰→墨西哥兆驰，~3833万美元，刚投产爬坡）",
   chief="墨西哥兆驰(MTC Electronic Mexico)厂长/IT；深圳总部 ODM 制造",
   lev="ODM 智能终端，墨西哥厂刚投产、与越南基地分工；已在爬坡期，偏多基地标准化(FTPC)",
   hook="“墨西哥 ODM 厂刚投产爬坡，和越南厂怎么用同一套标准、客户验厂怎么快速过？”",
   steps=["找墨西哥兆驰厂长，投产即补追溯",
          "ODM 客户(品牌方)验厂要求高，主打快速合规",
          "越南+墨西哥双海外，转 FTPC 标准统一"]),
 "300818":dict(loc="连锁海外厂（蒙古已投产，智利2026底、秘鲁2027H2投产；墨西哥为销售子公司）",
   chief="耐普矿机海外连锁工厂项目组；智利/秘鲁厂长待配",
   lev="矿山耐磨备件，全球连锁工厂模式=多基地标准化(FTPC)，墨西哥是销售子公司不是工厂；工艺离散，PLEX 弱匹配",
   hook="（偏 FTPC/ThingWorx）“智利、秘鲁新厂要和蒙古厂统一标准，FTPC 连锁复制；备件产线数据用 ThingWorx。”",
   steps=["不打 PLEX，转 FTPC 多厂连锁统一",
          "盯智利(2026底)/秘鲁(2027H2)投产前的 MES 窗口",
          "客户紫金/洛钼等矿山，强调交付追溯"]),
 "605099":dict(loc="墨西哥新莱昂州（2026-04 已公告终止建设项目）",
   chief="—（项目已终止，非活跃线索）",
   lev="注意：共创草坪2026年4月已董事会决议终止墨西哥生产基地，改走越南/印尼基地——此条为负信号，不应再投入",
   hook="（无效线索）墨西哥项目已终止，无需接触；可关注其越南/印尼基地扩产",
   steps=["从高优先剔除，标注“墨西哥项目已终止”",
          "若要跟，转其越南/印尼草坪基地",
          "不再投入资源"]),
}

def plex_playbook(nm, title, ind):
    """按行业+海外目的地，自动还原原始 case 卡的打法：决策人/杠杆点/按顺序动作。"""
    t=title
    if "泰" in t:
        loc="泰国（罗勇/春武里 EEC）"; park="泰中罗勇工业园 TCR / Amata（中资园区运营方，握租户名单）"
    elif "越南" in t:
        loc="越南（海防/北宁/平阳）"; park="VSIP / 中资制造园区运营方"
    elif "墨西哥" in t or "墨西哥" in nm or "北美" in t:
        loc="墨西哥（Bajío 汽车带）"; park="Bajío 汽车产业园运营方（周边已有大量中国汽零厂）"
    elif "印尼" in t or "印度尼西亚" in t:
        loc="印尼（雅加达/卡拉旺）"; park="中资园区运营方"
    elif "马来" in t:
        loc="马来西亚（槟城/柔佛）"; park="中资园区运营方"
    else:
        loc="海外（国别见公告）"; park="当地中资/制造园区运营方"
    chief=f"海外新厂厂长 + IT(数字化)负责人 + 质量经理（IATF 体系）"
    lev=(f"{park}；该公司产线若已用罗克韦尔 PLC 则走存量客户渠道；公告董秘/证券部可代为转数字化负责人")
    steps=[
        f"脉脉/LinkedIn 搜【{nm}】+“{loc.split('（')[0]} IT / 海外工厂 IT”——在招这类岗=MES 选型正当时，直接加人聊",
        f"上门{park.split('（')[0]}运营方——一次上门可同时触达同园区多家出海企业",
        f"首通电话三问：①海外厂排产报工现在用系统还是 Excel？②当地买服务器招 IT 的预算/周期能接受吗？③北美/欧洲客户的质量追溯现在靠什么满足？",
    ]
    return chief, lev, steps, loc
def auto_dossier(nm, title, rule, ind):
    """无手写档案时，从公告原文自动抠真料（投资额/阶段/地点），不再套模板。"""
    import re as _re
    txt=(title or "")+" "+" ".join(rule or [])
    # 投资额
    amts=_re.findall(r'([0-9.]+)\s*(?:亿|万)\s*(?:元|人民币|美元|港币)?', txt)
    amt=("约"+amts[0]+"亿" if amts and "亿" in txt else ("约"+amts[0]+"万美元" if amts and "美元" in txt else ("约"+amts[0] if amts else "投资额见公告")))
    # 阶段
    if "终止" in txt or "终止投资" in txt: stage="项目已终止/调整（负信号，需复核是否仍跟进）"
    elif "封顶" in txt: stage="厂房已封顶，正处设备进场/安装阶段＝MES 选型最后窗口"
    elif "试产" in txt: stage="已进入试产/小批量阶段"
    elif "投产" in txt: stage="已投产，处产能爬坡期"
    elif "奠基" in txt: stage="刚奠基，土建阶段，IT 方案尚未锁定"
    elif "购买土地" in txt or "购地" in txt or "土地" in txt: stage="刚签购地协议/拿地，纯绿地最早期窗口"
    elif "增资" in txt: stage="对海外子公司增资扩产，处扩产/建设期"
    elif "筹建" in txt or "设立" in txt: stage="子公司刚设立/筹建，管理团队未到位＝切入窗口"
    else: stage="对外投资/设立海外项目，处早期"
    # 地点
    loc=plex_playbook(nm, title, ind)[3]
    chief=f"{nm} 总部：海外项目/IT（数字化）负责人；海外子公司管理层待配或刚到位"
    lev=f"公告显示{stage}；投资规模{amt}。这是从公告直接读出的窗口——决策人尚未在公开材料点名，需通过总部董秘办转海外项目组。"
    steps=[
        f"先读公告原文确认：{stage}、投资{amt}，判断是否真有 MES/MOM 选型需求",
        f"通过 {nm} 董秘办/证券部转“海外项目/数字化负责人”（公告里只有董秘联系方式，这是公开可达的入口）",
        f"脉脉/Boss 搜【{nm}】+“海外 IT / 海外工厂 IT / 制造工程师”——在招=选型正当时，直接加人",
    ]
    return chief,lev,steps,loc

def contact_links(code, nm, ask=None, role="", why="", grade=""):
    """分析先行的搭线块：先讲清找谁/为什么是他/为什么现在，再给把话递进去的合规通道。"""
    import urllib.parse as _u
    code=(code or "").replace(".SH","").replace(".SZ","").zfill(6)
    ask=ask or "关注到贵司海外新厂建设，希望15分钟对接信息化/海外项目负责人交流MES方案"
    if code.startswith("6"):
        ir=f"https://sns.sseinfo.com/company/company.do?stockcode={code}"
        ir_label="董秘/e互动直达（上交所）"
    else:
        ir=f"https://irm.cninfo.com.cn/ircs/company/companyDetail?stockcode={code}"
        ir_label="董秘/互动易直达（深交所·可直接提问）"
    job=f"https://www.zhipin.com/web/geek/job?query={_u.quote(nm)}"
    baidu=f"https://www.baidu.com/s?wd={_u.quote(nm+' 董秘 电话 年报')}"
    now={"高":"公告时点正好卡在选型/设备进场期，现在递话最容易被当回事",
         "中":"项目刚启动，先把关系种进去，等方案阶段再深入",
         "低":"大概率已定选型，只做存量/向上销售观察"}.get(grade,"")
    tgt=role.split("（")[0].split("/")[0]  # 转介目标，取最前的部门/人
    steps=[
      f'① <b>董秘/证券部公开问询转介</b>（合规、对方不能装没看见）：在 <a class="ev" href="{esc(ir)}" target="_blank">{ir_label}</a> 公开发商务问询，请其转{esc(tgt)}；别写成广告，写“多基地制造/数字化标准统一，求商务对接窗口”。',
      f'② <b>打证券部直线（不是总机）</b>：<a class="ev" href="{esc(baidu)}" target="_blank">年报查董秘办电话</a>，工作日上午一句“有个多基地方案想转{esc(tgt)}，给个商务对接邮箱”。',
      f'③ <b>行业论坛换名片</b>：查下一场智能制造/工业互联网大会，看{esc(nm)}谁上台；你方是赞助商，茶歇以“同议题”同行交流破冰，不推销。',
      f'④ <b>内部CRM找现有AM warm intro（你独有）</b>：查{esc(nm)}现有设备/经销商，让已在现场服务的AM转介{esc(tgt)}；并用 <a class="ev" href="{esc(job)}" target="_blank">BOSS看在招岗</a> 验证是否在建队（在招=窗口期）。',
    ]
    steps_html="".join('<div style="margin:3px 0">'+s+'</div>' for s in steps)
    return (f'<div style="margin-top:8px;padding:10px;background:#f4f8ff;border:1px solid #d6e4ff;border-radius:8px">'
            f'<b style="font-size:12.5px;color:#003DA5">🎯 搭线动作链（照着走就能拨通）</b>'
            f'<div style="font-size:12px;line-height:1.7;margin-top:5px;color:#243b53">'
            f'<div>👉 <b>该找谁：</b>{esc(role)}</div>'
            f'<div>🧭 <b>为什么是他：</b>{esc(why)}</div>'
            f'<div>⏱️ <b>为什么现在：</b>{esc(now)}</div>'
            f'</div>'
            f'<div style="margin-top:7px;padding-top:6px;border-top:1px dashed #c8d8ee;font-size:11.8px;line-height:1.8;color:#33414f">'
            f'{steps_html}'
            f'<div style="color:#5a6b7b;margin-top:4px">✍️ 开口话术：“我是罗克韦尔/PTC，{esc(ask)}。”</div>'
            f'</div></div>')

# 公司级真查档案（AI 逐家上网查证后写入；未入档的走产品级通用判断）
# 字段: role=该找谁 why=为什么是他 chief=已具名查到的对接人/部门 + 证据
PRODUCT_DOSSIERS={
 "600298":{"role":"信息中心（韩陈）+ 集团装备部（王秀峰部长）",
   "why":"安琪是流程行业多基地（德宏/普洱/崇左/济宁/赤峰/可克达拉），已拿两化融合贯标+国家级智能制造场景，MES不空白；FTPC 切的是多基地标准统一，找信息中心而非车间",
   "chief":"已查证：装备部部长王秀峰主谈2026技改（AI+自动化升级）；信息中心高工韩陈负责5G工业互联网标识项目。董秘留言请转信息中心/装备部。",
   "email":{"to":"wangxiufeng@angelyeast.com","alt":"wang.xiufeng@ / wangxf@angelyeast.com","note":"官网域名angelyeast.com；按A股制造业惯例推断，首封先测退信"}},
 "603313":{"role":"集团制造/IT负责人（南通零压/上海总部）",
   "why":"六大基地横跨中/美/塞尔维亚/西班牙/泰国，已用RFID+智能吊挂+立体仓；海外新基地（塞/泰/西）正是多工厂MES标准统一的选型窗，典型FTPC标的",
   "chief":"未具名：六大基地跨三洲，制造标准统一是集团议题；董秘留言请转'海外基地制造/信息化负责人'。"},
 "002812":{"role":"注意：FTPC优先级下调，转 ThingWorx/DataMosaix 向上销售",
   "why":"恩捷2026年8月已全面上线SCADA数采+能源管理系统，并与安脉盛做AI制膜厚度控制，OT数据层已自建；FTPC不是空白切入，数据上下文化/OEE才是切口",
   "chief":"已查证：SCADA生产数采2026-08全面上线、能源管理系统2026-09试运行。建议走数字化/智造团队而非MES采购口径。"},
}

HQ_REGION={
 "603166":"华南·广西","001380":"长三角·浙江","002454":"长三角·上海","300852":"珠三角·广东",
 "301039":"珠三角·广东","600261":"长三角·浙江","002543":"珠三角·广东","002815":"珠三角·广东",
 "001359":"华中·湖北","605277":"珠三角·广东","002843":"华中·湖南","300791":"珠三角·广东",
 "002429":"珠三角·广东","300818":"华中·江西","605099":"长三角·江苏",
 # FTPC
 "600459":"西南·云南","300318":"京津冀·北京","300811":"珠三角·广东","300441":"长三角·浙江",
 "688179":"长三角·上海","600298":"华中·湖北","301353":"长三角·浙江","603313":"长三角·江苏",
 "002988":"珠三角·广东","002812":"长三角·上海","601500":"长三角·江苏","600973":"长三角·江苏",
 # ThingWorx
 "002469":"华东·山东","600343":"西北·陕西","000657":"华中·湖南","600176":"长三角·浙江",
 "600819":"长三角·上海","000603":"西北·甘肃","002078":"华东·山东","002140":"华东·安徽",
 "600793":"西南·四川","603077":"西南·四川","000717":"珠三角·广东","600585":"华东·安徽",
 # Windchill
 "300831":"西北·陕西","301517":"西北·陕西","000532":"珠三角·广东","600262":"华北·内蒙古",
 "605259":"长三角·浙江","600980":"京津冀·北京","688385":"长三角·上海","000544":"华中·河南",
 "603906":"长三角·江苏","300183":"华东·山东",
 # DataMosaix
 "688529":"长三角·上海","603002":"珠三角·广东","688306":"长三角·浙江","603011":"华东·安徽",
}
def region_of(code):
    c=(code or "").replace(".SH","").replace(".SZ","").zfill(6)
    return HQ_REGION.get(c, "其他/待确认")


# ============================================================
# 6.4 AI 自动深查（GLM-4.7-Flash 免费 + Tavily 免费额度）
# 新线索自动查：信号判断 / 总部区域 / 具名负责人 / 邮箱推断
# 硬规则：人名必须能在检索到的原文里逐字命中且带可点开来源，否则一律"未具名"
# ============================================================
GLM_KEY = os.environ.get("GLM_KEY", "").strip()
TVLY_KEY = os.environ.get("TAVILY_KEY", "").strip()
AI_CACHE = DATA / "dossiers.json"
AI_MAX = int(os.environ.get("AI_MAX", "3") or 3)   # 每次运行最多深查几家，护住免费额度
AI_MODELS = ["glm-4-flash", "glm-4.7-flash"]
REGION_RE = re.compile(r"^(长三角|珠三角|京津冀|华中|华南|西南|西北|华东|华北|东北)·[\u4e00-\u9fa5]{2,4}$")
PROV_REGION = {
 "江苏":"长三角","浙江":"长三角","上海":"长三角","安徽":"华东","山东":"华东",
 "广东":"珠三角","广西":"华南","海南":"华南","福建":"华东",
 "北京":"京津冀","天津":"京津冀","河北":"京津冀","山西":"华北","内蒙古":"华北",
 "湖北":"华中","湖南":"华中","河南":"华中","江西":"华中",
 "四川":"西南","重庆":"西南","云南":"西南","贵州":"西南","西藏":"西南",
 "陕西":"西北","甘肃":"西北","宁夏":"西北","新疆":"西北","青海":"西北",
 "辽宁":"东北","吉林":"东北","黑龙江":"东北",
}
def norm_region(txt):
    """把模型给的区域说法归一成 大区·省；认不出来则返回空"""
    t = (txt or "").strip()
    if REGION_RE.match(t): return t
    for prov, reg in PROV_REGION.items():
        if prov in t:
            return f"{reg}·{prov}"
    return ""
_AI_CACHE = {}

def _py_full(name):
    """姓名全拼，用于生成邮箱。优先本地拼音库，避免模型自己拼错（如 蔡报贵→caobaoji）"""
    try:
        from pypinyin import lazy_pinyin
    except Exception:
        try:
            if os.environ.get("GITHUB_ACTIONS"):
                import subprocess
                subprocess.run([sys.executable, "-m", "pip", "install", "-q", "pypinyin"],
                               capture_output=True, timeout=120)
            from pypinyin import lazy_pinyin
        except Exception:
            return ""
    try:
        out = "".join(lazy_pinyin(name)).lower()
        return out if re.match(r"^[a-z]{2,20}$", out) else ""
    except Exception:
        return ""

def _http_json(url, payload, headers=None, timeout=40):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json", **(headers or {})})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))

def tavily(q, n=4):
    if not TVLY_KEY: return []
    for attempt in range(3):
        try:
            d = _http_json("https://api.tavily.com/search",
                           {"api_key": TVLY_KEY, "query": q, "max_results": n,
                            "search_depth": "basic", "include_answer": False}, timeout=30)
            return [{"url": x.get("url",""), "title": x.get("title",""), "content": x.get("content","")}
                    for x in (d.get("results") or [])]
        except Exception as e:
            log(f"    tavily 第{attempt+1}次失败：{str(e)[:70]}")
            time.sleep(2 + attempt * 3)
    return []

def glm_json(system, user):
    if not GLM_KEY: return None, None
    for i, m in enumerate(AI_MODELS):
        try:
            d = _http_json("https://open.bigmodel.cn/api/paas/v4/chat/completions",
                {"model": m, "messages": [{"role":"system","content":system},
                                          {"role":"user","content":user}],
                 "temperature": 0.1, "stream": False},
                {"Authorization": f"Bearer {GLM_KEY}"}, timeout=100)
            txt = d["choices"][0]["message"]["content"] or ""
            txt = re.sub(r"```json|```", "", txt).strip()
            s, e = txt.find("{"), txt.rfind("}")
            if s >= 0 and e > s:
                return json.loads(txt[s:e+1]), m
        except Exception as ex:
            log(f"    glm {m} 不可用：{str(ex)[:70]}")
            time.sleep(3 + i * 4)          # 指数退避，避开免费档限流
    return None, None

def _load_ai_cache():
    try:
        if AI_CACHE.exists():
            d = json.loads(AI_CACHE.read_text(encoding="utf-8"))
            for k, v in d.items():
                _AI_CACHE[k] = v
                PRODUCT_DOSSIERS.setdefault(k, v)
                if v.get("region") and REGION_RE.match(v["region"]):
                    HQ_REGION[k] = v["region"]
            log(f"AI 档案缓存载入 {len(d)} 家")
    except Exception as e:
        log(f"AI 档案缓存读取失败：{str(e)[:60]}")

def _save_ai_cache():
    if not _AI_CACHE: return
    try:
        AI_CACHE.parent.mkdir(exist_ok=True)
        AI_CACHE.write_text(json.dumps(_AI_CACHE, ensure_ascii=False, indent=1), encoding="utf-8")
        log(f"AI 档案缓存已写盘（{len(_AI_CACHE)} 家）")
    except Exception as e:
        log(f"AI 档案缓存写盘失败：{str(e)[:60]}")

_AI_SYS = (
 "你是罗克韦尔/PTC 售前线索分析员。只依据我给你的检索材料作答，禁止臆造。"
 "输出严格 JSON，不要任何解释文字。人名必须逐字出现在材料中才可填，否则填空字符串。"
)

def _same_subject(nm, c):
    """人名所在材料是否讲的就是这家公司：公司名后 4 字内出现'股份'，拦掉'浙大联科'这类同名不同家"""
    if not (nm and c): return False
    return re.search(re.escape(nm) + r"[^，。；\s]{0,4}股份", c) is not None

def ai_dossier(code, nm, p, finding):
    """对一家公司做一次自动深查，返回 dossier dict；证据不足则返回 None"""
    if not (GLM_KEY and TVLY_KEY): return None
    nm = (nm or "").strip()
    if len(nm) < 2: return None
    logs = []
    qs = [f"{nm} 海外 建厂 投资 新基地 扩产 2025 2026",
          f"{nm} 信息化 数字化 IT 负责人 MES 智能制造 项目",
          f"{nm} 总部 地址 年报 董事会秘书 联系电话",
          f'"{nm}" 副总经理 技术总监 生产总监 简历 高管',   # 高管/任命公告页实名带职务，实测命中最高
          f'"{nm}" 中标 智能制造 项目 负责人 联系人']      # 招投标材料常留项目联系人真名
    for q in qs:
        logs += tavily(q, 4)
    seen, hits = set(), []
    for h in logs:
        u = h.get("url") or ""
        if u and u not in seen:
            seen.add(u); hits.append(h)
    if not hits:
        log("    未检索到任何材料，跳过")
        return None
    urls = [h["url"] for h in hits]
    ctx = "\n".join(f"[{i+1}] {h['title']}\nURL: {h['url']}\n{h['content'][:700]}"
                    for i, h in enumerate(hits))[:11000]
    playbook = (f"产品：{p['name']}｜定位：{p['pos']}\n目标客户：{p['who']}\n"
                f"验证真需求的三问：{' / '.join(p['three'])}")
    user = f"""公司：{nm}（{code}），本次触发的公告标题：{finding.get('title','')}

{playbook}

检索材料（含编号与URL，只能引用这些URL）：
{ctx}

请输出 JSON，字段：
{{
 "demand":"是|否|待验证",           // 该信号是否会引出上述产品的选型需求（你自己的判断）
 "demand_reason":"60字内，依据哪条材料做出的判断",
 "signal":"40字内，公告/材料里最关键的一条事实（带时间）",
 "region":"大区·省，如 长三角·浙江；大区只能是 长三角|珠三角|京津冀|华中|华南|西南|西北|华东|华北|东北",
 "grade":"高|中|低",                  // 现在打它的优先级
 "role":"该找的对口部门/岗位",
 "why":"50字内，为什么找他而不是别人",
 "chief_name":"材料中逐字出现的负责人姓名；查不到填空字符串",
 "chief_title":"他的职务，如 信息中心主任",
 "chief_source":"chief_name 出现的那条材料的完整URL，必须来自上面列表；查不到填空字符串",
 "insider_name":"专利发明人或招标公告里的项目联系人真名（只能是材料中逐字出现的，且不是董秘/董事长）；查不到填空字符串",
 "insider_title":"他的身份，如 专利发明人 / 某项目联系人 / 技术中心主任",
 "insider_source":"insider_name 出现的那条材料的完整URL，必须来自上面列表；查不到填空字符串",
 "domain":"该公司企业邮箱域名，如 angelyeast.com；查不到填空字符串",
 "email_guess":"按该域名与中文姓名拼音推断的最可能邮箱；查不到填空字符串"
}}"""
    d, model = glm_json(_AI_SYS, user)
    if not d:
        log("    模型未返回有效 JSON")
        return None
    log(f"    模型={model} demand={d.get('demand')}")

    # —— 硬校验：人名必须在我们真实检索到的页面原文里逐字出现 ——
    name = (d.get("chief_name") or "").strip()
    src = (d.get("chief_source") or "").strip()
    bad_words = ("未具名", "未查到", "未知", "不详", "N/A")
    if name and (any(w in name for w in bad_words) or len(name) < 2):
        name = ""
    if name:
        # 优先用模型给的那条来源；否则在我们抓到的材料里找一条真的写了这个人名、且讲的是这家公司的
        cand = []
        if src in urls:
            cand.append(next(h for h in hits if h["url"] == src))
        cand += [h for h in hits if h["url"] != src]
        got = ""
        for h in cand:
            c = (h["content"] or "")
            if name in c and _same_subject(nm, c):
                got = h["url"]; break
        if got:
            src = got
        else:
            log(f"    人名「{name}」在检索原文中未命中 → 未具名")
            name = ""; src = ""

    # —— 内部线索人（专利发明人/招标项目联系人）：同样要求原文逐字命中 ——
    iname = (d.get("insider_name") or "").strip()
    isrc = ""
    if iname and not any(w in iname for w in bad_words) and len(iname) >= 2 and iname != name:
        for h in hits:
            c = (h["content"] or "")
            if iname in c and _same_subject(nm, c):
                isrc = h["url"]; break
        if not isrc:
            log(f"    内部线索人「{iname}」未在原文命中 → 丢弃"); iname = ""
    else:
        iname = ""
    insider = None
    if iname:
        insider = {"name": iname,
                   "title": (d.get("insider_title") or "内部线索人")[:40],
                   "src": isrc}

    # —— 域名校验：必须真实出现在材料里 ——
    domain = (d.get("domain") or "").strip().lower().lstrip("@")
    email = None
    if name and domain:
        dom_ok = any(domain in (h["content"] or "").lower() or domain in h["url"].lower() for h in hits)
        guess = (d.get("email_guess") or "").strip().lower()
        py = _py_full(name)                     # 自己算拼音，不信模型拼写
        target = (py + "@" + domain) if py else guess
        ok_fmt = bool(re.match(r"^[a-z0-9._-]+@[a-z0-9.-]+\.(com|cn|com\.cn|net)$", target or ""))
        if dom_ok and ok_fmt and target.endswith("@" + domain):
            email = {"to": target,
                     "alt": "其他常见格式：名.姓 / 姓+名首字母（如 shenzw）",
                     "note": f"官网域名{domain}；姓名全拼由程序生成，按A股制造业惯例推断，首封必测退信"}

    _title = (d.get("chief_title") or "").strip()
    referral = ""
    if name and any(k in _title for k in ("董秘", "董事会秘书", "证券事务")):
        chief = (f"未具名（决策人）：已锁定转介人 —— {_title} {name}。"
                 f"董秘只能转介，不是拍板人；先请他转制造/信息化负责人，再谈方案")
        referral = name
    elif name:
        chief = f"🤖 AI 查证：{_title} {name}（来源可点开核实）"
    elif insider and any(k in insider["title"] for k in ("联系人","项目负责","主任","经理","部长","总监")):
        chief = (f"未具名（决策人）：但从招投标材料挖到项目线具名人 —— "
                 f"{insider['title']} {insider['name']}。可直接按招标联系方式触达，比董秘更近一层")
    else:
        chief = "未具名：AI 已跑 5 路公开源（公告/官网/新闻/招聘/专利·招投标），未见可靠具名负责人；走董秘转介 + 年报证券部直线"

    region = norm_region(d.get("region"))
    # 注意：不做"正文里扫省份"的兜底——会抓到公告里的子公司/项目所在地，
    # 把总部区域judge错（曾把云南的贵研铂业判成山东）。认不出就留待确认。
    ds = {"name": nm,
          "role": (d.get("role") or PRODUCT_WHO_ROLE.get(p["key"], ("信息化/数字化负责人",""))[0])[:60],
          "why": (d.get("why") or "")[:100],
          "chief": chief,
          "demand": d.get("demand") or "待验证",
          "demand_reason": (d.get("demand_reason") or "")[:100],
          "signal": (d.get("signal") or "")[:80],
          "grade": d.get("grade") if d.get("grade") in ("高","中","低") else "中",
          "region": region if region else "",
          "auto": True, "model": model,
          "src": src if name else "",
          "referral": referral,
          "chief_name": (referral or name),
          "insider": insider,
          "researched": datetime.date.today().isoformat()}
    if email:
        if referral:
            email["note"] = "转介用邮箱：发给董秘请他转对口部门，不是直接谈方案"
        ds["email"] = email
    if not region: log("    区域未能归一（保留待确认）")
    return ds

PLEX_PSEUDO = dict(key="PLEX", emoji="🚀", name="罗克韦尔 PLEX（SaaS MES）",
    pos="海外绿地建厂/新基地投产的快速 MES 落地：几周上线、免运维、按订阅付费，绕开传统 MES 几百万+半年实施",
    who="有海外建厂/新基地/并购扩张的上市制造企业；多工厂要统一制造运营标准",
    three=["新厂投产时间表定了吗？老厂与新厂是不是一套标准？",
           "多工厂的排产与制造执行口径统一了吗？",
           "这次扩产是照搬老厂方案，还是重新选型？"],
    hook="“传统 MES 海外新厂第一年几百万、半年才上线；PLEX SaaS 一年几十万、几周上线、免运维。”")

def queue_pass():
    """即查模式接 AI：页面按钮在 queue/ 下建一个任务文件，CI 跑时自动查完入档并删除该文件"""
    _load_ai_cache()
    if not (GLM_KEY and TVLY_KEY): return
    qdir = ROOT / "queue"
    if not qdir.exists(): return
    files = sorted(qdir.glob("*.txt"))
    if not files: return
    log(f"即查队列：{len(files)} 个待查任务")
    for fp in files[:max(1, AI_MAX)]:
        try:
            raw = fp.read_text(encoding="utf-8").strip()
        except Exception:
            raw = ""
        if not raw:
            fp.unlink(missing_ok=True); continue
        parts = [x.strip() for x in raw.replace("｜","|").split("|")]
        nm = parts[0]
        pk = (parts[1].upper() if len(parts) > 1 and parts[1] else "PLEX")
        p = next((x for x in PRODUCTS if x["key"] == pk), None) or PLEX_PSEUDO
        code = ""
        m = re.search(r"\d{6}", nm)
        if m: code = m.group(0)
        if not code: code = "Q" + hashlib.md5(nm.encode("utf-8")).hexdigest()[:8]
        if code in PRODUCT_DOSSIERS:
            log(f"  队列：{nm} 已入档，跳过")
            fp.unlink(missing_ok=True); continue
        log(f"  队列深查：{nm} [{p['key']}]")
        ds = ai_dossier(code, nm, p, {"title": "（用户即查模式提交）"})
        if ds:
            PRODUCT_DOSSIERS[code] = ds; _AI_CACHE[code] = ds
            if ds.get("region"): HQ_REGION[code] = ds["region"]
            log(f"    ✓ 队列入档（{ds['demand']}/{ds['grade']}）")
            fp.unlink(missing_ok=True)
            _save_ai_cache()
        else:
            log(f"    队列：{nm} 未查到，保留文件待下次重试")

def ai_research_pass():
    """新线索自动查档；已入档的跳过，不重复烧额度"""
    _load_ai_cache()
    if not (GLM_KEY and TVLY_KEY):
        log("未配置 GLM_KEY / TAVILY_KEY，跳过 AI 自动深查（页面照常生成）")
        return
    todo = []
    _seen_code = set()
    for p in PRODUCTS:
        for h in product_findings.get(p["key"], []):
            c = (h.get("sec") or "").replace(".SH","").replace(".SZ","").zfill(6)
            if not c or c in PRODUCT_DOSSIERS: continue
            nm = (h.get("name") or c).strip()
            if len(nm) < 2 or c in _seen_code: continue
            _seen_code.add(c)
            try:
                r = analyze(p["key"], nm, h.get("title",""), h.get("ts"))
            except Exception:
                r = None
            if not r or r[0] != "高":      # 只深查"正好在选型窗口"的高优先线索，省钱省时
                continue
            d0 = age_info(h.get("ts"))[0]
            todo.append((d0 if d0 is not None else 999, c, nm, p, h))
    # PLEX 的线索来自 new_findings（公告扫描），不在 product_findings，必须单独入队
    for f in new_findings:
        if f.get("phase") != "window": continue
        c = (f.get("sec") or "").replace(".SH","").replace(".SZ","").zfill(6)
        if not c or c in PRODUCT_DOSSIERS or c in _seen_code: continue
        nm = (f.get("name") or c).strip()
        if len(nm) < 2: continue
        _seen_code.add(c)
        d0 = age_info(f.get("ts"))[0]
        todo.append((d0 if d0 is not None else 999, c, nm, PLEX_PSEUDO, f))
    # PLEX 是首页，优先深查：先 PLEX 窗口期，再按公告新旧
    todo.sort(key=lambda x: (0 if x[3].get("key") == "PLEX" else 1, x[0]))
    if todo:
        log(f"待 AI 深查（仅高优先窗口期）{len(todo)} 家，本次上限 {AI_MAX} 家")
    done = 0
    for _age, c, nm, p, h in todo:
        if done >= AI_MAX: break
        log(f"  AI 深查：{nm} [{p['key']}]")
        ds = ai_dossier(c, nm, p, h)
        done += 1
        if ds:
            PRODUCT_DOSSIERS[c] = ds
            _AI_CACHE[c] = ds
            if ds.get("region"): HQ_REGION[c] = ds["region"]
            log(f"    ✓ 入档（{ds['demand']}/{ds['grade']}）")
    _save_ai_cache()

def card_for(f):
    g=PHASE_GRADE.get(f.get("phase"),"中")
    ind=industry_of((f["name"] or ""), f["title"])
    judge=" ".join(f["rule"])
    nm=f["name"] or f["sec"]
    code=(f["sec"] or "").replace(".SH","").replace(".SZ","").zfill(6)
    d=DOSSIERS.get(code)
    ai=PRODUCT_DOSSIERS.get(code)
    if d:
        chief,lev,steps,loc=d["chief"],d["lev"],d["steps"],d["loc"]
        hook=d.get("hook") or "“传统 MES 海外新厂第一年几百万、半年才上线；PLEX SaaS 一年几十万、几周上线、免运维。”"
        badge='<span class="ev" style="color:#137333">✓ AI 已查证深挖</span>'
    else:
        chief,lev,steps,loc=auto_dossier(nm, f["title"], f["rule"], ind)
        hook="“传统 MES 海外新厂第一年几百万、半年才上线；PLEX SaaS 一年几十万、几周上线、免运维。”"
        badge='<span class="ev" style="color:#0b6bcb">⚙ AI 已自动解析公告（可深挖）</span>'
    ai_html=""
    if ai:
        _dc={"是":"#137333","否":"#9aa7b4","待验证":"#b26a00"}.get(ai.get("demand"),"#b26a00")
        _in=(ai.get("insider") or {})
        _inetxt=""
        if _in.get("name"):
            _inetxt=('<div style="font-size:12px;color:#7a4b00;line-height:1.6">🧩 内部线索人：'
                     + esc(_in.get("title") or "内部线索人") + ' <b>' + esc(_in.get("name","")) + '</b>'
                     + ' <a class="ev" style="color:#0b6bcb" href="' + esc(_in.get("src","")) + '" target="_blank">🔗 出处</a></div>')
        _srchtml=""
        if ai.get("src"):
            _srchtml=('<div style="font-size:11.5px"><a class="ev" style="color:#0b6bcb" href="'
                      + esc(ai.get("src","")) + '" target="_blank">🔗 点开核实人名出处</a></div>')
        ai_html=('<div style="margin-top:4px;background:#f7f5ff;border:1px solid #e3dcff;border-radius:8px;padding:8px 10px">'
                 '<div style="font-size:12.5px;line-height:1.7">'
                 '<span style="color:#7a5af8;font-weight:700">🤖 AI 自动查证</span>'
                 ' <span style="color:' + _dc + ';font-weight:700">选型判断：' + esc(ai.get("demand","")) + '</span>'
                 '<span style="color:#5a6b7b"> —— ' + esc(ai.get("demand_reason","")) + '</span></div>'
                 '<div style="font-size:12px;color:#33414f;line-height:1.6">🔎 信号：' + esc(ai.get("signal","")) + '</div>'
                 '<div style="font-size:12px;color:#0b6bcb;line-height:1.6">👤 ' + esc(ai.get("chief","")) + '</div>'
                 + _srchtml + _inetxt + '</div>')
        if ai.get("chief"): chief=ai["chief"]
        badge='<span class="ev" style="color:#7a5af8">🤖 AI 自动查证（已覆盖模板判断）</span>'
    steps_html="".join(f'<div class="step"><b>{i+1}.</b>{esc(m)}</div>' for i,m in enumerate(steps))
    mail_html=""
    if ai and ai.get("email"):
        import urllib.parse as _u3
        _e3=ai["email"]
        _sub3=_u3.quote(f"【{nm}】海外新厂 MES 快速上线交流——罗克韦尔 PLEX")
        _body3=_u3.quote(f"您好，\n\n关注到 {nm} 的产能扩张/海外基地布局。PLEX 是订阅制 SaaS MES，几周上线、免运维，"
                         f"适合新厂赶投产节点、老厂不想再投几百万的传统 MES。\n"
                         f"想用 15 分钟聊聊：新厂 MES 时间表、与老厂是否一套标准。\n\n罗克韦尔自动化")
        mail_html=(f'<div style="margin-top:6px"><a class="btn btn-mail" target="_blank" '
                   f'href="mailto:{esc(_e3["to"])}?subject={_sub3}&body={_body3}">📧 直接写邮件给 {esc(_e3["to"])}</a>'
                   f'<div style="font-size:11.5px;color:#8494a4;margin-top:2px">📮 推断邮箱 · {esc(_e3.get("note",""))} ｜ 备选：{esc(_e3.get("alt",""))}</div></div>')
    rg=region_of(code).split("·")[0]
    return (f'<div class="card open" data-rg="{esc(rg)}">'
      f'<div class="card-h" onclick="this.parentNode.classList.toggle(\'open\')">'
      f'<span class="cid">{g}</span><span class="nm">{esc(nm)}</span><span class="code">{esc(f["sec"] or "")}</span>'
      f'<span class="base">📍 {esc(loc)}</span>'
      f'<span class="right"><span class="grade">🏭 {esc(ind)}</span> <span class="grade">🗺️ {esc(region_of(code).split("·")[0])}</span></span></div>'
      '<div class="card-b">'
      f'<div class="row"><b>为什么是机会：</b>{esc(judge)}</div>'
      f'<div class="row">{badge}</div>'
      f'{ai_html}'
      f'<div class="row"><b>决策人：</b>{esc(chief)}</div>'
      f'<div class="row"><b>杠杆点：</b>{esc(lev)}</div>'
      f'<div class="row"><b>怎么开口：</b>{esc(hook)}</div>'
      f'{mail_html}'
      f'<div style="margin-top:8px"><b style="font-size:12.5px;color:#b26a00">📋 你要做（按顺序）：</b>{steps_html}</div>'
      f'{contact_links(code,nm)}'
      f'<details><summary>证据</summary><div class="ev"><a href="{esc(f["url"])}" target="_blank">公告原文 ↗</a></div></details>'
      f'<div style="margin-top:8px"><button class="btn fu-btn" data-fu="{esc(nm)}">📌 标记今天跟进</button></div>'
      '</div></div>')
# 先跑 AI 深查，再渲染卡片——否则 PLEX 卡片永远拿不到当天的档案
queue_pass()
ai_research_pass()
cards_win="".join(card_for(f) for f in new_findings if f.get("phase")=="window")
cards_fresh="".join(card_for(f) for f in new_findings if f.get("phase")=="fresh")
cards_old="".join(card_for(f) for f in new_findings if f.get("phase")=="old")
n_window=len([f for f in new_findings if f.get("phase")=="window"])
# PLEX 区域筛选条
_plex_rg=[]
for f in new_findings:
    r=region_of(f.get("sec")).split("·")[0]
    if r not in _plex_rg: _plex_rg.append(r)
_plex_chips='<button class="rg-chip active" data-rg="全部">全部</button>'+"".join(f'<button class="rg-chip" data-rg="{esc(r)}">{esc(r)}</button>' for r in _plex_rg)
_plex_filter_js="""
<style>.rg-chip{cursor:pointer;border:1px solid #c8d4e0;background:#fff;color:#33475b;border-radius:14px;padding:4px 12px;font-size:12px;margin:0 6px 6px 0}.rg-chip.active{background:#003DA5;color:#fff;border-color:#003DA5}.rg-bar{position:sticky;top:0;background:#fff;padding:8px 0;z-index:9;border-bottom:1px solid #eef2f7}.srch{width:100%;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:7px 11px;font-size:12.5px;color:#33475b;margin:2px 0 8px 0;outline:none}.srch:focus{border-color:#003DA5;box-shadow:0 0 0 3px rgba(0,61,165,.08)}.srch-empty{display:none;padding:10px;font-size:12px;color:#8494a4;text-align:center}</style>
<script>
(function(){var bar=document.getElementById('rgbar');if(!bar)return;
var inp=bar.querySelector('.srch'), rg='全部';
function apply(){
 var q=(inp&&inp.value||'').trim().toLowerCase(), n=0;
 document.querySelectorAll('.card[data-rg]').forEach(function(el){
   var okr=(rg==='全部'||el.getAttribute('data-rg')===rg);
   var oks=!q||(el.textContent||'').toLowerCase().indexOf(q)>=0;
   if(okr&&oks) n++;
   el.style.display=(okr&&oks)?'':'none';});
 var em=bar.parentNode.querySelector('.srch-empty');
 if(em){ em.style.display=(q&&n===0)?'block':'none'; }
}
bar.querySelectorAll('.rg-chip').forEach(function(c){c.onclick=function(){
 bar.querySelectorAll('.rg-chip').forEach(function(x){x.classList.remove('active')});c.classList.add('active');
 rg=c.getAttribute('data-rg'); apply();};});
if(inp){ inp.oninput=apply; }
})();
</script>"""

reject_html="".join(f"<tr><td>{r['id']}</td><td>{r['name']} {r['code']}</td><td>{esc(r['why'])}</td></tr>" for r in REJECTED)

# 标记跟进 + 隔天回访提醒（浏览器 localStorage，免后端）
FU_JS = """
<div id="fu-banner"></div>
<script>
(function(){
  var KEY='fu_marks_v1';
  function load(){try{return JSON.parse(localStorage.getItem(KEY)||'{}')}catch(e){return{}}}
  function save(m){localStorage.setItem(KEY,JSON.stringify(m))}
  function fmt(ts){var d=new Date(ts);return (d.getMonth()+1)+'/'+d.getDate();}
  function paintBtns(){
    document.querySelectorAll('.fu-btn').forEach(function(b){
      var m=load()[b.getAttribute('data-fu')];
      if(m){b.textContent='✅ 已标记 '+fmt(m.ts);b.classList.add('done');}
      else{b.textContent='📌 标记今天跟进';b.classList.remove('done');}
    });
  }
  function paintBanner(){
    var m=load(), box=document.getElementById('fu-banner'), now=Date.now(), html='';
    Object.keys(m).forEach(function(k){
      var e=m[k]; if(e.snooze && e.snooze>now) return;
      var days=(now-e.ts)/86400000;
      if(days>=2){
        html+='<div class="fu-item">🔔 <b>'+k+'</b> —— 你 '+Math.floor(days)+' 天前标记跟进，完成了吗？ '
          +'<button data-act="done" data-k="'+k+'">✅ 已完成</button> '
          +'<button data-act="snooze" data-k="'+k+'">⏰ 过两天再问</button></div>';
      }
    });
    box.innerHTML=html;
  }
  document.body.addEventListener('click',function(ev){
    var b=ev.target;
    if(!b||!b.getAttribute) return;
    var k=b.getAttribute('data-fu'), act=b.getAttribute('data-act');
    if(k && b.classList.contains('fu-btn')){
      var m=load();
      if(m[k]){delete m[k];}else{m[k]={name:k,ts:Date.now()};}
      save(m); paintBtns(); paintBanner();
    }
    else if(act){
      var m=load();
      if(act==='done'){delete m[k];}
      else if(act==='snooze'){if(m[k])m[k].snooze=Date.now()+2*86400000;}
      save(m); paintBanner();
    }
  });
  paintBtns(); paintBanner();
})();
</script>"""

def shell(title, active, body):
    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(title)}</title><style>{CSS}</style>
<style>.fu-btn{{cursor:pointer;border:1px solid #c8d4e0;background:#fff;color:#003DA5;border-radius:6px;padding:4px 10px;font-size:12px;margin-left:4px}}.fu-btn.done{{background:#e6f4ea;color:#137333;border-color:#137333}}.fu-item{{background:#fff8e1;border:1px solid #f0c36d;border-radius:8px;padding:10px 12px;margin:8px 0;font-size:13px}}.fu-item button{{margin-left:6px;cursor:pointer;border:none;border-radius:6px;padding:4px 10px;font-size:12px;background:#003DA5;color:#fff}}</style>
</head><body>
<div class="wrap">
{nav_html(active)}
{FU_JS}
{body}
</div></body></html>"""

# ---- PLEX 首页 ----
# 新线索自动深查（先查档，再渲染，页面里看到的就是查过的）
def _ai_archive():
    rows=[]
    for c,d in sorted(PRODUCT_DOSSIERS.items(), key=lambda x: x[1].get("researched",""), reverse=True):
        nm=d.get("name") or c
        grade=d.get("grade",""); dem=d.get("demand","")
        _gc={"高":"#d93025","中":"#b26a00","低":"#8494a4"}.get(grade,"#8494a4")
        _dc={"是":"#137333","否":"#9aa7b4","待验证":"#b26a00"}.get(dem,"#b26a00")
        ch=(d.get("chief") or "")[:90]
        em=(d.get("email") or {}).get("to","")
        rg=d.get("region") or "待确认"
        link=(f'<a class="ev" style="color:#0b6bcb" href="{esc(d["src"])}" target="_blank">出处 ↗</a>' if d.get("src") else "—")
        rows.append('<tr><td><b>'+esc(nm)+'</b><div style="font-size:11px;color:#8494a4">'+esc(c)+'</div></td>'
                    '<td style="white-space:nowrap">'+esc(rg)+'</td>'
                    '<td><span style="color:'+_gc+';font-weight:700">'+esc(grade)+'</span> · <span style="color:'+_dc+'">'+esc(dem)+'</span></td>'
                    '<td>'+esc(ch)+' '+link+'</td>'
                    '<td>'+((('<span style="color:#0b6bcb">'+esc(em)+'</span>')) if em else '<span style="color:#8494a4">—</span>')+'</td></tr>')
    return ('<table style="width:100%;font-size:12px;border-collapse:collapse">'
            '<tr><th style="text-align:left">公司</th><th style="text-align:left">总部大区</th>'
            '<th style="text-align:left">判断</th><th style="text-align:left">决策人</th><th style="text-align:left">邮箱</th></tr>'
            + "".join(rows) + '</table>')
ai_archive_html=_ai_archive()
plex_body=f"""
<div class="top"><h1><img class="hdrlogo" src="logo-plex.png" alt=""> PLEX 每日商机 · 今天该干什么</h1></div>
<div class="kpis">
  <div class="kpi"><div class="n">{NOW}</div><div class="l">上次更新</div></div>
  <div class="kpi"><div class="n">{len(new_findings)}</div><div class="l">近12月线索</div></div>
  <div class="kpi"><div class="n">{n_window}</div><div class="l">🟢 正好选型窗口</div></div>
  <div class="kpi"><div class="n">{manual_total}</div><div class="l">待手动跟进</div></div>
</div>
<div class="rg-bar" id="rgbar">📍 按总部区域筛选：{_plex_chips}<input class="srch" type="search" placeholder="🔍 搜公司名 / 股票代码 / 行业（如 均普、600699）" autocomplete="off"></div>
<div class="srch-empty">没有匹配的线索 —— 用下面的「即查模式」查这家公司</div>
{qk_panel("PLEX",
   "有海外建厂/新基地/并购扩张的上市制造企业；多工厂要统一制造运营标准",
   ["新厂投产时间表定了吗？老厂与新厂是不是一套标准？",
    "多工厂的排产与制造执行口径统一了吗？",
    "这次扩产是照搬老厂方案，还是重新选型？"],
   "PLEX 在汽车零部件与离散制造有大量多工厂落地；本页按公告后 1-9 个月选型窗口排序。",
   PRODUCT_WHO_ROLE["PLEX"][0])}{cards_win or '<div class="empty">今日无窗口期命中</div>'}</div>
<div class="sec"><div class="sec-h">🟡 中优先 · 刚公告（≤1个月，先观察）</div>{cards_fresh or '<div class="empty">无</div>'}</div>
<div class="sec"><div class="sec-h">🗄 低优先 · 超9个月（可能已选型，存量观察）</div>{cards_old or '<div class="empty">无</div>'}</div>
{_plex_filter_js}
<details class="sec"><summary style="cursor:pointer;font-weight:600;color:#5a6b7b">🗒 手动维护的重点客户名单（历史种子，不随每日公告刷新，点开看）</summary>
<div style="margin-top:10px"><b>高优先</b></div>{''.join(card(c) for c in high)}
<div style="margin-top:10px"><b>中优先</b></div>{''.join(card(c) for c in mid)}
<div style="margin-top:10px"><b>低优先</b></div>{''.join(card(c) for c in low)}</details>
<div class="sec"><div class="sec-h">🤖 AI 已深查档案（{len(PRODUCT_DOSSIERS)} 家 · 含即查队列结果，每天自动补）</div>{ai_archive_html or '<div class="empty">暂无</div>'}</div>
<div class="sec"><div class="sec-h">否决池</div><table><tr><th>ID</th><th>公司</th><th>原因</th></tr>{reject_html}</table></div>
<div class="foot">日志 logs/run_{datetime.date.today().strftime('%Y%m%d')}.log ｜ 标记跟进存在本机浏览器，换手机需重新标记</div>
"""
(ROOT/"dashboard.html").write_text(shell("PLEX 每日商机","PLEX",plex_body),encoding="utf-8")
log("dashboard.html 已生成（PLEX 首页）")

# ---- 每款产品一个独立页 ----
for p in PRODUCTS:
    pbody=f"""
<div class="top"><h1><img class="hdrlogo" src="{LOGO[p['key']]}" alt=""> {esc(p['name'])}</h1></div>
<div class="sub">{esc(p['pos'])}</div>
{render_product(p)}
<div class="foot">每日自动扫描巨潮公告 · 时间 {NOW}</div>
"""
    fn=FILEMAP[p["key"]]
    (ROOT/fn).write_text(shell(p["name"], p["key"], pbody), encoding="utf-8")
    log(f"{fn} 已生成")

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
def _send_push(sc, bk, title, body):
    if sc:
        try:
            data = urllib.parse.urlencode({"title": title, "desp": body}).encode("utf-8")
            req = urllib.request.Request(f"https://sctapi.ftqq.com/{sc}.send", data=data)
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

def notify():
    sc = os.environ.get("SERVERCHAN_KEY", "").strip()
    bk = os.environ.get("BARK_KEY", "").strip()
    site = os.environ.get("SITE_URL", "").strip()
    if not (sc or bk):
        log("未配置 SERVERCHAN_KEY / BARK_KEY，跳过推送")
        return
    # 事件驱动：近48小时内出现的新公告信号（不靠固定时钟，安静日不推）
    fresh=[]; _fs=set()
    for p in PRODUCTS:
        for h in product_findings.get(p["key"], []):
            d,_,_ = age_info(h.get("ts"))
            if d is not None and d <= 2:
                sec=h.get("sec")
                if sec in _fs: continue
                _fs.add(sec)
                fresh.append((p["name"], h.get("name") or sec, h.get("title",""), h.get("url","")))
    fresh=fresh[:6]
    today=datetime.date.today()
    if today.weekday() == 0:
        # 周一：全景周报；若恰逢有新信号，一并附上
        n_win = len([f for f in new_findings if f.get("phase")=="window"])
        top = [ (f["name"] or f["sec"]) for f in new_findings if f.get("phase")=="window" ][:5]
        prod_counts = "；".join(f"{p['key']} {len(product_findings.get(p['key'],[]))}条" for p in PRODUCTS)
        title = f"销售线索周报 {TODAY}"
        body = (f"本周扫描汇总（每周一推一次）：\n"
                f"- PLEX 海外建厂线索 {len(new_findings)} 条，其中正好选型窗口 {n_win} 条\n"
                f"- 各产品命中：{prod_counts}\n"
                f"- 本周最该打：{('、'.join(top) if top else '暂无窗口期命中')}\n"
                f"- 待手动跟进 {manual_total} 项\n"
                f"看板：{site or ''}")
        if fresh:
            body += "\n\n🆕 近48h新信号：\n" + "\n".join(f"· [{pn}] {nm}：{tt}" for pn,nm,tt,_ in fresh)
        _send_push(sc,bk,title,body)
    elif fresh:
        # 非周一：仅当出现新信号才即时推，安静日不打扰
        title = f"🆕 新销售信号 {len(fresh)} 条 {TODAY}"
        body = "\n".join(f"· [{pn}] {nm}\n  {tt}\n  {u}" for pn,nm,tt,u in fresh)
        body += f"\n\n打开看板看搭线动作链：{site or ''}"
        _send_push(sc,bk,title,body)
    else:
        log("非周一且近48h无新信号，静默不推送")

notify()

# ============================================================
# 7. 写日志
# ============================================================
log_path = LOGS / f"run_{datetime.date.today().strftime('%Y%m%d')}.log"
log_path.write_text("\n".join(log_lines), encoding="utf-8")
print(f"\n完成。首页：{ROOT/'dashboard.html'}  产品页：ftpc/thingworx/windchill/datamosaix.html")

# CI 里自动提交所有页面（workflow 的 commit 步骤只 add dashboard.html，这里补全多页）
if os.environ.get("GITHUB_ACTIONS"):
    try:
        import subprocess
        def sh(args):
            return subprocess.run(args, cwd=str(ROOT), capture_output=True, text=True)
        sh(["git","checkout","main"])
        sh(["git","add","-A"])
        chk=sh(["git","diff","--cached","--quiet"])
        if chk.returncode != 0:
            sh(["git","config","user.email","plex-bot@users.noreply.github.com"])
            sh(["git","config","user.name","plex-bot"])
            sh(["git","commit","-m","daily: update all product pages"])
            r=sh(["git","push","origin","main"])
            print("self-push:", r.returncode, (r.stderr or "")[-200:])
        else:
            print("self-push: nothing to commit")
    except Exception as e:
        print("self-push skipped:", e)

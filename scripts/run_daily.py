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
   scan_kw=["工业互联网","预测性维护","设备联网","OEE"]),
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
   scan_kw=["研发中心","PLM","产品数据","协同研发"]),
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
   scan_kw=["数据中台","数字孪生","工业大数据","数据治理"]),
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
        if any(w in t for w in ["数据中台","工业大数据","数字孪生","数据治理","工业互联网平台","数据湖"]):
            sig=("数据平台建设", f"{ind}企业建数据平台，DataMosaix 窗口")
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

def render_product(p):
    raw = product_findings.get(p["key"], [])
    _seen=set(); _raw=[]
    for h in raw:
        k=(h.get("sec"),h.get("title"))
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
    leads=leads[:10]
    cards=""
    for g,judge,lev,agelab,ind,nm,h in leads:
        import urllib.parse as _up
        sq="https://www.google.com/search?q="+_up.quote(nm+" 数字化 信息化 负责人 MES")
        cards+=('<div class="lead">'
          f'<div class="lead-top"><span class="badge {GRADE_CLS[g]}">{g}</span><b>{esc(nm)}</b><span class="lc">{esc(h["sec"] or "")}</span>'
          f'<span class="indtag">🏭 {esc(ind)}</span></div>'
          f'<a class="lead-t" href="{h["url"]}" target="_blank">{esc(h["title"])}</a>'
          f'<div class="why"><b>🔍 关键信息：</b>{esc(judge)}</div>'
          f'<div class="nxt"><b>🎯 下一步怎么做：</b>{esc(lev).replace("杠杆：","")}</div>'
          f'<a class="btn btn-link" href="{h["url"]}" target="_blank">📄 看公告原文</a> '
          f'<a class="btn btn-link" href="{sq}" target="_blank">🔍 找联系人</a>'
          '</div>')
    if not cards:
        cards='<div class="empty" style="padding:14px;font-size:12.5px;color:#8494a4">今日无合格线索（已过滤财报/治理类公告，只留有新产能·并购·技改信号的；雷达每天自动扫）</div>'
    qs="".join(f'<div class="q">问{i+1}：{esc(q)}</div>' for i,q in enumerate(p["three"]))
    return f"""
<div class="prod" id="{p['key']}">
  <div class="prod-h">{p['emoji']} {esc(p['name'])}<small>{esc(p['pos'])}</small></div>
  <div class="prod-b" style="grid-template-columns:1fr">
    <div>
      <div style="font-size:12.5px;color:#33475b;margin:2px 0 10px">🆕 今日自动线索 <b style="color:#003DA5">{total}</b> 条 · 已按优先级排序（高在前）</div>
      {cards}
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
def card_for(f):
    g=PHASE_GRADE.get(f.get("phase"),"中")
    ind=infer_industry(f["title"]+(f["name"] or ""))
    judge=" ".join(f["rule"])
    nm=f["name"] or f["sec"]
    sq="https://www.google.com/search?q="+_up2.quote(nm+" 海外工厂 IT 制造 负责人")
    # 自动补全打法：找谁 / 杠杆 / 话术（不靠用户自己探索）
    who="海外基地厂长 / 工厂 IT 经理 / 质量经理（新厂三巨头）"
    lever=("① 脉脉/LinkedIn 搜【"+esc(nm)+"】+“海外 IT / 墨西哥 / 东南亚 / 工厂 IT”岗——在招这类人=选型正当时；"
           "② 通过公告里的董秘/证券部转数字化负责人；"
           "③ 若该公司产线已用罗克韦尔 PLC，走存量客户渠道直接约")
    hook=("“海外新厂按传统 MES：当地服务器+机房+国内顾问飞过去，第一年几百万、半年才上线；"
          "PLEX 是 SaaS，一年几十万、几周上线，免运维，直接满足北美/欧洲客户的质量追溯审核。”")
    return ('<div class="lead">'
      f'<div class="lead-top"><span class="badge {GRADE_CLS[g]}">{g}</span><b>{esc(nm)}</b><span class="lc">{esc(f["sec"] or "")}</span>'
      f'<span class="indtag">🏭 {esc(ind)}</span></div>'
      f'<a class="lead-t" href="{esc(f["url"])}" target="_blank">{esc(f["title"])}</a>'
      f'<div class="why"><b>🔍 为什么是机会：</b>{esc(judge)}</div>'
      f'<div class="nxt"><b>🎯 找谁：</b>{who}<br><b>🧭 怎么搭上线：</b>{lever}<br><b>🎣 第一句怎么开口：</b>{hook}</div>'
      f'<a class="btn btn-link" href="{esc(f["url"])}" target="_blank">📄 看公告原文</a> '
      f'<button class="btn fu-btn" data-fu="{esc(nm)}">📌 标记今天跟进</button>'
      '</div>')
cards_win="".join(card_for(f) for f in new_findings if f.get("phase")=="window")
cards_fresh="".join(card_for(f) for f in new_findings if f.get("phase")=="fresh")
cards_old="".join(card_for(f) for f in new_findings if f.get("phase")=="old")
n_window=len([f for f in new_findings if f.get("phase")=="window"])

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
plex_body=f"""
<div class="top"><h1><img class="hdrlogo" src="logo-plex.png" alt=""> PLEX 每日商机 · 今天该干什么</h1></div>
<div class="sub">全免费纯云端｜每天早9点自动扫巨潮公告｜点卡片上"📌标记跟进"，隔天会再提醒你</div>
<div class="kpis">
  <div class="kpi"><div class="n">{NOW}</div><div class="l">上次更新</div></div>
  <div class="kpi"><div class="n">{len(new_findings)}</div><div class="l">近12月线索</div></div>
  <div class="kpi"><div class="n">{n_window}</div><div class="l">🟢 正好选型窗口</div></div>
  <div class="kpi"><div class="n">{manual_total}</div><div class="l">待手动跟进</div></div>
</div>
<div class="sec"><div class="sec-h">🟢 高优先 · 正好选型窗口（公告后1-9个月，今天就打）{n_window} 家</div>{cards_win or '<div class="empty">今日无窗口期命中</div>'}</div>
<div class="sec"><div class="sec-h">🟡 中优先 · 刚公告（≤1个月，先观察）</div>{cards_fresh or '<div class="empty">无</div>'}</div>
<div class="sec"><div class="sec-h">🗄 低优先 · 超9个月（可能已选型，存量观察）</div>{cards_old or '<div class="empty">无</div>'}</div>
<details class="sec"><summary style="cursor:pointer;font-weight:600;color:#5a6b7b">🗒 手动维护的重点客户名单（历史种子，不随每日公告刷新，点开看）</summary>
<div style="margin-top:10px"><b>高优先</b></div>{''.join(card(c) for c in high)}
<div style="margin-top:10px"><b>中优先</b></div>{''.join(card(c) for c in mid)}
<div style="margin-top:10px"><b>低优先</b></div>{''.join(card(c) for c in low)}</details>
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
def notify():
    sc = os.environ.get("SERVERCHAN_KEY", "").strip()
    bk = os.environ.get("BARK_KEY", "").strip()
    site = os.environ.get("SITE_URL", "").strip()
    if not (sc or bk):
        log("未配置 SERVERCHAN_KEY / BARK_KEY，跳过推送")
        return
    # 只在每周一推送一次汇总，平时静默不打扰
    if datetime.date.today().weekday() != 0:
        log("非周一，静默不推送")
        return
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
    if sc:
        try:
            data = urllib.parse.urlencode({"title": title, "desp": body}).encode("utf-8")
            req = urllib.request.Request(f"https://sctapi.ftqq.com/{sc}.send", data=data)
            urllib.request.urlopen(req, timeout=8).read()
            log("Server酱周报推送成功")
        except Exception as e:
            log(f"Server酱推送失败（已忽略）：{e}")
    if bk:
        try:
            url = f"https://api.day.app/{bk}/{urllib.parse.quote(title)}/{urllib.parse.quote(body)}"
            urllib.request.urlopen(url, timeout=8).read()
            log("Bark 周报推送成功")
        except Exception as e:
            log(f"Bark 推送失败（已忽略）：{e}")

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

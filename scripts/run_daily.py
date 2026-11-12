# -*- coding: utf-8 -*-
"""
PLEX 每日商机工具 · 本地每日运行脚本
- 全免费、本地运行：唯一外部源是巨潮资讯网免费公告查询接口（cninfo），失败优雅降级。
- 不爬任何登录墙/反爬平台（猎聘/BOSS/脉脉/领英/IATF customerportal 一律不碰，标人工）。
- 产物：dashboard.html（双击即开）+ leads.xlsx + logs/run_YYYYMMDD.log
"""
import json, os, re, sys, time, datetime, hashlib, urllib.request, urllib.parse
import html as html_lib
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
ACTION_KW = ["增资", "设立", "对外投资", "购地", "新建", "扩建", "投产", "子公司", "生产基地", "工厂",
             "开工", "竣工", "募投", "技改", "智能制造", "数字化车间", "搬迁", "二期", "三期", "新增产能"]
NOISE_PLEX = ["回购","减持","增持","分红","股东会","法律意见书","评级","股东大会","摘要","季度报告","年度报告","半年度报告","董事会","监事会","独董","声明","承诺","议案","决议","问询","关注函","担保","质押","会计","减值","审计","更正","补充","批复","处罚","提示","召开","通知","换届","选举","聘任","辞职","薪酬","制度","修订","停牌","复牌","龙虎榜","可转换","债券","票据","授信","贷款"]
# 单关键词检索（多词 AND 太严会漏），动作关键词在客户端过滤
AUTO_KEYWORDS = ["泰国", "墨西哥", "匈牙利", "塞尔维亚", "摩洛哥",
                 "越南", "印尼", "印度", "波兰", "德国", "美国", "巴西",
                 "沙特", "埃及", "土耳其", "马来西亚", "柬埔寨", "菲律宾",
                 "罗马尼亚", "西班牙", "海外基地", "境外投资"]

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


def _brief(txt, maxlen=70):
    """一句话摘要：优先取完整首句；非要截断时落在标点处并补省略号，
    避免出现"……用于"这种话说一半就断的效果。"""
    t = re.sub(r"[\s\n]+", " ", (txt or "")).strip()
    if not t: return ""
    m = re.match(r"^(.{6,}?[。；!？!?])", t)
    if m and len(m.group(1)) <= maxlen:
        return m.group(1)
    if len(t) <= maxlen:
        return t
    cut = t[:maxlen]
    k = max(cut.rfind("，"), cut.rfind("、"), cut.rfind("；"), cut.rfind("。"))
    if k >= int(maxlen * 0.55):
        cut = cut[:k]
    return cut.rstrip("，、；。") + "…"

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
.lead-row{display:flex;align-items:center;gap:10px;padding:10px 14px;border:1px solid #e3ebf3;border-radius:8px;margin-bottom:6px;background:#fff;text-decoration:none;font-size:12.5px;color:#33475b}
.lead-row:hover{background:#f7fafd;border-color:#c8d6e5}
.lead-row .lr-g{flex:0 0 26px;font-weight:700;color:#003DA5}
.lead-row .lr-nm{font-weight:700;color:#003DA5;white-space:nowrap}
.lead-row .lr-code{color:#8494a4;font-size:11px;font-family:ui-monospace,Consolas,monospace}
.lead-row .lr-base{color:#5a6b7d;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:20%}
.lead-row .lr-sum{flex:1;color:#7a8899;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
.lead-row .lr-rg{color:#8494a4;font-size:11.5px;white-space:nowrap}
.lead-row .lr-go{color:#00B5E2;font-weight:700;white-space:nowrap}
.ld-h{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:4px 0 8px}
.ld-nm{font-size:19px;color:#003DA5}
.ld-code{font-family:ui-monospace,Consolas,monospace;font-size:11.5px;color:#8494a4}
.ld-meta{font-size:12px;color:#5a6b7d}
.ld-back{margin-left:auto;font-size:12.5px;color:#00B5E2;font-weight:700;text-decoration:none}
.ld-sum{font-size:12.5px;color:#7a8899;background:#f7fafd;border-radius:8px;padding:8px 12px;margin-bottom:10px}
.ld-body{border:1px solid #e3ebf3;border-radius:10px;padding:14px 16px}
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
<div class="card">
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
    """即查模式：输入公司名 → 落到临时页（inquiry.html），不进主看板、不改代码"""
    tpl = """
<div class="qk" id="qk-__KEY__">
  <div style="font-size:13px;font-weight:700;color:#003DA5">🔎 查一家不在雷达里的公司（临时查询）</div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:7px">
    <input class="qk-name" type="text" placeholder="公司名 / 股票代码，如 上声电子、688533" autocomplete="off"
      style="flex:1;min-width:210px;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:8px 11px;font-size:13px;outline:none">
    <button class="qk-go" style="cursor:pointer;border:none;border-radius:8px;padding:8px 16px;font-size:13px;background:#7a5af8;color:#fff;font-weight:600">🚀 查这家</button>
  </div>
  <div style="margin-top:7px"><a href="inquiry.html" target="_blank" style="font-size:12.5px;color:#7a5af8;text-decoration:none;font-weight:600">🔎 看即查结果 →</a></div>
</div>
<style>.qk{margin:10px 0;padding:12px;background:#fff;border:1px solid #dbe4ef;border-radius:10px;box-shadow:0 1px 3px rgba(0,61,165,.06)}</style>
<script>
(function(){
 var box=document.getElementById('qk-__KEY__');if(!box)return;
 var inp=box.querySelector('.qk-name');
 function go(){
  var n=(inp.value||'').trim();
  if(!n){inp.focus();return;}
  window.open('inquiry.html?c='+encodeURIComponent(n)+'&p=__KEY__','_blank');
 }
 box.querySelector('.qk-go').onclick=go;
 inp.onkeydown=function(e){if(e.key==='Enter')go();};
})();
</script>
"""
    return tpl.replace("__KEY__", key)


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
                             f' <a class="ev" href="mailto:{_e["to"]}?subject={_sub}&body={_body}">{esc(_cta_label(_d))}</a></div>')
                    _chief += _tel_html(_d, esc)
            else:
                _role,_why=_rw[0],_rw[1]; _chief=''
            _lk = code if code not in LEADS else (code + "-" + p["key"])
            _loc = ((_d.get("hq") if _d else "") or "")
            _body = (f'<div class="why"><b>🔍 关键信息：</b>{esc(judge)}</div>'
                     f'<div class="nxt"><b>🎯 下一步怎么做：</b>{esc(lev).replace("杠杆：","")}</div>'
                     f'{_chief}'
                     f'{contact_links(code,nm,ask,role=_role,why=_why,grade=g)}'
                     f'<div style="margin-top:8px"><a class="btn btn-link" href="{h["url"]}" target="_blank">📄 看公告原文</a></div>')
            LEADS.setdefault(_lk, {"n": nm, "s": code, "g": g, "loc": _loc or "—",
                                   "rg": rg, "sum": _brief(judge, 90), "b": _body, "p": h["url"],
                                   "pk": p["key"]})
            cards+=(f'<a class="lead lead-row" data-rg="{esc(rg)}" href="lead.html?c={esc(_lk)}">'
              f'<span class="lr-g">{g}</span><span class="lr-nm">{esc(nm)}</span>'
              f'<span class="lr-code">{esc(code)}</span>'
              f'<span class="lr-base">🏭 {esc(ind)}</span>'
              f'<span class="lr-sum">{esc(_brief(judge, 90))}</span>'
              f'<span class="lr-rg">🗺️ {esc(rg)}</span>'
              f'<span class="lr-go">详情 →</span></a>')
    # 区域筛选条
    chips='<button class="rg-chip active" data-rg="全部">全部</button>'+"".join(
        f'<button class="rg-chip" data-rg="{esc(rg)}">{esc(rg)}</button>' for rg in rg_order)
    rg_filter_js="""
<style>.rg-chip{cursor:pointer;border:1px solid #c8d4e0;background:#fff;color:#33475b;border-radius:14px;padding:4px 12px;font-size:12px;margin:0 6px 6px 0}.rg-chip.active{background:#003DA5;color:#fff;border-color:#003DA5}.rg-bar{position:sticky;top:0;background:#fff;padding:8px 0;z-index:9;border-bottom:1px solid #eef2f7}.srch{width:100%;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:7px 11px;font-size:12.5px;color:#33475b;margin:2px 0 8px 0;outline:none}.srch:focus{border-color:#003DA5;box-shadow:0 0 0 3px rgba(0,61,165,.08)}.exp-all{margin:0 0 8px 0;border:1px solid #c8d4e0;background:#fff;color:#003DA5;border-radius:8px;padding:5px 12px;font-size:12px;cursor:pointer}.srch-empty{display:none;padding:10px;font-size:12px;color:#8494a4;text-align:center}</style>
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
AUTH_HASH = "3a244add8aa00cef959339a845be9cd2efb4d85a3b817126ff4506dbf0476e62"  # sha256("leon:123456")

AUTH_GUARD = """<script>
(function(){
 var ok=false;
 try{ ok=(localStorage.getItem('pd_auth')==='__AUTH_HASH__'); }catch(e){}
 if(ok) return;
 var back=location.pathname.split('/').pop()+location.search;
 location.replace('index.html?next='+encodeURIComponent(back));
})();
</script>"""

def nav_html(active):
    chips=""
    for f,k in NAVITEMS:
        cls="chip" if k==active else "chip alt"
        chips+=(f'<a class="{cls}" href="{f}">'
                f'<img class="navlogo" src="{LOGO[k]}" alt="">'
                f'<span>{SHORT[k]}</span></a>')
    chips+=('<a class="chip alt" href="#" onclick="try{localStorage.removeItem(\'pd_auth\')}catch(e){};location.href=\'index.html\';return false"'
            ' style="background:#f3f4f6;color:#6b7280">退出</a>')
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
_EXPAND_BTN = ('<button class="exp-all" type="button" onclick="var all=[].filter.call(document.querySelectorAll(\'.card[data-rg]\'),function(c){return c.style.display!==\'none\'});var any=all.some(function(c){return !c.classList.contains(\'open\')});all.forEach(function(c){c.classList.toggle(\'open\',any)});this.textContent=any?\'收起全部 ▴\':\'展开全部 ▾\';">展开全部 ▾</button>')
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
    """已停用：原来的"董秘转介→打证券部→论坛换名片→CRM warm intro"是每个 case 都一样的模板，
    用户明确要求删掉。现在的搭线路径由 AI 按这家公司的事实逐步生成（_plan_html）。"""
    return ""


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
AI_MAX = int(os.environ.get("AI_MAX", "3") or 3)
REFRESH_MAX = int(os.environ.get("REFRESH_MAX", "3") or 3)   # 每次运行滚动刷新几家老档案
# 智谱是主力通道（同一个 key，四档模型按质量排：plus 最好，flash 只写空话放最后）
AI_MODELS = ["glm-4-plus", "glm-4-air", "glm-4-flash-250414", "glm-4-flash"]
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

_TVLY_DEAD = {"v": False}     # 所有 Tavily key 都不可用后置位，本轮直接走免费源
_FREE_FAIL = {"n": 0}         # 免费源连续失败计数，>=2 就本轮不再尝试，省时间
_TVLY_KEYS = [k for k in [os.environ.get("TAVILY_KEY"), os.environ.get("TAVILY_KEY2"),
                          os.environ.get("TAVILY_KEY3")] if k]

def _ddg(q, n=4):
    """免费兜底搜索①：DuckDuckGo HTML 版，无需 key"""
    try:
        data = urllib.parse.urlencode({"q": q}).encode()
        req = urllib.request.Request("https://html.duckduckgo.com/html/", data=data,
                                     headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                                              "Content-Type": "application/x-www-form-urlencoded"})
        with urllib.request.urlopen(req, timeout=20) as r:
            html = r.read().decode("utf-8", "ignore")
    except Exception as e:
        log(f"    ddg 失败：{str(e)[:50]}")
        return []
    out = []
    for m in re.finditer(r'<a rel="nofollow" class="result__a" href="(.*?)">(.*?)</a>', html, re.S):
        u = html_lib.unescape(m.group(1)); t = re.sub(r"<[^>]+>", "", html_lib.unescape(m.group(2)))
        out.append({"url": u, "title": t.strip(), "content": ""})
        if len(out) >= n: break
    return out

def _bing(q, n=4):
    """免费兜底搜索②：Bing 网页，无需 key"""
    try:
        req = urllib.request.Request("https://www.bing.com/search?" + urllib.parse.urlencode({"q": q, "count": n}),
                                     headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
        with urllib.request.urlopen(req, timeout=20) as r:
            html = r.read().decode("utf-8", "ignore")
    except Exception as e:
        log(f"    bing 失败：{str(e)[:50]}")
        return []
    out = []
    for m in re.finditer(r'<li class="b_algo".*?<h2><a href="(.*?)".*?>(.*?)</a>.*?(?:<p[^>]*>(.*?)</p>)?', html, re.S):
        u = html_lib.unescape(m.group(1)); t = re.sub(r"<[^>]+>", "", html_lib.unescape(m.group(2)))
        c = re.sub(r"<[^>]+>", "", html_lib.unescape(m.group(3) or ""))
        out.append({"url": u, "title": t.strip(), "content": c.strip()})
        if len(out) >= n: break
    return out

def _free_search(q, n=4):
    if _FREE_FAIL["n"] >= 2:      # 免费源连续失败，本轮不再白耗时间
        return []
    for fn in (_bing, _ddg):
        r = fn(q, n)
        if r:
            _FREE_FAIL["n"] = 0
            log(f"    免费源兜底命中（{fn.__name__}）：{len(r)} 条")
            return r
    _FREE_FAIL["n"] += 1
    return []

def tavily(q, n=4):
    # 多个 Tavily key 轮换：一个额度打满自动换下一个；全挂了再走免费源
    while _TVLY_KEYS and not _TVLY_DEAD["v"]:
        key = _TVLY_KEYS[0]
        try:
            d = _http_json("https://api.tavily.com/search",
                           {"api_key": key, "query": q, "max_results": n,
                            "search_depth": "basic", "include_answer": False}, timeout=30)
            return [{"url": x.get("url",""), "title": x.get("title",""), "content": x.get("content","")}
                    for x in (d.get("results") or [])]
        except Exception as e:
            msg = str(e)[:80]
            if any(k in msg for k in ("432", "429", "401", "403", "quota", "usage")):
                _TVLY_KEYS.pop(0)
                if _TVLY_KEYS:
                    log(f"    Tavily 额度不可用（{msg}），自动切换到下一个 key")
                else:
                    log(f"    Tavily 全部 key 额度不可用（{msg}），本轮切免费源")
                    _TVLY_DEAD["v"] = True
                continue
            log(f"    tavily 失败：{msg}")
            time.sleep(2)
            break
    return _free_search(q, n)

_GLM_DEAD = {"v": False}      # GLM 额度打满后置位，本轮直接走备用模型

def _ai_providers():
    """按顺序尝试的大模型通道：GLM 优先，额度打满自动切备用（都走 OpenAI 兼容接口）"""
    ps = []
    if GLM_KEY:
        for m in AI_MODELS:
            ps.append(("glm", "https://open.bigmodel.cn/api/paas/v4/chat/completions", GLM_KEY, m))
    # 备用通道：只要配了对应 key 就自动可用，无需改代码
    if os.environ.get("DEEPSEEK_KEY"):
        ps.append(("deepseek", "https://api.deepseek.com/v1/chat/completions",
                   os.environ["DEEPSEEK_KEY"], os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")))
    if os.environ.get("SILICONFLOW_KEY"):
        ps.append(("siliconflow", "https://api.siliconflow.cn/v1/chat/completions",
                   os.environ["SILICONFLOW_KEY"], os.environ.get("SILICONFLOW_MODEL", "deepseek-ai/DeepSeek-V3")))
    if os.environ.get("DASHSCOPE_KEY"):
        ps.append(("qwen", "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
                   os.environ["DASHSCOPE_KEY"], os.environ.get("DASHSCOPE_MODEL", "qwen-plus")))
    # 免注册、免 key 的公共服务（匿名层），作为最后的兜底
    ps.append(("pollinations", "https://text.pollinations.ai/openai", "", "openai-fast"))
    if os.environ.get("OPENROUTER_KEY"):
        ps.append(("openrouter", "https://openrouter.ai/api/v1/chat/completions",
                   os.environ["OPENROUTER_KEY"], os.environ.get("OPENROUTER_MODEL", "deepseek/deepseek-chat-v3-0324:free")))
    return ps

def glm_json(system, user):
    ps = _ai_providers()
    if not ps: return None, None
    for i, (prov, url, key, model) in enumerate(ps):
        if prov == "glm" and _GLM_DEAD["v"]:
            continue
        try:
            d = _http_json(url,
                {"model": model, "messages": [{"role":"system","content":system},
                                              {"role":"user","content":user}],
                 "temperature": 0.1, "stream": False, "max_tokens": 3600},
                ({"Authorization": f"Bearer {key}"} if key else {}), timeout=120)
            msg = (d.get("choices") or [{}])[0].get("message") or {}
            txt = (msg.get("content") or msg.get("reasoning_content") or "")
            txt = re.sub(r"```json|```", "", txt).strip()
            s, e = txt.find("{"), txt.rfind("}")
            if s >= 0 and e > s:
                return json.loads(txt[s:e+1]), f"{prov}:{model}"
        except Exception as ex:
            msg = str(ex)[:80]
            log(f"    {prov}:{model} 不可用：{msg}")
            if prov == "glm" and any(k in msg for k in ("429", "1305", "额度", "402", "quota")):
                log("    GLM 额度/限流不可用，自动切备用模型通道")
                _GLM_DEAD["v"] = True
                continue
            time.sleep(2 + (i % 3) * 3)      # 轻退避，别把整轮跑成等待
    return None, None

def _load_ai_cache():
    try:
        if AI_CACHE.exists():
            d = json.loads(AI_CACHE.read_text(encoding="utf-8"))
            for k, v in d.items():
                _AI_CACHE[k] = v
                if not v.get("tier"):
                    c = v.get("chief") or ""
                    if "董事长" in c or "法定代表人" in c or "总经理" in c or "拍板" in c: v["tier"] = "拍板"
                    elif "董秘" in c or "证券事务" in c or "转介" in c: v["tier"] = "转介"
                    elif "总监" in c or "部长" in c or "主任" in c or "经理" in c: v["tier"] = "对口"
                    else: v["tier"] = "转介"
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
        # 保护：若本次只载入/新增了少数几家，先把盘上已有的并回来，避免把 50 家档案写丢
        try:
            old = json.loads(AI_CACHE.read_text(encoding="utf-8"))
            if isinstance(old, dict):
                for k, v in old.items():
                    _AI_CACHE.setdefault(k, v)
        except Exception:
            pass
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

def _tel_html(ds, escfn=None):
    """📞 电话线索块：公开可查的号码 + 30 秒开场（比冷邮件有效得多）"""
    ph = ds.get("phones") or []
    if not ph: return ""
    esc = escfn or (lambda x: x)
    rows = []
    for x in ph:
        rows.append('<div style="margin:1px 0">· <b>' + esc(x.get("num","")) + '</b>'
                    + ' <span style="color:#8494a4">（原文语境：' + esc((x.get("ctx") or "")[-28:]) + '）</span>'
                    + ' <a class="ev" style="color:#0b6bcb" href="' + esc(x.get("src","")) + '" target="_blank">🔗 出处</a></div>')
    return ('<div style="margin-top:4px;background:#fffdf5;border:1px solid #f0e0b6;border-radius:8px;padding:7px 10px">'
            '<div style="font-size:12px;font-weight:700;color:#8a6b1f">📞 电话直拨（比发邮件有效）</div>'
            + "".join(rows) +
            '<div style="font-size:11.5px;color:#5a6b7b;margin-top:3px">开场只说一句：“我在整理贵司新基地/扩产的制造系统方案，想确认该找制造还是信息中心，麻烦转一下。”——不要一上来讲产品。</div></div>')

def _cta_label(ds):
    t = ds.get("tier")
    if t == "对口": return "📧 直接约 15 分钟"
    if t == "内线": return "📧 问汇报线（内线）"
    return "📧 发转介请求（请转对口）"

def _tier(title):
    """联系人分级：不是谁官大，而是谁是这个产品的门"""
    t = title or ""
    if any(k in t for k in ("董事长","副董事长","法定代表人","法人","总经理","总裁","董事")):
        return "拍板"
    if any(k in t for k in ("董秘","董事会秘书","证券事务")):
        return "转介"
    if any(k in t for k in ("人事","人力资源","HR","招聘")):
        return "内线"
    if any(k in t for k in ("总监","部长","主任","经理","厂长","负责人","主管","工程师")):
        return "对口"
    return "对口"

def _phones(nm, hits, limit=2):
    """从抓回材料里抽真电话；必须讲的是这家公司，且有联系人/招聘/招标语境"""
    out=[]
    for h in hits:
        c = h.get("content") or ""
        if not _same_subject(nm, c): continue
        if not any(k in c for k in ("联系人","联系电话","联系方式","招聘","招标","询价","人力资源","HR")):
            continue
        for m in re.finditer(r"(0\d{2,3}[-－]?\d{7,8}|400[-－]\d{3}[-－]\d{4}|1[3-9]\d{9})", c):
            num = re.sub(r"[-－]", "-", m.group(1))
            if any(x["num"] == num for x in out): continue
            ctx = c[max(0, m.start()-45):m.start()].replace("\n", " ")[-45:]
            out.append({"num": num, "src": h["url"], "ctx": ctx})
            break
        if len(out) >= limit: break
    return out

_MAIL_NOISE = ("forbeschina.com", "sina.com", "sina.cn", "sohu.com", "163.com", "qq.com", "people.com.cn",
               "xinhuanet.com", "chinadaily.com.cn", "ce.cn", "eastmoney.com", "cnstock.com", "stcn.com",
               "yicai.com", "jiemian.com", "thepaper.cn", "36kr.com", "huxiu.com", "pedaily.cn")

def _mail_scan(nm, hits, limit=4):
    """从真实抓回材料里扫可用的公开邮箱（招聘JD的HR邮箱、招标公告联系人、官网公示等）"""
    out, seen = [], set()
    for h in hits:
        c = (h.get("content") or "")
        if not _same_subject(nm, c): continue
        for m in re.finditer(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.(?:com\.cn|com|cn|net|org)", c):
            em = m.group(0).lower()
            if em in seen: continue
            if any(x in em for x in (".png", ".jpg", "example.com", "test.com", "sentry", "@2x")): continue
            # 媒体/门户的转载联系邮箱不是这家公司的，别当线索发过去
            if any(em.endswith("@" + d) for d in _MAIL_NOISE): continue
            ctx = c[max(0, m.start()-38):m.start()].replace("\n", " ")[-38:]
            seen.add(em)
            out.append({"to": em, "kind": "原文邮箱", "note": "抓回原文里真实出现的（语境：" + ctx + "）", "src": h["url"]})
            if len(out) >= limit: break
        if len(out) >= limit: break
    return out

def _clean_ladder(x, limit=4, allowed=()):
    """搭线梯子：只留 4 步以内、字段齐全的，防止模型塞废话"""
    out = []
    if not isinstance(x, list): return out
    for it in x[:limit]:
        if not isinstance(it, dict): continue
        purpose = str(it.get("purpose") or "").split("·")[0].strip()[:14]
        to = str(it.get("to") or "").strip()[:60]
        via = str(it.get("via") or "").strip()[:60]
        say = str(it.get("say") or "").strip()[:80]
        get = str(it.get("get") or "").strip()[:60]
        if to or via:
            out.append({"purpose": purpose, "to": to, "via": via, "say": say, "get": get})
    # 擦掉未核证的人名：先认出可疑人名，再在全篇梯子里统一替换成岗位指代
    bad_nm = []
    for st in out:
        t = st["to"]
        if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", t) and t not in allowed \
                and not any(k in t for k in ("部","中心","组","科","处","室","岗","经理","总监","主任","负责人","团队","事业")):
            bad_nm.append(t)
    # 从"XX分机号/XX的电话"这类句式里再揪一遍（注意只认分机号/电话，避免误伤"或邮箱"）
    for st in out:
        for f in ("to", "via", "say", "get"):
            for m in re.finditer(r"([\u4e00-\u9fa5]{2,3})(?=的分机号|分机号|的电话|手机号|手机)", st[f]):
                tk = m.group(1)
                if tk not in allowed and tk not in bad_nm \
                        and not any(k in tk for k in ("部","中心","岗","经理","总监","主任","负责人","机","话")):
                    bad_nm.append(tk)
    for st in out:
        for tk in bad_nm:
            if tk and tk not in allowed:
                for k in ("to","via","say","get"):
                    st[k] = st[k].replace(tk, "对口负责人")
                if st["to"].strip() == "对口负责人":
                    st["to"] = "对口负责人（姓名待电话确认）"
    return out

_DEPT_WORDS = ("信息中心","信息化","数字化","智能制造","制造运营","生产制造","生产计划","工艺","设备","IT部","IT 部","信息技术","生产管理","工程部","技术中心","供应链","物流","品质","质量")
_LADDER_BAN = ("访问官网","官网查找","查找联系方式","官网","约会议","开会","进一步了解","决策者","我们想了解","罗克韦尔/PTC 的","我们可以满足","最终决策","负责人（","电话沟通")

def _facts(nm, hits):
    """从真实抓回材料里抽“硬事实”：部门名/电话/招聘岗位/招标公示/海外基地/时间点。
    模型只能基于这些事实编搭线步骤，写不出具体通道就说明这家公司材料不够。"""
    depts, phones, recruit, tender, sites, times, orgs = [], [], [], [], [], [], []
    for h in hits:
        c = (h.get("content") or "")
        if not _same_subject(nm, c): continue
        title = (h.get("title") or "")
        for w in _DEPT_WORDS:
            if w in c and w not in depts: depts.append(w)
        for m in re.finditer(r"(0\d{2,3}[-－]?\d{7,8}|400[-－]\d{3}[-－]\d{4}|1[3-9]\d{9})", c):
            num = re.sub(r"[-－]", "-", m.group(1))
            if num not in [x[0] for x in phones]:
                phones.append((num, c[max(0, m.start()-40):m.start()].replace("\n"," ")[-40:], h["url"]))
        for kw in ("招聘","岗位","任职要求","汇报"):
            for m in re.finditer(kw, c):
                seg = c[max(0,m.start()-30):m.start()+70].replace("\n"," ")
                if any(x in seg for x in ("MES","信息化","IT","数字化","制造","主管","经理","主任")) and seg not in recruit:
                    recruit.append(seg[:90])
        for kw in ("招标","中标","公示","联系人","询价","采购"):
            for m in re.finditer(kw, c):
                seg = c[max(0,m.start()-30):m.start()+80].replace("\n"," ")
                if seg not in tender: tender.append(seg[:90])
        for kw in ("墨西哥","捷克","泰国","越南","印度","匈牙利","波兰","马来西亚","印尼","巴西","德国","美国","新基地","新工厂","扩产","扩建","投产"):
            if kw in c and kw not in sites: sites.append(kw)
        for m in re.finditer(r"20\d{{2}}\s*年\s*\d{{1,2}}\s*月|20\d{{2}}\s*年底|20\d{{2}}\s*年上半年|投产", c):
            seg = c[max(0,m.start()-24):m.start()+40].replace("\n"," ")
            if seg not in times: times.append(seg[:70])
    addr = []
    for h in hits:
        c = (h.get("content") or "")
        if not _same_subject(nm, c): continue
        for m in re.finditer(r"(总部|注册地址|办公地址|厂址|地址)[：: ]?\s*([^。；\n]{6,40})", c):
            seg = (m.group(1) + "：" + m.group(2)).strip()
            if seg not in addr: addr.append(seg[:60])
        for m in re.finditer(r"(省|市)[^。；\n]{0,10}(路|街|园区|工业园|开发区)[^。；\n]{0,12}", c):
            seg = m.group(0).strip()
            if seg not in addr: addr.append(seg[:40])
    f = []
    if addr: f.append("⓪ 总部/地址原文：" + " ｜ ".join(addr[:4]))
    if sites: f.append("① 海外/新基地词：" + "、".join(sites[:8]))
    if times: f.append("② 时间点原文：" + " ｜ ".join(times[:3]))
    if depts: f.append("③ 材料中出现的部门/职能词：" + "、".join(depts[:10]))
    if phones: f.append("④ 公开电话（含语境）：" + " ｜ ".join(n + "（" + ctx + "）" for n, ctx, _ in phones[:3]))
    if recruit: f.append("⑤ 招聘/岗位原文：" + " ｜ ".join(recruit[:3]))
    if tender: f.append("⑥ 招标/公示/联系人原文：" + " ｜ ".join(tender[:3]))
    return ("\n".join(f))[:1600]

def _ladder_bad(steps):
    bad = []
    for st in steps or []:
        t = ((st.get("to") or "") + " " + (st.get("via") or "") + " " + (st.get("say") or "") + " " + (st.get("get") or ""))
        for b in _LADDER_BAN:
            if b in t: bad.append(b)
        if not (st.get("via") or "").strip(): bad.append("空通道")
    return bad

_DUP_TXT = {}
_DUP_FIELDS = ("window", "dept_ev", "hook", "risk", "next", "hq", "lever")

def _norm_txt(x):
    return re.sub(r"[\s，。；、！？：·（）()\"“”'']", "", (x or "").lower())

def _init_dup():
    """同一句话出现在几家公司：>=2 就是模板话术，渲染时不再显示"""
    if _DUP_TXT: return
    seen = {}
    # 同一家公司在 PRODUCT_DOSSIERS / _AI_CACHE 里可能有两份拷贝，按"公司名+信号"去重，避免一句话被算两次
    seen_sig = set()
    pool = []
    for d in list(PRODUCT_DOSSIERS.values()) + list(_INQ.values()) + list(_AI_CACHE.values()):
        if not isinstance(d, dict): continue
        sig = (d.get("name") or "") + "|" + (d.get("signal") or "")[:30]
        if sig in seen_sig: continue
        seen_sig.add(sig); pool.append(d)
    for d in pool:
        for f in _DUP_FIELDS:
            v = _norm_txt(d.get(f))
            if len(v) >= 6: seen[v] = seen.get(v, 0) + 1
        for st in (d.get("ladder") or []):
            for f in ("via", "say"):
                v = _norm_txt(st.get(f))
                if len(v) >= 8: seen[v] = seen.get(v, 0) + 1
    _DUP_TXT.update(seen)

def _dup(x):
    return _DUP_TXT.get(_norm_txt(x), 0) >= 2

def _mailbox_html(ds, escfn=None, mailto=True):
    """📮 邮箱区：实名推断 / 原文真实 / 部门推测，三类分开标，别让人误发"""
    esc = escfn or (lambda x: x)
    rows = []
    for x in (ds.get("emails") or [])[:5]:
        if not x.get("to"): continue
        if any((x["to"] or "").lower().endswith("@" + d) for d in _MAIL_NOISE): continue
        k = x.get("kind") or "实名推断"
        col = {"实名推断": "#0b6bcb", "原文邮箱": "#137333"}.get(k, "#8494a4")
        link = ('<a class="ev" style="color:#0b6bcb" href="' + esc(x.get("src","")) + '" target="_blank">🔗</a>') if x.get("src") else ""
        rows.append('<div style="margin:1px 0;font-size:12px"><code style="background:#eef4fb;padding:1px 5px;border-radius:4px">'
                    + esc(x["to"]) + '</code> <span style="font-size:11px;color:' + col + '">' + k + '</span> '
                    + '<span style="font-size:11px;color:#8494a4">' + esc(x.get("note","")) + '</span> ' + link + '</div>')
    if not rows: return ""
    return ('<div style="margin-top:4px;background:#f6faff;border:1px solid #dcebfb;border-radius:8px;padding:7px 10px">'
            '<div style="font-size:12px;font-weight:700;color:#0b6bcb">📮 可用邮箱（从真到猜，按顺序试）</div>'
            + "".join(rows) + '</div>')

def _plan_html(ds, escfn=None):
    """怎么找人：窗口 → 该找哪个部门 → 搭线梯子 → 开场白 → 反证 → 今天做什么"""
    esc = escfn or (lambda x: x)
    rows = []
    _init_dup()
    for st in (ds.get("ladder") or []):
        if not st.get("to"): continue
        if _dup(st.get("via")) and _dup(st.get("say")):
            continue
        i = len(rows) + 1
        rows.append(
            '<div style="margin:5px 0;padding:7px 9px;background:#fbfcfe;border-left:3px solid #7a5af8;border-radius:0 6px 6px 0">'
            '<div style="font-size:12px;color:#243b53"><b>第' + str(i) + '步</b>'
            + ((' · ' + esc(st["purpose"])) if st.get("purpose") else '') + ' · ' + esc(st.get("to","")) + '</div>'
            '<div style="font-size:11.5px;color:#5a6b7b;line-height:1.7">通道：' + esc(st.get("via","")) + '</div>'
            + (('<div style="font-size:11.5px;color:#8a4b00;line-height:1.7">开口：' + esc(st.get("say","")) + '</div>') if st.get("say") else '')
            + (('<div style="font-size:11.5px;color:#137333;line-height:1.7">目标产出：' + esc(st.get("get","")) + '</div>') if st.get("get") else '')
            + '</div>')
    _init_dup()
    parts = []
    if ds.get("hq"):
        parts.append('<div style="font-size:12.5px;color:#243b53;margin-top:2px"><b>\U0001f4cd 总部/决策链位置：</b>' + esc(ds["hq"])
                     + ((' <span style="color:#8494a4">（依据：' + esc(ds.get("hq_ev","")) + '）</span>') if ds.get("hq_ev") else '')
                     + '<div style="font-size:11.5px;color:#5a6b7b;margin-top:1px">决定权通常在总部；若关键人在海外基地/新厂，说明这条线要先经总部再落到基地。</div></div>')
    if ds.get("window") and not _dup(ds.get("window")):
        parts.append('<div style="font-size:12.5px;color:#243b53"><b>\u23f1 窗口：</b>' + esc(ds["window"]) + '</div>')
    if ds.get("dept") and not (ds.get("dept_ev") and _dup(ds.get("dept_ev"))):
        parts.append('<div style="font-size:12.5px;color:#243b53;margin-top:3px"><b>\U0001f3af 该找的部门：</b>' + esc(ds["dept"])
                     + ((' <span style="color:#8494a4">（依据：' + esc(ds["dept_ev"]) + '）</span>') if ds.get("dept_ev") else '') + '</div>')
    if ds.get("lever") and not _dup(ds.get("lever")):
        parts.append('<div style="margin-top:7px;font-size:12.5px;color:#243b53;background:#fbf7ff;border-left:3px solid #7a5af8;border-radius:0 6px 6px 0;padding:7px 10px;line-height:1.75"><b>\U0001f9ed 杠杆点（为什么现在非动不可）：</b>' + esc(ds["lever"]) + '</div>')
    _dec = [x for x in (ds.get("deciders") or []) if x and not _dup(x)]
    if _dec:
        parts.append('<div style="margin-top:6px;font-size:12.5px;color:#243b53"><b>\U0001f5fa 决策链谁说了算：</b></div>'
                     + ''.join('<div style="font-size:12px;color:#33475b;line-height:1.7;padding-left:6px">· ' + esc(x) + '</div>' for x in _dec))
    _vf = [x for x in (ds.get("verify") or []) if (x.get("q") or "").strip()]
    if _vf:
        _vh = ''.join(
            '<div style="font-size:12px;color:#33475b;line-height:1.7;padding:5px 0;border-bottom:1px dashed #e8eef5">'
            '<b>' + str(i) + '. ' + esc(x.get("q","")) + '</b>'
            + (('<div style="color:#5a6b7b">怎么查：' + esc(x.get("how","")) + '</div>') if x.get("how") else '')
            + (('<div style="color:#137333">✅ 算真机会：' + esc(x.get("pass","")) + '</div>') if x.get("pass") else '')
            + (('<div style="color:#a04040">❌ 可以直接放：' + esc(x.get("kill","")) + '</div>') if x.get("kill") else '')
            + '</div>' for i, x in enumerate(_vf, 1))
        parts.append('<div style="margin-top:7px"><div style="font-size:12.5px;color:#003DA5;font-weight:700">✅ 先验证它是不是真需求（问完这三句再谈方案）</div>' + _vh + '</div>')
    if rows:
        parts.append('<div style="font-size:12.5px;color:#003DA5;font-weight:700;margin:8px 0 2px">\U0001fa9c 怎么搭上线（按顺序做）</div>' + ''.join(rows))
    if ds.get("hook") and not _dup(ds.get("hook")):
        parts.append('<div style="margin-top:6px;font-size:12.5px;color:#8a4b00;background:#fffdf5;border:1px solid #f0e0b6;border-radius:6px;padding:6px 9px"><b>\U0001f4ac 见到对口人第一句：</b>' + esc(ds["hook"]) + '</div>')
    if ds.get("next") and not _dup(ds.get("next")):
        parts.append('<div style="margin-top:5px;font-size:12.5px;color:#137333"><b>\u25b6 今天先做：</b>' + esc(ds["next"]) + '</div>')
    _act = [x for x in (ds.get("act") or []) if x and not _dup(x)]
    if _act:
        parts.append('<div style="margin-top:6px;font-size:12.5px;color:#243b53"><b>\u25b6 今天就做这几件：</b></div>'
                     + ''.join('<div style="font-size:12px;color:#33475b;line-height:1.7;padding-left:6px">' + str(i) + '. ' + esc(x) + '</div>' for i, x in enumerate(_act, 1)))
    if ds.get("risk") and not _dup(ds.get("risk")):
        parts.append('<div style="margin-top:4px;font-size:11.5px;color:#a04040"><b>\u26a0 反证：</b>' + esc(ds["risk"]) + '</div>')
    if not parts: return ""
    return '<div style="margin-top:6px;border-top:1px dashed #dbe4ef;padding-top:6px">' + ''.join(parts) + '</div>'

def _thin(d):
    """检查 AI 输出是否太薄：返回不合格原因，合格返回空字符串"""
    if not isinstance(d, dict): return "非 JSON"
    msgs = []
    lv = (d.get("lever") or "").strip()
    if len(lv) < 80: msgs.append("杠杆点不足80字")
    vf = [x for x in (d.get("verify") or []) if isinstance(x, dict) and (x.get("q") or "").strip()]
    if len(vf) < 3: msgs.append("验证问题不足3条")
    if any(not (x.get("pass") or "").strip() or not (x.get("kill") or "").strip() for x in vf):
        msgs.append("验证缺 pass/kill")
    act = [x for x in (d.get("act") or []) if str(x).strip()]
    if len(act) < 2: msgs.append("今日动作不足2条")
    dec = [x for x in (d.get("deciders") or []) if str(x).strip()]
    if len(dec) < 2: msgs.append("决策链不足2条")
    if not (d.get("hq") or "").strip(): msgs.append("缺总部")
    return "、".join(msgs)

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
          f'"{nm}" 中标 智能制造 项目 负责人 联系人',       # 招投标材料常留项目联系人真名
          f'"{nm}" 招聘 MES 信息化 主管 工程师 汇报 部门',   # 在招岗=正在建队，JD 常露部门名/HR电话
          f'"{nm}" 智能制造 试点示范 项目 负责人 公示']      # 政府/协会公示里是真实项目负责人
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
    facts = _facts(nm, hits)
    playbook = (f"产品：{p['name']}｜定位：{p['pos']}\n目标客户：{p['who']}\n"
                f"验证真需求的三问：{' / '.join(p['three'])}")
    user = f"""公司：{nm}（{code}），本次触发的公告标题：{finding.get('title','')}

{playbook}

材料硬事实（由程序从原文抽取，ladder 与 next 只能基于这些）：
{facts or "（这家公司没有抓到可用的部门名/电话/招聘/招标原文）"}

【怎么找人 · 硬规则】ladder 是这家公司专属的搭线路径，必须像老销售真的动手那样一步步写：
可用通道只有这几类，选最贴合硬事实的，并写清具体号码/岗位/人名：
 a 年报/公告里的证券事务部电话（只为要"转接"，不问业务）
 b 招聘 JD 里的 HR 联系方式或投递邮箱（借"应聘/了解岗位汇报线"问出部门）
 c 招标/中标/公示公告里的项目联系人（名正言顺问"这个项目")
 d 总机转分机（写清拨哪个号码、报什么由头才不会被退回前台）
 e 集团总部/母公司同类部门（子公司没入口时从上面压下来）
 f 行业协会/展会/政府园区对接人（写清具体活动或机构）
 g 已核证的具名联系人（只能是我给你的那位，不能新编）

四步结构（每步都要能今天做）：
 第1步 进门：拿到"该找哪个部门、谁管"，通道用 a/b/c/d
 第2步 拿到人：拿到姓氏+分机/邮箱，且说明用什么由头让他愿意接
 第3步 建立可信：给一个他能验证的同类事实或行业口径（不许报我们自己客户名字）
 第4步 约见：给一个 15 分钟的具体议题，不是"介绍一下产品"

say 的三条禁令：不提"我们想了解贵司需求"；不提我们产品名；不空泛寒暄。要说他此刻正头疼的具体事（爬坡、多厂标准不一、海外客户审核、追溯、换线效率），一句话。
to 不许写"负责人""决策者"这类泛称；写具体部门或具体岗位名（用材料里的叫法，材料没有就写"（推断）"）。
材料不够就少写，next 里说明先确认什么。宁可少，不许编，不许套模板。
6) ladder 里严禁出现任何人名（人名只能出现在 chief_name / insider_name 字段，且必须已在材料核证）。要指人时写"生产管理部负责人""信息中心主任"这类岗位，不许写"张三"。
检索材料（含编号与URL，只能引用这些URL）：
{ctx}

请输出 JSON，字段：
{{
 "demand":"是|否|待验证",           // 该信号是否会引出上述产品的选型需求（你自己的判断）
 "demand_reason":"60字内，依据哪条材料做出的判断",
 "signal":"100-180字，把关键事实串成一段：哪条公告/材料（时间）、在哪个基地或项目、产能/金额/产品、投产或推进节点、原文原话。只写材料里有的，不许概括成一句话。",
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
 "email_guess":"按该域名与中文姓名拼音推断的最可能邮箱；查不到填空字符串",

 "window":"这家公司材料里的具体时间节点（哪年哪月、哪条公告说了什么节点）；查不到就填空。禁止行业通则（如“投产前6个月定版”）。",
 "dept":"该找的具体部门（用这家公司自己的叫法，如 集团信息中心/制造运营部/工艺设备部）：必须是材料里出现过的部门名，没出现就写行业通称并标注（推断）",
 "dept_ev":"为什么是这个部门：依据哪条材料（招聘JD的部门名/公告署名/组织架构新闻），40字内",
 "ladder":[   // 必须正好4步，第4步是"给15分钟具体议题"，不是"见面聊产品"
   {{"purpose":"进门 · 拿到部门和分机","to":"找谁","via":"拨哪个号码、走哪条通道（招聘HR/招标联系人/证券事务部/总机+分机/集团总部）","say":"第一句（25字内，只说他的事）","get":"拿到什么（部门名称/姓氏/分机号/邮箱）"}},
   {{"purpose":"拿人 · 拿到姓名与直拨","to":"...","via":"...","say":"...","get":"..."}},
   {{"purpose":"建立可信 · 给一条他能验证的事实","to":"...","via":"...","say":"...","get":"..."}},
   {{"purpose":"约见 · 15分钟两页纸","to":"...","via":"...","say":"...","get":"..."}}
 ],
 "hook":"针对这家公司具体事实的一句开场（25字内，必须点到它的基地/产线/项目名或公告事实；禁止通用句、产品名、“我们想了解贵司需求”）",
 "risk":"这家的具体不利证据（公告只提意向未定金额/已延期/竞品已中标等）；写不出就填空。禁止“若计划推迟、预算削减”这类到处都能用的废话。",
 "next":"今天就能做的一个动作（20字内，含对象和动作，如 打0571总机转信息中心问MES归谁管）"
 "hq":"总部所在地：省+市+区/园区或街道，能细到哪一层就写到哪一层（依据注册地址/年报/官网联系地址）；只能写材料里出现的，查不到填空",
 "hq_ev":"总部地址依据哪条材料（公告/官网页面），20字内",
 "deciders":[
   "角色（这家公司自己的部门/岗位叫法）｜在哪（总部/哪个基地/哪个厂）｜他真正关心什么（产能爬坡、多厂口径统一、追溯审核、人工成本）"
 ],
 "lever":"杠杆点，2-3句说透：为什么这家现在非动不可、卡的是什么机制、不动的代价是什么；必须落到它的基地/产线/客户/审核要求上",
 "verify":[
   {{"q":"要验证的具体问题（一句）","how":"怎么查/问谁（哪条公开信息、打哪个电话问哪句话）","pass":"什么回答说明是真机会","kill":"什么回答可以直接放掉"}}
 ],
 "act":[
   "动作（谁—走哪条通道—第一句问什么—拿到什么）"
 ]
}}"""
    d, model = glm_json(_AI_SYS, user)
    # —— 内容太薄（验证没到3条/动作不到2条/杠杆点一句话）就重写一次 ——
    if d and _thin(d):
        log("    输出偏薄（" + _thin(d) + "），重写一次")
        d2, m2 = glm_json(_AI_SYS, user + "\n\n【上一次输出不合格】" + _thin(d) +
                          "。请补全：lever 至少 80 字且点名它的基地/产线/客户审核；verify 正好 3 条"
                          "（q/how/pass/kill 全填）；act 2-3 条，每条写清 谁—哪条通道—第一句问什么—拿到什么；"
                          "deciders 2-4 条并写清所在地。禁止写通用句。")
        if d2 and not _thin(d2):
            d, model = d2, m2
    # —— 搭线通道空转就重试一次，逼模型落到真实通道上 ——
    try:
        bad = _ladder_bad(_clean_ladder((d or {}).get("ladder")))
        if bad and (d or {}).get("demand"):
            log("    搭线通道空转（" + "、".join(sorted(set(bad)))[:40] + "），重写一次")
            fix = user + "\n\n【上一次输出不合格】你写的 ladder 里出现了空话：" + "、".join(sorted(set(bad))) + \
                  "。请重写 ladder：每一步必须落到材料硬事实里的具体电话/岗位/部门/公告联系人上，写不出就减少步数并在 next 说明要先确认什么。"
            d2, m2 = glm_json(_AI_SYS, fix)
            if d2 and not _ladder_bad(_clean_ladder(d2.get("ladder"))):
                d, model = d2, m2
    except Exception as ex:
        log("    搭线校验跳过：" + str(ex)[:50])
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
        _t = _tier(_title)
        if _t == "拍板":
            chief = f"🤖 AI 查证：{_title} {name}（拍板人｜他不会读冷邮件，只做顶层转发或请秘书转对口）"
        elif _t == "内线":
            chief = f"🤖 AI 查证：{_title} {name}（内线｜能问出这个岗的汇报线，比猜人准）"
        else:
            chief = f"🤖 AI 查证：{_title} {name}（对口人｜可直接谈方案，这是真正的门）"
    elif insider and any(k in insider["title"] for k in ("联系人","项目负责","主任","经理","部长","总监")):
        chief = (f"未具名（决策人）：但从招投标材料挖到项目线具名人 —— "
                 f"{insider['title']} {insider['name']}。可直接按招标联系方式触达，比董秘更近一层")
    else:
        chief = "未具名：AI 已跑 7 路公开源（公告/官网/新闻/招聘/高管/招投标/公示），未见可靠具名负责人；走总机转接 + 董秘转介"

    region = norm_region(d.get("region"))
    # 注意：不做"正文里扫省份"的兜底——会抓到公告里的子公司/项目所在地，
    # 把总部区域judge错（曾把云南的贵研铂业判成山东）。认不出就留待确认。
    ds = {"name": nm,
          "role": (d.get("role") or PRODUCT_WHO_ROLE.get(p["key"], ("信息化/数字化负责人",""))[0])[:60],
          "why": (d.get("why") or "")[:100],
          "chief": chief,
          "demand": d.get("demand") or "待验证",
          "demand_reason": (d.get("demand_reason") or "")[:100],
          "signal": (d.get("signal") or "")[:240],
          "grade": d.get("grade") if d.get("grade") in ("高","中","低") else "中",
          "region": region if region else "",
          "auto": True, "model": model,
          "src": src if name else "",
          "referral": referral,
          "chief_name": (referral or name),
          "insider": insider,
          "tier": _tier(_title if name else ""),
          "phones": _phones(nm, hits),
          "window": (d.get("window") or "")[:80],
          "dept": (d.get("dept") or "")[:40],
          "dept_ev": (d.get("dept_ev") or "")[:80],
          "ladder": _clean_ladder(d.get("ladder"), allowed=(name, iname) if iname else (name,)),
          "hook": (d.get("hook") or "")[:60],
          "risk": (d.get("risk") or "")[:80],
          "next": (d.get("next") or "")[:40],
          "hq": (d.get("hq") or "")[:60],
          "hq_ev": (d.get("hq_ev") or "")[:40],
          "deciders": [str(x)[:110] for x in (d.get("deciders") or []) if str(x).strip()][:4],
          "lever": (d.get("lever") or "")[:260],
          "verify": [{"q": str(x.get("q") or "")[:70], "how": str(x.get("how") or "")[:90],
                      "pass": str(x.get("pass") or "")[:60], "kill": str(x.get("kill") or "")[:60]}
                     for x in (d.get("verify") or []) if isinstance(x, dict) and (x.get("q") or "").strip()][:3],
          "act": [str(x)[:100] for x in (d.get("act") or []) if str(x).strip()][:3],
          "pk": p["key"],
          "researched": datetime.date.today().isoformat()}
    mails = (email and [dict(email, kind="实名推断")]) or []
    if not mails: mails = []
    if email: mails[0]["kind"] = "实名推断"
    scanned = _mail_scan(nm, hits)
    have = {x["to"] for x in mails}
    for x in scanned:
        if x["to"] not in have:
            mails.append(x); have.add(x["to"])
    if domain and not email and referral:
        py = _py_full(referral)
        t = (py + "@" + domain) if py else ""
        if py and re.match(r"^[a-z]+@", t) and t not in have:
            mails.insert(0, {"to": t, "kind": "实名推断（转介人）",
                             "note": "转介人姓名全拼由程序生成，转介用：只求转交对口部门"})
            have.add(t)
    if domain:
        for pre in ("hr", "info", "mes"):
            g = pre + "@" + domain
            if g not in have and len(mails) < 4:
                mails.append({"to": g, "kind": "推测待验证", "note": "按 " + pre + "@域名 惯例推断，发一封试探是否退信，别当主渠道"})
                have.add(g)
    ds["emails"] = mails[:5]
    if email:
        if referral:
            email["note"] = "转介用邮箱：只求转交对口部门，不要谈方案细节"
        elif ds.get("tier") == "拍板":
            email["note"] = "顶层邮箱：正文压到5行、只求转交制造/信息化负责人，附一页纸PDF"
        elif ds.get("tier") == "对口":
            email["note"] = "对口邮箱：可直接约15分钟，附同行业案例一页纸"
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

INQ_CACHE = ROOT / "data" / "inquiries.json"
_INQ = {}

def _load_inq():
    try:
        if INQ_CACHE.exists():
            d = json.loads(INQ_CACHE.read_text(encoding="utf-8"))
            for k, v in d.items():
                # 老记录里的人名同样过一遍防编造：梯子里只允许已核证的具名
                if v.get("ladder"):
                    v["ladder"] = _clean_ladder(v["ladder"], allowed=(v.get("chief_name") or "",))
                _INQ[k] = v
            log(f"即查结果载入 {len(d)} 家")
    except Exception as e:
        log(f"即查结果读取失败：{str(e)[:60]}")

def _save_inq():
    if not _INQ: return
    try:
        INQ_CACHE.parent.mkdir(exist_ok=True)
        INQ_CACHE.write_text(json.dumps(_INQ, ensure_ascii=False, indent=1), encoding="utf-8")
        log(f"即查结果已写盘（{len(_INQ)} 家）")
    except Exception as e:
        log(f"即查结果写盘失败：{str(e)[:60]}")

def queue_pass():
    """即查模式接 AI：页面按钮在 queue/ 下建一个任务文件，CI 跑时自动查完入档并删除该文件"""
    _load_ai_cache(); _load_inq()
    if not (GLM_KEY and TVLY_KEY): return
    qdir = ROOT / "queue"
    if not qdir.exists(): return
    files = sorted(qdir.glob("*.txt"))
    if not files: return
    log(f"即查队列：{len(files)} 个待查任务")
    done = 0
    for fp in files:
        if done >= max(1, AI_MAX): break
        try:
            raw = fp.read_text(encoding="utf-8").strip()
        except Exception:
            raw = ""
        if not raw:
            fp.unlink(missing_ok=True); continue
        if "|" not in raw.replace("｜", "|"):
            log(f"  队列：{fp.name} 非任务格式，跳过不删")
            continue
        parts = [x.strip() for x in raw.replace("｜","|").split("|")]
        nm = parts[0]
        pk = (parts[1].upper() if len(parts) > 1 and parts[1] else "PLEX")
        p = next((x for x in PRODUCTS if x["key"] == pk), None) or PLEX_PSEUDO
        code = ""
        m = re.search(r"\d{6}", nm)
        if m: code = m.group(0)
        if not code: code = "Q" + hashlib.md5(nm.encode("utf-8")).hexdigest()[:8]
        if code in PRODUCT_DOSSIERS or code in _INQ:
            log(f"  队列：{nm} 已查过，跳过")
            fp.unlink(missing_ok=True); continue
        log(f"  队列深查：{nm} [{p['key']}]")
        ds = ai_dossier(code, nm, p, {"title": "（用户即查模式提交）"})
        if ds:
            ds["product_key"] = p["key"]
            ds["inquiry"] = True
            _INQ[code] = ds
            done += 1
            log(f"    ✓ 即查完成（{ds['demand']}/{ds['grade']}）→ inquiry.html")
            fp.unlink(missing_ok=True)
            _save_inq()
        else:
            log(f"    队列：{nm} 未查到，保留文件待下次重试")

def _sec_name(code):
    """按股票代码反查公司简称（老档案 name 为空时用），查不到返回空串"""
    for col in ("szse", "sse"):
        try:
            data = fetch_cninfo(code, column=col, days=365, page_size=5)
            for a in (data.get("announcements") or []):
                if (a.get("secCode") or "").zfill(6) == str(code).zfill(6) and a.get("secName"):
                    return a["secName"].strip()
        except Exception:
            continue
    return ""

def _infer_pk(ds):
    """老档案没存产品 key 时按信号/岗位/理由推断，避免把 Windchill/FTPC 的线索当 PLEX 重查"""
    t = ((ds.get("signal") or "") + " " + (ds.get("role") or "") + " " + (ds.get("why") or "") + " "
         + (ds.get("demand_reason") or "") + " " + (ds.get("lever") or ""))
    if any(k in t for k in ("PLM", "Windchill", "BOM", "研发数据", "图纸", "EBOM")): return "WIND"
    if any(k in t for k in ("ThingWorx", "物联网", "设备联网", "IIoT", "数据采集", "远程监控")): return "TWX"
    if any(k in t for k in ("DataMosaix", "数据湖", "数据治理", "数据平台")): return "DMX"
    if any(k in t for k in ("FTPC", "MES", "MOM", "制造执行", "车间", "排产", "追溯")): return "FTPC"
    return "PLEX"

def _refresh_pass(AI_MAX, REFRESH_MAX, PRODUCT_DOSSIERS, _AI_CACHE, HQ_REGION):
    """滚动刷新：高优先档案按"最久没更新"轮流重查；外部信息天天在变，档案不能烂在缓存里"""
    done = 0
    # 滚动刷新保底：不管 AI_MAX 设多少，每轮至少刷 5 家高优先（一天两轮≈10 家，29 家高优先约 3 天一轮）
    _RF = max(5, min(REFRESH_MAX, AI_MAX))
    try:
        _PBK = {p0["key"]: p0 for p0 in PRODUCTS}
        _PBK.setdefault("PLEX", PLEX_PSEUDO)
        _stale = []
        for c0, ds0 in PRODUCT_DOSSIERS.items():
            if (ds0.get("grade") or "") != "高": continue
            r0 = (ds0.get("researched") or "")[:10]
            try:
                gap = (datetime.date.today() - datetime.date.fromisoformat(r0)).days
            except Exception:
                gap = 99
            if gap < (0 if os.environ.get("BACKFILL") else 1): continue
            _stale.append((gap, c0, ds0))
        _stale.sort(key=lambda x: -x[0])
        _n = 0
        for gap, c0, ds0 in _stale:
            if _n >= _RF: break
            p0 = _PBK.get(ds0.get("pk") or _infer_pk(ds0))
            if not p0: continue
            _nm0 = (ds0.get("name") or "").strip() or _sec_name(c0)
            if not _nm0:
                log(f"  跳过 {c0}：按代码查不到公司名")
                continue
            log(f"  滚动刷新（{gap}天未更新）：{_nm0} [{p0['key']}]")
            ds1 = ai_dossier(c0, _nm0, p0, {"title": ds0.get("signal") or ""})
            _n += 1; done += 1
            if ds1:
                PRODUCT_DOSSIERS[c0] = ds1
                _AI_CACHE[c0] = ds1
                if ds1.get("region"): HQ_REGION[c0] = ds1["region"]
                log(f"    ✓ 已刷新（{ds1.get('grade')}）")
            else:
                log("    刷新失败，保留旧档案")
    except Exception as ex:
        log("滚动刷新跳过：" + str(ex)[:60])
    return done

def ai_research_pass():
    """新线索自动查档；已入档的跳过，不重复烧额度"""
    _load_ai_cache()
    if not (GLM_KEY and TVLY_KEY):
        log("未配置 GLM_KEY / TAVILY_KEY，跳过 AI 自动深查（页面照常生成）")
        return
    done = _refresh_pass(AI_MAX, REFRESH_MAX, PRODUCT_DOSSIERS, _AI_CACHE, HQ_REGION)
    if done: _save_ai_cache()

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

LEADS = {}      # 详情页数据：code -> {n,s,g,loc,rg,sum,b,p}
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
                 + (('<div style="font-size:12px;color:#0b6bcb;line-height:1.6">👤 ' + esc(ai.get("chief","")) + '</div>')
                  if not (ai.get("deciders") and (ai.get("chief") or "").startswith("未具名")) else '')
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
                   f'href="mailto:{esc(_e3["to"])}?subject={_sub3}&body={_body3}">{esc(_cta_label(ai))} · {esc(_e3["to"])}</a></div>'
                   + _mailbox_html(ai, esc))
    # 关键：没有推断邮箱时也要出"总部/决策链/杠杆点/验证三问/今日动作"，否则卡片还是只有法人+邮箱
    if ai:
        mail_html += _plan_html(ai, esc)
        mail_html += _tel_html(ai, esc)
    if ai and (ai.get("lever") or ai.get("verify") or ai.get("act")):
        legacy_rows = ""      # 新四段已覆盖，老模板行整段不显示，避免每个 case 都重复同一套话
    else:
        legacy_rows = (f'<div class="row"><b>决策人：</b>{esc(chief)}</div>'
                       f'<div class="row"><b>杠杆点：</b>{esc(lev)}</div>'
                       f'<div class="row"><b>怎么开口：</b>{esc(hook)}</div>'
                       f'<div style="margin-top:8px"><b style="font-size:12.5px;color:#b26a00">📋 你要做（按顺序）：</b>{steps_html}</div>')
    rg=region_of(code).split("·")[0]
    _sum = _brief((ai or {}).get("signal") or f.get("title") or "", 90)
    rg=region_of(code).split("·")[0]
    _body = (f'<div class="row"><b>为什么是机会：</b>{esc(judge)}</div>'
             f'<div class="row">{badge}</div>'
             f'{ai_html}'
             f'{legacy_rows}'
             f'{mail_html}'
             f'{contact_links(code,nm)}'
             f'<details><summary>证据</summary><div class="ev"><a href="{esc(f["url"])}" target="_blank">公告原文 ↗</a></div></details>'
             f'<div style="margin-top:8px"><button class="btn fu-btn" data-fu="{esc(nm)}">📌 标记今天跟进</button></div>')
    LEADS[code] = {"n": nm, "s": f["sec"] or "", "g": g, "loc": loc, "rg": rg,
                   "sum": _sum, "b": _body, "p": f.get("url") or ""}
    return (f'<a class="lead-row" data-rg="{esc(rg)}" href="lead.html?c={esc(code)}">'
            f'<span class="lr-g">{g}</span>'
            f'<span class="lr-nm">{esc(nm)}</span>'
            f'<span class="lr-code">{esc(f["sec"] or "")}</span>'
            f'<span class="lr-base">📍 {esc(loc)}</span>'
            f'<span class="lr-sum">{esc(_sum)}</span>'
            f'<span class="lr-rg">🗺️ {esc(rg)}</span>'
            f'<span class="lr-go">详情 →</span></a>')
# 先跑 AI 深查，再渲染卡片——否则 PLEX 卡片永远拿不到当天的档案
def _ci_push_data():
    """CI 里把 data/（档案、队列、即查结果）回写仓库。workflow 的 commit 步骤不 add data/，
    若不自己推，每天的深查结果下一次运行就丢，档案永远停在旧版本。"""
    if os.environ.get("GITHUB_ACTIONS") != "true": return
    try:
        import subprocess
        if not Path("data").exists(): return
        subprocess.run(["git","add","data"], check=False, capture_output=True)
        r = subprocess.run(["git","commit","-m","data: 档案/队列自动回写"], capture_output=True, text=True)
        if r.returncode != 0:
            log("data 无变化，跳过回写")
            return
        pr = subprocess.run(["git","push"], capture_output=True, text=True)
        log("data 回写仓库：" + ("成功" if pr.returncode == 0 else ("失败 " + pr.stderr[:80])))
    except Exception as ex:
        log("data 回写失败：" + str(ex)[:60])

def _probe_sources():
    """CI 里探测一次免费检索源可达性，日志留证；不可达本轮直接不试，省时间"""
    if os.environ.get("GITHUB_ACTIONS") != "true": return
    for nm, fn in (("bing", _bing), ("ddg", _ddg)):
        try:
            r = fn("测试", 1)
            log(f"  CI 免费检索源 {nm}：{'可用 ' + str(len(r)) + ' 条' if r else '不可用'}")
        except Exception as ex:
            log(f"  CI 免费检索源 {nm}：异常 {str(ex)[:40]}")

_probe_sources()
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
<style>.rg-chip{cursor:pointer;border:1px solid #c8d4e0;background:#fff;color:#33475b;border-radius:14px;padding:4px 12px;font-size:12px;margin:0 6px 6px 0}.rg-chip.active{background:#003DA5;color:#fff;border-color:#003DA5}.rg-bar{position:sticky;top:0;background:#fff;padding:8px 0;z-index:9;border-bottom:1px solid #eef2f7}.srch{width:100%;box-sizing:border-box;border:1px solid #c8d4e0;border-radius:8px;padding:7px 11px;font-size:12.5px;color:#33475b;margin:2px 0 8px 0;outline:none}.srch:focus{border-color:#003DA5;box-shadow:0 0 0 3px rgba(0,61,165,.08)}.exp-all{margin:0 0 8px 0;border:1px solid #c8d4e0;background:#fff;color:#003DA5;border-radius:8px;padding:5px 12px;font-size:12px;cursor:pointer}.srch-empty{display:none;padding:10px;font-size:12px;color:#8494a4;text-align:center}</style>
<script>
(function(){var bar=document.getElementById('rgbar');if(!bar)return;
var inp=bar.querySelector('.srch'), rg='全部';
function apply(){
 var q=(inp&&inp.value||'').trim().toLowerCase(), n=0;
 document.querySelectorAll('.lead-row[data-rg]').forEach(function(el){
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
{AUTH_GUARD.replace("__AUTH_HASH__", AUTH_HASH)}
<div class="wrap">
{nav_html(active)}
{FU_JS}
{body}
</div></body></html>"""

# ---- 登录页（index.html）----
LOGIN_TPL = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>登录 · 销售线索雷达</title>
<style>
*{box-sizing:border-box}
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
background:linear-gradient(135deg,#003DA5 0%,#0b6bcb 55%,#00B5E2 100%);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif}
.card{width:min(92vw,360px);background:#fff;border-radius:16px;padding:30px 26px 24px;box-shadow:0 18px 50px rgba(0,20,60,.28)}
.logo{font-size:13px;letter-spacing:.14em;color:#003DA5;font-weight:700;text-align:center}
h1{margin:10px 0 3px;font-size:20px;color:#12263f;text-align:center;font-weight:700}
.sub{margin:0 0 22px;font-size:12.5px;color:#8494a4;text-align:center}
label{display:block;font-size:12px;color:#5a6b7b;margin:12px 0 5px}
input{width:100%;padding:10px 12px;font-size:14px;border:1px solid #d5dde7;border-radius:9px;outline:none;transition:.15s}
input:focus{border-color:#0b6bcb;box-shadow:0 0 0 3px rgba(11,107,203,.12)}
button{width:100%;margin-top:20px;padding:11px;border:none;border-radius:9px;background:#003DA5;color:#fff;
font-size:14.5px;font-weight:600;cursor:pointer;transition:.15s}
button:hover{background:#0b52c4}
button:disabled{opacity:.6;cursor:default}
.err{margin-top:11px;font-size:12.5px;color:#c62828;text-align:center;min-height:17px}
</style></head><body>
<div class="card">
  <div class="logo">ROCKWELL · PTC</div>
  <h1>销售线索雷达</h1>
  <p class="sub">请登录后查看</p>
  <form id="f" autocomplete="off">
    <label>用户名</label>
    <input id="u" type="text" placeholder="请输入用户名" required>
    <label>密码</label>
    <input id="p" type="password" placeholder="请输入密码" required>
    <button id="go" type="submit">登 录</button>
  </form>
  <div class="err" id="e"></div>
</div>
<script>
var HASH="__HASH__";
function hex(buf){return Array.prototype.map.call(new Uint8Array(buf),function(b){return ("0"+b.toString(16)).slice(-2)}).join("")}
document.getElementById("f").addEventListener("submit",function(ev){
 ev.preventDefault();
 var u=document.getElementById("u").value.trim(),p=document.getElementById("p").value;
 var btn=document.getElementById("go"),err=document.getElementById("e");
 err.textContent="";btn.disabled=true;btn.textContent="登录中…";
 var done=function(h){
   if(h!==HASH){err.textContent="用户名或密码不正确";btn.disabled=false;btn.textContent="登 录";return;}
   try{localStorage.setItem("pd_auth",HASH)}catch(e){}
   var n=new URLSearchParams(location.search).get("next")||"dashboard.html";
   location.replace(n);
 };
 if(window.crypto&&crypto.subtle&&crypto.subtle.digest){
   crypto.subtle.digest("SHA-256",new TextEncoder().encode(u+":"+p)).then(function(b){done(hex(b))},function(){err.textContent="当前环境不支持登录校验，请用 https 访问";btn.disabled=false});
 }else{err.textContent="当前环境不支持登录校验，请用 https 访问";btn.disabled=false;btn.textContent="登 录";}
});
</script></body></html>"""
(ROOT/"index.html").write_text(LOGIN_TPL.replace("__HASH__", AUTH_HASH), encoding="utf-8")
log("index.html 登录页已生成")

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


<div class="sec"><div class="sec-h">否决池</div><table><tr><th>ID</th><th>公司</th><th>原因</th></tr>{reject_html}</table></div>
<div class="foot">日志 logs/run_{datetime.date.today().strftime('%Y%m%d')}.log ｜ 标记跟进存在本机浏览器，换手机需重新标记</div>
"""
(ROOT/"dashboard.html").write_text(shell("PLEX 每日商机","PLEX",plex_body),encoding="utf-8")

# ---- 线索详情页：列表只放一行，点开看这一家的完整档案 ----
LEAD_JS = """
<script>
(function(){
  var m=location.search.match(/[?&]c=([^&]+)/), c=m?decodeURIComponent(m[1]):'';
  var box=document.getElementById('leadbox'), d=(window.PD_LEAD||{})[c];
  if(!d){box.innerHTML='<div style="padding:18px;font-size:13px;color:#8494a4">没找到这家公司（可能今天已被刷出名单）。<a href="dashboard.html" style="color:#00B5E2">← 返回列表</a></div>';return;}
  var h='<div class="ld-h"><span class="tag" style="background:#eef4fb;color:#003DA5">'+d.g+'优先</span>'
    +'<b class="ld-nm">'+d.n+'</b>'
    +'<span class="ld-code">'+d.s+'</span>'
    +'<span class="ld-meta">📍 '+d.loc+'</span>'
    +'<span class="ld-meta">🗺️ '+d.rg+'</span>'
    +'<a class="ld-back" href="dashboard.html">← 返回列表</a></div>'
    +'<div class="ld-sum">'+d.sum+'</div>'
    +'<div class="ld-body">'+d.b+'</div>';
  box.innerHTML=h;
  document.title=d.n+' · 销售线索详情';
})();
</script>
"""

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

# ---- 线索详情页（所有产品页渲染完，LEADS 才完整）----
def _write_lead_page():
    _lead_json = json.dumps(LEADS, ensure_ascii=False).replace("</", "<\\/")
    lead_body = ('<div id="leadbox" style="font-size:13px;color:#8494a4;padding:10px 0">正在加载…</div>'
                 '<script>window.PD_LEAD=' + _lead_json + ';</script>' + LEAD_JS)
    (ROOT/"lead.html").write_text(shell("线索详情","PLEX",lead_body),encoding="utf-8")
    log(f"lead.html 已生成（{len(LEADS)} 家详情）")
_write_lead_page()

# ---- 即查结果独立页（临时查的公司不进主看板）----
def _inq_pending():
    """queue 里还没跑的任务，先在临时页显示“排队中”，避免点了按钮一片空白"""
    qdir = ROOT / "queue"
    if not qdir.exists(): return []
    out = []
    for fp in sorted(qdir.glob("*.txt")):
        try: raw = fp.read_text(encoding="utf-8").strip()
        except Exception: continue
        if "|" not in raw.replace("｜", "|"): continue
        out.append(raw.replace("｜", "|").split("|")[0].strip())
    return out

def _inq_html():
    if not _INQ:
        return '<div class="empty">还没有即查记录。到任意产品页用「交给 AI 自动查（排队）」提交一家公司，下次自动任务跑完就会出现在这里。</div>'
    rows=[]
    for code, d in sorted(_INQ.items(), key=lambda kv: kv[1].get("name","")):
        nm=d.get("name",""); pk=d.get("product_key","PLEX")
        em=(d.get("email") or {}).get("to","")
        emh=""
        if em:
            emh=f'<code style="background:#eef4fb;padding:1px 5px;border-radius:4px">{esc(em)}</code>'
            emh+=f' · <span style="font-size:11.5px;color:#8494a4">{esc((d.get("email") or {}).get("note",""))}</span>'
        ph="".join('<div>· <b>%s</b> <a class="ev" href="%s" target="_blank">🔗</a></div>'%(esc(x.get("num","")),esc(x.get("src",""))) for x in (d.get("phones") or []))
        src=d.get("src") or ""
        srch=f' <a class="ev" href="{esc(src)}" target="_blank">🔗 出处</a>' if src else ""
        rows.append(f"""<div class="lead" data-rg="{esc(d.get('region','') or '')}">
<div class="lead-top"><span class="badge">{esc(d.get('grade','中'))}</span><b>{esc(nm)}</b><span class="lc">{esc(code)}</span>
<span class="indtag">{esc(pk)}</span><span class="indtag">🗺️ {esc(d.get('region','') or '待确认')}</span></div>
<div style="font-size:12.5px;color:#3d4c5c;line-height:1.75;margin-top:4px">
<b>判断：</b>{esc(d.get('demand','待验证'))} —— {esc(d.get('demand_reason',''))}<br>
<b>信号：</b>{esc(d.get('signal',''))}</div>
{_plan_html(d, esc)}
<div style="font-size:12.5px;color:#3d4c5c;line-height:1.75;margin-top:6px"><b>联系人：</b>{esc(d.get('chief',''))}{srch}</div>
{_mailbox_html(d, esc)}
{('<div style="font-size:12px;color:#8a6b1f;margin-top:3px">📞 '+ph+'</div>') if ph else ''}
</div>""")
    return ('<div class="srchwrap"><input class="srch" id="q2" placeholder="🔍 输入公司名 / 代码 / 产品" oninput="flt2(this.value)">'
            '<div class="rgbar" id="rgbar2"></div></div>'
            '<div id="list2">'+"".join(rows)+'</div>')

_INQ_JS = """<script>
(function(){
  var q=(new URLSearchParams(location.search)).get('c');
  if(!q) return;
  var box=document.getElementById('q2');
  if(box){ box.value=q; }
  if(window.flt2){ flt2(q); }
  setTimeout(function(){
    var list=document.getElementById('list2'); if(!list) return;
    var vis=[].slice.call(list.querySelectorAll('.lead')).filter(function(el){ return el.style.display!=='none'; });
    if(vis.length===0){
      var b=document.createElement('div'); b.className='sec';
      b.style.cssText='border-color:#e3dcff;background:#faf8ff';
      b.innerHTML='<div class="sec-h">⏳ 你查的 '+q+' 还没有结果</div><div style="font-size:12.5px;color:#5a6b7b;line-height:1.8">这家还没跑过 AI。点上面的「提交给 AI 查」，在打开的新页面里点一次绿色「Commit changes」才算真正入队；入队后每天 09:00 / 20:30 自动跑，跑完结果直接出现在本页。缩写/英文简称建议换成公司全称，检索命中率更高。</div>';
      list.parentNode.insertBefore(b,list);
    } else {
      vis.forEach(function(el){
        var t=el.querySelector('.lead-top');
        if(t && !t.querySelector('.qtag')){
          var sp=document.createElement('span'); sp.className='indtag qtag';
          sp.style.cssText='background:#1f7a49;color:#fff';
          sp.textContent='本次查询'; t.appendChild(sp);
        }
      });
    }
  },80);
})();
</script>"""

_pend = _inq_pending()
_pend_html = ""
if _pend:
    _pend_html = ('<div class="sec" style="border-color:#e3dcff;background:#faf8ff">'
                  '<div class="sec-h">⏳ 排队中（' + str(len(_pend)) + ' 家 · 等自动任务）</div>'
                  '<div style="font-size:12.5px;color:#5a6b7b;line-height:1.8">'
                  + esc("、".join(_pend)) + '</div></div>')
inq_body=f"""
<div class="top"><h1>🔎 即查结果（临时查询 · 不进主看板）</h1></div>
<div class="sec" id="qkbox" style="border-color:#e3dcff;background:#faf8ff">
  <div class="sec-h">🔍 查一家不在雷达里的公司（临时，不改任何代码）</div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:8px">
    <input id="qiName" placeholder="公司全称（写全称，别用缩写）" style="flex:1;min-width:240px;padding:8px 10px;border:1px solid #cdd9e5;border-radius:8px;font-size:13px">
    <select id="qiProd" style="padding:8px 10px;border:1px solid #cdd9e5;border-radius:8px;font-size:13px">
      <option value="PLEX">PLEX</option><option value="FTPC">FTPC</option><option value="TWX">ThingWorx</option><option value="WIND">Windchill</option><option value="DMX">DataMosaix</option>
    </select>
    <a id="qkgo" class="ev" href="#" target="_blank" style="display:inline-block;background:#7a5af8;color:#fff;text-decoration:none;border-radius:8px;padding:9px 18px;font-size:13px;font-weight:600">🚀 提交给 AI 查</a>
  </div>
  <div style="margin-top:10px;font-size:12.5px;color:#5a6b7b;line-height:1.9">
    提交只要 3 步：① 点上面按钮 → 新窗口打开 GitHub ② 拉到页面最下面点绿色 <b>Commit changes</b> ③ 弹窗里再点一次 <b>Commit changes</b>。<br>
    <b style="color:#b33a2f">不点这一步＝没提交</b>，本页永远不会出现它的结果。
    <span id="qkchk" style="color:#0b62c4;cursor:pointer;text-decoration:underline">提交完了，点这里查是否已入队 →</span>
  </div>
  <div id="qkstat" style="margin-top:6px;font-size:12.5px;line-height:1.85"></div>
</div>
<script>
(function(){{
 var q=new URLSearchParams(location.search), c0=q.get('c')||'', p0=q.get('p')||'PLEX';
 var nameEl=document.getElementById('qiName'), prodEl=document.getElementById('qiProd');
 var go=document.getElementById('qkgo'), chk=document.getElementById('qkchk'), stat=document.getElementById('qkstat');
 if(c0) nameEl.value=c0;
 for(var i=0;i<prodEl.options.length;i++){{ if(prodEl.options[i].value===p0) prodEl.selectedIndex=i; }}
 function upd(){{
   var n=(nameEl.value||'').trim();
   if(!n) return;
   go.href='https://github.com/nuoke2007/plex-daily/new/main?filename='+encodeURIComponent('queue/'+n+'-'+prodEl.value+'.txt')+'&value='+encodeURIComponent(n+' | '+prodEl.value);
 }}
 nameEl.addEventListener('input',upd); prodEl.addEventListener('change',upd); upd();
 function check(){{
   var n=(nameEl.value||'').trim()||c0;
   if(!n){{ stat.innerHTML='先在上面填公司名'; return; }}
   stat.innerHTML='正在查队列…';
   fetch('https://api.github.com/repos/nuoke2007/plex-daily/contents/queue?t='+Date.now(), {{cache:'no-store'}})
    .then(function(r){{ return r.json(); }})
    .then(function(list){{
       var names=(list||[]).map(function(x){{ return decodeURIComponent(x.name||''); }}).filter(function(x){{ return x.slice(-4)==='.txt'; }});
       var hit=names.filter(function(x){{ return x.indexOf(n)>=0; }});
       if(hit.length){{ stat.innerHTML='<b style="color:#1f7a49">✅ 已入队：'+hit.join('、')+'</b>　下一次自动任务（09:00 / 20:30）跑完，结果会自动出现在本页。'; }}
       else {{ stat.innerHTML='<b style="color:#b33a2f">❌ 队列里还没有「'+n+'」</b>　说明 GitHub 那一步没点成。回去点一次绿色 Commit changes，再回来复查。'; }}
    }})
    .catch(function(){{ stat.innerHTML='查队列失败（网络限制），稍后再点一次。'; }});
 }}
 chk.addEventListener('click',check);
 if(c0) check();
}})();
</script>
<div class="sub">这里只放你临时交上来的公司。雷达内的日常线索仍在五个产品页，互不干扰。</div>
{_pend_html}
{_INQ_JS}
<div class="sec"><div class="sec-h">历史已查记录 {len(_INQ)} 家（全部临时查询，与本次查询无关）</div>{_inq_html()}</div>
<div class="foot">每次自动任务（09:00 / 20:30）处理 queue 队列，查完即写入本页</div>
"""
(ROOT/"inquiry.html").write_text(shell("即查结果","PLEX",inq_body),encoding="utf-8")
log(f"inquiry.html 已生成（即查 {len(_INQ)} 家）")

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
        body += f"\n\n打开看板看今日线索和搭线路径：{site or ''}"
        _send_push(sc,bk,title,body)
    else:
        log("非周一且近48h无新信号，静默不推送")

notify()

# ============================================================
# 7. 写日志
# ============================================================
log_path = LOGS / f"run_{datetime.date.today().strftime('%Y%m%d')}.log"
log_path.write_text("\n".join(log_lines), encoding="utf-8")
_ci_push_data()
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

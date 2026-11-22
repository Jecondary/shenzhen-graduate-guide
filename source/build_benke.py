# -*- coding: utf-8 -*-
"""《深圳应届毕业生补贴与税收优惠指南·本科版》v2
顺序：市级人才支持 → 就业类补贴 → 税收优惠 → 住房类支持 → 创业类补贴 → 区级入户补贴
每项固定：项目名称 / 申请条件 / 申请所需资料 / 申请入口 / 申请方法（+ 必要的参考链接）
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"D:\superAgent\交付\深圳应届生补贴全景\深圳应届毕业生补贴与税收优惠指南_本科版_2026年11月.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x5C, 0x8A)
GREY = RGBColor(0x59, 0x59, 0x59)
DARK = RGBColor(0x26, 0x26, 0x26)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xB0, 0x1E, 0x1E)
LINK = RGBColor(0x1F, 0x4E, 0x79)

HEAD_FONT = "微软雅黑"
BODY_FONT = "宋体"

# ---------- 常用链接 ----------
L_ABC = "https://hrss.sz.gov.cn/xxgk/zcfgjjd/zcfg/gfxwjcx/gfxwj/content/post_12820489.html"
L_ABC_JD = "https://hrss.sz.gov.cn/xxgk/zcfgjjd/zcfg/zcjd/content/post_12551353.html"
L_ZJJ = "https://zjj.sz.gov.cn/szszfhjsjwzgkml/szszfhjsjwzgkml/zxft/content/post_12808111.html"
L_TRAIN = "https://hrss.sz.gov.cn/tzgg/content/post_12080136.html"
L_TRAIN_XLS = "https://hrss.sz.gov.cn/attachment/1/1671/1671037/12080136.xls"
L_YM = "https://www.sz.gov.cn/cn/xxgk/zfxxgj/zwdt/content/post_12714558.html"
L_IIT = "https://www.chinatax.gov.cn/chinatax/n810219/n810780/c3965653/content.html"
L_PENSION = "https://www.gov.cn/zhengce/zhengceku/202412/content_6992498.htm"
L_BONUS = "https://www.gov.cn/zhengce/zhengceku/202308/content_6900595.htm"
L_KEY = "https://www.gov.cn/zhengce/zhengceku/202308/content_6896453.htm"
L_TALENT = "https://ga.sz.gov.cn/ZT/LHLY/TZGG/content/post_12850011.html"
L_FT = "https://www.szft.gov.cn/bmxx_qt/qwzzb/ych/rcyj/qnrc/content/post_12149484.html"
L_CERT = "https://www.gov.cn/zhengce/zhengceku/2021-12/03/content_5655553.htm"
L_OSTA = "https://zscx.osta.org.cn"
L_SZT = "https://www.shenzhentong.com/fwzx/1.html"
L_LB_JD = "https://www.sz.gov.cn/cn/xxgk/zfxxgj/zcjd/content/post_12551438.html"
L_LB = "https://hrss.sz.gov.cn/xxgk/zcfgjjd/zcfg/rcfw/content/post_12551112.html"
L_SPORT = "https://www.baoan.gov.cn/hcjd/zmhd/ywzsk/content/post_13045419.html"
L_SZG = "https://www.workercn.cn/papers/grrb/2026/05/28/1/news-1.html"
L_HSA = "https://hsa.sz.gov.cn"
L_GM = "https://www.szgm.gov.cn/xxgk/xqgwhxxgkml/gzgg/content/post_12902221.html"


def cn(run, name):
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def cell_margins(table, top=28, bottom=28, left=85, right=85):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def table_borders(table, color="C9D2E0"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def add_field(paragraph, instr):
    run = paragraph.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1)
    run._r.append(it)
    run._r.append(f2)


doc = Document()
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.2)
sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.0)
sec.header_distance = Cm(1.2)
sec.footer_distance = Cm(1.2)
sec.different_first_page_header_footer = True

normal = doc.styles["Normal"]
normal.font.size = Pt(10.5)
normal.font.name = BODY_FONT
normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
normal.paragraph_format.line_spacing = 1.4
normal.paragraph_format.space_after = Pt(3)


def group_h(text, new_page=True):
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = new_page
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    cn(r, HEAD_FONT)
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = NAVY
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "10")
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), "1F3864")
    pbdr.append(bottom)
    pPr.append(pbdr)


def item_h(no, name, tag=None, newpage=False):
    p = doc.add_paragraph()
    if newpage:
        p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r0 = p.add_run(no + "　")
    cn(r0, HEAD_FONT)
    r0.font.size = Pt(12.5)
    r0.font.bold = True
    r0.font.color.rgb = RGBColor(0x9A, 0xB0, 0xCC)
    r1 = p.add_run(name)
    cn(r1, HEAD_FONT)
    r1.font.size = Pt(12.5)
    r1.font.bold = True
    r1.font.color.rgb = NAVY
    if tag:
        r2 = p.add_run("　" + tag)
        cn(r2, HEAD_FONT)
        r2.font.size = Pt(8)
        r2.font.bold = True
        r2.font.color.rgb = RED


def fill_cell(cell, lines, size=8, label=False):
    first = True
    for line in lines:
        if first:
            p = cell.paragraphs[0]
            first = False
        else:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0.5)
        p.paragraph_format.line_spacing = 1.1
        if label:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(line)
            cn(r, HEAD_FONT)
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = NAVY
        elif line.startswith("@"):
            p.paragraph_format.left_indent = Cm(0.42)
            p.paragraph_format.first_line_indent = Cm(-0.42)
            r = p.add_run("· ")
            cn(r, BODY_FONT)
            r.font.size = Pt(7)
            r.font.color.rgb = BLUE
            r = p.add_run(line[1:].strip())
            cn(r, BODY_FONT)
            r.font.size = Pt(6.5)
            r.font.color.rgb = LINK
        elif line.startswith("·"):
            p.paragraph_format.left_indent = Cm(0.42)
            p.paragraph_format.first_line_indent = Cm(-0.42)
            r = p.add_run("· ")
            cn(r, BODY_FONT)
            r.font.size = Pt(size)
            r.font.color.rgb = BLUE
            r = p.add_run(line[1:].strip())
            cn(r, BODY_FONT)
            r.font.size = Pt(size)
            r.font.color.rgb = DARK
        elif line.startswith("▲"):
            r = p.add_run(line[1:].strip())
            cn(r, BODY_FONT)
            r.font.size = Pt(size)
            r.font.bold = True
            r.font.color.rgb = RED
        elif line.startswith("■"):
            r = p.add_run(line[1:].strip())
            cn(r, HEAD_FONT)
            r.font.size = Pt(size)
            r.font.bold = True
            r.font.color.rgb = NAVY
        else:
            r = p.add_run(line)
            cn(r, BODY_FONT)
            r.font.size = Pt(size)
            r.font.color.rgb = DARK


def item_block(item, no):
    item_h(no, item["name"], item.get("tag"), item.get("newpage", False))
    if item.get("links"):
        for i, line in enumerate(item["links"]):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3) if i == len(item["links"]) - 1 else Pt(0)
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.left_indent = Cm(1.15)
            p.paragraph_format.first_line_indent = Cm(-1.15)
            p.paragraph_format.keep_with_next = True
            r0 = p.add_run("参考链接　" if i == 0 else "　　　　　")
            cn(r0, HEAD_FONT)
            r0.font.size = Pt(7)
            r0.font.bold = True
            r0.font.color.rgb = RGBColor(0x9A, 0xB0, 0xCC)
            r = p.add_run(line[1:].strip() if line.startswith("@") else line)
            cn(r, BODY_FONT)
            r.font.size = Pt(7)
            r.font.color.rgb = LINK
    rows = [
        ("申请条件", item["condition"]),
        ("申请所需资料", item["materials"]),
        ("申请入口", item["entry"]),
        ("申请方法", item["method"]),
    ]
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    table_borders(t)
    cell_margins(t)
    for label, lines in rows:
        row = t.add_row()
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        cs.set(qn("w:val"), "true")
        trPr.append(cs)
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Cm(2.7)
        c1.width = Cm(13.9)
        shade(c0, "EDF1F7")
        c0.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        c1.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        fill_cell(c0, [label], label=True)
        fill_cell(c1, lines, size=8)


# ---------------- 页眉页脚 ----------------
hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hr = hp.add_run("深圳应届毕业生补贴与税收优惠指南 · 本科版 · 2026年11月")
cn(hr, HEAD_FONT)
hr.font.size = Pt(8)
hr.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
for txt, fld in (("第 ", None), (None, " PAGE "), (" 页 / 共 ", None), (None, " NUMPAGES "), (" 页", None)):
    if fld:
        add_field(fp, fld)
    else:
        r = fp.add_run(txt)
        cn(r, HEAD_FONT)
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

# ---------------- 封面 ----------------
for _ in range(4):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("深圳应届毕业生")
cn(r, HEAD_FONT)
r.font.size = Pt(30)
r.font.bold = True
r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("补贴与税收优惠指南")
cn(r, HEAD_FONT)
r.font.size = Pt(30)
r.font.bold = True
r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(16)
r = p.add_run("本 科 版")
cn(r, HEAD_FONT)
r.font.size = Pt(16)
r.font.bold = True
r.font.color.rgb = BLUE

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(8)
r = p.add_run("41 项 · 每项含申请条件、所需资料、申请入口、申请方法")
cn(r, BODY_FONT)
r.font.size = Pt(10)
r.font.color.rgb = GREY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(60)
r = p.add_run("政策口径：2026 年 11 月")
cn(r, HEAD_FONT)
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = BLUE

doc.add_page_break()

# ---------------- 目录 ----------------
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(10)
r = p.add_run("目　录")
cn(r, HEAD_FONT)
r.font.size = Pt(18)
r.font.bold = True
r.font.color.rgb = NAVY

toc_slot = doc.add_paragraph()
toc_slot.paragraph_format.space_after = Pt(0)

# ================= 数据 =================
DIST_BASE = [
    "■ 基本门槛",
    "· 全日制本科，取得学士学位；年龄 35 周岁及以下（以引进时间为判断时点）",
    "· 首次入户本区，且非深圳市内迁移（深圳生源、学历晋升另有例外，需问窗口）",
    "· 在本区企业或机构全职工作，并连续缴纳社会保险（月数见本区条件）",
    "· 自《人才引进通知书》签发之日起 24 个月内提出申请，逾期视为放弃",
    "· 未享受过本市其他区同类新引进人才补贴或奖励；机关事业单位在编人员不享受",
]
DIST_MAT = [
    "《新引进优秀青年人才入户补贴申请表》（单位盖章）",
    "身份证、户口本（首页 + 本人页，需体现落户本区）",
    "毕业证、学位证；学信网《教育部学历证书电子注册备案表》",
    "《人才引进通知书》或引进审批文件",
    "劳动合同（首页、期限页、签字盖章页）",
    "本区社保缴费记录（连续月数见本区条件）",
    "单位营业执照复印件（加盖公章）",
    "本人银行账户（多为金融社保卡）",
]

ITEMS = [
    # ================= 一、市级人才支持 =================
    {
        "group": "一、市级人才支持",
        "name": "青年人才认定",
        "tag": "其他项目的前置资格",
        "condition": [
            "■ 三项都满足才能申请",
            "· 年龄 45 周岁以下（不含 45 岁）",
            "· 2026 年 1 月 1 日起首次在深圳就业创业，以首次缴纳职工五险的时间为准",
            "· 在教育背景、创新能力、创新贡献三类标准中符合任意一类",
            "■ 本科走“教育背景”，需满足下列任一",
            "· “985”或“211”高校本科及以上学历和学士及以上学位，且所学专业在《重点产业领域专业目录》内",
            "· 国家“双一流”建设学科本科及以上学历和学士及以上学位",
            "· 软科／QS／THE／U.S.News 任一榜单前 200 名高校学士及以上学位，且境外院校专业属 STEM、境内院校专业在《重点产业领域专业目录》内",
            "■ 走不通教育背景时，还有“创新能力”11 条路径（摘要）",
            "· 国家、省、市（限副省级市）发改、科创、工信、教育、卫健部门的纵向项目负责人，或重点项目前五名参与人",
            "· 有一定影响力的国内外奖项、知名赛事较好成绩",
            "· 境内外相关职业资格（清单由认定标准细则分批发布）",
            "· 以第一作者或唯一通讯作者在重点核心期刊发表高水平论文",
            "▲ 《重点产业领域专业目录》未全文公开，申报时在 i深圳专区按专业名称检索比对；不确定的先拨 12333 或问社保所在区人力资源局",
        ],
        "materials": [
            "身份证",
            "毕业证、学位证；学信网《教育部学历证书电子注册备案表》",
            "境外学历需教育部留学服务中心《国外学历学位认证书》",
            "劳动合同（在深就业）或营业执照（在深创业）",
            "深圳社保缴费记录（需体现首次参保时间）",
            "专业与《重点产业领域专业目录》的对应说明（专业名称不完全一致时）",
            "创新能力佐证材料（项目立项文件、获奖证书、职业资格证书、论文首页等，走该路径时提供）",
        ],
        "entry": [
            "■ i深圳 App → 青年人才服务保障专区",
            "· 由社会保险关系所在区人力资源部门受理",
        ],
        "method": [
            "· 步骤 1：实名登录 i深圳，进入“青年人才服务保障专区”，选择“青年人才认定”",
            "· 步骤 2：选择申报路径（教育背景／创新能力／创新贡献），按路径上传对应材料",
            "· 步骤 3：区人力资源部门受理，材料不齐会一次性告知补正；自受理之日起 10 个工作日内出具审核意见",
            "· 步骤 4：审核通过即认定为青年人才，结果在专区可查；对结果有异议可申请复核",
            "▲ 认定须在首次在深缴纳社保之日起 1 年内提交，逾期不再受理",
        ],
        "links": [
            "@《深圳市青年人才认定和管理办法》（深人社规〔2025〕7号）全文：" + L_ABC,
            "@政策解读：" + L_ABC_JD,
            "@《重点产业领域专业目录》无公开全文，以申报系统检索或区人社局答复为准",
        ],
    },
    {
        "group": "一、市级人才支持",
        "name": "青年人才安居补贴",
        "tag": "1250 元/月 × 24 个月",
        "condition": [
            "■ 补贴标准",
            "· 1250 元/月（税后 1000 元），发放期不超过 24 个月，合计 3 万元",
            "■ 申请条件（需同时满足）",
            "· 已通过青年人才认定",
            "· 在本市无自有住房",
            "· 未正在本市享受住房保障优惠政策",
            "· 未正在享受用人单位提供的优惠住房",
            "· 通过本市工作单位依法缴纳社会保险",
            "· 未租住过渡性住房，且已按规定办理居住登记",
            "· 经用人单位审核同意后申报",
            "▲ 与“过渡性住房”二选一，不得重复享受；先领补贴后不能改租过渡房",
        ],
        "materials": [
            "身份证",
            "青年人才认定结果（专区可查）",
            "居住登记信息（社区网格登记）",
            "本人银行账户",
            "用人单位审核意见（专区线上确认）",
            "诚信承诺书（系统内签署）",
        ],
        "entry": [
            "■ i深圳 App → 青年人才服务保障专区 → 安居补贴发放系统",
            "· 向社会保险关系所在区提出申请",
        ],
        "method": [
            "· 步骤 1：认定通过后 1 年内，在专区提交安居补贴申请，逾期视为放弃",
            "· 步骤 2：用人单位在系统内审核同意",
            "· 步骤 3：区人力资源部门审核，10 个工作日内反馈；系统比对通过的“秒批”",
            "· 步骤 4：首笔发放至本人银行账户；此后系统按月自动续发（免申即享），无需再次申请",
            "▲ 社保断缴即暂停发放，恢复缴纳后继续，但暂停期间不补发",
        ],
        "links": [
            "@《深圳市青年人才住房支持实施办法》政策解读（市住建局，2026-05-22）：" + L_ZJJ,
        ],
    },
    {
        "group": "一、市级人才支持",
        "name": "青年人才过渡性住房",
        "tag": "与安居补贴二选一",
        "condition": [
            "■ 支持标准",
            "· 租金约为市场参考租金的 60%，以一间房为主，最长租住 36 个月",
            "· 配备热水器、空调、衣柜、床、桌椅等家具家电",
            "■ 申请条件（需同时满足）",
            "· 已通过青年人才认定",
            "· 在本市无自有住房",
            "· 未正在本市享受住房保障优惠政策",
            "· 未正在享受用人单位提供的优惠住房",
            "· 通过本市工作单位依法缴纳社会保险",
            "▲ 租住期间不得转租、转借或擅自改变用途；中途自愿退出的，不得再次申请",
        ],
        "materials": [
            "身份证",
            "青年人才认定结果",
            "深圳社保缴费记录",
            "婚姻及家庭情况材料（如有配偶、未成年子女）",
            "用人单位在职证明",
        ],
        "entry": [
            "■ i深圳 App → 青年人才服务保障专区（住房支持—过渡性住房）",
        ],
        "method": [
            "· 步骤 1：认定通过后 1 年内提交申请，逾期视为放弃",
            "· 步骤 2：在专区查看房源，结合供给情况自行选择意向房源",
            "· 步骤 3：区住建部门审核资格",
            "· 步骤 4：签约入住，按合同约定使用并缴纳水电燃气物业费",
        ],
        "links": [
            "@《深圳市青年人才住房支持实施办法》政策解读：" + L_ZJJ,
        ],
    },

    # ================= 二、就业类补贴 =================
    {
        "group": "二、就业类补贴",
        "name": "一次性求职创业补贴",
        "tag": "3000 元，一次性",
        "condition": [
            "■ 补贴标准",
            "· 每人 3000 元，一次性发放",
            "■ 申请条件（需同时满足）",
            "· 本市普通高校、中等职业学校、技工院校毕业学年学生",
            "· 符合下列四类情形之一：城乡困难家庭成员（低保、残疾人、脱贫人口、特困职工家庭）、特困人员、残疾毕业生、在学期间已获得国家助学贷款",
            "▲ 必须在校期间申请，由学校代收代办，个人不能自行申报；已毕业的无法补办",
        ],
        "materials": [
            "《一次性求职创业补贴申请表》（学校提供模板）",
            "身份证、学生证或学籍证明",
            "困难情形证明：低保证、特困人员救助供养证、残疾人证、国家助学贷款合同（四选一）",
            "本人银行账户（多为金融社保卡）",
        ],
        "entry": [
            "■ 所在学校就业指导中心（学校统一申报）",
            "· 学校通过广东政务服务网或广东公共就业服务云平台报送",
        ],
        "method": [
            "· 步骤 1：毕业学年内向学校就业指导中心提交材料",
            "· 步骤 2：学校审核、公示",
            "· 步骤 3：学校代向区或街道公共就业服务机构申报",
            "· 步骤 4：审批通过后次月月底前拨付至个人金融社保卡账户",
        ],
    },
    {
        "group": "二、就业类补贴",
        "name": "灵活就业社会保险补贴",
        "tag": "≤600 元/月，最长 3 年",
        "condition": [
            "■ 补贴标准",
            "· 按实际缴费额的 2/3 补贴，每月最高 600 元，累计最长 3 年",
            "■ 申请条件（需同时满足）",
            "· 已办理灵活就业登记",
            "· 以灵活就业人员身份参加企业职工基本养老保险和基本医疗保险，并连续缴纳 3 个月以上",
            "· 每周工作时间累计 20 小时以上",
            "· 从事岗位属于规定范围：家政服务与看护、餐饮服务、社区保洁保绿、来料加工与再生资源回收、平台新业态",
            "▲ 须在满足条件之日起 6 个月内首次申请",
        ],
        "materials": [
            "身份证",
            "社会保障卡（或《就业创业证》）",
            "灵活就业登记证明",
            "社保缴费记录（连续 3 个月以上）",
            "本人银行账户",
            "灵活就业岗位及工时承诺书",
        ],
        "entry": [
            "■ 户籍地或经常居住地的区／街道公共就业服务机构",
            "· 线上：广东政务服务网、广东公共就业服务云平台",
        ],
        "method": [
            "· 步骤 1：先办理灵活就业登记（粤就业、粤省事或街道窗口）",
            "· 步骤 2：以灵活就业身份参保并连续缴费 3 个月",
            "· 步骤 3：向街道公共就业服务机构提交补贴申请",
            "· 步骤 4：审核通过后按月或按季拨付",
        ],
    },
    {
        "group": "二、就业类补贴",
        "name": "个人社会保险补贴",
        "tag": "个人缴费 ×25%，最长 12 个月",
        "condition": [
            "■ 补贴标准",
            "· 按个人缴纳的养老、医疗、失业三项保险费之和的 25% 补贴，最长 12 个月",
            "■ 申请条件（需同时满足）",
            "· 在重点行业领域中小微企业就业",
            "· 签订 1 年以上劳动合同",
            "· 2026 年内在本市缴纳社会保险",
            "· 离校 2 年内未就业的高校毕业生，须在省标识名单内；不在名单的，可凭毕业证由企业提交",
            "■ 两个名词",
            "· 重点行业中小微企业：按《中小企业划型标准规定》（工信部联企业〔2011〕300 号），以行业、从业人员、营业收入三项划分，由企业对照自查、人社部门核验",
            "· 省标识名单：省级人社部门比对学生信息形成的“离校 2 年内未就业高校毕业生”名单，系统自动校验；不在名单的可凭毕业证由企业提交",
            "▲ 政策执行至 2026 年 12 月 31 日",
        ],
        "materials": [
            "身份证、毕业证",
            "劳动合同",
            "社保缴费记录",
            "本人银行账户",
            "企业属于重点行业中小微企业的说明",
        ],
        "entry": [
            "■ 广东公共就业服务云平台 ggfw.hrss.gd.gov.cn",
        ],
        "method": [
            "· 步骤 1：确认企业属重点行业中小微、合同 1 年以上",
            "· 步骤 2：由企业或个人在云平台提交申请",
            "· 步骤 3：区人力资源部门审核",
            "· 步骤 4：补贴拨付至个人账户",
        ],
    },
    {
        "group": "二、就业类补贴",
        "name": "职业技能提升补贴",
        "tag": "按目录，最高一般 3000 元",
        "condition": [
            "■ 补贴标准",
            "· 按《补贴目录》对应工种、等级定额补贴；同一职业同一等级不可重复享受",
            "· 同一年度可享受 1 次，累计不超过 3 次",
            "■ 什么是《补贴目录》",
            "· 指《广东省职业技能培训补贴（指导）标准目录（2025 年版）》（粤人社发〔2025〕5号），深圳自 2025 年 3 月 1 日起按该目录的“职业技能等级补贴（指导）标准”执行；目录外的工种仍按《深圳市职业技能培训补贴办法》标准执行",
            "· 目录以 Excel 附件形式发布，逐工种列出等级与金额，可直接下载比对",
            "■ 证书目录在哪查",
            "· 《国家职业资格目录（2021年版）》是人社部发布的唯一国家级清单，共 72 项：专业技术人员职业资格 59 项（准入类 33 项、水平评价类 26 项）、技能人员职业资格 13 项",
            "· 目录之外的“XX 资格证”不属于国家职业资格，不能作为补贴或个税扣除依据",
            "· 证书真伪查询：技能人才评价证书全国联网查询（职业资格证书与职业技能等级证书都在此查）",
            "· 两个目录别混：《国家职业资格目录》规定“哪些证算数”，《补贴目录》规定“各证补多少钱”",
            "■ 申请条件（需同时满足）",
            "· 非全日制在校生",
            "· 取得《补贴目录》范围内的职业资格证书或职业技能等级证书、专项职业能力证书",
            "· 在证书核发之日起 12 个月内申请",
            "▲ 政策有效期已延长至 2026 年 12 月 31 日",
        ],
        "materials": [
            "身份证、社会保障卡",
            "职业资格证书或职业技能等级证书、专项职业能力证书",
            "证书可在技能人才评价证书全国联网查询系统查验",
            "本人银行账户",
        ],
        "entry": [
            "■ 广东政务服务网深圳分厅 gdzwfw.gov.cn/?region=440300 → 搜索“职业技能培训补贴”",
            "· 原“广东省职业技能培训和评价服务平台 ggfw.hrss.gd.gov.cn/OUPX”入口已失效（2026-11 核验为 404），统一从省政务服务网进入",
            "· 咨询：12333；市职训指导中心 0755-83699610",
        ],
        "method": [
            "· 步骤 1：先下载《补贴目录》，确认证书工种与等级在目录内、金额多少",
            "· 步骤 2：在平台注册并提交证书信息",
            "· 步骤 3：系统比对证书真伪与参保信息",
            "· 步骤 4：审核通过后拨付至社保卡金融账户",
        ],
        "links": [
            "@《关于我市职业技能培训补贴业务有关事项的温馨提示》（含省目录与深圳标准衔接）：" + L_TRAIN,
            "@《补贴目录》Excel 附件（广东省职业技能培训补贴指导标准 2025 年版）：" + L_TRAIN_XLS,
            "@《国家职业资格目录（2021年版）》（人社部公告）：" + L_CERT,
            "@技能人才评价证书全国联网查询：" + L_OSTA,
        ],
    },
    {
        "group": "二、就业类补贴",
        "name": "工会“圆梦计划”学历与技能补贴",
        "tag": "学历 1000—2000 元 / 技能最高 3000 元",
        "condition": [
            "■ 补贴标准（第十九届，2026 年 3 月启动）",
            "· 学历补贴：取得成人非全日制专科或本科毕业证书，补贴 1000 元；同时取得成人非全日制本科毕业证书与学士学位证书，补贴 2000 元",
            "· 技能证书补贴：在指定期间内取得工会选定的国家职业资格证书或职业技能等级证书，最高补贴 3000 元/人",
            "· 另有免费技能培训名额：三项工程 4000 人、“20+8”产业 1500 人、企业班组长 10000 人、质量体系 2000 人",
            "■ 申请条件（需同时满足）",
            "· 深圳在职职工，且在“深i工”平台完成实名认证的工会会员",
            "· 学历补贴限成人非全日制学历，全日制本科不适用",
            "· 技能证书须在工会指定期间内取得，且属于工会选定证书范围",
            "▲ 名额有限：学历补贴 4000 人、技能证书补贴 8500 人，按批次报名、先到先得",
        ],
        "materials": [
            "身份证",
            "“深i工”实名认证信息（工会会员身份）",
            "毕业证 + 学位证（学历补贴）",
            "职业资格证书或职业技能等级证书（技能补贴）",
            "社保缴费记录（在职证明）",
            "本人银行账户",
        ],
        "entry": [
            "■ “深i工”平台 → 圆梦计划报名通道",
            "· 深圳市总工会官网及公众号同步发布批次通知",
        ],
        "method": [
            "· 步骤 1：在“深i工”完成实名认证，确认工会会员身份",
            "· 步骤 2：按批次报名（名额有限，开放即报）",
            "· 步骤 3：上传学历或证书材料",
            "· 步骤 4：审核、公示后拨付至个人账户",
        ],
        "links": [
            "@我市启动新一轮“圆梦计划”（第十九届服务清单，深圳市总工会）：" + L_YM,
        ],
    },

    # ================= 三、税收优惠 =================
    {
        "group": "三、税收优惠",
        "name": "住房租金专项附加扣除",
        "tag": "1500 元/月",
        "condition": [
            "■ 扣除标准",
            "· 深圳属计划单列市，每月 1500 元定额扣除，全年最高抵扣 1.8 万元应税所得",
            "■ 申请条件",
            "· 在主要工作城市没有自有住房，且发生住房租金支出",
            "· 夫妻双方主要工作城市相同的，只能由一方扣除",
            "▲ 与“住房贷款利息”扣除不能同时享受",
        ],
        "materials": [
            "无需上传纸质材料，在个人所得税 App 填报即可",
            "留存备查：住房租赁合同、租金支付记录",
        ],
        "entry": [
            "■ 个人所得税 App → 专项附加扣除填报 → 住房租金",
        ],
        "method": [
            "· 步骤 1：登录个人所得税 App，选择“住房租金”",
            "· 步骤 2：填写租赁房屋地址、租赁期、出租方信息",
            "· 步骤 3：选择扣缴义务人（单位按月预扣）或年度汇算时扣除",
            "· 步骤 4：每年 12 月确认次年信息；漏填可在次年 3—6 月汇算清缴时补扣",
        ],
        "links": [
            "@《个人所得税专项附加扣除暂行办法》全文（国家税务总局）：" + L_IIT,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "继续教育专项附加扣除",
        "tag": "400 元/月 或 3600 元/年",
        "condition": [
            "■ 扣除标准",
            "· 学历（学位）继续教育：每月 400 元，同一学历最长 48 个月",
            "· 职业资格继续教育：取得证书当年按 3600 元定额扣除",
            "■ 申请条件",
            "· 在中国境内接受学历（学位）继续教育，或取得技能人员、专业技术人员职业资格证书",
            "▲ 全日制在读由父母按“子女教育”扣除，本人不能重复享受",
        ],
        "materials": [
            "无需上传纸质材料",
            "留存备查：学籍证明、职业资格证书",
        ],
        "entry": [
            "■ 个人所得税 App → 专项附加扣除填报 → 继续教育",
        ],
        "method": [
            "· 步骤 1：登录 App 选择“继续教育”",
            "· 步骤 2：选择“学历（学位）继续教育”或“职业资格继续教育”",
            "· 步骤 3：填写教育起止时间或证书信息",
            "· 步骤 4：选择单位按月预扣或年度汇算扣除",
        ],
        "links": [
            "@《个人所得税专项附加扣除暂行办法》全文：" + L_IIT,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "赡养老人专项附加扣除",
        "tag": "3000 元/月",
        "condition": [
            "■ 扣除标准",
            "· 独生子女：每月 3000 元",
            "· 非独生子女：与兄弟姐妹分摊每月 3000 元额度，每人不超过 1500 元",
            "■ 申请条件",
            "· 赡养年满 60 岁的父母，或子女均已去世的年满 60 岁的祖父母、外祖父母",
            "▲ 2023 年起标准由 2000 元/月提高至 3000 元/月，网上旧攻略仍写 2000 元",
        ],
        "materials": [
            "无需上传纸质材料",
            "留存备查：约定或指定分摊协议（非独生子女）",
        ],
        "entry": [
            "■ 个人所得税 App → 专项附加扣除填报 → 赡养老人",
        ],
        "method": [
            "· 步骤 1：登录 App 选择“赡养老人”",
            "· 步骤 2：选择“独生子女”或“非独生子女”，填写被赡养人信息",
            "· 步骤 3：非独生子女选择分摊方式（均摊、约定分摊、指定分摊）并填写本人额度",
            "· 步骤 4：选择单位按月预扣或年度汇算扣除",
        ],
        "links": [
            "@《个人所得税专项附加扣除暂行办法》全文（赡养老人标准后经国务院调整）：" + L_IIT,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "住房贷款利息专项附加扣除",
        "tag": "1000 元/月，最长 240 个月",
        "condition": [
            "■ 扣除标准",
            "· 每月 1000 元定额扣除，扣除期限最长不超过 240 个月（20 年）",
            "■ 申请条件",
            "· 本人或配偶购买中国境内住房发生的首套住房贷款利息支出",
            "· 可选择由夫妻一方扣除",
            "▲ 与“住房租金”扣除不能同时享受；认贷不认房，具体首套认定以贷款合同为准",
        ],
        "materials": [
            "无需上传纸质材料",
            "留存备查：住房贷款合同、贷款还款支出凭证",
        ],
        "entry": [
            "■ 个人所得税 App → 专项附加扣除填报 → 住房贷款利息",
        ],
        "method": [
            "· 步骤 1：登录 App 选择“住房贷款利息”",
            "· 步骤 2：填写房屋信息、贷款合同编号、首次还款日期、贷款期限",
            "· 步骤 3：选择本人或配偶扣除",
            "· 步骤 4：选择单位按月预扣或年度汇算扣除",
        ],
        "links": [
            "@《个人所得税专项附加扣除暂行办法》全文：" + L_IIT,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "大病医疗专项附加扣除",
        "tag": "自付超 1.5 万部分，限额 8 万",
        "condition": [
            "■ 扣除标准",
            "· 一个纳税年度内，医保目录范围内个人自付累计超过 15000 元的部分，在 80000 元限额内据实扣除",
            "■ 申请条件",
            "· 本人、配偶或未成年子女发生的医药费用支出",
            "▲ 只能在年度汇算清缴时扣除，不能在平时预扣预缴环节享受",
        ],
        "materials": [
            "留存备查：医疗服务收费票据、医保结算单（可在“国家医保服务平台”App 查询年度自付金额）",
        ],
        "entry": [
            "■ 个人所得税 App → 年度汇算 → 大病医疗",
        ],
        "method": [
            "· 步骤 1：在“国家医保服务平台”App 查询本人或家属年度医保目录内自付金额",
            "· 步骤 2：次年 3—6 月办理年度汇算时，在 App 中填报“大病医疗”",
            "· 步骤 3：填写患者信息、医药费用总金额与个人负担金额",
            "· 步骤 4：系统核验后并入汇算，多退少补",
        ],
        "links": [
            "@《个人所得税专项附加扣除暂行办法》全文：" + L_IIT,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "个人养老金税前扣除",
        "tag": "12000 元/年",
        "condition": [
            "■ 扣除标准",
            "· 每年缴费上限 12000 元，在综合所得或经营所得中据实扣除",
            "· 账户内投资收益暂不征收个人所得税",
            "· 领取时按 3% 单独计算缴纳个人所得税",
            "■ 申请条件",
            "· 已参加城镇职工基本养老保险或城乡居民基本养老保险",
            "· 自愿在商业银行开立个人养老金账户和资金账户",
            "▲ 资金锁定到退休才能领取；边际税率不超过 3% 时基本没有节税收益，适合税率 10% 以上的人",
            "▲ 节税测算：按 12000 元全额缴，边际税率 10% 约省 840 元、20% 约省 2040 元、30% 约省 3240 元（均已扣除领取时 3% 的税）",
        ],
        "materials": [
            "个人养老金账户开立信息（银行 App 可查）",
            "缴费凭证（银行 App 生成，含二维码）",
            "身份证",
        ],
        "entry": [
            "■ 开户：国家社会保险公共服务平台、电子社保卡、商业银行 App",
            "■ 扣除：个人所得税 App → 个人养老金扣除信息管理",
        ],
        "method": [
            "· 步骤 1：在银行 App 或电子社保卡开立个人养老金账户、资金账户",
            "· 步骤 2：年内缴费，累计不超过 12000 元（可按月或一次性）",
            "· 步骤 3：在个税 App“个人养老金扣除信息管理”中扫码录入缴费凭证",
            "· 步骤 4：选择单位按月预扣或年度汇算时扣除",
            "· 步骤 5：达到领取条件（退休、完全丧失劳动能力、出国定居等）后领取，按 3% 缴税",
        ],
        "links": [
            "@《关于在全国范围实施个人养老金个人所得税优惠政策的公告》（2024 年）：" + L_PENSION,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "全年一次性奖金单独计税",
        "tag": "年终奖不并入综合所得",
        "condition": [
            "■ 政策内容",
            "· 全年一次性奖金不并入当年综合所得，以奖金收入除以 12 个月得到的数额，按按月换算后的综合所得税率表单独计税",
            "■ 申请条件",
            "· 居民个人取得全年一次性奖金",
            "· 一个纳税年度内，该计税办法只允许采用一次",
            "▲ 政策延续至 2027 年 12 月 31 日；与“并入综合所得”相比哪个更省税，因人而异，年度汇算时 App 会同时算出两种结果",
        ],
        "materials": [
            "无需材料",
            "单位发放时在个税扣缴端选择“全年一次性奖金”所得项目",
        ],
        "entry": [
            "■ 单位代扣代缴环节选择；或个人所得税 App 年度汇算时切换比较",
        ],
        "method": [
            "· 步骤 1：发年终奖前告知 HR，按“全年一次性奖金”单独计税申报",
            "· 步骤 2：单位在扣缴端选择对应所得项目，不并入当月工资",
            "· 步骤 3：次年汇算时，在 App 中分别试算“单独计税”和“并入综合所得”",
            "· 步骤 4：选择税额较低的一种提交",
        ],
        "links": [
            "@《关于延续实施全年一次性奖金个人所得税政策的公告》（财政部 税务总局公告 2023 年第 30 号）：" + L_BONUS,
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "个人所得税年度汇算（退税）",
        "tag": "多退少补，3—6 月办理",
        "condition": [
            "■ 能拿到什么",
            "· 多退少补。应届生首年常见退税情形：全年收入不足 6 万元但被预扣过税；漏填住房租金、赡养老人等专项附加扣除；年中入职导致基本减除费用未扣足",
            "■ 申请条件",
            "· 居民个人取得工资薪金、劳务报酬、稿酬、特许权使用费四项综合所得",
            "· 办理期为次年 3 月 1 日至 6 月 30 日",
            "▲ 年度综合所得收入不超过 12 万元，或补税金额不超过 400 元的，可免于办理（但想退税仍需申报）",
        ],
        "materials": [
            "身份证",
            "本人 I 类银行借记卡（用于接收退税）",
            "收入纳税明细（App 可查）",
            "专项附加扣除信息",
        ],
        "entry": [
            "■ 个人所得税 App → 综合所得年度汇算",
        ],
        "method": [
            "· 步骤 1：3 月 1 日起登录 App，进入“综合所得年度汇算”",
            "· 步骤 2：选择“使用已申报数据填报”，核对收入与已缴税额",
            "· 步骤 3：补填或修正专项附加扣除、个人养老金等扣除项",
            "· 步骤 4：提交申报，绑定银行卡申请退税",
            "· 步骤 5：税务机关审核后由国库办理退库，一般 10 个工作日内到账",
        ],
    },
    {
        "group": "三、税收优惠",
        "name": "重点群体个体经营税收扣减",
        "tag": "24000 元/户/年 × 3 年",
        "condition": [
            "■ 扣减标准",
            "· 每户每年 24000 元为限额（国家标准 20000 元，广东上浮 20%），依次扣减增值税、城建税、教育费附加、地方教育附加和个人所得税",
            "· 自办理个体工商户登记当月起 36 个月内享受",
            "■ 申请条件（需同时满足）",
            "· 毕业年度内高校毕业生（毕业所在自然年 1 月 1 日—12 月 31 日）",
            "· 从事个体经营并办理个体工商户登记",
            "· 持《就业创业证》，并注明“毕业年度内自主创业税收政策”",
            "▲ 政策执行至 2027 年 12 月 31 日；当年不足扣减的不得结转以后年度",
        ],
        "materials": [
            "《重点群体或自主就业退役士兵创业信息表》",
            "毕业证，或中国高等教育学历认证报告／国（境）外学历学位认证书",
            "《就业创业证》（深圳已推行社保卡替代，可用电子证照）",
            "个体工商户营业执照",
        ],
        "entry": [
            "■ 电子税务局（申报纳税时填报即享）",
            "· 《就业创业证》向区就业服务中心或街道窗口申领，也可在粤就业、粤省事线上申领",
        ],
        "method": [
            "· 步骤 1：办理个体工商户登记",
            "· 步骤 2：申领《就业创业证》并办理“毕业年度内自主创业税收政策”标注",
            "· 步骤 3：在电子税务局申报时填写《重点群体或自主就业退役士兵创业信息表》",
            "· 步骤 4：系统按限额自动扣减，资料留存备查",
        ],
        "links": [
            "@《关于进一步支持重点群体创业就业有关税收政策的公告》（2023 年第 15 号）：" + L_KEY,
        ],
    },

    # ================= 四、住房类支持 =================
    {
        "group": "四、住房类支持",
        "name": "青年驿站（免费住宿）",
        "tag": "短期免费住",
        "condition": [
            "■ 支持内容",
            "· 短期免费住宿 + 就业指导服务",
            "■ 申请条件（需同时满足）",
            "· 全日制大专及以上学历",
            "· 应届毕业生或毕业 2 年内",
            "· 在深圳市内暂无住所",
            "· 未申请入住过深圳市“青年驿站”，每人只可申请一次",
            "▲ 部分站点只接受毕业未满 1 年的毕业生，申请前先看站点须知",
        ],
        "materials": [
            "身份证",
            "毕业证或学生证（应届）",
            "面试通知、求职证明（部分站点需要）",
        ],
        "entry": [
            "■ 深圳青年驿站 home.szyouth.cn",
            "· 或微信小程序“深圳青年驿站”",
        ],
        "method": [
            "· 步骤 1：在官网或小程序选择站点，查看该站点的毕业年限与床位要求",
            "· 步骤 2：提前 3—15 天在线提交申请",
            "· 步骤 3：站点审核并通知入住",
            "· 步骤 4：按约定时间入住，遵守驿站管理规定",
        ],
    },
    {
        "group": "四、住房类支持",
        "name": "深梦扬帆（15 天免费住宿）",
        "tag": "报名截止 2026-12-31",
        "condition": [
            "■ 支持内容",
            "· 15 天免费住宿，另配租金券",
            "■ 申请条件（需同时满足）",
            "· 2026 届或 2027 届毕业生（大专及以上）",
            "· 来深求职",
            "▲ 报名截止 2026 年 12 月 31 日；报名截止不等于入住截止，截止前成功报名的可在 2027 年入住",
        ],
        "materials": [
            "身份证",
            "毕业证（2026 届）或学籍证明（2027 届）",
            "求职相关证明（部分批次需要）",
        ],
        "entry": [
            "■ 微信小程序（“安居乐寓”或“安居乐租租”，以实际入口为准）",
        ],
        "method": [
            "· 步骤 1：在小程序内找到“深梦扬帆”活动页",
            "· 步骤 2：按届别提交报名，上传学历材料",
            "· 步骤 3：审核通过后获得入住资格",
            "· 步骤 4：按通知办理入住，享受最长 15 天免费住宿",
        ],
    },
    {
        "group": "四、住房类支持",
        "name": "保障性租赁住房（政府组织配租）",
        "tag": "市场租金约 60%",
        "condition": [
            "■ 支持内容",
            "· 租金约为市场参考租金的 60%，建筑面积不超过 70 平方米，合同期限不超过 3 年",
            "· 市场参考租金指市住建部门发布的同地段、同品质住房租赁参考价格",
            "■ 申请条件（需同时满足）",
            "· 在本市无自有住房；未正在享受住房保障优惠政策；正常缴纳社会保险",
            "· 具备人才引进迁户核准条件",
            "■ 什么是“人才引进迁户核准条件”",
            "· 指《深圳市人才引进实施办法》第六条所列条件之一，本科对应的核心一条是：具有普通高等教育本科以上学历，且年龄在 45 周岁以下",
            "· 其他并列条件：专科以上学历且 35 周岁以下；中级专业技术资格 45 周岁以下；高级技师 45 周岁以下等",
            "▲ 按批次配租，需关注配租通告的受理时间",
        ],
        "materials": [
            "身份证、户口本",
            "学历学位证书（用于核验人才引进条件）",
            "社保缴费记录",
            "婚姻及家庭情况材料",
            "无房承诺书",
        ],
        "entry": [
            "■ 深圳市住房和建设局官网 zjj.sz.gov.cn（保障性租赁住房配租通告）",
            "· 各区住建局同步发布本区批次",
        ],
        "method": [
            "· 步骤 1：关注市、区住建局配租通告，确认本批次房源与条件",
            "· 步骤 2：在通告期内在线提交认租申请",
            "· 步骤 3：资格审核、排序",
            "· 步骤 4：选房、签约入住",
        ],
        "links": [
            "@《深圳市人才引进实施办法》（深人社规〔2016〕22 号，第六条即“迁户核准条件”）：" + L_TALENT,
        ],
    },
    {
        "group": "四、住房类支持",
        "name": "保障性租赁住房（社会主体出租）",
        "tag": "≤市场租金 90%，无学历户籍门槛",
        "condition": [
            "■ 支持内容",
            "· 租金不高于市场参考租金的 90%",
            "■ 申请条件（需同时满足）",
            "· 本人及配偶、未成年子女未租住保障性住房",
            "· 房源所在区无自有住房",
            "▲ 不要求深圳户籍、不要求学历、不要求社保，是门槛最低的居住成本削减方式",
        ],
        "materials": [
            "身份证",
            "无房承诺（部分项目需不动产登记查询结果）",
            "婚姻情况材料（如已婚）",
        ],
        "entry": [
            "■ 市场化租赁平台（安居乐寓、泊寓等运营机构）",
            "· 房源信息可在深圳市住房租赁监管服务平台查询",
        ],
        "method": [
            "· 步骤 1：在运营机构平台筛选已纳入保障性租赁住房的房源",
            "· 步骤 2：在线提交承租申请及无房承诺",
            "· 步骤 3：运营机构核验后签约",
            "· 步骤 4：按合同租金入住，租金低于同地段市场价",
        ],
    },
    {
        "group": "四、住房类支持",
        "name": "公共租赁住房",
        "tag": "深户 + 社保累计 3 年",
        "condition": [
            "■ 支持内容",
            "· 租金低于市场价，按轮候顺序配租",
            "■ 申请条件（需同时满足）",
            "· 具有深圳户籍",
            "· 年满 18 周岁",
            "· 参加本市社会保险累计缴费 3 年以上",
            "· 家庭收入、财产符合限额标准",
            "· 在本市无自有住房，且未享受其他住房保障",
        ],
        "materials": [
            "身份证、户口本",
            "社保缴费记录（累计 3 年）",
            "收入、财产申报材料",
            "婚姻及家庭情况材料",
        ],
        "entry": [
            "■ 深圳市住房和建设局官网（公租房轮候库）",
        ],
        "method": [
            "· 步骤 1：在线提交轮候申请，进入轮候库",
            "· 步骤 2：关注认租通告，按批次认租",
            "· 步骤 3：资格复核、排序",
            "· 步骤 4：选房、签约入住",
        ],
    },
    {
        "group": "四、住房类支持",
        "name": "住房公积金租房提取",
        "tag": "最高可提 80%",
        "condition": [
            "■ 支持内容",
            "· 每月提取额不超过本人应缴存额的 65%",
            "· 2025 年 11 月 1 日至 2027 年 10 月 31 日阶段性提高至 80%；多子女家庭 100%",
            "· 已办理租赁备案的，可按实际租金提取",
            "■ 申请条件",
            "· 本人及配偶在本市无自有住房",
            "· 公积金账户正常缴存",
        ],
        "materials": [
            "身份证",
            "本人银行账户（公积金联名卡或指定账户）",
            "无房承诺（系统自动核验，一般无需纸质材料）",
        ],
        "entry": [
            "■ i深圳 App、粤省事、深圳公积金微信公众号",
        ],
        "method": [
            "· 步骤 1：在线签订租房提取协议",
            "· 步骤 2：选择提取比例与到账账户",
            "· 步骤 3：系统核验无房信息",
            "· 步骤 4：每月自动划转，可随时变更或终止协议",
        ],
    },

    # ================= 五、创业类补贴 =================
    {
        "group": "五、创业类补贴",
        "name": "免租金创业空间",
        "tag": "个人 10 ㎡ 免租 6 个月 / 团队 50 ㎡ 免租 1 年",
        "condition": [
            "■ 支持标准",
            "· 个人创业者：10 平方米以内、免租金 6 个月的创业空间",
            "· 创业团队（至少 2 名青年人才）：50 平方米以内、免租金 1 年的创业空间",
            "· 免租期满后，可按约定的优惠租金续租；面积可根据团队人数适当调整",
            "· 申请创新型产业用房的，予以优先支持",
            "■ 申请条件（需同时满足）",
            "· 属于经认定的青年人才（个人申请或团队成员均需符合）",
            "· 在深圳创业，有具体创业项目",
            "· 入驻载体须符合“20+8”产业集群等产业导向",
            "▲ 该措施为 2026 年新增，自 2026 年 1 月 1 日起施行、有效期 5 年",
        ],
        "materials": [
            "身份证",
            "青年人才认定结果",
            "营业执照，或创业项目说明（尚未注册时可先提交项目计划）",
            "团队名单及成员认定结果（团队申请时）",
            "入驻申请表（载体提供模板）",
        ],
        "entry": [
            "■ i深圳 App → 青年人才服务保障专区 → 创业支持",
            "· 各区创新型产业用房管理机构、认定的创业孵化载体",
        ],
        "method": [
            "· 步骤 1：先完成青年人才认定",
            "· 步骤 2：在专区提交创业空间申请及项目材料",
            "· 步骤 3：载体运营机构匹配房源、审核资格",
            "· 步骤 4：签约入驻，免租期内不缴租金",
            "· 步骤 5：期满前申请续租，按约定优惠租金执行",
        ],
        "links": [
            "@《关于加强青年人才来深发展服务保障的若干措施》政策解读（创业支持 6 条）：" + L_LB_JD,
            "@《若干措施》原文：" + L_LB,
        ],
    },
    {
        "group": "五、创业类补贴",
        "name": "初创企业补贴",
        "tag": "1 万元（合伙合计 ≤10 万）",
        "condition": [
            "■ 补贴标准",
            "· 每人 1 万元，合伙经营合计不超过 10 万元；每人只能享受一次",
            "■ 申请条件（需同时满足）",
            "· 毕业 5 年内的高校毕业生",
            "· 企业注册 3 年内",
            "· 正常经营满 6 个月",
            "· 本人连续缴纳社会保险 6 个月",
            "▲ 所创办企业须未享受过创业扶持补贴；分公司不得申请",
        ],
        "materials": [
            "身份证、毕业证",
            "营业执照",
            "经营场所证明（租赁合同或产权证明）",
            "社保缴费记录（连续 6 个月）",
            "纳税记录或经营流水",
            "本人银行账户",
        ],
        "entry": [
            "■ 广东政务服务网、广东公共就业服务云平台",
            "· 或企业注册地街道公共就业服务机构",
        ],
        "method": [
            "· 步骤 1：确认注册时间、经营时长、社保月数达标",
            "· 步骤 2：在线或窗口提交申请",
            "· 步骤 3：街道受理、区人社部门审核",
            "· 步骤 4：公示后拨付",
        ],
    },
    {
        "group": "五、创业类补贴",
        "name": "创业场租补贴",
        "tag": "大学生上浮 30%",
        "condition": [
            "■ 补贴标准",
            "· 认定载体：1200／1000／700 元/月；载体外不超过 500 元/月（年不超过 6000 元）",
            "· 高校毕业生上浮 30%，即 1560／1300／910 元/月；载体外不超过 650 元/月（年不超过 7800 元）",
            "· 累计不超过 36 个月",
            "■ 申请条件（需同时满足）",
            "· 毕业 5 年内高校毕业生",
            "· 实际缴纳租金满 3 个月",
            "· 本人连续缴纳社会保险满 3 个月",
            "· 租赁场地与注册地址一致",
        ],
        "materials": [
            "身份证、毕业证",
            "营业执照",
            "场地租赁合同及租金发票（或收据）",
            "社保缴费记录（连续 3 个月）",
            "经营场所照片",
            "本人银行账户",
        ],
        "entry": [
            "■ 广东政务服务网、广东公共就业服务云平台",
        ],
        "method": [
            "· 步骤 1：确认租赁地与注册地一致、已缴租满 3 个月",
            "· 步骤 2：在线提交租赁合同、发票等材料",
            "· 步骤 3：街道受理并现场核查",
            "· 步骤 4：审核通过后按季度或按月拨付",
        ],
    },
    {
        "group": "五、创业类补贴",
        "name": "创业社会保险补贴",
        "tag": "单位部分 100%，≤36 个月",
        "condition": [
            "■ 补贴标准",
            "· 按最低缴费标准计算的单位缴纳部分给予 100% 补贴，累计不超过 36 个月",
            "■ 申请条件（需同时满足）",
            "· 毕业 5 年内高校毕业生创办企业",
            "· 本人及合伙人在该企业缴纳社会保险",
            "▲ 合伙人须一并申请，受理后不得增加人员",
        ],
        "materials": [
            "身份证、毕业证",
            "营业执照",
            "社保缴费记录（单位部分明细）",
            "合伙人名单及身份材料",
        ],
        "entry": [
            "■ 广东政务服务网、广东公共就业服务云平台",
        ],
        "method": [
            "· 步骤 1：确认企业及合伙人参保情况",
            "· 步骤 2：一次性提交全部合伙人材料",
            "· 步骤 3：区人社部门审核",
            "· 步骤 4：按月或按季补贴至企业账户",
        ],
    },
    {
        "group": "五、创业类补贴",
        "name": "创业带动就业补贴",
        "tag": "最高 3 万元",
        "condition": [
            "■ 补贴标准",
            "· 招用 3 人及以下，按每人 2000 元补贴；4 人及以上，每净增 1 人按 3000 元补贴",
            "· 总额不超过 3 万元",
            "■ 申请条件（需同时满足）",
            "· 毕业 5 年内高校毕业生创办企业",
            "· 与员工签订 1 年以上劳动合同",
            "· 为员工连续缴纳社会保险满 6 个月",
            "▲ 不得与一次性扩岗补助、一次性吸纳就业补贴叠加享受",
        ],
        "materials": [
            "身份证、毕业证",
            "营业执照",
            "员工劳动合同",
            "员工社保缴费记录（连续 6 个月）",
            "员工花名册",
        ],
        "entry": [
            "■ 广东政务服务网、广东公共就业服务云平台",
        ],
        "method": [
            "· 步骤 1：统计符合条件员工人数与净增人数",
            "· 步骤 2：在线提交劳动合同与社保记录",
            "· 步骤 3：区人社部门审核、公示",
            "· 步骤 4：拨付至企业账户",
        ],
    },
    {
        "group": "五、创业类补贴",
        "name": "创业担保贷款及贴息",
        "tag": "个人最高 30 万，贴息 50%",
        "condition": [
            "■ 支持标准",
            "· 个人最高 30 万元；带动就业的可提高至 60 万元；合伙经营合计不超过 500 万元",
            "· 小微企业最高 500 万元",
            "· 利率不超过 LPR+50BP，财政贴息 50%（LPR 即贷款市场报价利率，每月 20 日由全国银行间同业拆借中心公布）",
            "■ 申请条件（需同时满足）",
            "· 在深圳注册经营",
            "· 有具体经营项目",
            "· 除房贷、车贷等以外无其他贷款",
            "· 累计申请不超过 3 次",
        ],
        "materials": [
            "身份证、户口本或居住证",
            "营业执照",
            "经营场所证明",
            "创业项目说明或经营计划",
            "征信及负债情况材料",
            "担保或抵押材料（按银行要求）",
        ],
        "entry": [
            "■ 经办银行网点（深圳农商行、中国银行等合作银行）",
            "· 创业担保贷款担保基金运营机构",
        ],
        "method": [
            "· 步骤 1：向合作银行提交贷款申请及创业材料",
            "· 步骤 2：担保机构审核、出具担保",
            "· 步骤 3：银行审批放款",
            "· 步骤 4：按期还款后申请财政贴息",
        ],
    },

    # ================= 六、区级入户补贴 =================
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "福田区",
        "tag": "本科 2 万 / 双学位 3 万",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在申报时最近一次 QS／THE／ARWU／U.S.News 四大榜单中任一榜单排名前 150",
            "· 引进时间须为 2025 年 5 月 1 日起",
            "· 在福田辖区企业或机构全职工作，并连续缴纳社会保险满 12 个月",
            "· 双学位档要求两个学位均符合院校排名要求，且为同一学历层次",
            "▲ 指南有效期至 2026 年 12 月 31 日，按年度更新；项目受年度财政预算总量控制",
        ],
        "materials": DIST_MAT + ["双学位另附两个学士学位证书"],
        "entry": ["■ 福田区企业服务平台 qfzx.szft.gov.cn（由用人单位账号申报）"],
        "method": [
            "· 步骤 1：用人单位确认院校排名与社保月数符合条件",
            "· 步骤 2：单位在福田区企业服务平台填报，汇总纸质材料",
            "· 步骤 3：全年受理、集中审批，每季度审核上一季度申报",
            "· 步骤 4：拟发放名单在“福田政府在线”公示 5 个工作日",
            "· 步骤 5：公示无异议后一次性拨付至个人账户（税前金额）",
        ],
        "links": ["@福田区新引进优秀青年入户补贴申请指南：" + L_FT],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "南山区",
        "tag": "本科 2 万 / 双学位 3 万",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 150，或为教育部“双一流”建设高校及建设学科",
            "· 引进时间须为 2025 年 3 月 25 日及之后",
            "· 分两次发放：在申报单位连续缴社保及申报工资薪金个税满 3 个月发 30%，满 12 个月发 70%；满 12 个月也可一次性申领",
            "· 双学位指同一学历层次的两个独立学位证书",
            "▲ 申报期为 2025 年 3 月 25 日至 2027 年 3 月 25 日；首发与续发均须在引进之日起 24 个月内、由同一申报单位提出",
        ],
        "materials": DIST_MAT + [
            "工资薪金个税申报记录（续发需体现满 12 个月）",
            "双学位另附两个学士学位证书",
        ],
        "entry": ["■ i南山企业服务平台（由用人单位申报）"],
        "method": [
            "· 步骤 1：单位核对院校名单（四大榜单前 150 或双一流）",
            "· 步骤 2：单位在线填报并上传材料",
            "· 步骤 3：区人力资源局每月集中受理、审核",
            "· 步骤 4：满 3 个月领首发 30%，满 12 个月领续发 70%",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "龙华区",
        "tag": "本科 3 万（一次性）",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 150，或国内“双一流”高校及学科",
            "· 毕业 2 年内视为应届",
            "· 已入户龙华，与龙华区企业签订 3 年以上劳动合同",
            "· 社保与个税连续缴纳 3 个月以上",
            "▲ 每年 5 月 1—31 日集中受理，错过等一年；补贴为税前金额，发放前代扣代缴个税",
        ],
        "materials": DIST_MAT + ["个人所得税纳税记录（连续 3 个月）"],
        "entry": [
            "■ 龙华区人才服务系统（个人线上申报）",
            "· 也可从 i深圳“青年人才服务保障专区”进入",
        ],
        "method": [
            "· 步骤 1：确认入户龙华、合同 3 年以上、社保满 3 个月",
            "· 步骤 2：每年 5 月 1—31 日，个人在线填报并上传材料",
            "· 步骤 3：区人力资源部门审核",
            "· 步骤 4：审核通过后一次性发放 3 万元",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "龙岗区",
        "tag": "本科 2 万 / 双学位 3 万，另有岗位奖励 1 万",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 150，或“双一流”高校及学科",
            "· 已取得龙岗区户籍",
            "· 分两次发放：首发需连续缴社保满 1 年，续发需满 2 年",
            "■ 岗位奖励（另 1 万元）",
            "· 已全额领取（或符合全额领取）入户补贴",
            "· 自入户补贴续发申请之日起，在该单位连续全职工作满 1 年，且距续发申请之日不超过 2 年",
        ],
        "materials": DIST_MAT + ["岗位奖励另需：续发申请回执、连续在职证明"],
        "entry": ["■ 龙岗区人力资源局（由用人单位申报）"],
        "method": [
            "· 步骤 1：单位核对院校、户籍与社保年限",
            "· 步骤 2：单位在线填报，提交至区人力资源局",
            "· 步骤 3：满 1 年领首发，满 2 年领续发",
            "· 步骤 4：满足在职年限后，另行申请 1 万元岗位奖励",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "宝安区",
        "tag": "本科 1.5 万",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 300",
            "· 已取得宝安区户籍",
            "· 本科未满 30 周岁",
            "· 在宝安区用人单位连续缴纳社会保险 12 个月以上",
            "· 引进时间须为 2022 年 8 月 24 日及以后",
        ],
        "materials": DIST_MAT,
        "entry": ["■ 宝安亲清政企服务直达平台"],
        "method": [
            "· 步骤 1：确认院校前 300、宝安户籍、年龄与社保月数",
            "· 步骤 2：在亲清政企平台注册并填报",
            "· 步骤 3：区人力资源局审核、公示",
            "· 步骤 4：拨付至个人账户（税前，代扣代缴个税）",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "坪山区",
        "tag": "本科 3 万（分两次）",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 150",
            "· 已取得坪山区户籍",
            "· 全职在坪山工作满 3 个月",
            "· 所在单位属于七个产业方向之一：软件与信息服务、新一代电子信息、新能源和智能网联汽车、生物医药和医疗器械、人工智能、高端装备与仪器、低空经济与空天",
            "· 分两次发放（如 1 万 + 2 万），第二次须首次资助满 12 个月",
        ],
        "materials": DIST_MAT + ["单位所属产业方向的证明（营业执照经营范围、园区证明或企业认定文件）"],
        "entry": ["■ 坪山区人力资源局（由企业推荐申报）"],
        "method": [
            "· 步骤 1：确认单位属于七个产业方向、本人院校前 150",
            "· 步骤 2：由企业推荐并在区申报系统提交",
            "· 步骤 3：区人力资源局审核、公示",
            "· 步骤 4：首次资助到账满 12 个月后，按批次申请第二次",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "光明区",
        "tag": "本科 3 万（生活 + 就业合计）",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校为四大榜单任一排名前 300，或“双一流”高校及学科，或辖区高校",
            "· 已落户光明",
            "· 连续缴纳社会保险 6 个月",
            "· 须在首份人才引进审批文件签发之日起 9 个月内申报",
            "▲ 2026 年 9 月 1 日新版指南，标准较上一版提高",
        ],
        "materials": DIST_MAT,
        "entry": ["■ 光明区人才服务系统（个人线上申报）"],
        "method": [
            "· 步骤 1：核对院校口径（前 300／双一流／辖区高校）",
            "· 步骤 2：个人在线填报，上传材料",
            "· 步骤 3：区人力资源局审核、公示",
            "· 步骤 4：拨付，生活补贴与就业补贴合计 3 万元",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "大鹏新区",
        "tag": "本科 2 万",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校为四大榜单任一排名前 300，或 211／985／“双一流”",
            "· 已落户大鹏新区",
            "· 连续缴纳社会保险 12 个月",
            "· 引进后 2 年内申请",
            "▲ 政策期至 2026 年 12 月 31 日，到期后是否续期待定",
        ],
        "materials": DIST_MAT,
        "entry": ["■ 大鹏新区组织人事局（以区指南公布的受理窗口为准）"],
        "method": [
            "· 步骤 1：确认院校前 300、落户大鹏、社保满 12 个月",
            "· 步骤 2：按区指南在线或窗口提交",
            "· 步骤 3：审核、公示",
            "· 步骤 4：拨付 2 万元",
        ],
    },
    {
        "group": "六、区级新引进本科入户补贴",
        "name": "罗湖区",
        "tag": "本科最高 3 万（系统时有维护）",
        "condition": DIST_BASE + [
            "■ 本区专属条件",
            "· 毕业院校在四大榜单任一排名前 150，或内地 985 高校",
            "· 已落户罗湖",
            "· 连续缴纳社会保险 6 个月",
            "· 面向罗湖“3+3+4”重点产业青年人才",
            "▲ 申报金额为分档上限；区人才奖励申报系统 jyta.szlh.org.cn 时有维护（2026-11-22 核验首页可打开），申报前先电话确认通道",
        ],
        "materials": DIST_MAT,
        "entry": [
            "■ 罗湖区人才奖励申报系统 jyta.szlh.org.cn（维护中）",
            "· 咨询电话 0755-25438390",
        ],
        "method": [
            "· 步骤 1：先电话确认系统是否恢复、当前批次是否开放",
            "· 步骤 2：系统开放后在线填报",
            "· 步骤 3：区人力资源局审核、公示",
            "· 步骤 4：按分档标准拨付，最高 3 万元",
        ],
    },

    # ================= 七、交通与日常礼遇 =================
    {
        "group": "七、生活优惠",
        "name": "公共交通优惠（深圳通）",
        "tag": "地铁 9.5 折、公交最低 6.5 折",
        "condition": [
            "■ 常规折扣（仅限深圳通实体卡、NFC 手机深圳通）",
            "· 地铁、有轨电车、坪山云巴：9.5 折",
            "· 常规公交：3 元及以下 8 折；3—6 元的部分 7.5 折；6 元以上的部分 6.5 折",
            "· 90 分钟内换乘：在刷卡优惠基础上再减 0.4 元/人次",
            "■ 深圳通乘车码（微信、支付宝、深圳通 App）",
            "· 公交同享 8／7.5／6.5 折",
            "· 地铁与单程票同价，没有 9.5 折；换乘优惠同样是 0.4 元/人次，但“公交换乘地铁”暂不享受",
            "■ 不定期全市活动",
            "· 2026 年 11 月 16 日—20 日，深圳地铁、公交所有计程类票种五折，无需领券、过闸自动打折（活动已结束）",
            "· 此类活动由深圳地铁、深圳通不定期推出，关注“深圳地铁”“深圳通”公众号可第一时间获知",
            "▲ 日常通勤想省钱：地铁用实体卡或 NFC 手机深圳通享 9.5 折，只用乘车码会少这一档",
        ],
        "materials": [
            "深圳通实体卡（押金卡押金可退，或纪念卡），或在手机钱包开通 NFC 手机深圳通（开卡免费）",
            "乘车码（微信、支付宝或深圳通 App，无需实体卡但地铁无折扣）",
        ],
        "entry": [
            "■ 深圳通客服网点、各地铁站客服中心",
            "■ “深圳通”App、微信“深圳通+”小程序",
        ],
        "method": [
            "· 步骤 1：办理深圳通实体卡，或在手机钱包开通 NFC 手机深圳通",
            "· 步骤 2：地铁刷卡享 9.5 折，公交按票价区间享阶梯折扣",
            "· 步骤 3：90 分钟内换乘自动减 0.4 元/人次，无需额外操作",
            "· 步骤 4：关注“深圳地铁”“深圳通”公众号，遇全市活动直接刷卡即享",
        ],
        "links": [
            "@深圳通乘车优惠官方说明（折扣规则原文）：" + L_SZT,
        ],
    },
    {
        "group": "七、生活优惠",
        "name": "青年人才来深礼包",
        "tag": "交通、观光、生活便利",
        "newpage": True,
        "condition": [
            "■ 礼包内容",
            "· 包含交通资讯、城市观光、生活便利等服务礼遇，具体清单以 i深圳 专区当期展示为准",
            "· 与“秒批”入户、新引进优秀青年人才入户补贴、大学生就业专区服务、往来港澳人才签注同属“生活服务”5 条措施",
            "■ 申请条件",
            "· 已通过青年人才认定",
            "· 须在认定之日起 1 年内申请，与过渡性住房、安居补贴同一时限，逾期视为放弃",
            "▲ 礼包为服务类权益，不是现金；不同批次内容可能调整",
        ],
        "materials": [
            "身份证",
            "青年人才认定结果",
            "i深圳实名账号",
        ],
        "entry": [
            "■ i深圳 App → 青年人才服务保障专区 → 生活服务 → 来深礼包",
        ],
        "method": [
            "· 步骤 1：完成青年人才认定",
            "· 步骤 2：在专区进入“来深礼包”，查看当期权益清单",
            "· 步骤 3：按清单逐项领取或预约服务",
            "· 步骤 4：注意各项权益的有效期，过期不补",
        ],
        "links": [
            "@《若干措施》政策解读（生活服务 5 条）：" + L_LB_JD,
            "@《若干措施》原文：" + L_LB,
        ],
    },

    {
        "group": "七、生活优惠",
        "name": "公共体育场馆免费开放",
        "tag": "分时段免费，i深圳预约",
        "newpage": True,
        "condition": [
            "■ 支持内容",
            "· 全市公共体育场馆分时段免费或低收费开放，各区、各场馆时段不同",
            "· 以宝安区航城街道为例：领航文体中心室外一、二号综合运动场每日 7:00—22:00 免费；三、四号综合运动场及空中篮球场工作日 7:00—15:00 免费；室内综合馆（羽毛球、匹克球、三人篮球）工作日 7:00—15:00 免费",
            "· 狮子山公园足球场每日 8:30—22:30 免费；部分场地工作日 8:30—16:30 免费",
            "■ 申请条件",
            "· 面向全体市民，不限户籍、学历、社保",
            "· 多数场馆需提前线上预约，少数场馆现场登记即可使用",
            "▲ 各场馆免费时段与预约方式不同，以预约页当日信息为准",
        ],
        "materials": [
            "身份证（部分场馆现场核验）",
            "i深圳账号（线上预约）",
        ],
        "entry": [
            "■ i深圳 App → 体育场馆一键预约",
            "· 各场馆微信公众号（如“领航文体中心”）",
        ],
        "method": [
            "· 步骤 1：在 i深圳搜索“体育场馆”，进入预约平台",
            "· 步骤 2：选择场馆、日期与免费时段",
            "· 步骤 3：提交预约，按时到场扫码核销",
            "· 步骤 4：无需预约的场馆（如部分社区健身房）现场登记即可使用",
        ],
        "links": [
            "@宝安区公共体育场馆及预约方式官方问答（免费时段示例）：" + L_SPORT,
        ],
    },
]

# ---- 把“来深礼包”并入市级人才支持，置于过渡性住房之后 ----
_lb = [x for x in ITEMS if x["name"] == "青年人才来深礼包"][0]
ITEMS.remove(_lb)
_lb["group"] = "一、市级人才支持"
_lb.pop("newpage", None)
_i = next(i for i, x in enumerate(ITEMS) if x["name"] == "青年人才过渡性住房")
ITEMS.insert(_i + 1, _lb)

# ---------------- 生成目录 ----------------
NUMBERED = [(i, it) for i, it in enumerate(ITEMS, 1)]
_seen = []
for _, it in NUMBERED:
    if it["group"] not in _seen:
        _seen.append(it["group"])
for g in _seen:
    gp = doc.add_paragraph()
    gp.paragraph_format.space_before = Pt(4)
    gp.paragraph_format.space_after = Pt(0)
    gr = gp.add_run(g)
    cn(gr, HEAD_FONT)
    gr.font.size = Pt(9.5)
    gr.font.bold = True
    gr.font.color.rgb = BLUE
    toc_slot._p.addprevious(gp._p)
    for num, it in [(n, x) for n, x in NUMBERED if x["group"] == g]:
        ip = doc.add_paragraph()
        ip.paragraph_format.left_indent = Cm(0.7)
        ip.paragraph_format.space_after = Pt(0)
        ip.paragraph_format.line_spacing = 0.9
        r1 = ip.add_run("%02d　" % num)
        cn(r1, BODY_FONT)
        r1.font.size = Pt(8)
        r1.font.color.rgb = RGBColor(0x9A, 0xB0, 0xCC)
        r2 = ip.add_run(it["name"])
        cn(r2, BODY_FONT)
        r2.font.size = Pt(8)
        r2.font.color.rgb = DARK
        toc_slot._p.addprevious(ip._p)
toc_slot._p.getparent().remove(toc_slot._p)

# ---------------- 正文 ----------------
cur = None
for i, it in enumerate(ITEMS, 1):
    if it["group"] != cur:
        cur = it["group"]
        group_h(cur)
    item_block(it, "%02d" % i)

doc.save(OUT)
print("saved:", OUT, "items:", len(ITEMS))

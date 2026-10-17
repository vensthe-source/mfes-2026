# -*- coding: utf-8 -*-
"""MFES-2026 发布包生成：白皮书、评分手册、Excel 模板、案例 CSV。"""
import csv
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_boards import (AUTHOR, AFFILIATION, SYS_NAME, VERSION, BOARDS, TRACK_MAX,
                         DEFAULT_COEF, USE_PRESETS, GRADES, TEMPLATES, EXTERNAL_ARCHIVE,
                         CASE_ROWS)

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(OUT, exist_ok=True)

def set_run(run, east="宋体", ascii_font="Times New Roman", size=12, bold=False):
    run.font.name = ascii_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    run.font.size = Pt(size)
    run.font.bold = bold

def para(doc, text, east="宋体", size=12, bold=False, align="justify", indent=True, after=6, line=1.4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, east=east, size=size, bold=bold)
    p.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.85)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    return p

def heading(doc, text, level=1, center=False):
    sizes = {0: 18, 1: 14, 2: 12.5}
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, east="黑体", ascii_font="Arial", size=sizes[level], bold=True)
    if center or level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.3
    return p

def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.paragraphs[0].text = ""
        r = c.paragraphs[0].add_run(h)
        set_run(r, east="黑体", ascii_font="Arial", size=10.5, bold=True)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].paragraphs[0].text = ""
            r = cells[i].paragraphs[0].add_run(str(v))
            set_run(r, east="宋体", size=10.5)
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.5)
    sec.left_margin = sec.right_margin = Cm(2.6)
    return doc

def cover(doc, subtitle):
    heading(doc, "MFES-2026 电影多维评价体系", 0)
    heading(doc, subtitle, 0)
    para(doc, "版本 %s · 2026 年 10 月" % VERSION, align="center", indent=False)
    para(doc, "作者：%s（%s）" % (AUTHOR, AFFILIATION), align="center", indent=False, east="楷体")
    para(doc, "许可证：CC BY 4.0（署名 4.0 国际许可）", align="center", indent=False)
    para(doc, "本文件及配套评分手册、数据与工具构成 MFES-2026 的公开版本。", align="center", indent=False)

# ================= 白皮书 =================
doc = new_doc()
cover(doc, "白皮书")
heading(doc, "一、定位与目的", 1)
para(doc, "MFES-2026 是一套面向电影的多维加权评价体系，旨在为影片价值评价提供公开、可审计、可复算的工具。体系的定位是学术研究与行业参考工具：它为研究者提供跨影片比较的共同量表，为评奖、资助与社会效益考核提供可公开的评分细则草案，为国际传播选片提供测量基础。需要声明：v1.0 是研究者个人提出的学术方案，不构成任何官方标准；其全部规则公开，欢迎复核、质疑与修订建议。")
heading(doc, "二、体系总纲：一体两翼", 1)
para(doc, "体系由三轨构成：本体翼回答“电影本身好不好”，面向一切国别与时代的影片，是体系的科学内核；价值翼回答“是否符合中国特色社会主义发展观”，体现方向性；文明翼回答“是否回应人类命运共同体”，体现世界性。三轨独立评分、独立报告、按用途加权。综合加权总分定义为：")
para(doc, "S = 1000 × [ α·(B/500) + β·(V/300) + γ·(C/200) ]", align="center", indent=False, bold=True)
para(doc, "其中 B、V、C 分别为本体翼、价值翼、文明翼原始得分（满分 500、300、200），α、β、γ 为用途预设系数。默认综合用途 α=0.50、β=0.30、γ=0.20，此时 S = B + V + C，满分 1000 分。等级锚定：900–1000 传世之作；800–899 优秀；700–799 良好；600–699 合格；低于 600 一般。")
heading(doc, "三、权重分配", 1)
table(doc, ["轨道", "权重（默认）", "满分", "功能定位"],
      [["本体翼（九板块）", "50%", "500", "科学性：通用、客观、可跨国比较"],
       ["价值翼（八板块）", "30%", "300", "方向性：中国特色社会主义发展观"],
       ["文明翼（四板块）", "20%", "200", "世界性：人类命运共同体"]],
      widths=[3.6, 2.8, 1.8, 7.4])
table(doc, ["板块", "满分", "板块", "满分"],
      [("本体翼·剧作文本", 75, "价值翼·人民性", 75),
       ("本体翼·艺术美学", 75, "价值翼·时代性", 45),
       ("本体翼·观众接受", 75, "价值翼·文化自信", 45),
       ("本体翼·技术工艺", 60, "价值翼·创新性", 45),
       ("本体翼·历史定位", 50, "价值翼·协调性", 30),
       ("本体翼·思想哲学", 50, "价值翼·开放", 30),
       ("本体翼·社会文化", 50, "价值翼·共享", 15),
       ("本体翼·市场产业", 40, "价值翼·绿色", 15),
       ("本体翼·政治维度", 25, "文明翼·人类关怀", 60),
       ("文明翼·跨文化可读性", 60, "文明翼·文明互鉴", 50),
       ("文明翼·全球责任", 30, "（合计）", 1000)],
      widths=[4.4, 1.8, 4.6, 1.8])
para(doc, "权重由德尔菲法专家征询与层次分析法（AHP）确定（一致性比率 CR＜0.1）。设计取向：观众接受与剧作、艺术并列为最高权重；政治维度压低至 25 分以防止本体评价政治化；人民性 75 分落实“以人民为中心”的创作导向。")
heading(doc, "四、五类动态机制", 1)
para(doc, "其一，类型模板：不同片种在本体翼内再分配权重（调整合计为零）。其二，N/A 再分配：不适用板块不评分，其分值按比例再分配给同轨其余板块，报告披露“适用覆盖率”。其三，时代双轨：默认口径为语境相对分（与同时代同类型比较），另记绝对分供影史研究；上映不足五年按“潜质”评历史定位，不足一年整板块 N/A。其四，用途预设：学术研究 α=0.80/0.10/0.10，综合默认 0.50/0.30/0.20，国际传播选片 0.40/0.10/0.50。其五，本体原则与外部因素档案：排片、档期、票价、营销、放映设备、公益放映安排、获奖公关、审查环境等影片制作方不可控因素，只记录、不评分。")
heading(doc, "五、评分规则总则", 1)
para(doc, "全部指标采用“全量程五级锚定制”。每项指标给出 0–10 全域的五级锚定标准：10 分＝标杆（同类参照系顶级）；7.5 分＝优秀；5 分＝合格（同类平均）；2.5 分＝不足；0 分＝严重缺陷（结构性失败或违背规范）。评分程序两步：第一步对照锚点定档（允许落档间，步长 0.5）；第二步对照微调点清单浮动——每命中一处缺陷 −0.5（合计 −1 封顶）、每命中一处加分点 +0.5（合计 +1 封顶）。最终得分 = 档位 ± 微调，恒在 [0,10]，不存在超出满分。标注 −2 及以上或“结构性失败/底线联动”的扣分点直接压入低档。")
para(doc, "配套规则三条。其一，中性处理：指标对象不存在时（未上映、未参与合拍、无政治表达、未涉全球议题、未国际发行等），该项按中性 5 分或 N/A 再分配处理——总原则是不得以不存在之事扣分。其二，证据背书：每个定档依据与微调点必须附可核查证据（镜头时间码、数据来源、文献出处），无证据不得动分。其三，取整：板块得分 = Σ（指标得分/10 × 指标满分），结果四舍五入取整。负向指标（避免狭隘民族主义、避免自我东方化）默认满分，出现负向表现按细则扣减并举证。")
heading(doc, "六、外部因素档案", 1)
table(doc, ["外部因素", "记录内容", "用途"],
      EXTERNAL_ARCHIVE, widths=[2.8, 7.4, 5.4])
heading(doc, "七、主体性控制与质量控制", 1)
para(doc, "四层控制：事实层（指标化与证据背书）、判断层（锚定样例、标定测试、信度、仲裁）、偏见层（利益申报、盲评、多角化评审主体、重测稳定性）、立场层（权重与细则全文公开、版本化修订公示）。质量指标：每部影片至少 3 名经标定的评分员独立评分；组内相关系数 ICC≥0.75 或 Cohen's κ≥0.61 为合格；分差超过 2.0 分的指标进入仲裁；标定测试偏差须小于 1.5 分。文明翼评分必须包含至少 1 名非中文母语评分员。")
heading(doc, "八、四部影片试评", 1)
para(doc, "按上述规则对《霸王别姬》（1993）、《给阿嬷的情书》（2026）、《流浪地球2》（2023）、《功夫女足》（2026）四部影片试评，板块级得分如下（新片历史定位按 N/A 再分配，《功夫女足》全球责任按 N/A 处理）：")
table(doc, ["影片", "本体(500)", "价值(300)", "文明(200)", "总分", "等级", "豆瓣参照"],
      [list(r) for r in CASE_ROWS], widths=[4.6, 1.8, 1.8, 1.8, 1.6, 1.5, 1.7])
para(doc, "三点发现：其一，四片得分与公开口碑完全同序（856＞819＞809＞560 对 9.6＞9.0＞8.3＞6.5），构成效标效度的初步证据；其二，《功夫女足》票房约 17 亿元、市场板块得分不低，但总分落入“一般”档——体系能把“品质”与“票房”分开诊断；其三，老片的价值翼按当代标准评存在系统性偏低（《霸王别姬》绿色与数字资产复用两项因时代所限损失 13 分），提示后续版本需对老片启用价值翼语境分。四例试评为体系可操作性的演示，不构成体系有效性的证明。")
heading(doc, "九、限度声明", 1)
para(doc, "本体系公开声明六项限度：其一，价值判断不可消除——权重与锚点的选择本身是价值承诺，体系只保证其可见与可挑战；其二，千分制总分必须与三轨剖面、置信区间同报，禁止脱离剖面引用总分；其三，锚定样例天然偏向已有范式，先锋作品须以“形式即内容”规则缓冲；其四，底线准入仅适用于评奖与资助序列，学术用途不设底线；其五，语境相对分牺牲跨时代通约，体系不用于“有史以来最佳影片”式排名；其六，个案试评不构成有效性证明，有效性有待大样本信度研究检验。")
heading(doc, "十、版本与维护", 1)
para(doc, "本文件为 v1.0（2026 年 10 月发布）。版本管理规则：每五年全面修订一次，重大修订须公示变更清单与理由；评分手册、权重表、试评数据与复算工具随版本同步更新。反馈与挑战请通过 GitHub 仓库 issue 提交（仓库地址见 README）。")
heading(doc, "十一、引用格式", 1)
para(doc, "%s. MFES-2026 电影多维评价体系：白皮书（%s）[R]. 北京：%s，2026." % (AUTHOR, VERSION, AFFILIATION), indent=False)
doc.save(os.path.join(OUT, "白皮书_MFES-2026_v1.0.docx"))
print("saved 白皮书")

# ================= 评分手册 =================
doc = new_doc()
cover(doc, "评分手册")
heading(doc, "一、使用说明", 1)
para(doc, "评分流程五步：①确认类型模板与用途预设，计算各板块调整后满分（N/A 板块按比例再分配）；②逐项指标对照五级锚点定档（0/2.5/5/7.5/10，允许落档间）；③对照微调点清单浮动（每处 ±0.5，合计 ±1 封顶，−2 及以上直接压档）；④每个定档依据与浮动必须附证据（时间码、数据、文献），无证据不得动分；⑤汇总：板块得分 = Σ（指标得分/10 × 指标满分）四舍五入取整，再按用途预设系数计算加权总分 S。")
heading(doc, "二、通用规则", 1)
para(doc, "中性处理：指标对象不存在（未上映、未参与合拍、无政治表达、未涉全球议题、未国际发行、无特效、无新技术运用等）时，按中性 5 分或 N/A 再分配处理，不得以不存在之事扣分。负向指标：默认满分，出现负向表现按细则扣减，须附时间码与情节描述。底线联动：标注“底线联动”的扣分项触发时，同步记入底线核查（史实与历史虚无主义、公序良俗、违法违规），不通过者不得进入评奖序列。")

seq = 0
for track, board, bmax, inds in BOARDS:
    seq += 1
    heading(doc, "%d.%d %s（%s，%d 分）" % (2 + (seq - 1) // 20, seq % 20 if seq % 20 else 20, board, track, bmax), 2)
    table(doc, ["指标", "满分", "主要观测点"],
          [[n, m, o] for n, m, o, r in inds], widths=[3.2, 1.4, 11.0])
    table(doc, ["指标", "满分", "全量程锚定（0–10）与微调点"],
          [[n, m, r] for n, m, o, r in inds], widths=[3.2, 1.4, 11.0])

heading(doc, "三、外部因素档案", 1)
table(doc, ["外部因素", "记录内容", "用途"],
      EXTERNAL_ARCHIVE, widths=[2.8, 7.4, 5.4])

heading(doc, "四、空白评分表", 1)
for tno, (track, boards) in enumerate([("本体翼", [b for t, b, m, i in BOARDS if t == "本体翼"]),
                                       ("价值翼", [b for t, b, m, i in BOARDS if t == "价值翼"]),
                                       ("文明翼", [b for t, b, m, i in BOARDS if t == "文明翼"])]):
    rows = [[b, "—", "", "", ""] for b in boards]
    rows.append(["%s合计" % track, "—", TRACK_MAX[track], "", "（各项得分相加）"])
    table(doc, ["板块", "指标（满分）", "调整后满分", "得分", "证据与备注"], rows,
          widths=[2.6, 6.6, 1.9, 1.4, 3.1])
table(doc, ["项目", "记录"],
      [["本体翼 B（满分 500）", ""],
       ["价值翼 V（满分 300）", ""],
       ["文明翼 C（满分 200）", ""],
       ["加权总分 S = 1000×[α·(B/500)+β·(V/300)+γ·(C/200)]", "用途预设 α/β/γ：□综合 0.50/0.30/0.20　□学术 0.80/0.10/0.10　□国际传播 0.40/0.10/0.50"],
       ["等级", "□传世之作（≥900）　□优秀（800–899）　□良好（700–799）　□合格（600–699）　□一般（＜600）"],
       ["信度检验", "ICC＝　　　κ＝　　　仲裁项数："],
       ["适用覆盖率与类型模板", "覆盖率　　　%　模板：□剧情/文艺　□动作/科幻　□喜剧/合家欢　□纪录片　□其他"],
       ["口径与签名", "口径：□语境分　□绝对分　评分员签名：　　　　日期："]],
      widths=[7.6, 8.0])
doc.save(os.path.join(OUT, "评分手册_MFES-2026_v1.0.docx"))
print("saved 评分手册")

# ================= Excel 模板 =================
wb = Workbook()
ws = wb.active
ws.title = "使用说明"
notes = [
    "MFES-2026 空白评分表模板（%s）　作者：%s（%s）　许可：CC BY 4.0" % (VERSION, AUTHOR, AFFILIATION),
    "",
    "使用方法：",
    "1. 在“评分表”工作表中为每项指标填入 0–10 得分（对照评分手册五级锚点与微调点，步长 0.5），并在证据列注明时间码或数据来源；",
    "2. 各板块得分由公式自动计算（=ROUND(SUMPRODUCT(指标满分,得分)/10,0)），三轨合计与加权总分在“汇总”工作表自动生成；",
    "3. 类型模板调整：在板块行的“满分”单元格手动修改调整后满分（指标满分不变，调整合计为零）；",
    "4. N/A 处理：不适用指标留空并在证据列注明“N/A”，其分值按比例并入同轨其余板块（手动调整）；",
    "5. 负向指标默认 10 分，出现负向表现按手册扣减并附时间码。",
]
for i, t in enumerate(notes, 1):
    ws.cell(row=i, column=1, value=t).font = Font(size=11, bold=(i == 1))
ws.column_dimensions["A"].width = 110

ws2 = wb.create_sheet("评分表")
headers = ["板块", "指标", "满分", "得分(0-10)", "证据与备注"]
hfont = Font(bold=True, size=11)
fill = PatternFill("solid", fgColor="DDEBF7")
r = 1
for c, h in enumerate(headers, 1):
    cell = ws2.cell(row=r, column=c, value=h)
    cell.font = hfont
    cell.fill = fill
track_rows = {}
for track, board, bmax, inds in BOARDS:
    r += 1
    if track not in track_rows:
        track_rows[track] = {"start": r}
        ws2.cell(row=r, column=1, value="══ %s ══" % track).font = Font(bold=True, size=12)
        r += 1
    bstart = r + 1
    for n, m, o, rule in inds:
        r += 1
        ws2.cell(row=r, column=1, value=board)
        ws2.cell(row=r, column=2, value=n)
        ws2.cell(row=r, column=3, value=m)
    bend = r
    r += 1
    ws2.cell(row=r, column=1, value="%s 板块得分" % board).font = Font(bold=True)
    ws2.cell(row=r, column=3, value=bmax).font = Font(bold=True)
    ws2.cell(row=r, column=4, value="=ROUND(SUMPRODUCT(C%d:C%d,D%d:D%d)/10,0)" % (bstart, bend, bstart, bend)).font = Font(bold=True)
    track_rows[track]["board_end_row"] = r
for track in ["本体翼", "价值翼", "文明翼"]:
    r += 1
    ws2.cell(row=r, column=1, value="%s合计" % track).font = Font(bold=True)
    ws2.cell(row=r, column=3, value=TRACK_MAX[track]).font = Font(bold=True)
    ws2.cell(row=r, column=4, value="=SUM(D%d:D%d)" % (track_rows[track]["start"] + 1, track_rows[track]["board_end_row"])).font = Font(bold=True)
    track_rows[track]["total_row"] = r
for col, w in zip("ABCDE", [18, 22, 8, 12, 44]):
    ws2.column_dimensions[col].width = w

ws3 = wb.create_sheet("汇总")
rB, rV, rC = [track_rows[t]["total_row"] for t in ["本体翼", "价值翼", "文明翼"]]
rows3 = [
    ("本体翼 B（满分 500）", "=评分表!D%d" % rB),
    ("价值翼 V（满分 300）", "=评分表!D%d" % rV),
    ("文明翼 C（满分 200）", "=评分表!D%d" % rC),
    ("用途预设 α（本体）", 0.5),
    ("用途预设 β（价值）", 0.3),
    ("用途预设 γ（文明）", 0.2),
    ("加权总分 S", "=ROUND(1000*(B4*B1/500+B5*B2/300+B6*B3/200),0)"),
    ("等级", '=IF(B7>=900,"传世之作",IF(B7>=800,"优秀",IF(B7>=700,"良好",IF(B7>=600,"合格","一般"))))'),
]
for i, (k, v) in enumerate(rows3, 1):
    ws3.cell(row=i, column=1, value=k).font = Font(bold=True)
    ws3.cell(row=i, column=2, value=v)
ws3.column_dimensions["A"].width = 28
ws3.column_dimensions["B"].width = 16

ws4 = wb.create_sheet("外部因素档案")
for c, h in enumerate(["外部因素", "记录内容", "用途（只记录、不评分）"], 1):
    cell = ws4.cell(row=1, column=c, value=h)
    cell.font = hfont
    cell.fill = fill
for i, (a, b, c) in enumerate(EXTERNAL_ARCHIVE, 2):
    ws4.cell(row=i, column=1, value=a)
    ws4.cell(row=i, column=2, value=b)
    ws4.cell(row=i, column=3, value=c)
for col, w in zip("ABC", [16, 60, 40]):
    ws4.column_dimensions[col].width = w

wb.save(os.path.join(OUT, "空白评分表模板_MFES-2026_v1.0.xlsx"))
print("saved xlsx")

# ================= 案例 CSV =================
CASE_DATA = {
    "《霸王别姬》（1993）": {
        "本体翼": [("剧作文本", 75, 68), ("技术工艺", 60, 47), ("艺术美学", 95, 90), ("思想哲学", 60, 55),
                 ("社会文化", 50, 46), ("政治维度", 25, 19), ("市场产业", 10, 7), ("观众接受", 75, 63), ("历史定位", 50, 47)],
        "价值翼": [("人民性", 75, 62), ("时代性", 45, 40), ("文化自信", 45, 43), ("创新性", 45, 32),
                 ("协调性", 30, 22), ("开放", 30, 28), ("共享", 15, 12), ("绿色", 15, 5)],
        "文明翼": [("人类关怀", 60, 55), ("跨文化可读性", 60, 50), ("文明互鉴", 50, 45), ("全球责任", 30, 20)],
    },
    "《给阿嬷的情书》（2026）": {
        "本体翼": [("剧作文本", 83, 71), ("艺术美学", 105, 82), ("观众接受", 83, 75), ("技术工艺", 67, 52),
                 ("思想哲学", 67, 57), ("社会文化", 56, 53), ("市场产业", 11, 10), ("政治维度", 28, 24)],
        "价值翼": [("人民性", 75, 65), ("时代性", 45, 34), ("文化自信", 45, 43), ("创新性", 45, 31),
                 ("协调性", 30, 25), ("开放", 30, 19), ("共享", 15, 12), ("绿色", 15, 8)],
        "文明翼": [("人类关怀", 60, 54), ("跨文化可读性", 60, 37), ("文明互鉴", 50, 46), ("全球责任", 30, 21)],
    },
    "《流浪地球2》（2023）": {
        "本体翼": [("剧作文本", 75, 61), ("技术工艺", 85, 81), ("艺术美学", 75, 58), ("思想哲学", 30, 26),
                 ("社会文化", 40, 30), ("政治维度", 25, 18), ("市场产业", 40, 34), ("观众接受", 80, 67), ("历史定位", 50, 32)],
        "价值翼": [("人民性", 75, 60), ("时代性", 45, 38), ("文化自信", 45, 36), ("创新性", 45, 41),
                 ("协调性", 30, 25), ("开放", 30, 21), ("共享", 15, 11), ("绿色", 15, 10)],
        "文明翼": [("人类关怀", 60, 52), ("跨文化可读性", 60, 43), ("文明互鉴", 50, 42), ("全球责任", 30, 23)],
    },
    "《功夫女足》（2026）": {
        "本体翼": [("剧作文本", 94, 42), ("艺术美学", 61, 27), ("观众接受", 105, 55), ("技术工艺", 67, 27),
                 ("思想哲学", 39, 19), ("社会文化", 56, 35), ("市场产业", 50, 40), ("政治维度", 28, 16)],
        "价值翼": [("人民性", 75, 48), ("时代性", 45, 27), ("文化自信", 45, 25), ("创新性", 45, 20),
                 ("协调性", 30, 18), ("开放", 30, 17), ("共享", 15, 8), ("绿色", 15, 8)],
        "文明翼": [("人类关怀", 71, 43), ("跨文化可读性", 71, 39), ("文明互鉴", 58, 46)],
    },
}
csv_path = os.path.join(OUT, "cases_board_scores.csv")
with io.open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f)
    w.writerow(["film", "track", "board", "max", "score", "note"])
    for film, tracks in CASE_DATA.items():
        for track, rows in tracks.items():
            for board, mx, sc in rows:
                w.writerow([film, track, board, mx, sc, ""])
print("saved csv")
print("ALL RELEASE FILES DONE")

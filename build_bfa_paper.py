# -*- coding: utf-8 -*-
"""《北京电影学院学报》投稿版论文生成。"""
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

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
    p.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.85)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line
    return p

def heading(doc, text, level=1, center=False):
    sizes = {0: 17, 1: 13.5, 2: 12}
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

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.5)
sec.left_margin = sec.right_margin = Cm(2.6)

heading(doc, "电影评价的“可复算性”建构", 0)
heading(doc, "——多维加权评价体系 MFES-2026 的设计与试评", 0)
para(doc, "康宁（北京电影学院电影学系副研究员）", align="center", indent=False, east="楷体")

para(doc, "摘　要：电影评价长期停留在“印象—口碑—权威”的非科学状态：豆瓣、IMDb 等聚合平台输出单一总分，既无维度结构也不公开算法；奖项评审以评委声望替代方法权威；学术影评则不可检验、不可比较。本文提出并试运行一套多维加权电影评价体系（MFES-2026）：以本体翼（九大板块，500 分）、价值翼（八大板块，300 分）、文明翼（四大板块，200 分）三轨分报，全部 71 项指标采用“全量程五级锚定制”（0–10 分五档公开标准＋档内微调），并以“外部因素档案”将排片、放映等影片之外的因素排除于评分。对《霸王别姬》《给阿嬷的情书》《流浪地球2》《功夫女足》四部影片的试评得分（856、819、809、560）与公开口碑完全同序，且能将“品质”与“票房”分开诊断。文章论证了体系的信度与效度检验方案，并讨论其限度：价值判断的不可消除性、精确性错觉与标准化风险。", east="楷体", size=11, indent=False)
para(doc, "关键词：电影评价；评价体系；锚定评分；信度；效度", east="黑体", size=11, indent=False)
para(doc, "Title: Toward Replicable Film Evaluation: The Design and Trial of MFES-2026, a Multidimensional Weighted System", size=11, indent=False)
para(doc, "Abstract: Film evaluation has long remained in a non-scientific state of \"impression—word-of-mouth—authority\": aggregating platforms such as Douban and IMDb produce a single opaque score, award juries substitute personal prestige for open methodology, and academic criticism is neither testable nor comparable. This article designs and pilots a multidimensional weighted film evaluation system (MFES-2026) consisting of three independently reported tracks: an ontological track (nine sections, 500 points), an axiological track (eight sections, 300 points) grounded in the socialist development outlook with Chinese characteristics, and a civilizational track (four sections, 200 points) grounded in the vision of a community with a shared future for mankind. All 71 indicators use a full-range five-level anchored rating (0-10 with five public anchors plus bounded fine-tuning), and an \"external factors archive\" excludes screening, distribution and other non-filmic factors from scoring. Trial scores for Farewell My Concubine, Love Letter to Grandma, The Wandering Earth II and Kung Fu Women's Football (856, 819, 809, 560) are strictly ordered with public word-of-mouth, and the system separates \"quality\" from \"box office\". The article specifies reliability and validity procedures and discusses the limits of such scientific reformation. Keywords: film evaluation; rating system; anchored rating; reliability; validity", size=11, indent=False)

heading(doc, "一、问题的提出：单一总分的三重盲区", 1)
para(doc, "观众走进影院前看分数，走出影院后争分数。但一个总分能回答什么？它说不清电影好在哪、差在哪；说不清分数是谁打的、按什么规则算出；更说不清换一批人来评，结论是否还成立。豆瓣与 IMDb 的加权算法不公开，样本又受自选择偏差影响，刷分与“锁分”争议不断；烂番茄把影评二值化为“新鲜/腐烂”，丢失了评价的全部强度信息；Metacritic 依赖的影评人样本量小且国别不均。奖项体系则以身份权威替代方法权威：奥斯卡、戛纳、金鸡奖均不公布评分细则，其结论不可复算。学术影评以个人感悟与修辞见长，两位评论家对同一影片的结论无法比较，任何“影史排名”都只是修辞的堆叠。三种路径各有所长，也各自失明——电影价值评价由此长期停留在“印象—口碑—权威”的非科学状态。本文要回答的问题是：电影价值能否被多维量化？如果能量化，评分过程能否被公开审计？")
para(doc, "需要先行澄清一个范畴区分：可通约与可复算是两回事。趣味之争不需要、也不可能被“解决”——审美判断的主观性不可消除；但一个评价结论“基于哪些证据、经过哪些步骤得出、换一批受训评分者结论是否落在可接受区间”，是可以检验的。本文所谓“科学化”，目标不是取消争辩，而是把争辩从“好不好”的重复申说，转化为“凭什么”的可核对论证。")

heading(doc, "二、体系设计：三轨分报与千分制权重", 1)
para(doc, "MFES-2026 由三轨构成。本体翼回答“电影本身好不好”，面向一切国别与时代，是体系的科学内核，下设剧作文本 75、艺术美学 75、观众接受 75、技术工艺 60、历史定位 50、思想哲学 50、社会文化 50、市场产业 40、政治维度 25，共 500 分。价值翼回答“是否符合中国特色社会主义发展观”，下设人民性 75、时代性 45、文化自信 45、创新性 45、协调性 30、开放 30、共享 15、绿色 15，共 300 分。文明翼回答“是否回应人类命运共同体”，下设人类关怀 60、跨文化可读性 60、文明互鉴 50、全球责任 30，共 200 分。三轨独立评分、独立报告、按用途加权——价值判断不污染本体评价，方向要求不取代品质评判。综合加权总分为：")
para(doc, "S = 1000 × [ α·(B/500) + β·(V/300) + γ·(C/200) ]", align="center", indent=False, bold=True)
para(doc, "默认综合用途取 α=0.50、β=0.30、γ=0.20，此时 S = B + V + C，满分 1000 分；学术研究取 0.80/0.10/0.10，国际传播选片取 0.40/0.10/0.50。采用千分制而非百分制，是为了让权重直接以满分分值呈现，任何板块的加减分一目了然，便于审计。权重经德尔菲法三轮专家征询与层次分析法（AHP）确定，一致性比率 CR＜0.1。三个设计取向值得说明：观众接受与剧作、艺术并列为最高权重，落实“作品价值在接受中完成”的接受美学立场；政治维度刻意压低至 25 分，防止本体评价政治化；人民性 75 分与本体最高板块同权，落实“以人民为中心”的创作导向。")
para(doc, "体系配设五类动态机制。其一，类型模板：不同片种在本体翼内再分配权重（调整合计为零），如动作/科幻片技术工艺＋25、思想哲学－20、社会文化－10；文艺片艺术美学＋20、市场产业－30；调整值按类型相关性分配到板块内指标。其二，N/A 再分配：不适用板块（默片的录音、抽象电影的剧作项）不评分，分值按比例再分配给同轨板块，并披露“适用覆盖率”。其三，时代双轨：默认口径为语境相对分（与同时代同类型比较），另记绝对分供影史研究；上映不足五年的影片，历史定位按“潜质”评，不足一年整板块 N/A。其四，用途预设：同一影片按用途取不同系数，三轨原始分不变，保证可追溯。其五，本体原则与外部因素档案：排片、档期、票价、营销投入、放映设备、公益放映安排、获奖公关、审查环境等影片制作方不可控因素，只记录、不评分——这一设计纠正了评价实践中长期存在的对象错位，小成本佳片不会因排片少而吃亏。")

heading(doc, "三、评分细则：全量程五级锚定制", 1)
para(doc, "评分细则经历了从“满分起评扣分制”到“全量程五级锚定制”的修正。扣分制起评 10 分、逐项扣减，存在两个逻辑缺陷：其一，加分点在无扣分时被 10 分上限吞掉，沦为“扣分补偿券”；其二，“起评 10 分”将“无缺陷”错误等同于“标杆”，平庸之作必须靠“无表意”“平淡”这类别扭的缺失型扣分压分。修正后的规则为每项指标给出 0–10 全域的五级锚定标准：10 分为标杆（同类参照系顶级，可作样例）、7.5 分为优秀、5 分为合格（同类平均）、2.5 分为不足、0 分为严重缺陷（结构性失败或违背规范）。评分分两步：第一步对照锚点定档，允许落在两档之间（步长 0.5）；第二步对照微调点清单浮动——每命中一处缺陷 −0.5（合计 −1 封顶），每命中一处加分点 +0.5（合计 +1 封顶）；标注 −2 及以上或“结构性失败/底线联动”的扣分点直接压入低档。最终得分 = 档位 ± 微调，恒在 [0,10]，数学上不存在超出满分。以摄影指标为例：技艺娴熟、表意基本实现，落 7.5 档；一处失焦（−0.5）、一处光线层次平淡（−0.5），得 6.5 分；若有运镜表意创新（+0.5），得 7.0 分。")
para(doc, "两条辅助规则保证了评分边界。其一，中性处理：指标对象不存在时（未上映、未参与合拍、无政治表达、未涉全球议题、未国际发行等），该项按中性 5 分或 N/A 再分配处理——总原则是不得以不存在之事扣分；无特效的小成本剧情片不会因此输给特效大片。其二，证据背书：每个定档依据与微调点必须附可核查证据（镜头时间码、数据来源、文献出处），无证据不得动分。板块得分 = Σ（指标得分/10 × 指标满分），四舍五入取整。负向指标（避免狭隘民族主义、避免自我东方化）默认满分，出现负向表现按细则扣减并附时间码举证。")

heading(doc, "四、主体性的控制：信度与效度的制度设计", 1)
para(doc, "工具再严密也要由人执行，主体性控制因此必须是制度而非技术修补。体系设四层控制。第一层，以事实判断替代价值判断：把“剪辑节奏好”拆为镜头时长与节奏曲线等可观测事实，主观只允许出现在证据综合环节。第二层，约束残余判断：锚定样例库为每个档位提供标准影片片段，评分员经标定测试（与标准样例偏差小于 1.5 分）方可上岗；每部影片至少三名评分员独立评分，组内相关系数 ICC≥0.75 或 Cohen's κ≥0.61 为合格，分差超过 2.0 分的指标进入仲裁。第三层，控制系统性偏见：利益申报与回避、盲评环节、多角化评审主体（文明翼强制包含非中文母语评分员；观众接受分层报告影评人/大众/类型影迷三组分数）、重测稳定性检验。第四层，暴露体系设计者的立场：权重表、指标库、评分手册全文公开，体系版本化，修订必须公示理由。需要警惕三种伪客观：把票房与平台分当客观的数据崇拜、指标越多越科学的指标堆砌、多人平均保留共同偏见的平均幻觉。效度检验分三层：内容效度经德尔菲专家确认，结构效度以因子分析验证板块聚合结构，效标效度以权威奖项、影史教材与公开口碑为效标进行校准研究。")

heading(doc, "五、四部影片试评", 1)
para(doc, "按上述规则，对《霸王别姬》（1993）、《给阿嬷的情书》（2026）、《流浪地球2》（2023）、《功夫女足》（2026）四部影片试评。新片的历史定位板块按 N/A 再分配；《功夫女足》的全球责任板块因未涉全球议题同样按 N/A 处理；两片未参与合拍均按中性 5 分处理。结果如下：")
table(doc, ["影片", "本体(500)", "价值(300)", "文明(200)", "总分", "等级", "豆瓣参照"],
      [["《霸王别姬》（1993）", "442", "244", "170", "856", "优秀", "9.6"],
       ["《给阿嬷的情书》（2026）", "424", "237", "158", "819", "优秀", "9.0＋"],
       ["《流浪地球2》（2023）", "407", "242", "160", "809", "优秀", "8.3"],
       ["《功夫女足》（2026）", "261", "171", "128", "560", "一般", "6.5"]],
      widths=[4.6, 1.8, 1.8, 1.8, 1.6, 1.5, 1.7])
para(doc, "试评得出三点发现。其一，四片得分与公开口碑完全同序（856＞819＞809＞560 对 9.6＞9.0＞8.3＞6.5），构成效标效度的初步证据。其二，《功夫女足》票房约 17 亿元，市场产业板块得分 40/50 不低，但总分仅 560、落入“一般”档——技术工艺（特效、剪辑被广泛批评）与艺术美学（表演小品化）两块拖累整体。剖面上“高票房、低品质”一目了然，这是单一总分无法给出的诊断：票房测的是市场，总分测的是价值，二者分开，才能各得其所。其三，《霸王别姬》的价值翼按 2026 年标准评出现系统性偏低（绿色、数字资产复用两项因时代所限损失 13 分），提示后续版本需对老片启用价值翼语境分；“历史书写责任”指标触发 1.8 分的评分员间分差并进入仲裁，验证了争议标注机制的必要性。")

heading(doc, "六、限度与讨论", 1)
para(doc, "对测量学改造的限度必须有清醒的自觉。第一，价值判断不可消除：权重与锚点的选择本身就是价值判断，科学化改变的是这些判断的可见性，而非其存在性；宣称完全中立的评价体系只是把立场藏了起来。第二，精确性错觉：千分制总分必须与三轨剖面、置信区间同报，任何脱离剖面引用的总分都是误导。第三，标准化与创新的张力：锚定样例天然偏向已有范式，先锋作品易被低估，须以“形式即内容”规则与样例库的持续更新加以缓冲。第四，制度化风险：任何可操作的评价工具都可能被用作审查与准入工具，对策是把底线准入严格限定于评奖与资助序列，学术用途不设任何底线。第五，语境相对性的代价：语境相对分保证了跨时代公平，却牺牲了跨时代通约，体系不应被用于“有史以来最佳影片”式的总体排名。第六，个案试评的有限性：四例试评是体系可操作性的演示，不是有效性的证明，有效性有待大样本信度研究。还需澄清科学化与批评传统的关系：测量学改造不是要取代印象式批评——批评的洞察与修辞仍不可替代，恰恰是它们不断为指标库与锚定样例提供新资源；体系提供的只是可检验的公共基础，让批评的洞见与批评的偏见有办法被区分开来。二者是分工而非替代关系。")

heading(doc, "七、结语", 1)
para(doc, "电影评价的终极问题不是“能不能打分”，而是“分数能不能被追问”。当一套体系能把 71 项指标的五档标准、全部扣分理由与证据公之于众，任何人对任何一分都可以复核挑战时，评价就不再是权威的独白，而成为可检验的公共知识。文艺评价的规范理想，应当从“客观性”改写为“受约束的主观性”：指标可观测、定档可复核、结论可复算，分歧可度量、偏差可追溯、立场可挑战。用科学性赢得信任，用人民性站稳立场，用世界性赢得对话——这是 MFES-2026 的设计理想，也应当是中国电影评价走向成熟的方向。")

heading(doc, "参考文献", 1)
refs = [
    "〔美〕大卫·波德维尔，克里斯汀·汤普森．电影艺术：形式与风格[M]．曾伟祯，译．北京：北京联合出版公司，2015．",
    "〔法〕安德烈·巴赞．电影是什么？[M]．崔君衍，译．北京：商务印书馆，2017．",
    "〔美〕托马斯·沙茨．好莱坞类型电影[M]．冯欣，译．上海：上海人民出版社，2009．",
    "〔德〕H.R.姚斯，〔美〕R.C.霍拉勃．接受美学与接受理论[M]．周宁，金元浦，译．沈阳：辽宁人民出版社，1987．",
    "习近平．在文艺工作座谈会上的讲话（2014 年 10 月 15 日）[M]．北京：人民出版社，2015．",
    "全国人民代表大会常务委员会．中华人民共和国电影产业促进法[Z]．2017-03-01 施行．",
    "豆瓣电影．豆瓣电影评分说明[EB/OL]．https://movie.douban.com/about（查询于 2026 年 10 月）．",
    "British Film Institute. The Greatest Films of All Time: Sight and Sound 2022 Poll[EB/OL]. https://www.bfi.org.uk/sight-and-sound/greatest-films-all-time（查询于 2026 年 10 月）.",
    "Smith P C, Kendall L M. Retranslation of expectations: An approach to the construction of unambiguous anchors for rating scales[J]. Journal of Applied Psychology, 1963, 47(3): 149-155.",
    "Saaty T L. The Analytic Hierarchy Process[M]. New York: McGraw-Hill, 1980.",
    "Cohen J. A coefficient of agreement for nominal scales[J]. Educational and Psychological Measurement, 1960, 20(1): 37-46.",
    "Cronbach L J. Coefficient alpha and the internal structure of tests[J]. Psychometrika, 1951, 16(3): 297-334.",
]
for i, r in enumerate(refs, 1):
    para(doc, "[%d] %s" % (i, r), size=10.5, indent=False, after=3, line=1.25)

doc.save(os.path.join(OUT, "投稿_北京电影学院学报_MFES.docx"))
print("saved 学报论文")

# -*- coding: utf-8 -*-
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import shutil

OUT = Path("dist")
OUT.mkdir(exist_ok=True)

def p(text="", bold=False, size=22, center=False, first=True, font="宋体", space_after=80):
    text = "" if text is None else str(text)
    jc = '<w:jc w:val="center"/>' if center else ''
    ind = '<w:ind w:firstLine="420"/>' if first and not center else ''
    ppr = f'<w:pPr>{jc}{ind}<w:spacing w:line="360" w:lineRule="auto" w:after="{space_after}"/></w:pPr>'
    b = '<w:b/>' if bold else ''
    rpr = f'<w:rPr>{b}<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'

def heading(text, level=1):
    size = 30 if level == 1 else 26
    return p(text, bold=True, size=size, center=False, first=False, font="黑体", space_after=120)

def make_docx(path, title, subtitle, sections):
    body = [p(title, bold=True, size=36, center=True, first=False, font="黑体", space_after=220)]
    if subtitle:
        body.append(p(subtitle, size=24, center=True, first=False, font="宋体", space_after=220))
    for head, text in sections:
        body.append(heading(head, 1))
        for line in str(text).split("\n"):
            line=line.strip()
            if not line:
                continue
            if line.startswith("【") and line.endswith("】"):
                body.append(heading(line.strip("【】"),2))
            elif line.startswith("• "):
                body.append(p("• "+line[2:], size=22, first=False))
            else:
                body.append(p(line, size=22, first=True))
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + \
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + ''.join(body) + sect + '</w:body></w:document>'
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
</w:styles>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    docrels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/_rels/document.xml.rels", docrels)

opening_sections = [
("一、课题基本信息", """课题名称：初中物理知识图谱支持下的学情诊断与循证辅导
学科：初中物理
研究周期：2026年9月—2027年8月
研究性质：校级、一人承担的实践性预研课题
学校：龙岗区吉华街道怡翠实验学校
说明：现有学情分析工具与知识图谱仅作为前期技术可行性基础和原型探索，不作为本课题已经完成的正式研究成果。"""),

("二、问题提出与研究背景", """日常考试能够持续产生分数、题目得分、错题、排名变化等数据，但一线教师真正困难的并不是“有没有数据”，而是怎样把分散的数据转化为可以解释、可以行动、可以验证的学情证据。相同的一道错题，可能来自知识遗忘、前置知识缺口、概念混淆、稳定误概念、方法性错误，也可能只是偶发失误。如果只依据一次错误就安排重复练习，容易把“发现错误”等同于“解释错误”，也难以判断后续辅导是否真正修复了学生的认知问题。
本课题因此把研究重点从“做一个更完整的平台”收窄为一条教师能够长期执行的诊断链：日常考试证据筛查—锁定个体异常—教师轻量题目知识点标注—诊断微题组复诊—必要时访谈核验—知识图谱回溯—按错因辅导—短周期再测—后续考试观察。学情平台主要承担筛查和证据汇集，知识图谱主要承担解释和干预支架，两者不是两个并列的软件成果。
物理学习具有明显的概念关联性。学生对“浮力、压强、惯性、平衡力与相互作用力、电流与电压、电功与电功率”等内容出现错误时，错误往往与前置概念、相邻概念或生活经验中的直觉解释有关。因此，本研究将“为什么错—如何辅导”作为核心，把知识图谱中的前置、易混、关联关系用于诊断后的回溯，而不是把知识图谱仅作为知识展示工具。"""),

("三、核心概念界定", """1. 学情诊断：不是对一次考试成绩作一般性描述，而是依据多次考试、题目作答、理由、置信度、跨题一致性以及必要的访谈证据，对学生当前认知状态作出可修订的判断。
2. 误概念：学生对物理概念、规律或条件形成的相对稳定且与科学解释不一致的认知。研究中避免因单题错误直接给学生贴“误概念”标签，而使用“稳定误概念候选”“脆弱理解”“知识缺口”等审慎表述。
3. 诊断微题组：围绕一个高频易错概念设计的少量关联题，综合“作答结果+选择理由+置信程度+同概念多题一致性”判断认知状态。必要时辅以教师追问。
4. 知识图谱支持：利用知识点之间的前置、后续、易混和跨概念关系帮助教师和学生追溯错误可能涉及的认知链条，并为后续辅导提供结构化线索。
5. 循证辅导：依据学生实际作答证据、诊断证据和相关研究证据选择辅导策略，并通过再测和后续正式考试检验判断与干预是否有效。"""),

("四、国内外研究现状与本课题切入点", """科学教育中的误概念诊断已经形成较成熟的工具传统。Treagust提出两层诊断测验，把“答案”和“理由”结合起来；Caleon与Subramaniam进一步在机械波研究中引入答案与理由的置信度，形成四层诊断思路[1-2]。Kaltakci-Gurel等对科学误概念诊断工具的综述表明，访谈、开放题、选择题和多层诊断测验各有优势与成本，说明一线教学不宜只依赖单一题型或单次作答[3]。2017年的几何光学四层诊断研究也显示，多层信息有助于区分“答错”与较稳定的错误认识[4]。这些研究为本课题采用“答案—理由—置信度—跨题一致性”的轻量诊断证据组合提供了方法依据。
概念转变研究强调，学生并非带着空白头脑进入课堂。已有概念结构会影响新知识解释；有效干预需要识别原有解释，并通过认知冲突、证据比较、模型重建等方式促进概念重组[5-6]。因此，本课题不把“多做几道同类题”默认视为辅导，而是根据错误原因选择实验冲突、概念辨析、前置知识补偿、表征转换或方法训练。
形成性评价和反馈研究进一步提示：反馈的价值不在于简单告知对错，而在于提供能改变后续学习行动的信息。Hattie与Timperley、Shute以及Wisniewski等的研究都强调反馈内容、层次和可行动性的重要性[7-9]。这支持本课题把诊断结果转化为具体辅导建议，并通过短周期再测验证。
题目与知识属性的对应关系是本课题落地的关键。认知诊断研究中的Q矩阵用于描述题目需要哪些认知属性。传统做法通常先由领域专家依据课程与题目内容定义题目—属性关系，再利用作答数据检验、修订Q矩阵。de la Torre与Chiu明确指出，Q矩阵通常由领域专家构建，但专家判断具有主观性，因此需要经验数据进行验证[10]。这为本课题“教师先标注、数据后验证”的做法提供了直接方法依据。
近年教育知识图谱研究表明，知识图谱能够组织概念及其关系，并被用于个性化学习、资源推荐、学情诊断等场景；同时也存在知识建模成本高、数据资源不足、应用评价不足等问题[11-13]。因此，本课题不追求大规模自动构图或复杂算法，而采用教师可维护的轻量知识关系作为解释性支架。
综合来看，已有研究分别为误概念诊断、概念转变、反馈、Q矩阵和教育知识图谱提供了较充分依据；本课题的实践切入点不是宣称这些领域“无人研究”，而是探索一名普通初中物理教师如何在日常考试情境中，把这些研究传统压缩成低负担、可持续、可复用的一条证据链。"""),

("五、研究目标与研究问题", """研究目标：形成一套适合一人、一年校级课题实施的初中物理个体学情诊断与循证辅导流程，使日常考试中的“这个学生又错了”能够进一步转化为“他为什么错”“下一步怎么辅导”“辅导后是否真正改变”。
研究问题1：如何将已有轻量知识图谱从知识展示工具转化为诊断后的解释与辅导支架，使其能够支持前置知识、易混概念和关联概念的回溯？
研究问题2：如何以较低教师负担建立可靠的“题目—主知识点—学生作答”联系，并在需要时增加关联知识点和典型错误模型，使日常考试数据能够进入个体诊断链？
研究问题3：围绕若干高频误概念，答案、理由、置信度与跨题一致性怎样组合，才能较稳妥地区分知识缺失、脆弱理解、稳定误概念候选、方法性错误和偶发错误？
研究问题4：依据诊断结果实施针对性辅导后，学生在短周期再测和后续正式考试中的表现是否与原诊断判断相互印证？"""),

("六、研究内容", """【（一）建立高频误概念母表】
不追求覆盖全部知识点，先选择对初中物理学习影响较大、在教学中高频出现且适合诊断的概念，形成“知识点—前置知识—典型错误表征—可能误概念—诊断证据—辅导策略—再测方式”母表。
【（二）设计教师轻量题目—知识点标注规则】
每次考试原则上每题只完成1个主知识点标签；只有确有诊断意义时，才补充1—2个关联知识点或一个典型错误模型。标签以题目内容和主要认知要求为依据，不追求把一道综合题拆成大量属性。相同题目或稳定题型形成可复用标签库，后续考试优先复用，只在命题要求发生变化时修订。
【（三）开发诊断微题组】
围绕每个目标误概念设计2—4道互相关联但表面情境不同的题目。每题不仅记录答案，还尽量记录选择理由和置信度；同一概念用变式题观察一致性。单题错误只作为筛查信号，不直接判定稳定误概念。
【（四）形成认知状态判断规则】
初步采用五类状态：正确答案+正确理由+高置信且跨题稳定，视为相对稳定理解；正确但理由含糊或低置信，视为脆弱理解；错误且理由不稳定、低置信，优先考虑知识缺口或不确定；错误+具体错误理由+高置信且跨变式重复，列为稳定误概念候选；单题错误但后续同概念题正确且理由正确，优先考虑偶发或执行性错误。该规则在研究过程中依据案例修订。
【（五）知识图谱回溯与按因辅导】
当诊断指向某个薄弱点时，利用知识图谱查看其前置知识、易混概念和关联关系，再由教师决定辅导路径。例如学生在漂浮条件题中反复表现出“物体越深浮力越大”的高置信错误时，不立即增加计算题，而先核查其对排开液体体积、液体压强与浮力关系的理解，再设计等体积浸没等认知冲突活动。
【（六）再测与证据回流】
每次重点辅导后安排短周期再测，并在后续正式考试继续观察同一概念。若学生只在原题上改善而在变式题中仍重复原错误，则不视为概念修复完成。诊断结果、教师追问和再测结果共同用于修订题组、标签和图谱关系。"""),

("七、教师轻量标注：本课题的关键操作方案", """本研究不把“题目自动标知识点”作为技术开发任务，而把教师专业判断正式纳入诊断链。Q矩阵研究中常见的思路是由领域专家先定义题目需要的认知属性，再用学生作答数据检验和修订[10]。在校级课题中可将这一传统简化为“专家先定义、数据后验证”的教师版本。
第一步，考试后只做主标签。教师对每道题回答一个问题：“学生要正确解决这道题，最核心、最不可替代的物理知识是什么？”只记录一个主知识点。例如“浮力—物体浮沉条件”。不要一开始同时标“密度、质量、体积、受力分析、阿基米德原理”等所有可能关系。
第二步，只有出现诊断价值时再加深。若某道题错误率异常，或某名学生在同一概念上连续异常，再补充关联知识点、主要认知要求和典型错误模型。例如：主知识点“浮力”；关联“阿基米德原理、浮沉条件”；错误模型“把浸没深度作为浮力决定因素”。
第三步，建立复用库。标签对象不只保存“第12题”，还保存题型特征或题目指纹。下一次出现同题、改数题或结构稳定的同型题，系统或教师直接复用原标签，只需确认是否改变了核心认知要求。这样教师工作量随使用次数增加而下降，而不是每次考试重新从零标注。
第四步，用数据做后验检查。若被标为同一主知识点的一组题在大量学生中的表现长期完全不一致，或访谈显示学生错误原因主要来自另一个知识点，就把它作为“标签需要复核”的信号，而不是立即让算法自动改标签。数据的作用是提醒和验证，最终修订仍由教师结合题目内容决定。
第五步，建立最低可行数据结构。建议每题只保存：考试ID、题号、主知识点ID、可选关联知识点ID、可选错误模型ID、题型/来源、标签版本、教师确认时间。学生层面再关联得分、选项、理由、置信度即可。这样已经足够支撑一年期课题，不必开发复杂Q矩阵算法。
这一方案的研究价值在于：它承认题目—知识点映射需要学科专业判断，同时通过“单主标签、按需加深、长期复用、数据复核”把教师负担控制在可接受范围。"""),

("八、研究方法及证据对应", """1. 文献研究法：用于确定误概念诊断、概念转变、形成性反馈、Q矩阵与知识图谱的理论和方法边界。
2. 行动研究法：作为一年期研究的主要实践框架，在真实考试—诊断—辅导—再测循环中持续修订流程。
3. 学习数据分析：利用多次考试成绩、题目得分、错题、趋势等信息进行初筛，回答“谁在什么位置持续出现异常”。
4. 诊断性访谈/半结构访谈：只对有代表性的学生和关键概念使用，用于核验学生为什么这样想，不作为每名学生每次考试的常规任务。
5. 个案研究法：选取若干典型学生，记录“考试证据—诊断—图谱回溯—辅导—再测—后续考试”的纵向证据链。
本研究不把准实验作为核心。若研究条件允许，可用前后测、同概念再测或小范围比较作为辅助量化证据，但不为追求统计显著而扩大样本和工作量。"""),

("九、技术路线", """文献与理论框架 → 筛选高频易错易混概念 → 建立“误概念—前置知识—错误表征”母表 → 形成教师轻量题目—知识点标注规则 → 日常考试多源证据筛查 → 锁定个体异常 → 诊断微题组（答案/理由/置信度/跨题一致性） → 必要时教师追问 → 认知状态判断 → 知识图谱回溯前置/易混/关联关系 → 按错因辅导 → 短周期再测与后续正式考试 → 证据回流，修订标签、题组与图谱关系。"""),

("十、实施步骤与进度安排", """第一阶段（2026.9—2026.11）：完成文献梳理和研究框架；确定首批高频误概念；形成轻量标注字段、认知状态初步规则和首批诊断微题组。此阶段成果是“规则和原型”，不宣称形成成熟系统。
第二阶段（2026.12—2027.3）：在真实考试中试行。优先选择约6—8个高价值误概念，形成约20—30道诊断题；积累若干“考试异常—微题组—访谈—图谱回溯”的完整案例；根据实际工作量调整标注规则。
第三阶段（2027.4—2027.6）：对第一轮规则进行修订，适度扩展到约12—15个高频概念；重点积累纵向个案，观察辅导后的短周期再测和后续考试表现；完善题型标签复用机制。
第四阶段（2027.7—2027.8）：整理数据与案例，形成“筛查—复诊—归因—图谱回溯—辅导—再测”的轻量流程，完成结题报告、诊断微题组样例和典型案例集。"""),

("十一、预期成果", """1. 一份初中物理高频误概念—错误表征—前置知识—诊断证据—干预策略母表。
2. 一套教师轻量题目—知识点标注规则与可复用标签样例。
3. 一批围绕高频误概念的诊断微题组及认知状态判断规则。
4. 若干完整的学生纵向诊断与辅导案例。
5. 一套教师可持续执行的“考试筛查—诊断—图谱回溯—按因辅导—再测验证”流程。
6. 校级课题结题报告。"""),

("十二、可能创新点", """本课题不把软件功能数量作为创新，而把创新控制在三个可验证的实践点：一是把日常考试中的错题证据继续追问到“为什么错”，避免单题错误直接等同于知识不会；二是把答案、理由、置信度、跨题一致性与必要访谈组合成适合一线教师的小型诊断证据链；三是借鉴Q矩阵“专家先定义、数据后验证”的思想，形成“单主标签—按需加深—长期复用—数据复核”的低负担题目知识点映射办法，并把映射结果与知识图谱的前置、易混关系连接起来。"""),

("十三、研究基础、条件与边界", """前期已对考试数据分析、错题下钻、多次考试趋势、学生个体报告以及初中物理知识点关系进行了技术可行性探索，说明相关数据处理和图谱展示在普通教师条件下具有实现可能。正式课题研究从开题后开始，前期原型只用于减少技术不确定性。
研究者为一线初中物理教师，能够持续获得真实教学情境中的考试和辅导证据，但同时也是单人课题。因此研究主动限制范围：不追求覆盖全部知识点，不开发复杂认知诊断算法，不研究自动题目知识点标注，不以大规模准实验为主要任务。研究质量主要通过证据链完整性、规则可复用性和案例的纵向验证来保证。"""),

("十四、课题组分工", """本课题由负责人一人承担。负责人完成文献研究、研究设计、题目知识点标注、诊断微题组开发、学生访谈、辅导实施、再测、数据整理、案例分析和报告撰写。必要时邀请同备课组教师对个别题目标签或误概念判断进行非正式同行复核，但不将其列为课题组成员。"""),

("十五、参考文献", """[1] Treagust D F. Development and use of diagnostic tests to evaluate students' misconceptions in science[J]. International Journal of Science Education, 1988, 10(2): 159-169.
[2] Caleon I S, Subramaniam R. Do students know what they know and what they don't know? Using a four-tier diagnostic test to assess the nature of students' alternative conceptions[J]. Research in Science Education, 2010, 40(3): 313-337.
[3] Gurel D K, Eryilmaz A, McDermott L C. A review and comparison of diagnostic instruments to identify students' misconceptions in science[J]. EURASIA Journal of Mathematics, Science and Technology Education, 2015, 11(5): 989-1008.
[4] Kaltakci-Gurel D, Eryilmaz A, McDermott L C. Development and application of a four-tier test to assess pre-service physics teachers' misconceptions about geometrical optics[J]. Research in Science & Technological Education, 2017, 35(2): 238-260.
[5] Posner G J, Strike K A, Hewson P W, Gertzog W A. Accommodation of a scientific conception: Toward a theory of conceptual change[J]. Science Education, 1982, 66(2): 211-227.
[6] Duit R, Treagust D F. Conceptual change: A powerful framework for improving science teaching and learning[J]. International Journal of Science Education, 2003, 25(6): 671-688.
[7] Hattie J, Timperley H. The power of feedback[J]. Review of Educational Research, 2007, 77(1): 81-112.
[8] Shute V J. Focus on formative feedback[J]. Review of Educational Research, 2008, 78(1): 153-189.
[9] Wisniewski B, Zierer K, Hattie J. The power of feedback revisited: A meta-analysis of educational feedback research[J]. Frontiers in Psychology, 2020, 10: 3087. DOI:10.3389/fpsyg.2019.03087.
[10] de la Torre J, Chiu C Y. A general method of empirical Q-matrix validation[J]. Psychometrika, 2016, 81(2): 253-273. DOI:10.1007/s11336-015-9467-8.
[11] Abu-Salih B, Alotaibi S. A systematic literature review of knowledge graph construction and application in education[J]. Heliyon, 2024, 10(3): e25383. DOI:10.1016/j.heliyon.2024.e25383.
[12] 沈红叶, 肖婉, 季一木, 等. 教育知识图谱的类型、应用及挑战[J]. 软件导刊, 2023, 22(10): 237-243. DOI:10.11907/rjdk.222110.
[13] Qu K, Li K C, Wong B T M, et al. A survey of knowledge graph approaches and applications in education[J]. Electronics, 2024, 13(13): 2537.
[14] Black P, Wiliam D. Developing the theory of formative assessment[J]. Educational Assessment, Evaluation and Accountability, 2009, 21: 5-31.
[15] Zimmerman B J. Becoming a self-regulated learner: An overview[J]. Theory Into Practice, 2002, 41(2): 64-70.
说明：本开题稿优先保留可明确核验的文献。原申报材料中个别作者、年份或题名无法公开稳定核验的条目未写入本稿，避免形成不可追溯引用。"""),

("十六、开题后需要负责人确认的信息", """1. 学校正式开题表中负责人姓名、职称、联系方式等个人信息。
2. 学校对“研究周期”和“中期检查”具体月份的格式要求。
3. 首批进入诊断微题组的6—8个高频误概念名单。
4. 学生数据匿名化方式和校内数据使用规范。
5. 是否需要在正式开题会上展示前期原型；若展示，应明确表述为“前期可行性探索”，不作为已完成成果。""")
]

lit_sections = [
("一、文献调研结论摘要", """前期文献调研对本课题最有价值的结论不是“知识图谱很先进”，而是三条可直接转化为研究设计的证据链。第一，误概念诊断不能只看一道题的对错，多层诊断测验把答案、理由和置信度结合起来，可以更细致地区分“不知道”“猜对”“稳定错误解释”等状态。第二，诊断的目的不是给学生贴标签，而是为概念转变和后续反馈提供依据；辅导后需要通过变式再测和后续表现验证。第三，题目—知识点映射不必追求全自动，Q矩阵研究本来就常由领域专家先定义属性关系，再由经验数据进行验证和修订。这三条正好对应本课题“为什么错—如何辅导—如何验证”的主线。"""),
("二、对教师轻量标注的直接方法依据", """de la Torre与Chiu（2016）在Q矩阵验证研究中指出，Q矩阵用于规定每道题所需的认知属性，通常由领域专家构建；但专家判断存在主观性，因此需要经验数据检验。对一线教师而言，可将其转化为更轻量的实践规则：先由教师依据题目核心认知要求给出一个主知识点标签；只有在诊断需要时再补关联知识点或错误模型；随着题库积累复用既有标签；学生作答数据用于发现需要复核的映射，而不是让算法自动替代教师判断。
这种设计的优势是把“教师介入”从技术不足的补救，转化为研究方法的一部分。研究对象也因此不是“自动标注算法准确率”，而是“在可接受教师负担下，专家标注—数据复核机制能否稳定支撑个体诊断”。"""),
("三、诊断微题组的证据结构", """建议把常规诊断证据固定为“3+1”：作答结果、选择理由、置信程度，加上同概念多题的一致性。访谈不作为常规第四层，而只用于关键案例核验。这样既继承多层诊断测验的思想，又控制一人课题的实施成本。
单题错误只能触发“复诊”，不能直接输出“稳定误概念”。当学生在多个表面不同但核心概念相同的题目中重复相同错误理由，并且置信度较高时，才把它列为稳定误概念候选。"""),
("四、研究空白的稳妥表述", """不宜写“目前国内外尚无将知识图谱用于初中物理学情诊断的研究”之类绝对判断。更稳妥的表述是：已有研究分别覆盖误概念诊断、形成性评价、认知诊断和教育知识图谱，但在普通中学教师日常考试情境中，如何以低负担方式把错题筛查、教师题目知识点标注、误概念复诊、图谱回溯、针对性辅导和再测验证整合为可持续流程，仍有较大的实践探索空间。"""),
("五、首批建议研究对象", """为控制一年期单人课题工作量，首批可优先从以下类型中选6—8个，而不是一次铺开全部知识点：惯性与力；平衡力与相互作用力；压力与压强；浮力决定因素与漂浮条件；内能、热量与温度；电流与电压；电功与电功率；质量与重力。最终名单应结合本校学生真实错题频率确定。"""),
("六、已核验的核心参考文献", """1. Caleon I S, Subramaniam R. Research in Science Education, 2010, 40(3):313-337。四层诊断测验：答案、理由及相应置信度。
2. Gurel D K, Eryilmaz A, McDermott L C. EURASIA Journal of Mathematics, Science and Technology Education, 2015, 11(5):989-1008。综述273篇研究，对访谈、开放题、选择题和多层诊断工具进行比较。
3. Kaltakci-Gurel D, Eryilmaz A, McDermott L C. Research in Science & Technological Education, 2017, 35(2):238-260。几何光学四层诊断工具。
4. de la Torre J, Chiu C Y. Psychometrika, 2016, 81(2):253-273。Q矩阵通常由领域专家构建，并提出经验验证方法。DOI:10.1007/s11336-015-9467-8。
5. Wisniewski B, Zierer K, Hattie J. Frontiers in Psychology, 2020, 10:3087。反馈元分析，强调反馈效果受信息内容与形式影响。DOI:10.3389/fpsyg.2019.03087。
6. Abu-Salih B, Alotaibi S. Heliyon, 2024, 10(3):e25383。教育知识图谱构建与应用系统综述。DOI:10.1016/j.heliyon.2024.e25383。
7. 沈红叶, 肖婉, 季一木, 等. 软件导刊, 2023, 22(10):237-243。教育知识图谱类型、应用与挑战。DOI:10.11907/rjdk.222110。
其余经典理论来源见开题报告参考文献。公开平台被引次数会随时间变化，本下载包不把动态被引数写入正式开题正文，以避免后续失真；如学校要求，可在最终提交前另做一次日期化核验。"""),
("七、对原申报材料的处理", """原申报材料中出现但当前无法稳定核验的个别中文作者、年份、模型名称，不应继续作为开题报告的关键依据。正式稿应宁可减少条目，也不要保留无法追溯的文献。已有软件功能只写作前期技术可行性探索；正式研究成果必须在开题后通过诊断微题组、真实案例、规则修订和再测证据逐步形成。""")
]

make_docx(OUT/"初中物理知识图谱支持下的学情诊断与循证辅导_开题报告.docx",
          "初中物理知识图谱支持下的学情诊断与循证辅导",
          "校级课题开题报告（讨论稿）", opening_sections)
make_docx(OUT/"文献调研与研究设计核验说明.docx",
          "文献调研与研究设计核验说明",
          "配套材料", lit_sections)

readme = """下载包说明
==========
1. 初中物理知识图谱支持下的学情诊断与循证辅导_开题报告.docx
2. 文献调研与研究设计核验说明.docx

版本：2026-09-27
本包不含ChatGPT/Deep Research内部链接；正文引用采用普通编号。
本稿按“一人、一年期校级预研”控制范围，现有工具仅作为前期技术可行性基础。
"""
(OUT/"README.txt").write_text(readme, encoding="utf-8")
shutil.make_archive(str(OUT/"初中物理课题开题报告_下载包"), "zip", OUT)
print("built:", [p.name for p in OUT.iterdir()])

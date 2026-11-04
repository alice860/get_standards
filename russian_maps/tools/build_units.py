"""P3 批量单元生成器：tools/units/*.py 的数据 -> 每单元一张 HTML + 索引页 + 数据表 xlsx。

用法：python tools/build_units.py <russian_maps 目录> <用户标黄的覆盖统计.xlsx>
"""
import sys, os, html, re
sys.path.insert(0, os.path.dirname(__file__))
from units import t01, t02, t04, t05, t06
from units.common import PT
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT, MARKED = sys.argv[1], sys.argv[2]
OUTDIR = os.path.join(ROOT, "units")
os.makedirs(OUTDIR, exist_ok=True)
E = html.escape
UNITS = t01.UNITS + t02.UNITS + t04.UNITS + t05.UNITS + t06.UNITS
UNITS.sort(key=lambda u: u["id"])
THEMES = {"T01": "交流与信息传达", "T02": "思维与认知", "T03": "情感、态度与意愿", "T04": "行动：阶段、能力与实现", "T05": "存在、关系与变化（学术核心）",
          "T06": "社会互动", "T07": "拥有、给予与交换", "T08": "运动与位置", "T09": "日常生活与身体"}

def plain(s):
    return s.replace("́", "").replace("ё", "е")

def heads(card):
    s = plain(card["hw"])
    s = re.sub(r"\(([^)]*)\)", r" \1 ", s)
    return {w.strip() for w in re.split(r"[ /·,]+", s) if w.strip()}

# ---------------------------------------------------------- 覆盖核对：用户标黄的动词都要有卡片
ws = load_workbook(MARKED)["动词使用情况"]
marked = []
for r in ws.iter_rows(min_row=2):
    if r[0].fill.fill_type and r[0].fill.fgColor.rgb == "FFFFF7CC":
        marked.append((r[0].value, r[1].value or "", r[3].value, "学术高频（原有）" if r[5].value else "你新增"))
FIXPF = {"возратиться": "возвратиться", "разделиться": "поделиться", "приготовиться": "подготовиться"}
coverage, missing = [], []
for nsv, sv, unit, src in marked:
    found = None
    for u in UNITS:
        for _, cards in u["clusters"]:
            for c in cards:
                if c["kind"] == "full" and plain(nsv) in heads(c):
                    found = (u["id"], c["hw"]); break
            if found: break
        if found: break
    if not found:
        missing.append(nsv)
    coverage.append((nsv, sv, unit, src, found[0] if found else "", found[1] if found else "未找到"))
print("marked", len(marked), "missing", missing)

# ---------------------------------------------------------- HTML
base = open(os.path.join(ROOT, "govor_unit.html"), encoding="utf-8").read()
style_block = base[base.index("<link rel=\"preconnect\""):base.index("</style>")]
script = base[base.index("<script>"):]
EXTRA = """
.nav{display:flex;flex-wrap:wrap;gap:8px 16px;font-size:13px}.nav a{color:var(--aspect)}
.kind{font-size:10.5px;font-weight:700;border-radius:3px;padding:0 5px;letter-spacing:.03em}
.k-new{color:var(--cross);background:var(--cross-bg)}.k-short{color:var(--muted);background:var(--soft)}
.card.short,.card.new{gap:3px;padding:7px 10px}.card.new{border-style:dashed}
.pt{font-size:11px;font-weight:700;color:var(--aspect)}
.purpose{font-size:13.5px;margin-top:4px}
.tg{display:grid;gap:6px;grid-template-columns:repeat(auto-fit,minmax(min(100%,300px),1fr))}
.ix{display:grid;gap:16px}.ix .th{display:grid;gap:8px}.ix h2{font-family:var(--f-zh);font-size:16px}
.ix ul{display:grid;gap:6px;grid-template-columns:repeat(auto-fit,minmax(min(100%,320px),1fr))}
.ix li a{display:block;border:1px solid var(--line);border-radius:6px;padding:9px 11px;text-decoration:none;color:var(--ink);background:var(--sheet)}
.ix li a:hover,.ix li a:focus-visible{border-color:var(--aspect)}
.ix li b{font-family:var(--f-ru)}.ix li small{display:block;color:var(--muted);font-size:12px;margin-top:2px}
"""
PTAG = {"S": "体", "Y": "体", "I": "体", "L": "限", "Q": "起", "O": "一", "B": "两", "N": "无", "D": "定"}

def page_head(title):
    return f"<!doctype html>\n<html lang=\"zh\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n<title>{E(title)}</title>\n<meta name=\"claude:viewport-fit-sides\" content=\"cover\">\n" + style_block + EXTRA + "</style></head><body>\n"

def card_html(c):
    cls = {"full": "card", "short": "card short", "new": "card new"}[c["kind"]]
    badge = {"full": "", "short": '<span class="kind k-short">同单元</span>', "new": '<span class="kind k-new">表外补充</span>'}[c["kind"]]
    pt = f'<span class="pt">{E(PT[c["pt"]])}</span>' if c["pt"] else ""
    o = [f'<article class="{cls}"><div class="head"><span class="hw">{E(c["hw"])}</span>{pt}{badge}</div><p class="mean">{E(c["zh"])}</p><ul class="rel">']
    if c["gov"]: o.append(f'<li><i class="tag t-gov">支</i><span class="ru">{E(c["gov"])}</span></li>')
    if c["col"]: o.append(f'<li><i class="tag t-deriv">搭</i><span class="ru">{E(c["col"])}</span></li>')
    if c["der"]: o.append(f'<li><i class="tag t-deriv">派</i><span class="ru">{E(c["der"])}</span></li>')
    if c["trap"]: o.append(f'<li><i class="tag t-trap">⚠</i><span>{E(c["trap"])}</span></li>')
    if c["acad"]: o.append(f'<li><i class="tag t-acad">学</i><span class="ru">{E(c["acad"])}</span></li>')
    o.append("</ul>")
    if c["ex"]:
        ru, zh = (c["ex"].split(" || ") + [""])[:2]
        o.append(f'<p class="ex"><span class="ru">{E(ru)}</span><span class="zh">{E(zh)}</span></p>')
    o.append("</article>")
    return "".join(o)

def unit_page(u, prev, nxt):
    n_full = sum(1 for _, cs in u["clusters"] for c in cs if c["kind"] == "full")
    n_new = sum(1 for _, cs in u["clusters"] for c in cs if c["kind"] == "new")
    o = [page_head(f"{u['name']} · {u['id']}"), '<div class="wrap">\n']
    nav = '<nav class="nav" aria-label="单元导航"><a href="../units_index_page.html">← 全部单元</a>'
    if prev: nav += f'<a href="{prev["id"]}.html">上一单元：{E(prev["name"])}</a>'
    if nxt: nav += f'<a href="{nxt["id"]}.html">下一单元：{E(nxt["name"])}</a>'
    o.append(nav + "</nav>\n")
    o.append(f'<header class="top"><div><h1>{E(u["name"])}<small>{E(THEMES[u["id"][:3]])} · 单元图 {E(u["id"])} · 重点卡 {n_full} 张 · 表外补充 {n_new} 个</small></h1><p class="purpose">意图：{E(u["purpose"])}</p></div>')
    o.append('<div class="tools"><button type="button" id="quizBtn" aria-pressed="false">自测模式：遮住俄语</button></div></header>\n')
    o.append('<div class="legend"><span><i class="tag t-gov">支</i>支配</span><span><i class="tag t-deriv">搭</i>搭配</span><span><i class="tag t-deriv">派</i>派生词</span>'
             '<span><i class="tag t-trap">⚠</i>易错</span><span><i class="tag t-acad">学</i>学术 / 正式用法</span><span><span class="kind k-short">同单元</span>单元里的其他词（简写）</span>'
             '<span><span class="kind k-new">表外补充</span>原词表没有、按主题补进来的词</span><span>蓝字为体对类型</span></div>\n')
    o.append('<section class="sec" aria-labelledby="hA"><h2 id="hA"><span class="L">A</span>词卡 <em>按功能分簇；支配和学术用法写在卡片里</em></h2><div class="clusters">\n')
    for name, cards in u["clusters"]:
        o.append(f'<div class="cluster"><h3>{E(name)}</h3>' + "".join(card_html(c) for c in cards) + "</div>\n")
    o.append("</div></section>\n")
    if u["C"]:
        o.append('<section class="sec" aria-labelledby="hC"><h2 id="hC"><span class="L">C</span>关键辨析</h2><div class="tbl"><table><thead><tr><th>对比</th><th>区别</th><th>例</th></tr></thead><tbody>')
        for a, b, c in u["C"]:
            o.append(f'<tr><td class="hw">{E(a)}</td><td>{E(b)}</td><td class="ru">{E(c)}</td></tr>')
        o.append("</tbody></table></div></section>\n")
    if u["D"]:
        o.append('<section class="sec" aria-labelledby="hD"><h2 id="hD"><span class="L">D</span>易错与陷阱</h2><div class="traps"><div class="trapbox">')
        for t, bad, good in u["D"]:
            o.append(f'<div class="pair"><span class="zh">{E(t)}</span><span class="bad ru">{E(bad)}</span><span class="good ru">{E(good)}</span></div>')
        o.append("</div></div></section>\n")
    o.append('<div class="row2">\n<section class="sec" aria-labelledby="hG"><h2 id="hG"><span class="L">G</span>选词小决策 <em>自测模式下点击俄语看答案</em></h2><div class="intent">')
    for zh, ru in u["G"]:
        o.append(f'<div class="iv"><span class="zh">{E(zh)}</span><span class="arr">→</span><span class="q ru" tabindex="0">{E(ru)}</span></div>')
    o.append('</div></section>\n<section class="sec" aria-labelledby="hH"><h2 id="hH"><span class="L">H</span>输出自测 <em>先说或写，再看答案</em></h2><div class="drill">')
    for zh, ru in u["H"]:
        o.append(f'<div class="dr"><span class="zh">{E(zh)}</span><span class="q ru" tabindex="0">{E(ru)}</span></div>')
    o.append("</div></section>\n</div>\n")
    o.append('<footer class="foot"><p>核验状态：全部为“待核实”。词义、支配、体对和例句是按常规词典知识整理的教学材料，尚未逐条对照 Грамота.ру 或 НКРЯ；“表外补充”的词是按主题补入的，原词表中没有。</p>'
             '<p>本页由 tools/build_units.py 从 tools/units/ 的数据生成，数据同时写入 P3_单元数据.xlsx；修改内容请改数据再重新生成。打印按 A3 横向排版。</p></footer>\n</div>\n')
    o.append(script.replace("</script>", "</script>\n</body></html>", 1) if "</body>" not in script else script)
    return "".join(o)

for i, u in enumerate(UNITS):
    prev = UNITS[i - 1] if i > 0 else None
    nxt = UNITS[i + 1] if i < len(UNITS) - 1 else None
    with open(os.path.join(OUTDIR, f"{u['id']}.html"), "w", encoding="utf-8") as f:
        f.write(unit_page(u, prev, nxt))

# index（作为 artifact 主页面，不带 doctype，由发布时包裹）
tot_full = sum(1 for u in UNITS for _, cs in u["clusters"] for c in cs if c["kind"] == "full")
tot_new = sum(1 for u in UNITS for _, cs in u["clusters"] for c in cs if c["kind"] == "new")
ix = [f"<title>动词单元图谱</title>\n<meta name=\"claude:viewport-fit-sides\" content=\"cover\">\n", style_block + EXTRA + "</style>\n", '<div class="wrap">']
ix.append(f'<header class="top"><h1>动词单元图谱<small>P3 第一批 · {len(UNITS)} 个单元 · 重点卡 {tot_full} 张 · 表外补充 {tot_new} 个 · 覆盖你标黄的 {len(marked) - len(missing)} / {len(marked)} 个动词</small></h1>'
          '<p class="meta">每个单元一页，按表达意图组织：A 词卡、C 辨析、D 易错、G 选词、H 自测。已完成的 говор- 和 T11 学术单元另见单独的页面。</p></header>\n<div class="ix">')
for t, tname in THEMES.items():
    us = [u for u in UNITS if u["id"].startswith(t)]
    if not us: continue
    ix.append(f'<section class="th"><h2>{E(t)} {E(tname)}</h2><ul>')
    for u in us:
        nf = sum(1 for _, cs in u["clusters"] for c in cs if c["kind"] == "full")
        first = next(c["hw"] for _, cs in u["clusters"] for c in cs if c["kind"] == "full")
        ix.append(f'<li><a href="units/{u["id"]}.html">{E(u["id"])} {E(u["name"])}<small><b>{E(first)}</b> 等 {nf} 张重点卡</small></a></li>')
    ix.append("</ul></section>")
ix.append('</div><footer class="foot"><p>核验状态：全部为“待核实”。</p></footer></div>')
with open(os.path.join(ROOT, "units_index.html"), "w", encoding="utf-8") as f:
    f.write("".join(ix))


# 完整文档版的索引页，供单元页的“返回”链接使用（artifact 主页面本身不带 doctype）
_t = re.search(r"<title>.*?</title>", "".join(ix)).group(0)
_b = "".join(ix).replace(_t, "")
_style, _rest = _b.split("</style>", 1)
with open(os.path.join(ROOT, "units_index_page.html"), "w", encoding="utf-8") as f:
    f.write('<!doctype html>\n<html lang="zh"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">' + _t + _style + "</style></head><body>" + _rest + "</body></html>")

# ---------------------------------------------------------- XLSX
FN = "Arial"
HF = PatternFill("solid", fgColor="1F3A5F"); HFont = Font(name=FN, bold=True, color="FFFFFF", size=10); BF = Font(name=FN, size=10)
th = Side(style="thin", color="C9CED8"); BD = Border(left=th, right=th, top=th, bottom=th); WR = Alignment(wrap_text=True, vertical="top")
NEWF = PatternFill("solid", fgColor="FBF1E1"); MISS = PatternFill("solid", fgColor="FBE4E2")
wb = Workbook(); wb.remove(wb.active)

def sheet(name, hdr, wd, rows, fillcol=None):
    s = wb.create_sheet(name)
    for j, h in enumerate(hdr, 1):
        c = s.cell(row=1, column=j, value=h); c.font = HFont; c.fill = HF; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="center")
        s.column_dimensions[get_column_letter(j)].width = wd[j - 1]
    for i, r in enumerate(rows, 2):
        for j, v in enumerate(r, 1):
            c = s.cell(row=i, column=j, value=v); c.font = BF; c.alignment = WR; c.border = BD
            if fillcol is not None and fillcol(r): c.fill = fillcol(r)
    s.freeze_panes = "C2"; s.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{len(rows) + 1}"
    return s

notes = sheet("说明", ["项目", "内容"], [26, 110], [
    ("范围", f"你在覆盖统计表中标黄的 {len(marked)} 个动词（{sum(1 for m in marked if m[3] == '你新增')} 个你新增 + {sum(1 for m in marked if m[3] != '你新增')} 个原有的学术高频词），分布在 {len(UNITS)} 个单元。"),
    ("卡片类型", "重点卡 = 你标黄的词，写全义项、体对类型、支配、搭配、派生、易错、学术用法和例句；同单元 = 单元里没标黄的词，只写义项和支配；表外补充 = 原词表没有、按主题补入的相关词。"),
    ("各表", "3b_义项：所有卡片，一行一张；4_关系：C 区辨析；5_意图选词：G 区；6_陷阱：D 区；7_自测题：H 区；覆盖核对：每个标黄动词对应的单元和卡片。列结构与主工作簿一致，可以粘贴进去。"),
    ("核验状态", "全部为“待核实”。"),
    ("生成方式", "由 tools/build_units.py 从 tools/units/ 的数据生成；HTML 单元图在 units/ 目录。"),
])
KIND = {"full": "重点卡", "short": "同单元", "new": "表外补充"}
srows = []
k = 1
for u in UNITS:
    for cname, cards in u["clusters"]:
        for c in cards:
            ex_ru, ex_zh = ((c["ex"].split(" || ") + [""])[:2]) if c["ex"] else ("", "")
            srows.append((f"S{k:04d}", u["id"], u["name"], cname, KIND[c["kind"]], c["hw"], PT[c["pt"]], c["zh"], c["gov"], c["col"], c["der"], c["trap"], c["acad"], ex_ru, ex_zh,
                          "教学推论" if c["kind"] == "new" else "词典事实（待核）", "待核实"))
            k += 1
sheet("3b_义项", ["sense_id", "单元ID", "单元", "簇", "卡片类型", "词头", "体对类型", "中文义", "支配", "典型搭配", "派生词", "易错提示", "学术 / 正式用法", "例句（俄）", "例句（中）", "证据类型", "核验状态"],
      [8, 9, 18, 16, 9, 30, 12, 30, 30, 32, 26, 34, 30, 40, 26, 14, 9], srows, lambda r: NEWF if r[4] == "表外补充" else None)
sheet("4_关系", ["单元ID", "对比", "区别说明", "例"], [9, 36, 60, 56], [(u["id"], a, b, c) for u in UNITS for a, b, c in u["C"]])
sheet("5_意图选词", ["单元ID", "中文意图", "首选俄语"], [9, 30, 50], [(u["id"], zh, ru) for u in UNITS for zh, ru in u["G"]])
sheet("6_陷阱", ["单元ID", "类型", "错误说法", "正确说法"], [9, 14, 46, 56], [(u["id"], t, b, g) for u in UNITS for t, b, g in u["D"]])
sheet("7_自测题", ["单元ID", "中文提示", "参考答案", "我的答案", "结果", "错误类型（五类）", "日期"], [9, 40, 60, 30, 8, 14, 12], [(u["id"], zh, ru, "", "", "", "") for u in UNITS for zh, ru in u["H"]])
sheet("覆盖核对", ["НСВ（原表）", "СВ（原表）", "原表单元", "来源", "所在单元", "对应卡片"], [18, 18, 10, 14, 10, 40], coverage, lambda r: MISS if r[5] == "未找到" else None)
summ = wb.create_sheet("汇总", 1)
summ.column_dimensions["A"].width = 30; summ.column_dimensions["B"].width = 14
items = [("单元数", len(UNITS)), ("卡片总数", "=COUNTA('3b_义项'!A2:A5000)"), ("重点卡", "=COUNTIF('3b_义项'!E2:E5000,\"重点卡\")"),
         ("同单元", "=COUNTIF('3b_义项'!E2:E5000,\"同单元\")"), ("表外补充", "=COUNTIF('3b_义项'!E2:E5000,\"表外补充\")"),
         ("标黄动词", "=COUNTA('覆盖核对'!A2:A1000)"), ("其中找到卡片", "=COUNTA('覆盖核对'!F2:F1000)-COUNTIF('覆盖核对'!F2:F1000,\"未找到\")"),
         ("辨析条目", "=COUNTA('4_关系'!A2:A2000)"), ("陷阱条目", "=COUNTA('6_陷阱'!A2:A2000)"), ("意图条目", "=COUNTA('5_意图选词'!A2:A2000)"), ("自测题", "=COUNTA('7_自测题'!A2:A2000)")]
for i, (a, b) in enumerate(items, 1):
    summ.cell(row=i, column=1, value=a).font = Font(name=FN, bold=True, size=10); summ.cell(row=i, column=2, value=b).font = BF
wb.save(os.path.join(ROOT, "P3_单元数据.xlsx"))
print("units", len(UNITS), "cards", len(srows), "full", tot_full, "new", tot_new)

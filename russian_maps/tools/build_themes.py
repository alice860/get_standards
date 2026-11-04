"""主题图生成器：每个主题一张图（A 区按功能分组，G 区按意图选词，F 区功能 × 语域），外加索引页和数据表。

用法：python tools/build_themes.py <russian_maps 目录> <用户标黄的覆盖统计.xlsx>
数据：tools/units/t01…t06.py（各分组的词卡）+ tools/units/extra.py（补充分组和主题级内容）
"""
import sys, os, html, re, copy
sys.path.insert(0, os.path.dirname(__file__))
from units import t01, t02, t04, t05, t06, extra
from units.common import PT
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT, MARKED = sys.argv[1], sys.argv[2]
OUT = os.path.join(ROOT, "themes")
os.makedirs(OUT, exist_ok=True)
E = html.escape
LINKS = {"govor": "https://claude.ai/artifact/6pqnxHjjrGGnDoKom6kKQB", "T11-U01": "https://claude.ai/artifact/6A9BACCUy994KKQfobhtct",
         "T11-U02": "https://claude.ai/artifact/1tf1c2LiNgjqvScZJBhLmw"}

# ------------------------------------------------------------ 组装主题
G = {u["id"]: copy.deepcopy(u) for u in t01.UNITS + t02.UNITS + t04.UNITS + t05.UNITS + t06.UNITS
     + extra.T01_EXTRA + extra.T04_EXTRA + extra.T06_EXTRA + extra.T07_EXTRA + extra.T08_EXTRA}
G["T01-U07"]["clusters"].append(extra.T01_CONTACT)
G["T04-U03"]["name"] = "能力、尝试与追求"
G["T04-U03"]["purpose"] = "我想说能、会、试图、尽力、成功做到、追求目标"
G["T04-U03"]["clusters"].append(extra.STREMIT)
G["T05-U02"]["clusters"].append(extra.VKLYUCHAT)
G["T05-U04"]["clusters"][-1][1].append(extra.SKAZYVATSYA)
G["T05-U02"]["G"].append(("包括", "включать в себя"))
G["T04-U03"]["G"].append(("力求做到", "стремиться + 不定式"))
THEMES = []
for tid, m in extra.META.items():
    THEMES.append(dict(id=tid, **m, groups=[G[g] for g in m["order"]]))

def plain(s):
    return s.replace("́", "").replace("ё", "е")

def heads(card):
    s = re.sub(r"\(([^)]*)\)", r" \1 ", plain(card["hw"]))
    return {w.strip() for w in re.split(r"[ /·,]+", s) if w.strip()}

def cards(theme, kinds=("full", "short", "new")):
    for g in theme["groups"]:
        for cname, cs in g["clusters"]:
            for c in cs:
                if c["kind"] in kinds:
                    yield g, cname, c

# ------------------------------------------------------------ 覆盖核对
ws = load_workbook(MARKED)["动词使用情况"]
marked = [(r[0].value, r[1].value or "", r[3].value, "学术高频（原有）" if r[5].value else "你新增")
          for r in ws.iter_rows(min_row=2) if r[0].fill.fill_type and r[0].fill.fgColor.rgb == "FFFFF7CC"]
coverage, missing = [], []
for nsv, sv, unit, src in marked:
    hit = next(((t["id"], g["name"], c["hw"]) for t in THEMES for g, _, c in cards(t, ("full",)) if plain(nsv) in heads(c)), None)
    if not hit: missing.append(nsv)
    coverage.append((nsv, sv, unit, src) + (hit if hit else ("", "", "未找到")))
print("marked", len(marked), "missing", missing)

# ------------------------------------------------------------ HTML
base = open(os.path.join(ROOT, "govor_unit.html"), encoding="utf-8").read()
style = base[base.index("<link rel=\"preconnect\""):base.index("</style>")]
script = base[base.index("<script>"):]
EXTRA = """
.nav{display:flex;flex-wrap:wrap;gap:6px 14px;font-size:13px}.nav a{color:var(--aspect)}.nav b{color:var(--ink)}
.kind{font-size:10.5px;font-weight:700;border-radius:3px;padding:0 5px}
.k-new{color:var(--cross);background:var(--cross-bg)}.k-short{color:var(--muted);background:var(--soft)}
.card.short,.card.new{gap:3px;padding:7px 10px}.card.new{border-style:dashed}
.pt{font-size:11px;font-weight:700;color:var(--aspect)}
.grp{display:grid;gap:10px;padding:14px 0;border-top:1px solid var(--line)}.grp:first-of-type{border-top:0;padding-top:0}
.gh{font-family:var(--f-zh);font-size:15px;display:flex;gap:8px;align-items:baseline;flex-wrap:wrap}
.gh .no{font-family:var(--f-ru);color:var(--sheet);background:var(--aspect);border-radius:3px;padding:0 6px;font-size:13px}
.gp{font-size:12.5px;color:var(--muted)}
.grp .clusters{grid-template-columns:1fr;gap:12px}
.grp .cluster{grid-template-columns:repeat(auto-fill,minmax(min(100%,270px),1fr));gap:10px}
.grp .cluster>h3{grid-column:1/-1}
.grp .card{align-self:start;align-content:start}
@media print{.grp .cluster{grid-template-columns:repeat(5,1fr)}}
.gsub{font-family:var(--f-zh);font-size:12.5px;font-weight:700;color:var(--aspect);margin:10px 0 4px;grid-column:1/-1}
.toc{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px}.toc a{font-size:12.5px;border:1px solid var(--line);border-radius:4px;padding:2px 8px;color:var(--ink);text-decoration:none}
.ix ul{display:grid;gap:8px;grid-template-columns:repeat(auto-fit,minmax(min(100%,330px),1fr))}
.ix li a{display:block;border:1px solid var(--line);border-radius:6px;padding:10px 12px;text-decoration:none;color:var(--ink);background:var(--sheet)}
.ix li a:hover,.ix li a:focus-visible{border-color:var(--aspect)}.ix li small{display:block;color:var(--muted);font-size:12px;margin-top:3px}
.ix h2{font-family:var(--f-zh);font-size:16px;margin:8px 0}
"""
NO = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭"

def doc(title, body, with_script=True):
    return ("<!doctype html>\n<html lang=\"zh\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">"
            f"<title>{E(title)}</title><meta name=\"claude:viewport-fit-sides\" content=\"cover\">" + style + EXTRA + "</style></head><body>" + body
            + (script if with_script else "") + "</body></html>")

def card_html(c):
    cls = {"full": "card", "short": "card short", "new": "card new"}[c["kind"]]
    badge = {"full": "", "short": '<span class="kind k-short">同主题</span>', "new": '<span class="kind k-new">表外补充</span>'}[c["kind"]]
    pt = f'<span class="pt">{E(PT[c["pt"]])}</span>' if c["pt"] else ""
    o = [f'<article class="{cls}"><div class="head"><span class="hw">{E(c["hw"])}</span>{pt}{badge}</div><p class="mean">{E(c["zh"])}</p><ul class="rel">']
    for key, tag, tcls, ru in (("gov", "支", "t-gov", True), ("col", "搭", "t-deriv", True), ("der", "派", "t-deriv", True), ("trap", "⚠", "t-trap", False), ("acad", "学", "t-acad", True)):
        if c[key]:
            attr = ' class="ru"' if ru else ""
            o.append(f'<li><i class="tag {tcls}">{tag}</i><span{attr}>{E(c[key])}</span></li>')
    o.append("</ul>")
    if c["ex"]:
        ru, zh = (c["ex"].split(" || ") + [""])[:2]
        o.append(f'<p class="ex"><span class="ru">{E(ru)}</span><span class="zh">{E(zh)}</span></p>')
    return "".join(o) + "</article>"

def theme_page(t, i):
    o = ['<div class="wrap">']
    nav = '<nav class="nav" aria-label="主题导航"><a href="../index_page.html">← 全部主题</a>'
    for tt in THEMES:
        nav += f'<b>{E(tt["id"])} {E(tt["name"])}</b>' if tt is t else f'<a href="{tt["id"]}.html">{E(tt["id"])} {E(tt["name"])}</a>'
    o.append(nav + "</nav>")
    nf = sum(1 for _ in cards(t, ("full",))); ns = sum(1 for _ in cards(t, ("short",))); nn = sum(1 for _ in cards(t, ("new",)))
    o.append(f'<header class="top"><h1 class="ru">{E(t["h1"])}<small>{E(t["id"])} {E(t["name"])} · {len(t["groups"])} 个功能分组 · 重点卡 {nf} · 同主题 {ns} · 表外补充 {nn}</small></h1>'
             f'<p class="meta"><b>意图：</b>{E(t["purpose"])}。<b>读法：</b>A 区按功能分组，同一组是意思相近、要互相比较的词；G 区从中文意图出发选词；F 区看同一功能在不同场合的说法。</p>'
             '<div class="tools"><button type="button" id="quizBtn" aria-pressed="false">自测模式：遮住俄语</button></div></header>')
    o.append('<div class="legend"><span><i class="tag t-gov">支</i>支配</span><span><i class="tag t-deriv">搭</i>搭配</span><span><i class="tag t-deriv">派</i>派生词</span>'
             '<span><i class="tag t-trap">⚠</i>易错</span><span><i class="tag t-acad">学</i>学术 / 正式用法</span><span><span class="kind k-short">同主题</span>原词表中本主题的其他词（简写）</span>'
             '<span><span class="kind k-new">表外补充</span>原词表没有、为完整补入的词</span><span>蓝字为体对类型</span><span>核验状态：全部“待核实”</span></div>')
    # A
    toc = "".join(f'<a href="#g{k}">{NO[k]} {E(g["name"])}</a>' for k, g in enumerate(t["groups"]))
    o.append(f'<section class="sec" aria-labelledby="hA"><h2 id="hA"><span class="L">A</span>功能分组 <em>同组的词意思相近，要放在一起比较</em></h2><div class="toc">{toc}</div>')
    for k, g in enumerate(t["groups"]):
        o.append(f'<div class="grp" id="g{k}"><h3 class="gh"><span class="no">{NO[k]}</span>{E(g["name"])}</h3><p class="gp">意图：{E(g["purpose"])}</p><div class="clusters">')
        for cname, cs in g["clusters"]:
            o.append(f'<div class="cluster"><h3>{E(cname)}</h3>' + "".join(card_html(c) for c in cs) + "</div>")
        o.append("</div></div>")
    o.append("</section>")
    # B + C
    o.append('<div class="row2"><section class="sec" aria-labelledby="hB"><h2 id="hB"><span class="L">B</span>相关主题与单元</h2><div class="xlist">')
    for name, zh in t["B"]:
        key = "govor" if name.startswith("говор") else ("T11-U01" if name.startswith("T11") else None)
        label = f'<a href="{LINKS[key]}">{E(name)}</a>' if key else (f'<a href="{name[:3]}.html">{E(name)}</a>' if name[:3] in [x["id"] for x in THEMES] else E(name))
        o.append(f'<div class="x"><span class="hw">{label}</span><span class="zh">{E(zh)}</span></div>')
    o.append('</div></section><section class="sec" aria-labelledby="hC"><h2 id="hC"><span class="L">C</span>关键辨析 <em>先看跨分组的，再看组内的</em></h2><div class="tbl"><table><thead><tr><th>分组</th><th>对比</th><th>区别</th><th>例</th></tr></thead><tbody>')
    for a, b, c in t["C"]:
        o.append(f'<tr><td class="zh">跨分组</td><td class="hw">{E(a)}</td><td>{E(b)}</td><td class="ru">{E(c)}</td></tr>')
    for k, g in enumerate(t["groups"]):
        for a, b, c in g["C"]:
            o.append(f'<tr><td class="zh">{NO[k]}</td><td class="hw">{E(a)}</td><td>{E(b)}</td><td class="ru">{E(c)}</td></tr>')
    o.append("</tbody></table></div></section></div>")
    # D grouped by type
    dd = {}
    for typ, bad, good in t["D"]:
        dd.setdefault(typ, []).append((bad, good, "跨分组"))
    for k, g in enumerate(t["groups"]):
        for typ, bad, good in g["D"]:
            dd.setdefault(typ, []).append((bad, good, NO[k]))
    o.append('<section class="sec" aria-labelledby="hD"><h2 id="hD"><span class="L">D</span>易错与陷阱 <em>按错误类型归类，圈号是所在分组</em></h2><div class="traps">')
    for typ, items in dd.items():
        o.append(f'<div class="trapbox"><h3>{E(typ)}</h3>' + "".join(f'<div class="pair"><span class="bad ru">{E(b)}</span><span class="good ru">{E(gd)}</span><span class="zh">{E(w)}</span></div>' for b, gd, w in items) + "</div>")
    o.append("</div></section>")
    # E + F
    o.append('<div class="row2"><section class="sec" aria-labelledby="hE"><h2 id="hE"><span class="L">E</span>支配速查</h2><div class="tbl"><table><thead><tr><th>分组</th><th>词</th><th>支配</th></tr></thead><tbody>')
    for k, g in enumerate(t["groups"]):
        for cname, cs in g["clusters"]:
            for c in cs:
                if c["gov"]:
                    o.append(f'<tr><td class="zh">{NO[k]}</td><td class="hw">{E(c["hw"])}</td><td class="ru">{E(c["gov"])}</td></tr>')
    cols = "".join(f"<th>{E(x)}</th>" for x in t["F_cols"])
    o.append(f'</tbody></table></div></section><section class="sec" aria-labelledby="hF"><h2 id="hF"><span class="L">F</span>{"前缀 × 动作" if t["id"] == "T08" else "交际功能 × 语域"} '
             f'<em>{"同一前缀，意义在四类运动动词中一致" if t["id"] == "T08" else "同一行 = 同一功能，左右只差场合"}</em></h2><div class="tbl"><table><thead><tr>{cols}</tr></thead><tbody>')
    for row in t["F"]:
        o.append(f'<tr><td class="zh">{E(row[0])}</td>' + "".join(f'<td class="ru">{E(x)}</td>' for x in row[1:]) + "</tr>")
    o.append("</tbody></table></div></section></div>")
    # G + H
    o.append('<div class="row2"><section class="sec" aria-labelledby="hG"><h2 id="hG"><span class="L">G</span>选词小决策 <em>从中文意图出发；自测模式下点击俄语看答案</em></h2><div class="intent">')
    if t["G"]:
        o.append('<p class="gsub">跨分组：同一个中文词，按意图选不同的俄语词</p>')
        o += [f'<div class="iv"><span class="zh">{E(zh)}</span><span class="arr">→</span><span class="q ru" tabindex="0">{E(ru)}</span></div>' for zh, ru in t["G"]]
    for k, g in enumerate(t["groups"]):
        if g["G"]:
            o.append(f'<p class="gsub">{NO[k]} {E(g["name"])}</p>')
            o += [f'<div class="iv"><span class="zh">{E(zh)}</span><span class="arr">→</span><span class="q ru" tabindex="0">{E(ru)}</span></div>' for zh, ru in g["G"]]
    o.append('</div></section><section class="sec" aria-labelledby="hH"><h2 id="hH"><span class="L">H</span>输出自测 <em>先说或写，再看答案</em></h2><div class="drill">')
    for zh, ru in t["H"] + [x for g in t["groups"] for x in g["H"]]:
        o.append(f'<div class="dr"><span class="zh">{E(zh)}</span><span class="q ru" tabindex="0">{E(ru)}</span></div>')
    o.append("</div></section></div>")
    o.append('<footer class="foot"><p>核验状态：全部为“待核实”。词义、支配、体对和例句是按常规词典知识整理的教学材料，尚未逐条对照 Грамота.ру 或 НКРЯ。“表外补充”是为使主题完整而补入的词。</p>'
             '<p>本页由 tools/build_themes.py 从 tools/units/ 的数据生成，数据同时写入 P3_主题数据.xlsx。打印按 A3 横向排版。</p></footer></div>')
    return doc(f"{t['name']} · {t['id']}", "".join(o))

for i, t in enumerate(THEMES):
    open(os.path.join(OUT, f"{t['id']}.html"), "w", encoding="utf-8").write(theme_page(t, i))

def index_body(prefix):
    tot = sum(1 for t in THEMES for _ in cards(t)); nf = sum(1 for t in THEMES for _ in cards(t, ("full",)))
    o = ['<div class="wrap"><header class="top"><h1>动词主题图谱<small>'
         f'{len(THEMES)} 个主题 · 卡片 {tot} 张（重点卡 {nf}）· 覆盖你标黄的 {len(marked) - len(missing)} / {len(marked)} 个动词</small></h1>'
         '<p class="meta">每个主题一张图，和 говор- 词族图的结构一致：A 区按功能分组，C 辨析，D 易错，E 支配，F 功能 × 语域，G 从中文意图选词，H 自测。</p></header><div class="ix"><ul>']
    for t in THEMES:
        n = sum(1 for _ in cards(t))
        o.append(f'<li><a href="{prefix}{t["id"]}.html"><b>{E(t["id"])} {E(t["name"])}</b><small>{len(t["groups"])} 个功能分组 · {n} 张卡片 · {E(t["purpose"])}</small></a></li>')
    o.append('</ul><h2>已完成的单独单元</h2><ul>'
             f'<li><a href="{LINKS["govor"]}"><b>говор- 词族图</b><small>交流主题的核心词族</small></a></li>'
             f'<li><a href="{LINKS["T11-U01"]}"><b>T11-U01 证据与结论的强度表达</b><small>学术论证</small></a></li>'
             f'<li><a href="{LINKS["T11-U02"]}"><b>T11-U02 回应提问、质疑与异议</b><small>答辩互动</small></a></li></ul>'
             '<p class="foot">原 T03（情感）和 T09（日常生活）这次只标了一个词：стремиться 放进 T04“能力、尝试与追求”，включать（包括义）放进 T05“关系、属性与依存”。这两个主题以后可以单独做成完整的图。</p></div></div>')
    return "".join(o)

frag = "<title>动词主题图谱</title>\n<meta name=\"claude:viewport-fit-sides\" content=\"cover\">\n" + style + EXTRA + "</style>\n" + index_body("themes/")
open(os.path.join(ROOT, "themes_index.html"), "w", encoding="utf-8").write(frag)
open(os.path.join(ROOT, "index_page.html"), "w", encoding="utf-8").write(doc("动词主题图谱", index_body("themes/"), with_script=False))

# ------------------------------------------------------------ XLSX
FN = "Arial"; HF = PatternFill("solid", fgColor="1F3A5F"); HFont = Font(name=FN, bold=True, color="FFFFFF", size=10); BF = Font(name=FN, size=10)
th = Side(style="thin", color="C9CED8"); BD = Border(left=th, right=th, top=th, bottom=th); WR = Alignment(wrap_text=True, vertical="top")
NEWF = PatternFill("solid", fgColor="FBF1E1"); MISS = PatternFill("solid", fgColor="FBE4E2")
wb = Workbook(); wb.remove(wb.active)

def sheet(name, hdr, wd, rows, fill=None):
    s = wb.create_sheet(name)
    for j, h in enumerate(hdr, 1):
        c = s.cell(row=1, column=j, value=h); c.font = HFont; c.fill = HF; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical="center")
        s.column_dimensions[get_column_letter(j)].width = wd[j - 1]
    for i, r in enumerate(rows, 2):
        for j, v in enumerate(r, 1):
            c = s.cell(row=i, column=j, value=v); c.font = BF; c.alignment = WR; c.border = BD
            if fill and fill(r): c.fill = fill(r)
    s.freeze_panes = "C2"; s.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{max(len(rows), 1) + 1}"

sheet("说明", ["项目", "内容"], [24, 110], [
    ("结构", "每个主题一张图；原来的“单元”变成主题图 A 区里的功能分组。辨析、易错、意图选词、自测都在主题层面汇总。"),
    ("范围", f"覆盖你标黄的 {len(marked)} 个动词；每个主题还包含原词表中本主题的其他词（同主题，简写）和为完整补入的词（表外补充）。"),
    ("核验状态", "全部为“待核实”。"),
    ("生成方式", "tools/build_themes.py 从 tools/units/ 的数据生成；HTML 在 themes/ 目录。")])
KIND = {"full": "重点卡", "short": "同主题", "new": "表外补充"}
rows = []
for t in THEMES:
    for k, g in enumerate(t["groups"]):
        for cname, cs in g["clusters"]:
            for c in cs:
                ru, zh = ((c["ex"].split(" || ") + [""])[:2]) if c["ex"] else ("", "")
                rows.append((t["id"], t["name"], f"{NO[k]} {g['name']}", cname, KIND[c["kind"]], c["hw"], PT[c["pt"]], c["zh"], c["gov"], c["col"], c["der"], c["trap"], c["acad"], ru, zh, "待核实"))
sheet("3b_词卡", ["主题ID", "主题", "功能分组", "簇", "卡片类型", "词头", "体对类型", "中文义", "支配", "搭配", "派生词", "易错", "学术 / 正式用法", "例句（俄）", "例句（中）", "核验状态"],
      [7, 16, 22, 18, 9, 30, 12, 30, 30, 30, 24, 32, 28, 40, 26, 9], rows, lambda r: NEWF if r[4] == "表外补充" else None)
def grp_rows(key, conv):
    out = []
    for t in THEMES:
        for x in t[key]: out.append((t["id"], "跨分组") + conv(x))
        for k, g in enumerate(t["groups"]):
            for x in g[key]: out.append((t["id"], f"{NO[k]} {g['name']}") + conv(x))
    return out
sheet("4_辨析", ["主题ID", "分组", "对比", "区别", "例"], [7, 22, 34, 56, 50], grp_rows("C", lambda x: tuple(x)))
sheet("5_意图选词", ["主题ID", "分组", "中文意图", "俄语"], [7, 22, 30, 50], grp_rows("G", lambda x: tuple(x)))
sheet("6_陷阱", ["主题ID", "分组", "类型", "错误说法", "正确说法"], [7, 22, 14, 42, 52], grp_rows("D", lambda x: tuple(x)))
sheet("7_自测题", ["主题ID", "分组", "中文提示", "参考答案", "我的答案", "结果", "错误类型（五类）"], [7, 22, 40, 56, 26, 7, 14], grp_rows("H", lambda x: tuple(x) + ("", "", "")))
sheet("F_功能语域", ["主题ID", "功能", "列 1", "列 2", "列 3", "列 4", "列 5"], [7, 12, 26, 26, 30, 30, 26], [(t["id"],) + tuple(r) for t in THEMES for r in t["F"]])
sheet("覆盖核对", ["НСВ（原表）", "СВ（原表）", "原表单元", "来源", "所在主题", "所在分组", "对应卡片"], [18, 18, 10, 14, 8, 26, 36], coverage, lambda r: MISS if r[6] == "未找到" else None)
s = wb.create_sheet("汇总", 1); s.column_dimensions["A"].width = 28; s.column_dimensions["B"].width = 12
for i, (a, b) in enumerate([("主题数", len(THEMES)), ("卡片总数", "=COUNTA('3b_词卡'!A2:A5000)"), ("重点卡", "=COUNTIF('3b_词卡'!E2:E5000,\"重点卡\")"),
                            ("同主题", "=COUNTIF('3b_词卡'!E2:E5000,\"同主题\")"), ("表外补充", "=COUNTIF('3b_词卡'!E2:E5000,\"表外补充\")"),
                            ("标黄动词", "=COUNTA('覆盖核对'!A2:A1000)"), ("其中找到卡片", "=COUNTA('覆盖核对'!G2:G1000)-COUNTIF('覆盖核对'!G2:G1000,\"未找到\")"),
                            ("辨析", "=COUNTA('4_辨析'!A2:A3000)"), ("陷阱", "=COUNTA('6_陷阱'!A2:A3000)"), ("意图", "=COUNTA('5_意图选词'!A2:A3000)"), ("自测题", "=COUNTA('7_自测题'!A2:A3000)")], 1):
    s.cell(row=i, column=1, value=a).font = Font(name=FN, bold=True, size=10); s.cell(row=i, column=2, value=b).font = BF
wb.save(os.path.join(ROOT, "P3_主题数据.xlsx"))
for t in THEMES:
    print(t["id"], len(t["groups"]), sum(1 for _ in cards(t, ("full",))), sum(1 for _ in cards(t, ("short",))), sum(1 for _ in cards(t, ("new",))))

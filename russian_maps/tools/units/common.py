"""单元数据的公共结构。

卡片字段：
  hw    词头（带重音），体对写成 “НСВ / СВ”
  pt    体对类型代码：S 严格体对应 · Y 异根体对应 · I 按义项体对应 · L 限定（по-）
        Q 起始（за-）· O 一次（-нуть）· B 两体 · N 无对 · D 定向/不定向 · '' 非动词
  zh    中文义（多个义项用 ①②）
  gov   支配
  col   典型搭配
  der   派生词
  trap  易错提示
  ex    例句，“俄语 || 中文”
  acad  学术或正式用法句块
  kind  full（重点卡）/ short（同单元其他成员）/ new（表外补充词）
"""

PT = {"S": "严格体对应", "Y": "异根体对应", "I": "按义项体对应", "L": "限定（по-）", "Q": "起始（за-/по-）",
      "O": "一次（-нуть）", "B": "两体", "N": "无对", "D": "定向 / 不定向", "": ""}


def V(hw, pt, zh, gov="", col="", der="", trap="", ex="", acad=""):
    return dict(kind="full", hw=hw, pt=pt, zh=zh, gov=gov, col=col, der=der, trap=trap, ex=ex, acad=acad)


def S(hw, pt, zh, gov=""):
    return dict(kind="short", hw=hw, pt=pt, zh=zh, gov=gov, col="", der="", trap="", ex="", acad="")


def N(hw, pt, zh, gov="", acad=""):
    return dict(kind="new", hw=hw, pt=pt, zh=zh, gov=gov, col="", der="", trap="", ex="", acad=acad)


def U(uid, name, purpose, clusters, C=(), D=(), G=(), H=(), note=""):
    return dict(id=uid, name=name, purpose=purpose, clusters=list(clusters), C=list(C), D=list(D), G=list(G), H=list(H), note=note)

"""厚労省「薬価基準収載品目リスト及び後発医薬品に関する情報」の最新Excelを data/ に取得する。"""
import json, re, sys, time, urllib.request
from datetime import datetime, timezone, timedelta
from html import unescape
from pathlib import Path
from urllib.parse import urljoin

INDEX = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000078916.html"
FALLBACK = "https://www.mhlw.go.jp/topics/2026/04/tp20260401-01.html"
# 診療報酬情報提供サービス「医薬品マスター」（全件）: 銘柄別の品名・カナ・一般名コード
YMASTER = "https://shinryohoshu.mhlw.go.jp/shinryohoshu/downloadMenu/yFile"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "ja,en;q=0.8"}
DATA = Path(__file__).parent / "data"
JST = timezone(timedelta(hours=9))


def get(url, tries=3):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
                return r.read()
        except Exception as e:  # noqa
            print(f"retry {i+1}: {url} {e}", file=sys.stderr)
            time.sleep(5 * (i + 1))
    raise SystemExit(f"取得失敗: {url}")


def text(url):
    b = get(url)
    for enc in ("utf-8", "cp932"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode("cp932", "replace")


def candidates():
    out = []
    try:
        html = text(INDEX)
        for href, text in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, re.S):
            t = re.sub(r"<[^>]+>", "", text)
            if "薬価基準収載品目リスト" in t and "まで" not in t:
                out.append(urljoin(INDEX, href))
    except SystemExit:
        pass
    year = datetime.now(JST).year
    for y in (year, year - 1):
        out.append(f"https://www.mhlw.go.jp/topics/{y}/04/tp{y}0401-01.html")
    out.append(FALLBACK)
    return list(dict.fromkeys(out))


def xlsx_links(page, html):
    urls = {}
    for href, n in re.findall(r"""href=["']([^"']+?_0([1-5])\.xlsx)["']""", html):
        urls.setdefault(n, urljoin(page, href))   # 目次（先頭側）の最初のリンクが最新版
    return urls


def main():
    page = html = urls = None
    for c in candidates():
        try:
            h = text(c)
        except SystemExit:
            continue
        u = xlsx_links(c, h)
        print("candidate", c, len(h), sorted(u))
        if set(u) == set("12345"):
            page, html, urls = c, h, u
            break
    if not page:
        raise SystemExit("Excelリンクが揃うページが見つかりません")
    title = re.search(r"<title>(.*?)</title>", html, re.S)
    title = unescape(title.group(1)).strip() if title else ""
    m = re.search(r"（(令和[^）]*適用)）", title)
    applied = m.group(1) if m else ""
    DATA.mkdir(exist_ok=True)
    for n, u in sorted(urls.items()):
        b = get(u)
        if b[:2] != b"PK":
            raise SystemExit(f"Excelではありません: {u}")
        (DATA / f"0{n}.xlsx").write_bytes(b)
        print(n, u, len(b))
    try:
        b = get(YMASTER)
        if b[:2] == b"PK":
            (DATA / "y.zip").write_bytes(b)
            print("y master", len(b))
        else:
            print("医薬品マスターがzipではありません", b[:80], file=sys.stderr)
    except SystemExit as e:   # 取れない日は前回分を使う
        print(e, file=sys.stderr)
    old = {}
    sp = DATA / "source.json"
    if sp.exists():
        old = json.loads(sp.read_text(encoding="utf-8"))
    new = {"page": page, "title": title, "applied": applied, "urls": urls}
    if {k: old.get(k) for k in new} != new:
        new["fetched"] = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
        sp.write_text(json.dumps(new, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()

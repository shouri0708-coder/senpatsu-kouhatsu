"""厚労省「薬価基準収載品目リスト及び後発医薬品に関する情報」の最新Excelを data/ に取得する。"""
import json, re, sys, time, urllib.request
from datetime import datetime, timezone, timedelta
from html import unescape
from pathlib import Path
from urllib.parse import urljoin

INDEX = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000078916.html"
FALLBACK = "https://www.mhlw.go.jp/topics/2026/04/tp20260401-01.html"
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


def find_list_page():
    try:
        html = get(INDEX).decode("utf-8", "replace")
        for href, text in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', html, re.S):
            t = re.sub(r"<[^>]+>", "", text)
            if "薬価基準収載品目リスト" in t and "まで" not in t:
                return urljoin(INDEX, href)
    except SystemExit:
        pass
    return FALLBACK


def main():
    page = find_list_page()
    html = get(page).decode("utf-8", "replace")
    title = re.search(r"<title>(.*?)</title>", html, re.S)
    title = unescape(title.group(1)).strip() if title else ""
    m = re.search(r"（(令和[^）]*適用)）", title)
    applied = m.group(1) if m else ""
    links = re.findall(r'href="([^"]+?_0([1-5])\.xlsx)"', html)
    urls = {}
    for href, n in links:          # 目次（ページ先頭側）の最初のリンクが最新版
        urls.setdefault(n, urljoin(page, href))
    if set(urls) != set("12345"):
        raise SystemExit(f"Excelリンクが揃いません: {urls}")
    DATA.mkdir(exist_ok=True)
    for n, u in sorted(urls.items()):
        b = get(u)
        if b[:2] != b"PK":
            raise SystemExit(f"Excelではありません: {u}")
        (DATA / f"0{n}.xlsx").write_bytes(b)
        print(n, u, len(b))
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

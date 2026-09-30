# 先発・後発 検索

品名（先発でも後発でも）や成分名を入力すると、同じ成分・剤形・規格の先発品と後発品が並んで表示されるサイト。

- 公開URL: https://shouri0708-coder.github.io/senpatsu-kouhatsu/
- データ: 厚生労働省「薬価基準収載品目リスト及び後発医薬品に関する情報」（先発・後発の区分）＋ 診療報酬情報提供サービス「医薬品マスター」（銘柄別の品名・カナ・薬価・選定療養）
- 更新: GitHub Actions `update` が毎日 7:17(JST) に `fetch.py` → `build.py` を実行し、変更があれば `docs/index.html` を更新
- 画面の修正は `template.html`、データの作り方は `build.py`

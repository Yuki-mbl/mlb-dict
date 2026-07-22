# build/ — アプリのビルド一式

`index.html` を生成するためのソースです。**元の Word (.docx) は不要**。

## 生成方法
```
python build/generate_app.py
```
- `build/ej_data.json`（英和2491）/ `build/je_data.json`（和英1111）を素材に読み込み、
  リポジトリ直下の `index.html` を書き出します。
- 辞書の中身を直す時は `ej_data.json` / `je_data.json` を編集。

## データ復旧
万一 JSON を失っても、`index.html` 内の `const EJ = [...];` / `const JE = [...];`
（json.dumps 出力＝そのまま有効なJSON）から再抽出できます。

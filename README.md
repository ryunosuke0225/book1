# book1

Microsoft Word のコメント機能で記載された内容を Excel に抽出するサンプルです。

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使い方

```bash
python scripts/extract_comments_to_excel.py input.docx output.xlsx
```

出力される Excel には以下の列が含まれます。

- Comment ID
- Author
- Date
- Comment
- Commented Text (コメント範囲の本文)

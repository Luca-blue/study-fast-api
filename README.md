# study-fast-api

FastAPI 入門のハンズオン資料と、スライドで使うサンプルコードです。

## スライド

`slides/fastapi-intro.html` をブラウザで開いてください（35枚・1ファイル完結）。

- `←` `→` / `Space` / マウスホイール / スワイプでページ送り、`Home` `End` で最初・最後へ
- URL の `#13` のように番号を付けると、そのスライドから開く
- `E` キー（または画面左上にマウスを乗せる）で編集モード。文字をクリックして直接書き換え、`Ctrl/Cmd + S` で HTML を保存
- 印刷（PDF 保存）すると 1 スライド 1 ページで出力

| 章 | 内容 |
|---|---|
| 01 | Web API のしくみ（クライアント・サーバー・DB、リクエスト/レスポンス、JSON、ステータスコード） |
| 02 | FastAPI の立ち位置（Django / Flask との比較、特徴） |
| 03 | 準備とインポート（fastapi・uvicorn・pydantic・sqlite3 の役目、インストール） |
| 04 | 書いてみる（最小コード、起動、パス/クエリパラメータ、トラブル対処） |
| 05 | HTTP メソッド（GET / POST / PUT / DELETE を ToDo アプリで） |
| 06 | データを受け取る（Pydantic、メモリ版 ToDo API、/docs） |
| 07 | 画面をつなぐ（fetch、StaticFiles、メモリ版 API を画面から操作） |
| 08 | SQLite につなぐ（SQL と HTTP メソッドの対応、DB 版 API、完成形） |

## サンプルコード

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install fastapi uvicorn

cd examples/01_hello
uvicorn main:app --reload        # → http://127.0.0.1:8000/ と /docs
```

| フォルダ | 内容 | 対応する章 |
|---|---|---|
| `examples/01_hello` | いちばん小さな API | 4章 |
| `examples/02_params` | パスパラメータ・クエリパラメータ | 4章 |
| `examples/03_todo_memory` | Pydantic + メモリに保存する ToDo API（CRUD） | 6章 |
| `examples/04_todo_frontend` | メモリ版 API + 画面（`static/index.html`） | 7章 |
| `examples/05_todo_sqlite` | SQLite に保存する版（画面は 04 と同じ） | 8章 |

`04` と `05` は起動後に http://127.0.0.1:8000/ を開くと ToDo 画面が表示されます。

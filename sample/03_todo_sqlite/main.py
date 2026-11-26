# SQLite（ファイル1つで動くデータベース）を使うための部品。Python に最初から入っている
import sqlite3

# FastAPI 本体
from fastapi import FastAPI
# フォルダの HTML などを、そのままブラウザに配るための部品
from fastapi.staticfiles import StaticFiles
# 受け取る JSON の形（項目名と型）を決めるための部品
from pydantic import BaseModel

# アプリ本体を作る（`uvicorn main:app` の app はこれ）
app = FastAPI()


# 受け取るデータの形（02_todo と同じ）
class TodoIn(BaseModel):
    title: str           # やることの内容（必須）
    done: bool = False   # 完了したか。送られてこなければ False（未完了）


# ===== データベースの準備 =====
# todo.db につなぐ。ファイルが無ければ、つないだときに自動で作られる
def get_conn():
    conn = sqlite3.connect("todo.db")
    conn.row_factory = sqlite3.Row  # row["title"] のように列名で取り出せる
    return conn


# 起動したときに1回だけ実行される部分。todos 表が無ければ作る（IF NOT EXISTS: すでにあれば何もしない）
# with get_conn() as conn: のブロックを抜けると、中で行った変更が保存（コミット）される
with get_conn() as conn:
    conn.execute(
        "CREATE TABLE IF NOT EXISTS todos ("
        # id: 1, 2, 3… と自動で振られる番号。PRIMARY KEY は「この列で1行を見分ける」という意味
        " id INTEGER PRIMARY KEY AUTOINCREMENT,"
        # title: やることの内容。NOT NULL で空（値なし）を禁止する
        " title TEXT NOT NULL,"
        # done: 完了なら 1、未完了なら 0。SQLite には真偽値の型が無いので整数で持つ
        " done INTEGER NOT NULL DEFAULT 0)"
    )


# ===== API（5章の関数を SQL に置き換えたもの） =====
# GET /todos : 一覧を返す
@app.get("/todos")
def list_todos():
    with get_conn() as conn:
        # fetchall() で、結果の行を全部リストとして受け取る
        rows = conn.execute("SELECT id, title, done FROM todos").fetchall()
    # データベースの行を、画面が受け取れる辞書のリストに直す
    todo_list = []
    for row in rows:
        # done は 0 / 1 で入っているので、bool() で False / True に戻す
        todo_list.append({"id": row["id"], "title": row["title"], "done": bool(row["done"])})
    return todo_list


# POST /todos : 新しく作る
@app.post("/todos")
def create_todo(todo: TodoIn):
    with get_conn() as conn:
        # ? の所に、後ろのタプルの値が順番に入る。
        # 値を SQL の文字列に直接つなげず ? を使うと、入力に記号が含まれていても安全に扱える
        cur = conn.execute(
            "INSERT INTO todos (title, done) VALUES (?, ?)",
            (todo.title, todo.done),
        )
    # cur.lastrowid は、今 INSERT した行に振られた id
    return {"id": cur.lastrowid, "title": todo.title, "done": todo.done}


# PUT /todos/{todo_id} : 書き換える
# todo_id は URL から、todo（新しい title と done）はリクエストの本文から受け取る
@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    with get_conn() as conn:
        # WHERE id = ? で、書き換える行を1つに絞る（付け忘れると全部の行が書き換わる）
        # 無い id を指定した場合は、どの行も変わらず、エラーにもならない
        conn.execute(
            "UPDATE todos SET title = ?, done = ? WHERE id = ?",
            (todo.title, todo.done, todo_id),
        )
    return {"id": todo_id, "title": todo.title, "done": todo.done}


# DELETE /todos/{todo_id} : 消す
@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    with get_conn() as conn:
        # (todo_id,) の最後のカンマは「要素が1つのタプル」を作るために必要
        conn.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
    # 無い id を指定しても、何も消さずに同じ値を返す（エラーにはならない）
    return {"ok": True}


# ===== 画面 ../frontend/index.html を配信する =====
# "/" 以下へのアクセスで、隣の frontend フォルダ（../frontend）のファイルを返す。02_todo と 03_todo_sqlite で同じ画面を使う。
# ../frontend は「起動したフォルダ」から見た場所なので、main.py のあるフォルダで uvicorn を起動する。
# html=True で / に来たら index.html を返す。
# app.mount は最後に置く。先に置くと "/" がすべてのアクセスを受け取ってしまい、上の /todos に届かなくなる。
# 画面を選ぶ: 使う方の行の先頭の # を外し、もう一方の行に # を付ける
app.mount("/", StaticFiles(directory="../frontend", html=True), name="static")  # シンプル版
# app.mount("/", StaticFiles(directory="../frontend/notebook", html=True), name="static")  # 手帳風
# app.mount("/", StaticFiles(directory="../frontend/hum", html=True), name="static")  # Hum 版

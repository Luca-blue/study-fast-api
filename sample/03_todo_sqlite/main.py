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
    """画面から送られてくる ToDo 1件分のデータ。

    Attributes:
        title: やることの内容。必須。
        done: 完了したかどうか。送られてこなければ False（未完了）になる。
    """

    title: str
    done: bool = False


# ===== データベースの準備 =====
def get_conn():
    """データベースにつなぐ。

    Returns:
        sqlite3.Connection: todo.db への接続。ファイルが無ければ、つないだときに自動で作られる。
    """
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
    """ToDo を全部返す。

    Returns:
        list[dict]: ToDo の一覧。データベースの行を、画面が受け取れる辞書の形に直して返す。
    """
    with get_conn() as conn:
        # fetchall() で、結果の行を全部リストとして受け取る
        rows = conn.execute("SELECT id, title, done FROM todos").fetchall()
    todo_list = []
    for row in rows:
        # done は 0 / 1 で入っているので、bool() で False / True に戻す
        todo_list.append({"id": row["id"], "title": row["title"], "done": bool(row["done"])})
    return todo_list


# POST /todos : 新しく作る
@app.post("/todos")
def create_todo(todo: TodoIn):
    """ToDo を1件追加する。

    Args:
        todo: 画面から送られてきた ToDo（title と done）。

    Returns:
        dict: 追加した ToDo（データベースが振った id 付き）。
    """
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
@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    """指定した id の ToDo を、送られてきた内容で置き換える。

    Args:
        todo_id: 書き換える ToDo の id（URL から受け取る）。
        todo: 新しい内容（title と done。リクエストの本文から受け取る）。

    Returns:
        dict: 書き換えた後の ToDo。
    """
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
    """指定した id の ToDo を削除する。

    Args:
        todo_id: 削除する ToDo の id。

    Returns:
        dict: {"ok": True}。無い id を指定しても、何も消さずに同じ値を返す。
    """
    with get_conn() as conn:
        # (todo_id,) の最後のカンマは「要素が1つのタプル」を作るために必要
        conn.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
    return {"ok": True}


# ===== 画面 ../frontend/index.html を配信する =====
# "/" 以下へのアクセスで、隣の frontend フォルダ（../frontend）のファイルを返す。02_todo と 03_todo_sqlite で同じ画面を使う。
# ../frontend は「起動したフォルダ」から見た場所なので、main.py のあるフォルダで uvicorn を起動する。
# html=True で / に来たら index.html を返す。
# この行は最後に置く。先に置くと "/" がすべてのアクセスを受け取ってしまい、上の /todos に届かなくなる。
app.mount("/", StaticFiles(directory="../frontend", html=True), name="static")

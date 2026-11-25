import sqlite3

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

app = FastAPI()


class TodoIn(BaseModel):
    title: str
    done: bool = False


# ===== データベースの準備 =====
def get_conn():
    conn = sqlite3.connect("todo.db")
    conn.row_factory = sqlite3.Row  # row["title"] のように列名で取り出せる
    return conn


with get_conn() as conn:
    conn.execute(
        "CREATE TABLE IF NOT EXISTS todos ("
        " id INTEGER PRIMARY KEY AUTOINCREMENT,"
        " title TEXT NOT NULL,"
        " done INTEGER NOT NULL DEFAULT 0)"
    )


# ===== API（5章の関数を SQL に置き換えたもの） =====
@app.get("/todos")
def list_todos():
    with get_conn() as conn:
        rows = conn.execute("SELECT id, title, done FROM todos").fetchall()
    todo_list = []
    for row in rows:
        todo_list.append({"id": row["id"], "title": row["title"], "done": bool(row["done"])})
    return todo_list


@app.post("/todos")
def create_todo(todo: TodoIn):
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO todos (title, done) VALUES (?, ?)",
            (todo.title, todo.done),
        )
    return {"id": cur.lastrowid, "title": todo.title, "done": todo.done}


@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    with get_conn() as conn:
        conn.execute(
            "UPDATE todos SET title = ?, done = ? WHERE id = ?",
            (todo.title, todo.done, todo_id),
        )
    return {"id": todo_id, "title": todo.title, "done": todo.done}


@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    with get_conn() as conn:
        conn.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
    return {"ok": True}


# ===== 画面 static/index.html を配信する =====
app.mount("/", StaticFiles(directory="static", html=True), name="static")

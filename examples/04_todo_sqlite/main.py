import sqlite3

from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

DB_PATH = "todo.db"

app = FastAPI()


# --- データベースの準備 -------------------------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # 行を row["title"] のように取り出せる
    return conn


with get_conn() as conn:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS todos (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT    NOT NULL,
            done  INTEGER NOT NULL DEFAULT 0
        )
        """
    )


def to_dict(row):
    return {"id": row["id"], "title": row["title"], "done": bool(row["done"])}


class TodoIn(BaseModel):
    title: str
    done: bool = False


# --- API -------------------------------------------------------------
@app.get("/todos")
def list_todos():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM todos").fetchall()
    return [to_dict(r) for r in rows]


@app.post("/todos", status_code=201)
def create_todo(todo: TodoIn):
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO todos (title, done) VALUES (?, ?)",
            (todo.title, todo.done),
        )
    return {"id": cur.lastrowid, **todo.model_dump()}


@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    with get_conn() as conn:
        cur = conn.execute(
            "UPDATE todos SET title = ?, done = ? WHERE id = ?",
            (todo.title, todo.done, todo_id),
        )
    if cur.rowcount == 0:
        raise HTTPException(status_code=404, detail="Todo not found")
    return {"id": todo_id, **todo.model_dump()}


@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
    if cur.rowcount == 0:
        raise HTTPException(status_code=404, detail="Todo not found")
    return {"ok": True}


# --- 画面（static/index.html）を配信する ※API より後ろに書く --------
app.mount("/", StaticFiles(directory="static", html=True), name="static")

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

app = FastAPI()


# 受け取るデータの形
class TodoIn(BaseModel):
    title: str
    done: bool = False


# ToDo の置き場所  例: {1: {"id": 1, "title": "牛乳を買う", "done": False}}
todos = {}


# 次に使う id を返す
def new_id():
    return max(todos, default=0) + 1


# ===== ここから下に API の関数を書く（5章） =====
@app.get("/todos")
def list_todos():
    return list(todos.values())


@app.post("/todos")
def create_todo(todo: TodoIn):
    todo_id = new_id()
    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}
    return todos[todo_id]


@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}
    return todos[todo_id]


@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    del todos[todo_id]
    return {"ok": True}


# ===== この行より上に書く（下の行は画面 static/index.html を配信する） =====
app.mount("/", StaticFiles(directory="static", html=True), name="static")

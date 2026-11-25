from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

app = FastAPI()


# 受け取るデータの「形」を決める
class TodoIn(BaseModel):
    title: str
    done: bool = False


# データ置き場（メモリ上の辞書）: {id: {"id": 1, "title": "...", "done": False}}
todos = {}


@app.get("/todos")
def list_todos():
    return list(todos.values())


@app.get("/todos/{todo_id}")
def get_todo(todo_id: int):
    if todo_id not in todos:
        raise HTTPException(status_code=404, detail="Todo not found")
    return todos[todo_id]


@app.post("/todos", status_code=201)
def create_todo(todo: TodoIn):
    new_id = max(todos, default=0) + 1
    todos[new_id] = {"id": new_id, **todo.model_dump()}
    return todos[new_id]


@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    if todo_id not in todos:
        raise HTTPException(status_code=404, detail="Todo not found")
    todos[todo_id] = {"id": todo_id, **todo.model_dump()}
    return todos[todo_id]


@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    if todo_id not in todos:
        raise HTTPException(status_code=404, detail="Todo not found")
    del todos[todo_id]
    return {"ok": True}


# --- 画面（static/index.html）を配信する ※API より後ろに書く --------
app.mount("/", StaticFiles(directory="static", html=True), name="static")

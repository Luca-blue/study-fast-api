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



# ===== この行より上に書く（下の行は画面 static/index.html を配信する） =====
app.mount("/", StaticFiles(directory="static", html=True), name="static")

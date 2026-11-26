# FastAPI 本体
from fastapi import FastAPI
# フォルダの HTML などを、そのままブラウザに配るための部品
from fastapi.staticfiles import StaticFiles
# 受け取る JSON の形（項目名と型）を決めるための部品
from pydantic import BaseModel

# アプリ本体を作る（`uvicorn main:app` の app はこれ）
app = FastAPI()


# 受け取るデータの形
# POST や PUT で送られてきた JSON がこの形に合っているかを、FastAPI が自動で確かめる。
# 合わない（例: title が無い）と、関数を呼ぶ前に 422 エラーを返す。
class TodoIn(BaseModel):
    """画面から送られてくる ToDo 1件分のデータ。

    Attributes:
        title: やることの内容。必須。
        done: 完了したかどうか。送られてこなければ False（未完了）になる。
    """

    title: str
    done: bool = False


# ToDo の置き場所  例: {1: {"id": 1, "title": "牛乳を買う", "done": False}}
# キーが id、値が ToDo 1件分の辞書。
# メモリ上にあるだけなので、サーバーを再起動すると（--reload でファイルを保存したときも）空に戻る。
todos = {}


# 次に使う id を返す
def new_id():
    """まだ使われていない id を返す。

    Returns:
        int: 今ある id のいちばん大きい数 + 1。todos が空なら 1。
            max(todos) は辞書のキー（id）の最大値。default=0 は空のときの値。
    """
    return max(todos, default=0) + 1


# ===== ここから下に API の関数を書く（5章） =====



# ===== この行より上に書く（下の行は画面を配信する） =====
# "/" 以下へのアクセスで、隣の frontend フォルダ（../frontend）のファイルを返す。02_todo と 03_todo_sqlite で同じ画面を使う。
# ../frontend は「起動したフォルダ」から見た場所なので、main.py のあるフォルダで uvicorn を起動する。
# html=True で / に来たら index.html を返す。
# app.mount は最後に置く。先に置くと "/" がすべてのアクセスを受け取ってしまい、上の /todos に届かなくなる。
# 画面を選ぶ: 使う方の行の先頭の # を外し、もう一方の行に # を付ける
app.mount("/", StaticFiles(directory="../frontend", html=True), name="static")  # シンプル版
# app.mount("/", StaticFiles(directory="../frontend/notebook", html=True), name="static")  # 手帳風
# app.mount("/", StaticFiles(directory="../frontend/hum", html=True), name="static")  # Hum 版

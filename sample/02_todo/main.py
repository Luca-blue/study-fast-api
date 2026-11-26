# FastAPI 本体
from fastapi import FastAPI
# static フォルダの HTML などを、そのままブラウザに配るための部品
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



# ===== この行より上に書く（下の行は画面 static/index.html を配信する） =====
# "/" 以下へのアクセスで static フォルダのファイルを返す。html=True で / に来たら index.html を返す。
# この行は最後に置く。先に置くと "/" がすべてのアクセスを受け取ってしまい、上の /todos に届かなくなる。
app.mount("/", StaticFiles(directory="static", html=True), name="static")

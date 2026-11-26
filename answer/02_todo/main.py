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
# GET /todos : 一覧を返す
@app.get("/todos")
def list_todos():
    """ToDo を全部返す。

    Returns:
        list[dict]: ToDo の一覧。todos.values() で辞書の値（ToDo）だけを取り出し、リストにして返す。
    """
    return list(todos.values())


# POST /todos : 新しく作る
# 引数 todo の型を TodoIn にすると、リクエストの本文（JSON）が TodoIn に変換されて入ってくる
@app.post("/todos")
def create_todo(todo: TodoIn):
    """ToDo を1件追加する。

    Args:
        todo: 画面から送られてきた ToDo（title と done）。

    Returns:
        dict: 追加した ToDo（割り振った id 付き）。
    """
    todo_id = new_id()
    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}
    return todos[todo_id]


# PUT /todos/{todo_id} : 書き換える
# todo_id は URL から（パスパラメータ）、todo はリクエストの本文から受け取る
@app.put("/todos/{todo_id}")
def update_todo(todo_id: int, todo: TodoIn):
    """指定した id の ToDo を、送られてきた内容で置き換える。

    Args:
        todo_id: 書き換える ToDo の id。
        todo: 新しい内容（title と done）。

    Returns:
        dict: 書き換えた後の ToDo。
    """
    # 注意: todos に無い id を指定すると、エラーにはならず新しい ToDo として追加される
    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}
    return todos[todo_id]


# DELETE /todos/{todo_id} : 消す
@app.delete("/todos/{todo_id}")
def delete_todo(todo_id: int):
    """指定した id の ToDo を削除する。

    Args:
        todo_id: 削除する ToDo の id。

    Returns:
        dict: {"ok": True}。

    Raises:
        KeyError: todos に無い id を指定したとき。FastAPI は 500 エラーとして返す。
    """
    # del は辞書からキーごと取り除く
    del todos[todo_id]
    return {"ok": True}


# ===== この行より上に書く（下の行は画面 ../frontend/index.html を配信する） =====
# "/" 以下へのアクセスで、隣の frontend フォルダ（../frontend）のファイルを返す（answer/frontend は sample/frontend と同じ内容のコピー）。
# ../frontend は「起動したフォルダ」から見た場所なので、main.py のあるフォルダで uvicorn を起動する。
# html=True で / に来たら index.html を返す。
# この行は最後に置く。先に置くと "/" がすべてのアクセスを受け取ってしまい、上の /todos に届かなくなる。
app.mount("/", StaticFiles(directory="../frontend", html=True), name="static")

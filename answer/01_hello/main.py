# FastAPI 本体を読み込む
from fastapi import FastAPI

# アプリ本体を作る。
# 起動コマンド `uvicorn main:app` の「main:app」は「main.py の中の app」という意味
app = FastAPI()


# @app.get("/") は「GET メソッドで / にアクセスが来たら、すぐ下の関数を呼ぶ」という目印（デコレータ）
@app.get("/")
def hello():
    """トップページにあいさつを返す。

    Returns:
        dict: {"message": "Hello, FastAPI!"}。返した辞書は FastAPI が JSON に変換してブラウザに送る。
    """
    return {"message": "Hello, FastAPI!"}


# パスパラメータ: URL の一部を変数として受け取る
# 例: /hello/たろう にアクセスすると、{name} の部分の「たろう」が引数 name に入る
@app.get("/hello/{name}")
def hello_name(name: str):
    """URL に含まれた名前であいさつを返す。

    Args:
        name: URL の {name} の部分。`: str` と書いたので文字列として受け取る。

    Returns:
        dict: {"message": "こんにちは、〇〇さん"}。
    """
    # f"...{name}..." は f 文字列。{} の中に変数の値が埋め込まれる
    return {"message": f"こんにちは、{name}さん"}


# クエリパラメータ: URL の ? 以降を受け取る
# 例: /add?a=1&b=2 にアクセスすると a=1, b=2 が入る
# 引数名が URL の {} に無いので、FastAPI はクエリパラメータだと判断する
@app.get("/add")
def add(a: int, b: int):
    """2つの数を足した結果を返す。

    Args:
        a: 1つ目の数。`: int` と書いたので、"1" のような文字列は自動で整数に変換される。
            数字以外（例: a=abc）を送ると、FastAPI が 422 エラーを返す。
        b: 2つ目の数。

    Returns:
        dict: {"result": a + b}。
    """
    return {"result": a + b}

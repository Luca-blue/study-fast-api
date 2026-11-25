from fastapi import FastAPI

app = FastAPI()


# パスパラメータ: URL の一部を変数として受け取る
@app.get("/hello/{name}")
def hello(name: str):
    return {"message": f"こんにちは、{name}さん"}


# クエリパラメータ: URL の ? 以降を受け取る
@app.get("/add")
def add(a: int, b: int):
    return {"result": a + b}

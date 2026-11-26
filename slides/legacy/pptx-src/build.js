// Build "はじめての FastAPI" as an editable PowerPoint deck.
const pptxgen = require("pptxgenjs");
const L = require("./lib");
const { IN, COL, FONT, MONO, LINE, FRAME, HEAD_H, BODY, grid, inset, rect, text, md2runs, label, badge, code, term, box, notes, task, ibox } = L;
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const THEME = {
  name: "Raw Grid FastAPI", headFontFace: FONT, bodyFontFace: FONT,
  colors: { dk1: "0A0A0A", lt1: "FFFFFF", dk2: "333333", lt2: "F5F5F5", accent1: "F2D4CF", accent2: "E5EDD6", accent3: "0A0A0A",
    accent4: "A8D8EA", accent5: "C7DDB0", accent6: "9A9A9A", hlink: "0A0A0A", folHlink: "333333" },
};

(async () => {
  await L.loadIcons(["MdDevices", "MdApi", "MdStorage", "MdComputer", "MdCloudOff", "MdAltRoute", "MdDns", "MdVerified", "MdVisibility", "MdAddCircleOutline", "MdEdit", "MdDeleteOutline"]);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "はじめての FastAPI";

  // ---------- layouts: one per chapter ----------
  const frameObjs = () => [
    { rect: { x: IN(FRAME.x), y: IN(FRAME.y), w: IN(FRAME.w), h: IN(FRAME.h), fill: { color: COL.white }, line: { color: COL.black, width: LINE },
      shadow: { type: "outer", color: COL.black, opacity: 1, blur: 0, offset: 4.5, angle: 45 } } },
  ];
  const CH = {
    INDEX: ["INDEX", "目次"], C1: ["CH.01", "しくみ"], C2: ["CH.02", "準備"], C3: ["CH.03", "書いてみる"], C4: ["CH.04", "メソッド"],
    C5: ["CH.05", "データ"], C6: ["CH.06", "画面"], C7: ["CH.07", "DB"], C7F: ["CH.07", "完成"], FIN: ["FIN.", "まとめ"],
  };
  for (const [key, [no, name]] of Object.entries(CH)) {
    pres.defineSlideMaster({
      title: key,
      background: { color: COL.white },
      objects: [
        ...frameObjs(),
        { rect: { x: IN(FRAME.x), y: IN(FRAME.y), w: IN(230), h: IN(HEAD_H), fill: { color: COL.pink }, line: { color: COL.black, width: LINE } } },
        { text: { text: no, options: { x: IN(FRAME.x + 34), y: IN(FRAME.y + 26), w: IN(180), h: IN(40), fontFace: MONO, fontSize: 15, bold: true, color: COL.black, margin: 0 } } },
        { text: { text: name, options: { x: IN(FRAME.x + 34), y: IN(FRAME.y + 70), w: IN(190), h: IN(34), fontFace: FONT, fontSize: 11, bold: true, color: COL.black, margin: 0 } } },
        { rect: { x: IN(FRAME.x + 230), y: IN(FRAME.y), w: IN(FRAME.w - 400), h: IN(HEAD_H), fill: { color: COL.white }, line: { color: COL.black, width: LINE } } },
        { rect: { x: IN(FRAME.x + FRAME.w - 170), y: IN(FRAME.y), w: IN(170), h: IN(HEAD_H), fill: { color: COL.white }, line: { color: COL.black, width: LINE } } },
        { placeholder: { options: { name: "title", type: "title", x: IN(FRAME.x + 230 + 46), y: IN(FRAME.y), w: IN(FRAME.w - 400 - 70), h: IN(HEAD_H),
          fontFace: FONT, fontSize: 25, bold: true, color: COL.black, valign: "middle", align: "left", margin: 0 }, text: "" } },
      ],
      slideNumber: { x: IN(FRAME.x + FRAME.w - 170), y: IN(FRAME.y + 44), w: IN(170), h: IN(44), fontFace: MONO, fontSize: 14, bold: true, color: COL.black, align: "center" },
    });
  }
  pres.defineSlideMaster({ title: "COVER", background: { color: COL.white }, objects: frameObjs() });

  let section = null;
  const S = (key, title, sec) => {
    if (sec && sec !== section) { pres.addSection({ title: sec }); section = sec; }
    const s = pres.addSlide({ masterName: key, sectionTitle: section });
    if (title !== undefined) s.addText(title, { placeholder: "title" });
    return s;
  };
  const cell = (s, r, fill = COL.white) => rect(s, r, fill);

  // =====================================================================
  // 1. COVER
  // =====================================================================
  {
    const s = S("COVER", undefined, "表紙・目次");
    const g = grid(["1.35fr", "1fr"], [110, "1fr", 240], FRAME);
    cell(s, g(0, 0, 0, 1));
    label(s, "PYTHON WEB API 入門", 88, 34 + 36, { px: 22 });
    text(s, "fastapi + uvicorn + sqlite", { x: 1100, y: 34 + 38, w: 724, h: 40 }, { px: 26, mono: true, bold: true, align: "right" });
    const hero = g(1, 0); cell(s, hero);
    label(s, "ハンズオン資料", hero.x + 60, hero.y + 52);
    text(s, "はじめての", { x: hero.x + 60, y: hero.y + 110, w: 900, h: 150 }, { px: 120, bold: true, lh: 1.0 });
    text(s, "FastAPI", { x: hero.x + 60, y: hero.y + 250, w: 900, h: 170 }, { px: 150, bold: true, mono: true, lh: 1.0 });
    text(s, "しくみを知る → 動かす → つなげる。\n小さな ToDo アプリを作りながら学ぶ。", { x: hero.x + 60, y: hero.y + 450, w: 900, h: 110 }, { px: 32, color: COL.dark, lh: 1.5 });
    const side = g(1, 1); cell(s, side, COL.green);
    label(s, "今日つくるもの", side.x + 56, side.y + 50);
    const items = [["ブラウザ", "client"], ["FastAPI", "server"], ["SQLite", "database"]];
    const arrows = ["↓ GET /todos", "↓ SELECT *"];
    let y = side.y + 120;
    items.forEach(([a, b], i) => {
      const r = { x: side.x + 56, y, w: side.w - 112, h: 84 };
      box(s, r);
      text(s, a, { x: r.x + 28, y: r.y, w: 300, h: r.h }, { px: 34, bold: true, valign: "middle" });
      text(s, b, { x: r.x + r.w - 228, y: r.y, w: 200, h: r.h }, { px: 22, mono: true, color: COL.dark, align: "right", valign: "middle" });
      y += 84;
      if (arrows[i]) { text(s, arrows[i], { x: side.x + 86, y: y + 10, w: 400, h: 40 }, { px: 24, mono: true, bold: true }); y += 60; }
    });
    const chaps = [["01", "Web API の\nしくみ"], ["02—03", "準備と\nはじめてのコード"], ["04—05", "HTTPメソッドと\nデータの受け渡し"], ["06—07", "画面と\nデータベースをつなぐ"]];
    const cg = grid(["1fr", "1fr", "1fr", "1fr"], ["1fr"], g(2, 0, 2, 1));
    chaps.forEach(([n, t], i) => {
      const r = cg(0, i); cell(s, r, COL.pink);
      text(s, n, { x: r.x + 40, y: r.y + 34, w: r.w - 80, h: 34 }, { px: 22, mono: true, bold: true });
      text(s, t, { x: r.x + 40, y: r.y + 76, w: r.w - 60, h: 110 }, { px: 32, bold: true, lh: 1.35 });
    });
  }

  // =====================================================================
  // 2. INDEX
  // =====================================================================
  {
    const s = S("INDEX", "今日の流れ：7つの章");
    const g = grid(["1fr", "1fr", "1fr", "1fr"], ["1fr", "1fr"]);
    const cards = [
      ["01", "Web API のしくみ", "クライアント・サーバー・データベースの関係", COL.white],
      ["02", "準備", "ライブラリの役目とインストール", COL.white],
      ["03", "書いてみる", "最小のコードを動かして、ブラウザで確かめる", COL.pink],
      ["04", "HTTP メソッド", "GET・POST・PUT・DELETE を具体例で", COL.pink],
      ["05", "データを受け取る", "Pydantic とメモリ版 ToDo API", COL.green],
      ["06", "画面をつなぐ", "fetch で API を呼んで ToDo アプリを動かす", COL.green],
      ["07", "SQLite につなぐ", "再起動しても消えないデータで完成", COL.green],
      ["GOAL", "ToDo アプリ完成", "ブラウザ → FastAPI → SQLite が自分のコードでつながる", COL.black],
    ];
    cards.forEach(([n, t, d, f], i) => {
      const r = g(Math.floor(i / 4), i % 4); cell(s, r, f);
      const c = f === COL.black ? COL.white : COL.black;
      text(s, n, { x: r.x + 40, y: r.y + 36, w: r.w - 80, h: 80 }, { px: 64, mono: true, bold: true, color: c });
      text(s, t, { x: r.x + 40, y: r.y + 130, w: r.w - 70, h: 50 }, { px: 32, bold: true, color: c });
      text(s, d, { x: r.x + 40, y: r.y + 190, w: r.w - 70, h: 120 }, { px: 24, color: f === COL.white ? COL.dark : c, lh: 1.45 });
    });
  }

  // =====================================================================
  // CH.01
  // =====================================================================
  {
    const s = S("C1", "Web API の役割：画面とデータの間をつなぐ", "1章 Web API のしくみ");
    const g = grid(["1fr"], ["1fr", 290]);
    cell(s, g(0, 0));
    const top = inset(g(0, 0), 48, 70);
    const bw = (top.w - 2 * 150) / 3;
    const actors = [
      ["MdDevices", "クライアント", "ブラウザ・スマホアプリなど。データを**要求し**、受け取った結果を画面に表示する。", "例：ToDo アプリの一覧画面", COL.white],
      ["MdApi", "Web API", "リクエストを受け取り、データを読み書きして結果を返す。**この講座で作る部分。**", "例：GET /todos → ToDo 一覧を返す", COL.green],
      ["MdStorage", "データベース", "データを保存する場所。クライアントからは**直接アクセスしない**。", "例：todo.db の todos テーブル", COL.white],
    ];
    const conns = [["リクエスト →", "← レスポンス"], ["読み書き →", "← データ"]];
    actors.forEach(([ic, h, p, ex, f], i) => {
      const x = top.x + i * (bw + 150);
      const r = { x, y: top.y, w: bw, h: top.h };
      box(s, r, f);
      ibox(s, ic, x + 28, top.y + 28, 96);
      text(s, h, { x: x + 144, y: top.y + 28, w: bw - 160, h: 96 }, { px: 38, bold: true, valign: "middle" });
      text(s, md2runs(p), { x: x + 28, y: top.y + 150, w: bw - 56, h: 130 }, { px: 25, lh: 1.45 });
      s.addShape("line", { x: IN(x + 28), y: IN(top.y + top.h - 70), w: IN(bw - 56), h: 0, line: { color: COL.black, width: LINE } });
      text(s, ex, { x: x + 28, y: top.y + top.h - 56, w: bw - 56, h: 36 }, { px: 21, mono: true });
      if (conns[i]) {
        const cx = x + bw + 10;
        text(s, conns[i][0] + "\n" + conns[i][1], { x: cx, y: top.y + top.h / 2 - 40, w: 130, h: 80 }, { px: 19, bold: true, align: "center", valign: "middle", lh: 1.6 });
      }
    });
    const bg = grid(["1fr", "1fr", "1fr", "1fr"], ["1fr"], g(1, 0));
    const ex = [["天気アプリ", "「東京の天気」を要求\n→「晴れ 18℃」が返る"], ["SNS", "新着投稿を要求\n→ 投稿の一覧が返る"], ["ネットショップ", "カートへの追加を要求\n→「追加しました」が返る"], ["API を通す理由", "決まった URL とメソッドでしか操作できないため、データを安全に、決められたルールどおりに扱える。"]];
    ex.forEach(([h, p], i) => {
      const r = bg(0, i); cell(s, r, i === 3 ? COL.pink : COL.white);
      text(s, h, { x: r.x + 40, y: r.y + 40, w: r.w - 80, h: 46 }, { px: 30, bold: true });
      text(s, p, { x: r.x + 40, y: r.y + 100, w: r.w - 70, h: 150 }, { px: i === 3 ? 23 : 26, lh: 1.5 });
    });
  }
  {
    const s = S("C1", "リクエストとレスポンス：要求と応答");
    const g = grid(["1fr", "1fr"], ["1fr", 190]);
    const L1 = g(0, 0), R1 = g(0, 1);
    cell(s, L1); cell(s, R1, COL.gray);
    label(s, "REQUEST ／ リクエスト（要求）", L1.x + 44, L1.y + 36);
    term(s, "!!GET /todos/1 HTTP/1.1\nHost: 127.0.0.1:8000\nAccept: application/json", { x: L1.x + 44, y: L1.y + 96, w: L1.w - 88, h: 168 }, { bar: "クライアント → サーバー", px: 24 });
    notes(s, [
      { n: "1", h: "メソッド（何をするか）", p: "`GET` = 取得。ほかに作成・更新・削除がある" },
      { n: "2", h: "URL（対象）", p: "`/todos/1` = id が 1 の ToDo" },
      { n: "3", h: "ヘッダー・ボディ（付加情報・送信データ）", p: "作成・更新するデータはボディに入れて送る" },
    ], L1.x + 44, L1.y + 300, L1.w - 88, { gap: 18 });
    label(s, "RESPONSE ／ レスポンス（応答）", R1.x + 44, R1.y + 36);
    term(s, 'HTTP/1.1 200 OK\ncontent-type: application/json\n\n{"id": 1, "title": "牛乳を買う", "done": false}', { x: R1.x + 44, y: R1.y + 96, w: R1.w - 88, h: 204 }, { bar: "サーバー → クライアント", px: 24 });
    notes(s, [
      { n: "4", h: "ステータスコード（処理結果）", p: "`200` = 成功。見つからなければ `404`" },
      { n: "5", h: "ボディ（データ本体）", p: "Web API では **JSON** という形式で返すのが定番" },
    ], R1.x + 44, R1.y + 336, R1.w - 88, { gap: 18 });
    const b = g(1, 0, 1, 1); cell(s, b, COL.pink);
    text(s, md2runs("**1回のリクエストに、必ず1回のレスポンス。** この往復が HTTP（ブラウザとサーバーの通信のルール）の基本。"), inset(b, 48, 40), { px: 30, valign: "middle" });
  }
  {
    const s = S("C1", "全体像：3つの登場人物と、つなぐもの");
    const g = grid(["1fr", "1fr"], ["1fr", 210]);
    cell(s, g(0, 0, 0, 1));
    const top = inset(g(0, 0, 0, 1), 48, 56);
    const bw = (top.w - 2 * 170) / 3;
    const A = [
      ["MdDevices", "CLIENT", "クライアント", ["ブラウザ（HTML + JavaScript）", "スマホアプリ", "ほかのプログラム"], "この講座では → 6章", COL.white],
      ["MdApi", "WEB API", "Python + FastAPI", ["URL とメソッドごとに処理を決める", "届いたデータをチェックする", "DB に読み書きしてレスポンスを作る"], "この講座では → 2〜5章", COL.green],
      ["MdStorage", "DATABASE", "データベース", ["SQLite（ファイル1つ）", "表（テーブル）で保存", "SQL という言葉で操作"], "この講座では → 7章", COL.white],
    ];
    const C = [["HTTP", "JSON"], ["SQL", "行データ"]];
    A.forEach(([ic, lb, h, li, foot, f], i) => {
      const x = top.x + i * (bw + 170), r = { x, y: top.y, w: bw, h: top.h };
      box(s, r, f);
      ibox(s, ic, x + 28, top.y + 28, 96);
      label(s, lb, x + 144, top.y + 40, { fill: i === 1 ? COL.black : COL.pink, color: i === 1 ? COL.white : COL.black });
      text(s, h, { x: x + 28, y: top.y + 140, w: bw - 56, h: 50 }, { px: 34, bold: true });
      text(s, li.map((t, k) => ({ text: "→ " + t, options: { breakLine: k < li.length - 1 } })), { x: x + 28, y: top.y + 200, w: bw - 40, h: 150 }, { px: 23, lh: 1.6 });
      text(s, foot, { x: x + 28, y: top.y + top.h - 50, w: bw - 56, h: 30 }, { px: 20, color: COL.dark });
      if (C[i]) text(s, `${C[i][0]}\n→\n←\n${C[i][1]}`, { x: x + bw + 20, y: top.y + top.h / 2 - 80, w: 130, h: 160 }, { px: 20, bold: true, align: "center", valign: "middle", lh: 1.4 });
    });
    [["クライアント ⇄ API は「HTTP」", "URL・メソッド・ステータスコードでやりとり。データは JSON。"], ["API ⇄ DB は「SQL」", "「この表からこの行を取り出して」と命令する専用の言葉。"]].forEach(([h, p], i) => {
      const r = g(1, i); cell(s, r);
      text(s, h, { x: r.x + 44, y: r.y + 44, w: r.w - 88, h: 44 }, { px: 30, bold: true });
      text(s, p, { x: r.x + 44, y: r.y + 100, w: r.w - 88, h: 70 }, { px: 25 });
    });
  }

  // =====================================================================
  // CH.02
  // =====================================================================
  {
    const s = S("C2", "主要ライブラリと、それぞれの役目", "2章 準備");
    const g = grid(["1fr", "1fr", "1fr", "1fr"], ["1fr", 230]);
    const libs = [
      ["MdAltRoute", "fastapi", "URL と処理の対応づけ", "「どの URL に、どのメソッドでリクエストが来たら、どの関数を実行するか」を書くためのライブラリ。", "pip install が必要", COL.green],
      ["MdDns", "uvicorn", "リクエストを受け付けるサーバー", "ネットワークから届いたリクエストを受け取り、FastAPI に渡す**サーバー**。アプリの起動に使う。", "pip install が必要", COL.pink],
      ["MdVerified", "pydantic", "データの型チェック", "届いたデータが決めた形（型）どおりか検証する。合わなければ自動で 422 を返す。", "fastapi と一緒に自動で入る", COL.white],
      ["MdStorage", "sqlite3", "データベース操作", "SQLite データベースを読み書きするための標準モジュール。7章で使う。", "Python に最初から入っている", COL.gray],
    ];
    libs.forEach(([ic, lb, h, p, foot, f], i) => {
      const r = g(0, i); cell(s, r, f);
      ibox(s, ic, r.x + 40, r.y + 40, 96);
      label(s, lb, r.x + 156, r.y + 70, { px: 22 });
      text(s, h, { x: r.x + 40, y: r.y + 166, w: r.w - 70, h: 90 }, { px: 30, bold: true, lh: 1.3 });
      text(s, md2runs(p), { x: r.x + 40, y: r.y + 262, w: r.w - 70, h: 220 }, { px: 23, lh: 1.55 });
      text(s, foot, { x: r.x + 40, y: r.y + r.h - 56, w: r.w - 70, h: 30 }, { px: 20, color: COL.dark });
    });
    const b = g(1, 0, 1, 3); cell(s, b);
    const steps = [["ブラウザ", "", COL.white], ["uvicorn", "受け取る", COL.pink], ["FastAPI（app）", "振り分ける", COL.green], ["あなたの関数", "チェック済みのデータを受け取る", COL.white]];
    const sw = (b.w - 96 - 3 * 80) / 4;
    steps.forEach(([h, p, f], i) => {
      const x = b.x + 48 + i * (sw + 80);
      box(s, { x, y: b.y + 50, w: sw, h: 130 }, f);
      text(s, h, { x: x + 24, y: b.y + 70, w: sw - 48, h: 44 }, { px: 30, bold: true });
      if (p) text(s, p, { x: x + 24, y: b.y + 120, w: sw - 40, h: 50 }, { px: 21 });
      if (i < 3) text(s, "→", { x: x + sw + 20, y: b.y + 95, w: 40, h: 40 }, { px: 30, bold: true, align: "center", color: COL.white, fill: COL.black, valign: "middle" });
    });
  }
  {
    const s = S("C2", "教材を取得してインストール：3ステップ");
    const g = grid(["1.3fr", "1fr", "1fr"], ["1fr", 220]);
    const c1 = g(0, 0), c2 = g(0, 1), c3 = g(0, 2);
    cell(s, c1); cell(s, c2, COL.green); cell(s, c3, COL.pink);
    label(s, "STEP 1 ／ 教材を取得する", c1.x + 40, c1.y + 36);
    term(s, "$ git clone \\\n  https://github.com/luca-blue/study-fast-api.git\n$ cd study-fast-api", { x: c1.x + 40, y: c1.y + 96, w: c1.w - 80, h: 140 }, { px: 20 });
    text(s, "git がない場合は、GitHub の「Code」→「Download ZIP」で取得して展開する。", { x: c1.x + 40, y: c1.y + 270, w: c1.w - 80, h: 90 }, { px: 25, lh: 1.5 });
    label(s, "STEP 2 ／ 仮想環境を作る", c2.x + 40, c2.y + 36);
    term(s, "$ python -m venv .venv\n# Mac / Linux\n$ source .venv/bin/activate\n# Windows\n> .venv\\Scripts\\activate", { x: c2.x + 40, y: c2.y + 96, w: c2.w - 80, h: 210 }, { px: 21 });
    text(s, md2runs("先頭に `(.venv)` と出れば OK。"), { x: c2.x + 40, y: c2.y + 340, w: c2.w - 80, h: 50 }, { px: 25 });
    label(s, "STEP 3 ／ ライブラリを入れる", c3.x + 40, c3.y + 36);
    term(s, "$ pip install fastapi uvicorn\n$ pip list\n!!fastapi    0.x.x\n!!pydantic   2.x.x\n!!uvicorn    0.x.x", { x: c3.x + 40, y: c3.y + 96, w: c3.w - 80, h: 210 }, { px: 21 });
    text(s, "一覧に3つ出れば準備完了。", { x: c3.x + 40, y: c3.y + 340, w: c3.w - 80, h: 50 }, { px: 25 });
    const b1 = g(1, 0, 1, 1), b2 = g(1, 2); cell(s, b1, COL.gray); cell(s, b2);
    text(s, "仮想環境（venv）とは", { x: b1.x + 44, y: b1.y + 36, w: b1.w - 88, h: 44 }, { px: 30, bold: true });
    text(s, "プロジェクト専用の Python 環境。ここに入れたライブラリはこのフォルダの中でだけ使われるので、ほかのプロジェクトとバージョンがぶつからない。新しくターミナルを開いたら、もう一度 activate する。", { x: b1.x + 44, y: b1.y + 90, w: b1.w - 88, h: 110 }, { px: 24, lh: 1.5 });
    text(s, "エディタで開いておく", { x: b2.x + 40, y: b2.y + 36, w: b2.w - 80, h: 44 }, { px: 30, bold: true });
    text(s, md2runs("VS Code などで `study-fast-api` フォルダを開く。Python は **3.9 以上**。"), { x: b2.x + 40, y: b2.y + 90, w: b2.w - 80, h: 110 }, { px: 24, lh: 1.5 });
  }
  {
    const s = S("C2", "教材フォルダ sample/ の中身と進め方");
    const g = grid([640, "1fr"], ["1fr", 150]);
    const l = g(0, 0), r = g(0, 1);
    cell(s, l, COL.gray); cell(s, r);
    label(s, "study-fast-api/", l.x + 40, l.y + 36);
    const tree = [
      ["study-fast-api/", ""], ["├── slides/", "← このスライド"], ["├── sample/", "← ここを開いて編集する"], ["│   ├── 01_hello/", ""], ["│   │   └── main.py", ""],
      ["│   ├── 02_todo/", ""], ["│   │   ├── main.py", ""], ["│   │   └── static/index.html", ""], ["│   └── 03_todo_sqlite/", ""],
      ["│       ├── main.py", ""], ["│       └── static/index.html", ""], ["└── answer/", "← 答え合わせ用"],
    ];
    text(s, tree.flatMap(([a, b], i) => [
      { text: a, options: { fontFace: MONO, bold: true, highlight: i === 2 ? COL.pink : undefined } },
      { text: b ? "  " + b : "", options: { fontFace: FONT, color: COL.dark, breakLine: i < tree.length - 1 } },
    ]), { x: l.x + 40, y: l.y + 100, w: l.w - 60, h: 620 }, { px: 23, lh: 1.55 });
    const rows = [
      ["ファイル", "最初の状態", "やること", "章"],
      ["01_hello/main.py", "空", "すべて自分で書く", "3章"],
      ["02_todo/main.py", "import〜todos 辞書まで用意済み", "API の関数だけ書く", "5章"],
      ["02_todo/static/index.html", "完成品", "動かして、重要なところを読む", "6章"],
      ["03_todo_sqlite/", "完成品", "動かして、5章との違いを読む", "7章"],
    ];
    s.addTable(rows.map((row, ri) => row.map((c, ci) => ({
      text: c, options: {
        fontFace: ci === 0 && ri > 0 ? MONO : FONT, bold: ri === 0 || (ci === 2 && ri <= 2), fontSize: 12.5,
        color: ri === 0 ? COL.white : COL.black, fill: { color: ri === 0 ? COL.black : ci === 2 && ri <= 2 ? COL.green : ri % 2 === 0 ? COL.gray : COL.white },
      },
    }))), { x: IN(r.x + 40), y: IN(r.y + 90), w: IN(r.w - 80), colW: [IN(330), IN(300), IN(r.w - 80 - 330 - 300 - 90), IN(90)].map((v) => v),
      border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(86), valign: "middle" });
    const b = g(1, 0, 1, 1); cell(s, b, COL.pink);
    text(s, md2runs("**動かないときは、**`answer/`**の同じ名前のファイルと見比べる。**　各スライドの「作業」の手順どおりに進めれば OK。"), inset(b, 48, 30), { px: 29, valign: "middle" });
  }

  // =====================================================================
  // CH.03
  // =====================================================================
  {
    const s = S("C3", "いちばん小さな API：main.py", "3章 書いてみる");
    const g = grid(["1fr", "1fr"], ["1fr"]);
    const l = g(0, 0), r = g(0, 1);
    cell(s, l, COL.black); cell(s, r);
    label(s, "sample/01_hello/main.py（6行）", l.x + 44, l.y + 36, { fill: COL.white, color: COL.black });
    code(s, 'from fastapi import FastAPI\n\napp = FastAPI()\n\n\n@app.get("/")\ndef hello():\n    return {"message": "Hello, FastAPI!"}',
      { x: l.x + 44, y: l.y + 120, w: l.w - 88, h: 520 }, { px: 30, lh: 1.75, badges: { 0: "1", 2: "2", 5: "3", 7: "4" } });
    const t = task(s, "作業", ["`sample/01_hello/main.py` を開く（コメント1行だけの空ファイル）", "左のコードを**すべて自分で**書いて保存する"], r.x + 44, r.y + 40, r.w - 88);
    notes(s, [
      { n: "1", h: "ライブラリを読み込む", p: "fastapi から FastAPI クラスを取り出す。" },
      { n: "2", h: "アプリ本体を作る", p: "変数 `app` がこの API そのもの。起動のときに名前を使う。" },
      { n: "3", h: "「どの URL で動くか」の目印", p: "`@` で始まる行は**デコレータ**。「GET で `/` にリクエストが来たら、直下の関数を実行する」という登録。" },
      { n: "4", h: "レスポンスを返す", p: "辞書を return すると、FastAPI が自動で JSON にして返す。" },
    ], r.x + 44, r.y + 40 + t.h + 34, r.w - 88, { gap: 16, ppx: 23 });
  }
  {
    const s = S("C3", "起動する：uvicorn main:app --reload");
    const g = grid(["1.45fr", "1fr"], [270, "1fr"]);
    const top = g(0, 0, 0, 1); cell(s, top);
    const parts = [["uvicorn", "サーバーを起動する\nコマンド", COL.white], ["main", "main.py の\nファイル名", COL.green], [":app", "その中の\n変数 app", COL.pink], ["--reload", "保存すると自動で再起動\n（開発中だけ使う）", COL.white]];
    const pw = (top.w - 96 - 3 * 24) / 4;
    parts.forEach(([a, b, f], i) => {
      const r = { x: top.x + 48 + i * (pw + 24), y: top.y + 40, w: pw, h: top.h - 80 };
      box(s, r, f);
      text(s, a, { x: r.x + 26, y: r.y + 22, w: r.w - 52, h: 56 }, { px: 38, mono: true, bold: true });
      text(s, b, { x: r.x + 26, y: r.y + 88, w: r.w - 40, h: 90 }, { px: 24, lh: 1.4 });
    });
    const l = g(1, 0), r = g(1, 1); cell(s, l, COL.gray); cell(s, r);
    label(s, "実行結果", l.x + 40, l.y + 30);
    term(s, "$ cd sample/01_hello\n$ uvicorn main:app --reload\nINFO:     Will watch for changes in these directories: [...]\n!!INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)\nINFO:     Started reloader process [12345]\nINFO:     Started server process [12347]\nINFO:     Waiting for application startup.\n!!INFO:     Application startup complete.",
      { x: l.x + 40, y: l.y + 90, w: l.w - 80, h: 410 }, { bar: "main.py があるフォルダで実行", px: 18.5 });
    label(s, "ポイント", r.x + 40, r.y + 30);
    text(s, md2runs("→ `Uvicorn running on …` が出たら成功\n→ 起動中のサーバーは**動き続ける**。ほかの作業は別のターミナルで\n→ 止めるときは `Ctrl + C`\n→ ファイル名が app.py なら `uvicorn app:app`"), { x: r.x + 40, y: r.y + 90, w: r.w - 80, h: 420 }, { px: 25, lh: 1.55, psa: 10 });
  }
  {
    const s = S("C3", "ブラウザで確かめる");
    const g = grid(["1fr", "1fr"], ["1fr", 220]);
    const a = g(0, 0), b = g(0, 1);
    cell(s, a, COL.green); cell(s, b, COL.pink);
    const browser = (r, url, fill) => {
      box(s, r);
      rect(s, { x: r.x, y: r.y, w: r.w, h: 56 }, COL.gray);
      text(s, url, { x: r.x + 70, y: r.y + 10, w: r.w - 90, h: 36 }, { px: 21, mono: true, bold: true, fill: COL.white, valign: "middle" });
    };
    label(s, "① API を直接開く", a.x + 40, a.y + 32);
    browser({ x: a.x + 40, y: a.y + 92, w: a.w - 80, h: 300 }, " http://127.0.0.1:8000/");
    text(s, '{"message":"Hello, FastAPI!"}', { x: a.x + 70, y: a.y + 172, w: a.w - 120, h: 50 }, { px: 30, mono: true, bold: true });
    text(s, md2runs("return した辞書が JSON で表示される。\nブラウザの URL 欄に打つ＝ **GET リクエスト**を送っている。"), { x: a.x + 40, y: a.y + 420, w: a.w - 80, h: 100 }, { px: 25, lh: 1.5 });
    label(s, "② 自動ドキュメントを開く", b.x + 40, b.y + 32);
    browser({ x: b.x + 40, y: b.y + 92, w: b.w - 80, h: 300 }, " http://127.0.0.1:8000/docs");
    text(s, "FastAPI", { x: b.x + 70, y: b.y + 166, w: 300, h: 44 }, { px: 30, bold: true });
    rect(s, { x: b.x + 70, y: b.y + 226, w: b.w - 140, h: 64 }, COL.green, { lineW: 1.5 });
    text(s, "GET", { x: b.x + 90, y: b.y + 240, w: 120, h: 36 }, { px: 22, mono: true, bold: true, align: "center", valign: "middle", fill: COL.green });
    text(s, "/   Hello", { x: b.x + 230, y: b.y + 240, w: 300, h: 36 }, { px: 24, mono: true, bold: true, valign: "middle" });
    text(s, "クリック → Try it out → Execute で、その場で試せる", { x: b.x + 70, y: b.y + 310, w: b.w - 140, h: 40 }, { px: 21, color: COL.dark });
    text(s, md2runs("コードを書いただけで、この画面が**自動で**できる。"), { x: b.x + 40, y: b.y + 420, w: b.w - 80, h: 60 }, { px: 25 });
    const bt = grid(["1fr", "1fr", "1fr"], ["1fr"], g(1, 0, 1, 1));
    [["127.0.0.1", "「自分のパソコン」を指す特別な住所。localhost も同じ意味。"], [":8000", "ポート番号。同じパソコン内で接続先のプログラムを区別する番号。"], ["変更を試す", "\"Hello\" を書き換えて保存 → ブラウザを再読み込みすると反映される。"]].forEach(([h, p], i) => {
      const r = bt(0, i); cell(s, r);
      text(s, h, { x: r.x + 40, y: r.y + 36, w: r.w - 80, h: 44 }, { px: 30, bold: true, mono: i < 2 });
      text(s, p, { x: r.x + 40, y: r.y + 92, w: r.w - 70, h: 100 }, { px: 23, lh: 1.5 });
    });
  }
  {
    const s = S("C3", "いま起きていること：PC の中でサーバーが動いている");
    const g = grid(["1.55fr", "1fr"], ["1fr", 250]);
    const l = g(0, 0), r = g(0, 1); cell(s, l, COL.gray); cell(s, r);
    const pc = { x: l.x + 40, y: l.y + 36, w: l.w - 80, h: 400 };
    box(s, pc);
    ibox(s, "MdComputer", pc.x + 26, pc.y + 24, 80, COL.pink);
    text(s, "あなたの PC", { x: pc.x + 126, y: pc.y + 24, w: 400, h: 80 }, { px: 38, bold: true, valign: "middle" });
    label(s, "127.0.0.1 ＝ この PC 自身", pc.x + pc.w - 360, pc.y + 44);
    const tw = (pc.w - 52 - 200) / 2;
    term(s, "$ uvicorn main:app --reload\n!!Uvicorn running on\n!!  http://127.0.0.1:8000\n… 8000番でリクエストを待機中", { x: pc.x + 26, y: pc.y + 140, w: tw, h: 220 }, { bar: "ターミナル（サーバー）", px: 19 });
    text(s, "GET /\n←\n→\n200 + JSON", { x: pc.x + 26 + tw, y: pc.y + 170, w: 200, h: 160 }, { px: 21, bold: true, align: "center", valign: "middle", lh: 1.4 });
    const br = { x: pc.x + 26 + tw + 200, y: pc.y + 140, w: tw, h: 220 };
    box(s, br);
    rect(s, { x: br.x, y: br.y, w: br.w, h: 50 }, COL.gray);
    text(s, " 127.0.0.1:8000/", { x: br.x + 20, y: br.y + 9, w: br.w - 40, h: 32 }, { px: 19, mono: true, bold: true, fill: COL.white, valign: "middle" });
    text(s, "ブラウザ（クライアント）", { x: br.x + 20, y: br.y + 66, w: br.w - 40, h: 34 }, { px: 21, bold: true });
    text(s, '{"message":\n "Hello, FastAPI!"}', { x: br.x + 20, y: br.y + 110, w: br.w - 40, h: 80 }, { px: 21, mono: true, bold: true });
    ibox(s, "MdCloudOff", l.x + 40, l.y + 470, 96);
    text(s, md2runs("**インターネットには出ていない。**\nサーバーもクライアントも同じ PC の中にあり、通信は PC の中だけで完結している。"), { x: l.x + 160, y: l.y + 466, w: l.w - 200, h: 110 }, { px: 25, lh: 1.5 });
    notes(s, [
      { n: "1", h: "サーバー ＝ uvicorn", p: "ターミナルで動いている間、8000番ポートでリクエストを待ち続ける。" },
      { n: "2", h: "クライアント ＝ ブラウザ", p: "URL を開くと、そのアドレスに GET リクエストを送る。" },
      { n: "3", h: "アドレスの読み方", p: "`127.0.0.1` は自分の PC、`:8000` はその中で待機しているサーバーの番号。" },
    ], r.x + 40, r.y + 60, r.w - 80, { gap: 26 });
    const b1 = g(1, 0), b2 = g(1, 1);
    const bb = grid(["1fr", "1fr"], ["1fr"], g(1, 0, 1, 1));
    [[bb(0, 0), COL.green, "確認 1：ログを見る", 'INFO:  127.0.0.1:53086 - "GET / HTTP/1.1" 200 OK', "ブラウザで開く・再読み込みするたびに、ターミナルに1行ずつ増える。"],
     [bb(0, 1), COL.pink, "確認 2：サーバーを止める", "Ctrl+C → 再読み込み → 「このサイトにアクセスできません」", "応答するサーバーがいなくなるため接続できない。もう一度起動すれば元に戻る。"]].forEach(([rr, f, lb, cd, p]) => {
      cell(s, rr, f);
      label(s, lb, rr.x + 36, rr.y + 26);
      term(s, cd, { x: rr.x + 36, y: rr.y + 82, w: rr.w - 72, h: 62 }, { px: 19, shadow: false });
      text(s, p, { x: rr.x + 36, y: rr.y + 162, w: rr.w - 72, h: 60 }, { px: 22, lh: 1.45 });
    });
  }
  {
    const s = S("C3", "URL から値を受け取る：2つの方法");
    const g = grid(["1fr", "1fr"], [90, "1fr", 370]);
    const t = g(0, 0, 0, 1); cell(s, t);
    label(s, "作業", t.x + 36, t.y + 26);
    text(s, md2runs("同じ `sample/01_hello/main.py` の**末尾に2つの関数を追記** → 保存（`--reload` で自動反映）→ 下の URL をブラウザで開く"), { x: t.x + 130, y: t.y, w: t.w - 160, h: t.h }, { px: 24, valign: "middle" });
    const a = g(1, 0), b = g(1, 1); cell(s, a, COL.black); cell(s, b, COL.black);
    label(s, "パスパラメータ：URL の「途中」を変数に", a.x + 40, a.y + 30, { fill: COL.white, color: COL.black });
    code(s, '@app.get("/hello/{name}")\ndef hello_name(name: str):\n    return {"message": f"こんにちは、{name}さん"}', { x: a.x + 40, y: a.y + 92, w: a.w - 80, h: 140 }, { px: 25 });
    text(s, md2runs("`{name}` に入った文字が、引数 `name` に入る。"), { x: a.x + 40, y: a.y + 250, w: a.w - 80, h: 50 }, { px: 24, color: COL.white });
    label(s, "クエリパラメータ：URL の「? 以降」を変数に", b.x + 40, b.y + 30, { fill: COL.white, color: COL.black });
    code(s, '@app.get("/add")\ndef add(a: int, b: int):\n    return {"result": a + b}', { x: b.x + 40, y: b.y + 92, w: b.w - 80, h: 140 }, { px: 25 });
    text(s, md2runs("URL に書いていない引数は `?a=1&b=2` から受け取る。"), { x: b.x + 40, y: b.y + 250, w: b.w - 80, h: 50 }, { px: 24, color: COL.white });
    const tbl = (r, rows) => s.addTable(rows.map((row, ri) => row.map((c) => ({ text: c, options: { fontFace: ri && !c.includes("エラー") ? MONO : FONT, bold: ri === 0 || c.includes("エラー"), fontSize: 12, color: ri ? COL.black : COL.white, fill: { color: ri ? COL.white : COL.black } } }))),
      { x: IN(r.x + 36), y: IN(r.y + 90), w: IN(r.w - 72), colW: [IN(270), IN(r.w - 72 - 270)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.08, rowH: IN(66), valign: "middle" });
    const c = g(2, 0), d = g(2, 1); cell(s, c, COL.green); cell(s, d, COL.pink);
    label(s, "実行例", c.x + 36, c.y + 28); label(s, "実行例", d.x + 36, d.y + 28);
    tbl(c, [["URL", "結果"], ["/hello/たろう", '{"message":"こんにちは、たろうさん"}'], ["/hello/Alice", '{"message":"こんにちは、Aliceさん"}']]);
    tbl(d, [["URL", "結果"], ["/add?a=1&b=2", '{"result":3}'], ["/add?a=x&b=2", "422 エラー：a は整数ではない\n（a: int と書いたので自動でチェック）"]]);
  }
  {
    const s = S("C3", "うまく動かないときのチェック表");
    const g = grid(["1fr"], ["1fr", 150]);
    const a = g(0, 0); cell(s, a);
    const rows = [
      ["こんなメッセージが出たら", "よくある原因", "こうする"],
      ["command not found: uvicorn", "仮想環境が有効になっていない", "activate してから、もう一度 pip install"],
      ['Could not import module "main"', "別のフォルダで実行している", "cd で main.py のあるフォルダへ移動"],
      ['Attribute "app" not found', "コードの変数名とコマンドがずれている", "uvicorn ファイル名:変数名 を合わせる"],
      ["address already in use", "前に起動したサーバーが動いたまま", "そのターミナルで Ctrl+C／--port 8001"],
      ['ブラウザに {"detail":"Not Found"}', "URL のパスのつづりが違う", "/docs で登録済みの URL を確認"],
      ["Internal Server Error", "関数の中でエラーが起きている", "ターミナルに出たエラー文を読む"],
    ];
    s.addTable(rows.map((row, ri) => row.map((c, ci) => ({ text: c, options: { fontFace: ci === 0 && ri ? MONO : FONT, bold: ri === 0, fontSize: 12.5, color: ri ? COL.black : COL.white, fill: { color: ri === 0 ? COL.black : ri % 2 ? COL.white : COL.gray } } }))),
      { x: IN(a.x + 48), y: IN(a.y + 50), w: IN(a.w - 96), colW: [IN(640), IN(470), IN(a.w - 96 - 1110)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(84), valign: "middle" });
    const b = g(1, 0); cell(s, b, COL.pink);
    text(s, md2runs("**困ったら、まずターミナルを見る。** サーバー側で何が起きたかは、すべてそこに表示される。"), inset(b, 48, 30), { px: 30, valign: "middle" });
  }

  // =====================================================================
  // CH.04
  // =====================================================================
  {
    const s = S("C4", "HTTP メソッド：「どうしてほしいか」を表す動詞", "4章 HTTP メソッド");
    const g = grid(["1fr", "1fr", "1fr", "1fr"], ["1fr", 160]);
    const M = [
      ["GET", "MdVisibility", "取得する", "データを取り出す。中身は**変わらない**。", "なし", "200 OK", COL.green, COL.green],
      ["POST", "MdAddCircleOutline", "作成する", "新しいデータを送って登録する。データが**増える**。", "あり（作る中身）", "200 OK", COL.pink, COL.pink],
      ["PUT", "MdEdit", "更新する", "すでにあるデータを、送った内容で**置き換える**。", "あり（新しい中身）", "200 OK", COL.gray, COL.gray],
      ["DELETE", "MdDeleteOutline", "削除する", "指定したデータを削除する。データが**減る**。", "なし", "200 OK", COL.white, COL.black],
    ];
    M.forEach(([m, ic, h, p, body, ok, f, bf], i) => {
      const r = g(0, i); cell(s, r, f);
      rect(s, { x: r.x + 40, y: r.y + 36, w: r.w - 80, h: 40 }, bf);
      text(s, m, { x: r.x + 40, y: r.y + 36, w: r.w - 80, h: 40 }, { px: 24, mono: true, bold: true, align: "center", valign: "middle", color: bf === COL.black ? COL.white : COL.black });
      ibox(s, ic, r.x + 40, r.y + 110, 96);
      text(s, h, { x: r.x + 156, y: r.y + 110, w: r.w - 180, h: 96 }, { px: 36, bold: true, valign: "middle" });
      text(s, md2runs(p), { x: r.x + 40, y: r.y + 240, w: r.w - 70, h: 110 }, { px: 25, lh: 1.5 });
      s.addTable([["ボディ", body], ["成功時", ok]].map((row) => row.map((c, ci) => ({ text: c, options: { fontFace: ci && c.includes("OK") ? MONO : FONT, bold: ci === 0, fontSize: 12, fill: { color: COL.white } } }))),
        { x: IN(r.x + 40), y: IN(r.y + 390), w: IN(r.w - 80), colW: [IN(120), IN(r.w - 200)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.08, rowH: IN(70), valign: "middle" });
    });
    const b = g(1, 0, 1, 3); cell(s, b, COL.black);
    text(s, [{ text: "URL は「何を」（名詞）、メソッドは「どうする」（動詞）。", options: { bold: true, color: COL.white } }, { text: "　この4つで、作成・取得・更新・削除（CRUD）がすべてできる。", options: { color: COL.muted, fontSize: 12.5 } }], inset(b, 48, 30), { px: 32, valign: "middle" });
  }
  {
    const s = S("C4", "ToDo アプリの操作を、メソッドに置き換えると");
    const g = grid(["1fr", "1fr"], ["1fr", 170]);
    const a = g(0, 0, 0, 1); cell(s, a);
    const rows = [
      ["画面での操作", "メソッド", "URL", "送るデータ（ボディ）", "返ってくるもの"],
      ["一覧を開く", "GET", "/todos", "なし", "ToDo のリスト"],
      ["「追加」ボタン", "POST", "/todos", '{"title": "牛乳を買う"}', "作られた ToDo"],
      ["チェックを付ける", "PUT", "/todos/1", '{"title": "牛乳を買う", "done": true}', "更新後の ToDo"],
      ["「削除」ボタン", "DELETE", "/todos/1", "なし", '{"ok": true}'],
    ];
    const mf = { GET: COL.green, POST: COL.pink, PUT: COL.gray, DELETE: COL.black };
    s.addTable(rows.map((row, ri) => row.map((c, ci) => ({ text: c, options: {
      fontFace: ri && (ci === 1 || ci === 2 || c.startsWith("{")) ? MONO : FONT, bold: ri === 0 || ci === 1, fontSize: 13, align: ci === 1 && ri ? "center" : "left",
      color: ri === 0 || (ci === 1 && c === "DELETE") ? COL.white : COL.black, fill: { color: ri === 0 ? COL.black : ci === 1 ? mf[c] : COL.white } } }))),
      { x: IN(a.x + 48), y: IN(a.y + 100), w: IN(a.w - 96), colW: [IN(330), IN(200), IN(230), IN(620), IN(a.w - 96 - 1380)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(84), valign: "middle" });
    const b1 = g(1, 0), b2 = g(1, 1); cell(s, b1, COL.green); cell(s, b2, COL.pink);
    text(s, md2runs("**同じ **`/todos/1`** でも、**メソッドが PUT か DELETE かで別々の処理になる。"), inset(b1, 44, 30), { px: 25, valign: "middle", lh: 1.5 });
    text(s, md2runs("**GET と DELETE はボディなし。**何をするかは URL だけで伝わる。"), inset(b2, 44, 30), { px: 25, valign: "middle", lh: 1.5 });
  }

  // =====================================================================
  // CH.05
  // =====================================================================
  {
    const s = S("C5", "Pydantic：受け取るデータの「形」を決める", "5章 データを受け取る");
    const g = grid([780, "1fr"], ["1fr", 400]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.black); cell(s, b);
    label(s, "用意済み（sample/02_todo/main.py）", a.x + 44, a.y + 34, { fill: COL.white, color: COL.black });
    code(s, "from pydantic import BaseModel\n\n\nclass TodoIn(BaseModel):\n    title: str\n    done: bool = False", { x: a.x + 44, y: a.y + 100, w: a.w - 88, h: 320 }, { px: 28, lh: 1.6, badges: { 4: "1", 5: "2" } });
    notes(s, [
      { n: "1", h: "title は文字列で、必ず必要", p: "なければ自動で 422 エラー。" },
      { n: "2", h: "done は true / false。省略したら False", p: "`= 初期値` を書くと「省略してもいい項目」になる。" },
      { n: "→", h: "使い方：関数の引数に型として書くだけ", p: "`def create_todo(todo: TodoIn):` と書くと、ボディの JSON がチェックされて `todo` に入る。" },
    ], b.x + 44, b.y + 30, b.w - 88, { gap: 14, ppx: 23 });
    const c = g(1, 0, 1, 1); cell(s, c, COL.green);
    const rows = [["送った JSON", "結果", "関数が受け取る値 ／ エラーの理由"], ['{"title": "牛乳"}', "OK", "title='牛乳' done=False（初期値が入る）"], ['{"title": "牛乳", "done": true}', "OK", "title='牛乳' done=True"], ['{"done": true}', "422", "Field required（title がない）"], ['{"title": "牛乳", "done": "はい"}', "422", "Input should be a valid boolean（真偽値ではない）"]];
    s.addTable(rows.map((row, ri) => row.map((c2, ci) => ({ text: c2, options: { fontFace: ri && ci < 2 ? MONO : FONT, bold: ri === 0 || ci === 1, fontSize: 12.5, color: ri ? COL.black : COL.white, fill: { color: ri ? (ri % 2 ? COL.white : COL.gray) : COL.black } } }))),
      { x: IN(c.x + 40), y: IN(c.y + 32), w: IN(c.w - 80), colW: [IN(560), IN(160), IN(c.w - 80 - 720)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(64), valign: "middle" });
  }
  {
    const s = S("C5", "用意済みのファイルを開く：sample/02_todo/main.py");
    const g = grid(["1.1fr", "1fr"], ["1fr"]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.black); cell(s, b);
    const src = 'from fastapi import FastAPI\nfrom fastapi.staticfiles import StaticFiles\nfrom pydantic import BaseModel\n\napp = FastAPI()\n\n# 受け取るデータの形\nclass TodoIn(BaseModel):\n    title: str\n    done: bool = False\n\n# ToDo の置き場所\ntodos = {}\n\n# 次に使う id を返す\ndef new_id():\n    return max(todos, default=0) + 1\n\n# ===== ここから下に API の関数を書く =====\n\n\n# ===== この行より上に書く =====\napp.mount("/", StaticFiles(directory="static", html=True), ...)';
    const lh = 21 * 1.5;
    rect(s, { x: a.x + 30, y: a.y + 40 + 18 * lh - 2, w: a.w - 60, h: lh * 3 + 4 }, "3A3330", { lineW: 0 });
    text(s, "← ここに書く", { x: a.x + a.w - 250, y: a.y + 40 + 19 * lh, w: 200, h: lh }, { px: 21, bold: true, color: COL.pink, align: "right" });
    code(s, src, { x: a.x + 44, y: a.y + 40, w: a.w - 88, h: a.h - 60 }, { px: 21, lh: 1.5, badges: { 12: "1", 15: "2", 22: "3" } });
    const t = task(s, "作業", ["`sample/02_todo/main.py` を開き、用意済みの部分を確認する", "「ここから下に API の関数を書く」の下に、次のスライドの4つの関数を書く", "一番下の `app.mount` の行は消さない"], b.x + 44, b.y + 40, b.w - 88);
    notes(s, [
      { n: "1", h: "todos：ToDo を入れる辞書", p: 'キーが id、値が ToDo 1件。\n`{1: {"id": 1, "title": "牛乳を買う", "done": False}}`' },
      { n: "2", h: "new_id()：次に使う id を返す", p: "中身は気にせず、呼ぶだけで OK。" },
      { n: "3", h: "app.mount：画面ファイルを配信する", p: "`static/index.html` を `http://127.0.0.1:8000/` で開けるようにする（6章で使う）。" },
    ], b.x + 44, b.y + 40 + t.h + 34, b.w - 88, { gap: 14, ppx: 22, hpx: 27 });
  }
  {
    const s = S("C5", "API の関数を書く：GET・POST・PUT・DELETE");
    const g = grid(["1.32fr", "1fr"], ["1fr"]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.black); cell(s, b);
    code(s, '@app.get("/todos")\ndef list_todos():\n    return list(todos.values())\n\n\n@app.post("/todos")\ndef create_todo(todo: TodoIn):\n    todo_id = new_id()\n    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}\n    return todos[todo_id]\n\n\n@app.put("/todos/{todo_id}")\ndef update_todo(todo_id: int, todo: TodoIn):\n    todos[todo_id] = {"id": todo_id, "title": todo.title, "done": todo.done}\n    return todos[todo_id]\n\n\n@app.delete("/todos/{todo_id}")\ndef delete_todo(todo_id: int):\n    del todos[todo_id]\n    return {"ok": True}',
      { x: a.x + 40, y: a.y + 40, w: a.w - 80, h: a.h - 60 }, { px: 20, lh: 1.6, badges: { 0: "1", 5: "2", 12: "3", 18: "4" } });
    const y = notes(s, [
      { n: "1", h: "GET：一覧を返す", p: "todos の値（ToDo）をリストにして返す。" },
      { n: "2", h: "POST：新しく追加する", p: "`new_id()` で id を決め、送られた title と done を辞書に入れる。" },
      { n: "3", h: "PUT：上書きする", p: "URL の id の場所を、送られた内容で置き換える。" },
      { n: "4", h: "DELETE：削除する", p: "`del` で辞書から消す。" },
    ], b.x + 40, b.y + 36, b.w - 80, { gap: 14, ppx: 23 });
    task(s, "作業", ["4つの関数を書いて保存する", "ターミナルで起動する\n`cd sample/02_todo`\n`uvicorn main:app --reload`"], b.x + 40, y + 6, b.w - 80, { px: 23 });
    text(s, "※ 存在しない id を指定したときのエラー処理は、今回は扱わない。", { x: b.x + 40, y: b.y + b.h - 56, w: b.w - 80, h: 34 }, { px: 20, color: COL.dark });
  }
  {
    const s = S("C5", "/docs で全部のメソッドを試す");
    const g = grid(["1.1fr", "1fr"], ["1fr", 160]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.gray); cell(s, b);
    const br = { x: a.x + 36, y: a.y + 32, w: a.w - 72, h: a.h - 64 };
    box(s, br);
    rect(s, { x: br.x, y: br.y, w: br.w, h: 54 }, COL.gray);
    text(s, " http://127.0.0.1:8000/docs", { x: br.x + 60, y: br.y + 10, w: br.w - 80, h: 34 }, { px: 21, mono: true, bold: true, fill: COL.white, valign: "middle" });
    const mrow = (y, m, path, f, h = 62) => {
      rect(s, { x: br.x + 24, y, w: br.w - 48, h }, f === COL.pink ? COL.pink : COL.white, { lineW: 1.5 });
      text(s, m, { x: br.x + 44, y: y + 13, w: 140, h: 36 }, { px: 22, mono: true, bold: true, align: "center", valign: "middle", fill: { GET: COL.green, POST: COL.pink, PUT: COL.gray, DELETE: COL.black }[m], color: m === "DELETE" ? COL.white : COL.black });
      text(s, path, { x: br.x + 210, y: y + 13, w: 400, h: 36 }, { px: 24, mono: true, bold: true, valign: "middle" });
    };
    mrow(br.y + 80, "GET", "/todos");
    mrow(br.y + 160, "POST", "/todos", COL.pink, 236);
    label(s, "Try it out", br.x + br.w - 200, br.y + 173);
    term(s, '{"title": "牛乳を買う"}', { x: br.x + 44, y: br.y + 240, w: br.w - 88, h: 80 }, { px: 21, shadow: false, bar: "Request body" });
    label(s, "Execute", br.x + 44, br.y + 338);
    text(s, '→ 200 {"id": 1, "title": "牛乳を買う", "done": false}', { x: br.x + 190, y: br.y + 340, w: br.w - 220, h: 36 }, { px: 19, mono: true, valign: "middle" });
    mrow(br.y + 416, "PUT", "/todos/{todo_id}");
    mrow(br.y + 496, "DELETE", "/todos/{todo_id}");
    task(s, "作業：http://127.0.0.1:8000/docs を開く", ["**POST /todos** → Try it out → Request body に JSON を書いて Execute。2〜3件作る", "**GET /todos** で、作った ToDo が並んでいることを確認", "**PUT** で done を true に、**DELETE** で1件削除", "title を消して POST → **422** と理由が返ってくることを確認"], b.x + 40, b.y + 90, b.w - 80);
    const c = g(1, 0, 1, 1); cell(s, c, COL.pink);
    text(s, md2runs("**/docs ですべて動作したら、次は画面から API を使う。**\n→ 6章：用意済みの画面 `static/index.html` を開いて操作する。"), inset(c, 48, 26), { px: 27, valign: "middle", lh: 1.5 });
  }

  // =====================================================================
  // CH.06
  // =====================================================================
  {
    const s = S("C6", "画面で動作を確認する：操作 ＝ HTTP リクエスト", "6章 画面をつなぐ");
    const g = grid(["1fr", "1fr"], ["1fr", 160]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.gray); cell(s, b);
    const t = task(s, "作業：sample/02_todo/static/index.html（完成品）", ["5章のサーバーを起動したまま `http://127.0.0.1:8000/` を開く", "ToDo を追加 → チェック → 削除 してみる", "再読み込みしても一覧が表示されることを確認", "右のように、ターミナルのログを確認"], a.x + 36, a.y + 30, a.w - 72, { px: 23 });
    const br = { x: a.x + 36, y: a.y + 30 + t.h + 26, w: a.w - 72, h: a.h - t.h - 90 };
    box(s, br);
    rect(s, { x: br.x, y: br.y, w: br.w, h: 50 }, COL.gray);
    text(s, " http://127.0.0.1:8000/", { x: br.x + 50, y: br.y + 8, w: br.w - 70, h: 34 }, { px: 20, mono: true, bold: true, fill: COL.white, valign: "middle" });
    text(s, "ToDo", { x: br.x + 26, y: br.y + 64, w: 200, h: 44 }, { px: 32, bold: true });
    rect(s, { x: br.x + 26, y: br.y + 114, w: br.w - 190, h: 44 }, COL.white, { lineW: 1.5 });
    text(s, "歯医者を予約", { x: br.x + 40, y: br.y + 118, w: 400, h: 36 }, { px: 22, color: COL.dark, valign: "middle" });
    text(s, "追加", { x: br.x + br.w - 150, y: br.y + 114, w: 124, h: 44 }, { px: 22, bold: true, align: "center", valign: "middle", fill: COL.pink });
    text(s, "☐ 牛乳を買う ［削除］　☑ レポート提出 ［削除］", { x: br.x + 26, y: br.y + 172, w: br.w - 52, h: 36 }, { px: 22 });
    term(s, 'INFO: 127.0.0.1 - "GET / HTTP/1.1" 200 OK\nINFO: 127.0.0.1 - "GET /todos HTTP/1.1" 200 OK\nINFO: 127.0.0.1 - "POST /todos HTTP/1.1" 200 OK\nINFO: 127.0.0.1 - "GET /todos HTTP/1.1" 200 OK\nINFO: 127.0.0.1 - "PUT /todos/1 HTTP/1.1" 200 OK\nINFO: 127.0.0.1 - "DELETE /todos/2 HTTP/1.1" 200 OK',
      { x: b.x + 36, y: b.y + 30, w: b.w - 72, h: 296 }, { bar: "ターミナルのログ（1行 ＝ 1リクエスト）", px: 20 });
    const mf = { GET: COL.green, POST: COL.pink, PUT: COL.gray, DELETE: COL.black };
    const rows = [["画面での操作", "送られるリクエスト", null], ["開く・再読み込み", "GET", "/todos"], ["「追加」", "POST", "/todos"], ["チェック", "PUT", "/todos/1"], ["「削除」", "DELETE", "/todos/2"]];
    s.addTable(rows.map((row, ri) => row.filter((c) => c !== null).map((c, ci) => ({ text: c, options: { colspan: ri === 0 && ci === 1 ? 2 : undefined, fontFace: ci && ri ? MONO : FONT, bold: ri === 0 || ci === 1, fontSize: 12.5, align: ci === 1 && ri ? "center" : "left",
      color: ri === 0 || c === "DELETE" ? COL.white : COL.black, fill: { color: ri === 0 ? COL.black : ci === 1 ? mf[c] : COL.white } } }))),
      { x: IN(b.x + 36), y: IN(b.y + 356), w: IN(b.w - 72), colW: [IN(330), IN(180), IN(b.w - 72 - 510)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.08, rowH: IN(54), valign: "middle" });
    const c = g(1, 0, 1, 1); cell(s, c, COL.pink);
    text(s, md2runs("**ただし、サーバーを Ctrl+C で止めて再起動すると、ToDo はすべて消える。**\nデータが Python の変数（メモリ）にしかないため。→ 7章で、保存先をデータベースに変える。"), inset(c, 48, 26), { px: 27, valign: "middle", lh: 1.5 });
  }
  {
    const s = S("C6", "コードの重要なところを読む：fetch");
    const g = grid(["1fr", "1fr"], ["1fr", "1fr", 180]);
    const blocks = [
      ["GET：一覧を取得して表示（load 関数）", 'const res = await fetch("/todos");\nconst todos = await res.json();', "返ってきた JSON の配列を、1件ずつ画面に並べる。", 22],
      ["POST：「追加」ボタン", 'await fetch("/todos", {\n  method: "POST",\n  headers: { "Content-Type": "application/json" },\n  body: JSON.stringify({ title: title.value }),\n});', "", 21],
      ["PUT：チェックボックス", 'await fetch(`/todos/${t.id}`, {\n  method: "PUT",\n  headers: { "Content-Type": "application/json" },\n  body: JSON.stringify({ title: t.title, done: e.target.checked }),\n});', "", 19],
      ["DELETE：「削除」ボタン", 'await fetch(`/todos/${t.id}`, { method: "DELETE" });', "消すだけなのでボディは送らない。", 22],
    ];
    blocks.forEach(([lb, src, p, px], i) => {
      const r = g(Math.floor(i / 2), i % 2); cell(s, r, COL.black);
      label(s, lb, r.x + 36, r.y + 26, { fill: COL.white, color: COL.black });
      code(s, src, { x: r.x + 36, y: r.y + 84, w: r.w - 60, h: 180 }, { px, lh: 1.5 });
      if (p) text(s, p, { x: r.x + 36, y: r.y + 84 + (src.split("\n").length) * px * 1.5 + 24, w: r.w - 72, h: 40 }, { px: 24, color: COL.white });
    });
    const c = grid(["1fr", "1fr", "1fr"], ["1fr"], g(2, 0, 2, 1));
    [["fetch(URL, {...})", "ブラウザから HTTP リクエストを送る。method にメソッド、body に送るデータ。"], ["JSON.stringify", "JavaScript のデータを JSON の文字列にする（Pydantic が受け取る形）。"], ["await", "レスポンスが届くまで待ってから、次の行に進む。"]].forEach(([h, p], i) => {
      const r = c(0, i); cell(s, r, COL.green);
      text(s, h, { x: r.x + 36, y: r.y + 30, w: r.w - 72, h: 40 }, { px: 27, mono: true, bold: true });
      text(s, p, { x: r.x + 36, y: r.y + 80, w: r.w - 60, h: 90 }, { px: 22, lh: 1.45 });
    });
  }

  // =====================================================================
  // CH.07
  // =====================================================================
  {
    const s = S("C7", "なぜデータベース？ そして SQLite とは", "7章 SQLite につなぐ");
    const g = grid(["1fr", "1fr"], ["1fr"]);
    const a = g(0, 0), b = g(0, 1); cell(s, a); cell(s, b, COL.green);
    label(s, "メモリ（変数）の弱点", a.x + 44, a.y + 36);
    box(s, { x: a.x + 44, y: a.y + 100, w: 330, h: 110 }, COL.green);
    text(s, "起動中\ntodos = {1: ..., 2: ...}", { x: a.x + 68, y: a.y + 114, w: 300, h: 90 }, { px: 22, bold: true, lh: 1.5 });
    text(s, "再起動\n→", { x: a.x + 384, y: a.y + 120, w: 120, h: 80 }, { px: 21, bold: true, align: "center", valign: "middle" });
    box(s, { x: a.x + 514, y: a.y + 100, w: a.w - 558, h: 110 }, COL.pink);
    text(s, "再起動後\ntodos = {}  ← 空っぽ", { x: a.x + 538, y: a.y + 114, w: a.w - 600, h: 90 }, { px: 22, bold: true, lh: 1.5 });
    text(s, md2runs("プログラムが終わると変数は消える。\nだから**ファイルに書き残す**仕組み＝データベースが必要。"), { x: a.x + 44, y: a.y + 240, w: a.w - 88, h: 90 }, { px: 25, lh: 1.5 });
    label(s, "データベースは「表（テーブル）」でしまう", a.x + 44, a.y + 352);
    s.addTable([["id", "title", "done"], ["1", "牛乳を買う", "0"], ["2", "レポート提出", "1"]].map((row, ri) => row.map((c) => ({ text: c, options: { fontFace: FONT, bold: ri === 0, fontSize: 13, color: ri ? COL.black : COL.white, fill: { color: ri ? COL.white : COL.black } } }))),
      { x: IN(a.x + 44), y: IN(a.y + 412), w: IN(a.w - 88), colW: [IN(140), IN(a.w - 88 - 290), IN(150)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(66), valign: "middle" });
    text(s, "1行 ＝ ToDo 1件。列（id・title・done）は Pydantic モデルの項目と対応。", { x: a.x + 44, y: a.y + 630, w: a.w - 88, h: 40 }, { px: 22, color: COL.dark });
    label(s, "SQLite（エスキューライト）", b.x + 44, b.y + 36);
    box(s, { x: b.x + 44, y: b.y + 100, w: b.w - 88, h: 180 });
    rect(s, { x: b.x + 72, y: b.y + 122, w: 110, h: 136 }, COL.white, { shadow: true });
    text(s, "todo.db", { x: b.x + 72, y: b.y + 210, w: 110, h: 40 }, { px: 20, mono: true, bold: true, align: "center" });
    text(s, md2runs("**データベースがファイル1つ。**\nサーバーを別に立てる必要がなく、コピーすれば持ち運べる。"), { x: b.x + 210, y: b.y + 130, w: b.w - 290, h: 130 }, { px: 25, lh: 1.5 });
    text(s, ["Python に最初から入っている（import sqlite3）", "インストールも設定もいらない", "スマホアプリやブラウザの中でも使われている", "練習〜小さなアプリなら十分な性能"].map((t, i, arr) => ({ text: "✓  " + t, options: { breakLine: i < arr.length - 1 } })), { x: b.x + 44, y: b.y + 320, w: b.w - 88, h: 230 }, { px: 26, lh: 1.7 });
    text(s, "本格的なサービスでは PostgreSQL や MySQL を使うことが多いが、SQL の書き方はほぼ同じ。", { x: b.x + 44, y: b.y + 600, w: b.w - 88, h: 70 }, { px: 21, color: COL.dark, lh: 1.5 });
  }
  {
    const s = S("C7", "SQL の4つの命令は、HTTP メソッドと対応する");
    const g = grid(["1fr", "1fr"], ["1fr", 230]);
    const a = g(0, 0, 0, 1); cell(s, a);
    const mf = { GET: COL.green, POST: COL.pink, PUT: COL.gray, DELETE: COL.black };
    const rows = [["HTTP", "SQL", "意味", "SQL の例"], ["GET", "SELECT", "取り出す", "SELECT * FROM todos"], ["POST", "INSERT", "追加する", "INSERT INTO todos (title, done) VALUES (?, ?)"], ["PUT", "UPDATE", "書き換える", "UPDATE todos SET title = ?, done = ? WHERE id = ?"], ["DELETE", "DELETE", "削除する", "DELETE FROM todos WHERE id = ?"]];
    s.addTable(rows.map((row, ri) => row.map((c, ci) => ({ text: c, options: { fontFace: ri && ci !== 2 ? MONO : FONT, bold: ri === 0 || ci < 2, fontSize: 13, align: ci === 0 && ri ? "center" : "left",
      color: ri === 0 || (ci === 0 && c === "DELETE") ? COL.white : COL.black, fill: { color: ri === 0 ? COL.black : ci === 0 ? mf[c] : COL.white } } }))),
      { x: IN(a.x + 48), y: IN(a.y + 60), w: IN(a.w - 96), colW: [IN(200), IN(200), IN(200), IN(a.w - 96 - 600)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(76), valign: "middle" });
    text(s, md2runs("`*` は「全部の列」、`WHERE id = ?` は「id が〇〇の行だけ」という条件。WHERE を忘れると**全行**が対象になるので注意。"), { x: a.x + 48, y: a.y + 480, w: a.w - 96, h: 60 }, { px: 25 });
    const b1 = g(1, 0), b2 = g(1, 1); cell(s, b1, COL.pink); cell(s, b2, COL.black);
    text(s, "? はプレースホルダ（値の差し込み口）", { x: b1.x + 44, y: b1.y + 40, w: b1.w - 88, h: 44 }, { px: 29, bold: true });
    text(s, md2runs("値は `?` の位置に、あとからタプルで渡す。\n`execute(sql, (title, done))`"), { x: b1.x + 44, y: b1.y + 100, w: b1.w - 88, h: 100 }, { px: 24, lh: 1.6 });
    text(s, "文字列をつなげて SQL を作らない", { x: b2.x + 44, y: b2.y + 40, w: b2.w - 88, h: 44 }, { px: 29, bold: true, color: COL.white });
    text(s, md2runs('`f"... WHERE id = {todo_id}"` は SQL インジェクション（乗っ取り）の原因。必ず `?` を使う。'), { x: b2.x + 44, y: b2.y + 100, w: b2.w - 88, h: 100 }, { px: 24, lh: 1.6, color: COL.white });
  }
  {
    const s = S("C7", "DB 版を読む（1/2）：データベースの準備");
    const g = grid(["1.2fr", "1fr"], ["1fr"]);
    const a = g(0, 0), b = g(0, 1); cell(s, a, COL.black); cell(s, b);
    code(s, 'import sqlite3\n# （FastAPI・StaticFiles・BaseModel の import と\n#   app・TodoIn は 02_todo と同じ）\n\n\ndef get_conn():\n    conn = sqlite3.connect("todo.db")\n    conn.row_factory = sqlite3.Row\n    return conn\n\n\nwith get_conn() as conn:\n    conn.execute(\n        "CREATE TABLE IF NOT EXISTS todos ("\n        " id INTEGER PRIMARY KEY AUTOINCREMENT,"\n        " title TEXT NOT NULL,"\n        " done INTEGER NOT NULL DEFAULT 0)"\n    )',
      { x: a.x + 44, y: a.y + 44, w: a.w - 88, h: a.h - 70 }, { px: 23, lh: 1.55, badges: { 5: "1", 11: "2" } });
    const t = task(s, "作業", ["`sample/03_todo_sqlite/main.py` を開いて読む（完成品）"], b.x + 44, b.y + 40, b.w - 88);
    const y = notes(s, [
      { n: "1", h: "DB につなぐ関数", p: "`todo.db` がなければ自動で作られる。`row_factory` で `row[\"title\"]` のように列名で取り出せる。" },
      { n: "2", h: "起動時にテーブルを作る", p: "`IF NOT EXISTS`＝なければ作る。id は自動で 1, 2, 3… と振られる。`with` を抜けると変更が保存される。" },
    ], b.x + 44, b.y + 40 + t.h + 36, b.w - 88, { gap: 22 });
    box(s, { x: b.x + 44, y: y + 10, w: b.w - 88, h: 110 }, COL.green);
    text(s, "SQLite には true/false の型がないので、done は 0 / 1 で保存する。", { x: b.x + 70, y: y + 26, w: b.w - 140, h: 80 }, { px: 24, lh: 1.5 });
  }
  {
    const s = S("C7", "DB 版を読む（2/2）：5章の関数と比べる");
    const g = grid([232, "1fr", "1.3fr"], [92, "1fr", "1fr", "1fr", "1fr"]);
    cell(s, g(0, 0), COL.gray);
    const h1 = g(0, 1), h2 = g(0, 2); cell(s, h1, COL.gray); cell(s, h2, COL.green);
    text(s, "5章：辞書に入れる（メモリ）", inset(h1, 30, 10), { px: 28, bold: true, valign: "middle" });
    text(s, [{ text: "7章：SQL で todo.db に保存", options: { bold: true } }, { text: "　（with get_conn() as conn: の中）", options: { fontSize: 10, color: COL.dark } }], inset(h2, 30, 10), { px: 28, valign: "middle" });
    const R = [
      ["GET", "一覧", "return list(todos.values())", 'rows = conn.execute(\n    "SELECT id, title, done FROM todos"\n).fetchall()   # → 辞書のリストにして返す'],
      ["POST", "追加", 'todo_id = new_id()\ntodos[todo_id] = {"id": todo_id,\n    "title": todo.title, "done": todo.done}', 'cur = conn.execute(\n    "INSERT INTO todos (title, done) VALUES (?, ?)",\n    (todo.title, todo.done),\n)   # id は cur.lastrowid'],
      ["PUT", "上書き", 'todos[todo_id] = {"id": todo_id,\n    "title": todo.title, "done": todo.done}', 'conn.execute(\n    "UPDATE todos SET title = ?, done = ? WHERE id = ?",\n    (todo.title, todo.done, todo_id),\n)'],
      ["DELETE", "削除", "del todos[todo_id]", 'conn.execute("DELETE FROM todos WHERE id = ?", (todo_id,))'],
    ];
    const mf = { GET: COL.green, POST: COL.pink, PUT: COL.gray, DELETE: COL.black };
    R.forEach(([m, j, mem, db], i) => {
      const c0 = g(i + 1, 0), c1 = g(i + 1, 1), c2 = g(i + 1, 2);
      cell(s, c0); cell(s, c1, COL.black); cell(s, c2, COL.black);
      text(s, m, { x: c0.x + 30, y: c0.y + c0.h / 2 - 44, w: c0.w - 60, h: 40 }, { px: 22, mono: true, bold: true, align: "center", valign: "middle", fill: mf[m], color: m === "DELETE" ? COL.white : COL.black });
      text(s, j, { x: c0.x + 30, y: c0.y + c0.h / 2 + 8, w: c0.w - 60, h: 34 }, { px: 22, align: "center" });
      const ml = mem.split("\n").length, dl = db.split("\n").length;
      code(s, mem, { x: c1.x + 30, y: c1.y + (c1.h - ml * 30) / 2, w: c1.w - 50, h: ml * 30 + 10 }, { px: 20, lh: 1.5 });
      code(s, db, { x: c2.x + 30, y: c2.y + (c2.h - dl * 30) / 2, w: c2.w - 50, h: dl * 30 + 10 }, { px: 20, lh: 1.5 });
    });
  }
  {
    const s = S("C7F", "完成！「追加」を押したとき、裏で起きていること");
    const g = grid(["1fr", 600], ["1fr", 200]);
    const a = g(0, 0), b = g(0, 1); cell(s, a); cell(s, b, COL.gray);
    const bw = 300, gap = (a.w - 96 - 3 * bw) / 2;
    const top = [["MdComputer", "1", "ブラウザ", "「追加」クリック\n→ fetch で POST", COL.white], ["MdApi", "2", "FastAPI", "Pydantic でチェック\n→ create_todo() 実行", COL.pink], ["MdStorage", "3", "SQLite", "todo.db に\n1行追加", COL.green]];
    top.forEach(([ic, n, h, p, f], i) => {
      const x = a.x + 48 + i * (bw + gap), y = a.y + 50;
      box(s, { x, y, w: bw, h: 300 }, f);
      ibox(s, ic, x + 24, y + 24, 88);
      badge(s, n, x + bw - 64, y + 24, 40);
      text(s, h, { x: x + 24, y: y + 130, w: bw - 48, h: 46 }, { px: 32, bold: true });
      text(s, p, { x: x + 24, y: y + 186, w: bw - 40, h: 90 }, { px: 22, lh: 1.45 });
      if (i < 2) text(s, ["POST /todos", "INSERT"][i] + "\n→", { x: x + bw + 6, y: y + 110, w: gap - 12, h: 80 }, { px: 20, bold: true, align: "center", valign: "middle", lh: 1.4 });
    });
    const bot = [["5", "ブラウザ", "load() で GET /todos\n→ 一覧を描き直す", COL.black], ["4", "FastAPI", "作った ToDo を\nJSON で返す", COL.pink]];
    bot.forEach(([n, h, p, f], i) => {
      const x = a.x + 48 + i * (bw + gap), y = a.y + 400;
      box(s, { x, y, w: bw, h: 220 }, f);
      const c = f === COL.black ? COL.white : COL.black;
      badge(s, n, x + 24, y + 24, 40);
      text(s, h, { x: x + 24, y: y + 80, w: bw - 48, h: 46 }, { px: 32, bold: true, color: c });
      text(s, p, { x: x + 24, y: y + 134, w: bw - 40, h: 80 }, { px: 22, lh: 1.45, color: c });
      text(s, ["←\n200 + JSON", "←\nlastrowid"][i], { x: x + bw + 6, y: y + 70, w: gap - 12, h: 80 }, { px: 20, bold: true, align: "center", valign: "middle", lh: 1.4 });
    });
    task(s, "作業：再起動テスト", ["5章のサーバーを Ctrl+C で止める", "`cd sample/03_todo_sqlite`\n`uvicorn main:app --reload`", "画面で ToDo を2件追加する", "Ctrl+C で止めて、もう一度起動", "再読み込み → **消えずに残っている**"], b.x + 36, b.y + 40, b.w - 72, { px: 23 });
    const c = g(1, 0, 1, 1); cell(s, c, COL.black);
    text(s, [{ text: "画面（index.html）は 6章と同じ。保存先を辞書から SQLite に変えただけで、データが残るようになった。", options: { bold: true, color: COL.white, breakLine: true } }, { text: "1章の全体図（クライアント → Web API → データベース）が、すべてつながった。", options: { color: COL.muted, fontSize: 12.5 } }], inset(c, 48, 30), { px: 29, valign: "middle", lh: 1.5 });
  }

  // =====================================================================
  // SUMMARY
  // =====================================================================
  {
    const s = S("FIN", "まとめと、次の一歩", "まとめ");
    const g = grid(["1fr", "1fr"], ["1fr", 150]);
    const a = g(0, 0), b = g(0, 1); cell(s, a); cell(s, b, COL.green);
    label(s, "今日できるようになったこと", a.x + 44, a.y + 36);
    const done = ["クライアント・Web API・DB の役割を説明できる", "HTTP の往復（リクエスト・レスポンス）が読める", "uvicorn main:app --reload で API を起動できる", "GET・POST・PUT・DELETE を使い分けられる", "Pydantic で受け取るデータをチェックできる", "SQLite に保存し、画面から fetch で呼び出せる"];
    done.forEach((t, i) => {
      const y = a.y + 110 + i * 82;
      rect(s, { x: a.x + 44, y, w: 44, h: 44 }, COL.green, { lineW: 2 });
      text(s, "✓", { x: a.x + 44, y, w: 44, h: 44 }, { px: 26, bold: true, align: "center", valign: "middle" });
      text(s, t, { x: a.x + 110, y, w: a.w - 160, h: 44 }, { px: 27, valign: "middle" });
    });
    label(s, "次に学ぶと良いこと", b.x + 44, b.y + 36);
    const nx = [["SQLModel / SQLAlchemy", "SQL を書かずに Python のクラスで DB を扱う（ORM）"], ["エラー処理", "存在しない id には 404 を返す（HTTPException）"], ["テスト", "TestClient で API を自動テスト"], ["認証", "ログインして自分の ToDo だけ見る"], ["async / await", "待ち時間の多い処理を効率よくさばく"], ["デプロイ", "インターネットに公開して誰でも使えるように"]];
    s.addTable(nx.map(([h, p]) => [{ text: h, options: { bold: true, fontFace: FONT, fontSize: 13 } }, { text: p, options: { fontFace: FONT, fontSize: 12.5 } }]),
      { x: IN(b.x + 44), y: IN(b.y + 100), w: IN(b.w - 88), colW: [IN(300), IN(b.w - 88 - 300)], border: { type: "solid", color: COL.black, pt: 2 }, margin: 0.1, rowH: IN(84), valign: "middle", fill: { color: COL.white } });
    const c = g(1, 0, 1, 1); cell(s, c, COL.pink);
    text(s, [{ text: "公式チュートリアル（日本語あり）：", options: { bold: true } }, { text: "fastapi.tiangolo.com/ja/", options: { bold: true, fontFace: MONO } }, { text: "　教材は sample/、答えは answer/ にある。", options: { fontSize: 12.5 } }], inset(c, 48, 30), { px: 30, valign: "middle" });
  }

  await pres.writeFile({ fileName: process.argv[2] });
  await applyTheme(process.argv[2], THEME);
  console.log("wrote", process.argv[2]);
})();

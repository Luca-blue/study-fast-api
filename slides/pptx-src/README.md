# PowerPoint 版の生成スクリプト

`slides/fastapi-intro.pptx` はこのスクリプトから生成しています（pptxgenjs）。

```bash
npm install pptxgenjs react react-dom react-icons sharp
SKILL=<pptx スキルのディレクトリ> node build.js ../fastapi-intro.pptx
```

`SKILL` はテーマ色を書き込む `apply_theme.js` の場所です。PowerPoint 上で直接編集する場合は、このスクリプトは使わなくて構いません。

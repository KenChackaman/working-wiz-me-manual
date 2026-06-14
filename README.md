# 私の取扱説明書ジェネレーター（Working with me）

ブラウザだけで入力し、**1枚のPNG画像**と **PowerPoint（.pptx）** をダウンロードできます。入力データはサーバーに送られません（オフラインでもローカルサーバー起動後に利用可能）。

## 使い方

1. このリポジトリをクローンする
2. 次のいずれかでローカルサーバーを起動する（`file://` 直開きは外部CDNのESM読み込みで失敗することがあります）
   - Python: `python -m http.server 8080`
   - Node がある場合: `npx --yes serve -p 8080`
3. ブラウザで `http://localhost:8080` を開く
4. フォームに入力 → **PNGを保存** / **PowerPointを保存**

## GitHub Pages で公開

1. このフォルダを GitHub リポジトリに push する（例：`main` ブランチ）
2. リポジトリの **Settings → Pages**
3. **Build and deployment → Source** を **GitHub Actions** にする
4. `main`（または `master`）に push すると `.github/workflows/pages.yml` が走り、`index.html` / `styles.css` / `app.js` だけを公開します

初回は Actions の完了後、Pages の URL が表示されるまで数分かかることがあります。

> **注意**: 本ツールは `esm.sh` から `pptxgenjs` と `html-to-image` を読み込みます。社内ポリシーで外部CDNが禁止の場合は、社内npmレジストリへベンダリングするか、バンドルを同梱する改修が必要です。

## 楽天ブランド（ReX）について

画面とスライドは **社内向けの簡易トンマナ**（赤アクセント＋余白多め）です。正式なロゴ・色数値・フォントは **社内の Rakuten Brand & ReX Guidelines** のテンプレートに差し替えてください（本ツールは雛形生成のみです）。

## ライセンス

MIT

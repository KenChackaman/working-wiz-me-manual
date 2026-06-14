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

ヘッダー／フッターに表示するワードマークは、**Wikimedia Commons の `Rakuten_logo.svg`** を `assets/rakuten-logo.svg` として同梱しています（出典・注意事項は `ATTRIBUTION.md`）。

- **著作権表記（Commons）**: PD-textlogo とされていますが、**商標**としての制限は別途あり得ます。
- **正式運用**では、社内の **Rakuten Brand & ReX Guidelines** から入手した最新アセットへ差し替えることを推奨します。

### 公式ロゴへ差し替える（任意）

社内手順で利用が許可されたファイルを `assets/` に配置し、`index.html` の `brand-lockup__logo` / `site-footer__logo` の `src` を差し替えてください。

## ライセンス

MIT

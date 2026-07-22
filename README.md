# zaiko-scanner

在庫商品のJANコードをスマホカメラで読み取り、ソロエルアリーナの商品ページへ直接遷移するツール（テスト運用中）。

## 構成

- `scan.html` / `admin.html` / `index.html`: GitHub Pagesで配信するフロントエンド
- `gas/`: JANコードと商品ページURLの対応表を管理するGAS Web App（バックエンドAPI、スプレッドシート連携）

## 使い方

- スキャン画面（`scan.html`）: JANコードを読み取ると、登録済みの商品ページへ自動遷移
- 管理画面（`admin.html`）: JANコードをスキャン or 手入力し、商品ページURLを登録

## デプロイ

- フロントエンド: GitHub Pages
- バックエンド: `cd gas && clasp push --force` → `clasp deploy` で更新後、`config.js` の `GAS_API_URL` を最新のデプロイURLに合わせる

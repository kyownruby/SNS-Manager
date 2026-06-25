# SNS Manager 仕様書

Threads（Meta）への投稿を管理する Web アプリ。
Google Apps Script（GAS）+ HTML フロントエンドで構成され、投稿の作成・予約・下書き・編集・即時公開を行う。

---

## 1. 概要

| 項目 | 内容 |
|------|------|
| アプリ名 | SNS Manager |
| 対象 SNS | Threads（Meta） |
| 実行基盤 | Google Apps Script（Web App） |
| フロントエンド | HTML / CSS / Vanilla JS（`Index.html` 単一ファイル） |
| データ保存先 | Google スプレッドシート（スタンドアロン管理） |
| 画像保存先 | Google Drive（`SNSManager_Images` フォルダ） |
| タイムゾーン | Asia/Tokyo |
| ランタイム | V8 |

---

## 2. システム構成

```
[ブラウザ / Index.html]
        │  google.script.run
        ▼
[GAS / Code.gs]
   ├─ スプレッドシート（posts シート）  … 投稿データ
   ├─ Google Drive（画像ファイル）       … 画像ホスティング
   └─ Threads Graph API                 … 投稿公開・トークン管理
```

- Web App は `doGet()` で `Index.html` を配信。
- フロントは `google.script.run` 経由で GAS 関数を呼び出す。
- デプロイは GitHub Actions + clasp で自動化（`main` ブランチへの push をトリガー）。

---

## 3. データモデル（posts シート）

スプレッドシートの `posts` シートに 1 投稿 = 1 行で保存。

| 列 | キー | 説明 |
|----|------|------|
| A | `id` | UUID（自動生成） |
| B | `status` | `draft` / `scheduled` / `posted` |
| C | `body` | 投稿本文 |
| D | `image_url` | 画像 URL（複数枚はカンマ区切り） |
| E | `scheduled_at` | 予約日時（`yyyy-MM-dd'T'HH:mm`） |
| F | `posted_at` | 投稿完了日時 |

- スプレッドシート ID はスクリプトプロパティ `SPREADSHEET_ID` で指定。
- `posts` シートが無い場合はヘッダー付きで自動生成。

---

## 4. 機能一覧

### 4.1 投稿作成（新規作成タブ）
- **即時投稿**：その場で Threads に公開し、`posted` で記録。
- **予約投稿**：日時を指定し `scheduled` で保存。1 分ごとのトリガーで自動公開。
- **下書き保存**：`draft` で保存（公開はしない）。
- 本文の文字数カウント（目安 500 文字、超過時は赤表示）。

### 4.2 画像（最大 4 枚）
- ドラッグ&ドロップ / クリック選択 / Ctrl+V 貼り付けに対応。
- プレビュー表示・個別削除が可能。
- アップロード先は Google Drive。「リンクを知っている全員が閲覧可」に設定し、`https://lh3.googleusercontent.com/d/{fileId}` の公開 URL を生成。
- 投稿形態は枚数で自動判定：
  - 0 枚 → `TEXT`
  - 1 枚 → `IMAGE`
  - 2〜4 枚 → `CAROUSEL`（カルーセル）

### 4.3 投稿一覧（一覧タブ）
- ステータスでフィルタ（すべて / 下書き / 予約 / 投稿済）。
- 各投稿をカード表示（ステータス・本文・画像サムネ・日時）。
- 画像サムネクリックでライトボックス拡大表示。
- `posted` 以外は **編集** / **削除** が可能。

### 4.4 投稿編集（編集モーダル）
- 本文・画像・予約日時・ステータスを編集。
- 既存画像の保持/削除 + 新規画像追加に対応（合計最大 4 枚）。
- **保存**（内容更新）または **投稿**（その場で公開）を選択可能。

---

## 5. GAS 関数仕様（Code.gs）

### Web App
| 関数 | 役割 |
|------|------|
| `doGet()` | `Index.html` を配信（iframe 埋め込み許可） |

### スプレッドシート操作
| 関数 | 役割 |
|------|------|
| `getSheet()` | `posts` シート取得（無ければ生成） |
| `generateId()` | UUID 生成 |
| `getPosts()` | 全投稿を新しい順で取得 |
| `saveDraft(body, imageUrl)` | 下書き保存 |
| `schedulePost(body, imageUrl, scheduledAt)` | 予約保存 + トリガー設定 |
| `postNow(body, imageUrl)` | 即時投稿 + 記録 |
| `updatePost(id, body, imageUrl, scheduledAt, status)` | 投稿更新 |
| `deletePost(id)` | 投稿削除 |
| `publishExistingPost(id, body, imageUrls)` | 既存投稿を即時公開 |

### 画像（Google Drive）
| 関数 | 役割 |
|------|------|
| `uploadImage(fileData, fileName, mimeType)` | Base64 画像を Drive に保存し公開 URL を返す |
| `getOrCreateImageFolder()` | 画像保存フォルダ取得/生成 |

### Threads API
| 関数 | 役割 |
|------|------|
| `getThreadsConfig()` | トークン・ユーザー ID を取得 |
| `publishToThreads(body, imageUrls)` | コンテナ作成 → 公開（TEXT/IMAGE/CAROUSEL 自動判定） |
| `createThreadsContainer(apiBase, token, params)` | メディアコンテナ作成 |
| `waitForContainerReady(containerId, token)` | カルーセル子要素の準備完了待ち（最大 15 回 × 2 秒） |

### 予約投稿トリガー
| 関数 | 役割 |
|------|------|
| `checkScheduledPosts()` | 公開時刻を過ぎた `scheduled` を公開 |
| `ensureTrigger()` | 1 分ごとの時間トリガーを設定（重複防止） |

### アクセストークン管理
| 関数 | 役割 |
|------|------|
| `exchangeForLongLivedToken()` | 短期 → 長期トークンへ交換 |
| `refreshLongLivedToken()` | 長期トークンを更新 |
| `ensureTokenRefreshTrigger()` | 30 日ごとの更新トリガーを設定 |

---

## 6. スクリプトプロパティ（要設定）

| キー | 用途 | 必須 |
|------|------|------|
| `SPREADSHEET_ID` | データ保存先スプレッドシート ID | ✅ |
| `THREADS_ACCESS_TOKEN` | Threads API アクセストークン | ✅ |
| `THREADS_USER_ID` | Threads ユーザー ID（未設定時は `me`） | 任意 |
| `THREADS_APP_SECRET` | トークン交換用アプリシークレット | トークン交換時 |

---

## 7. Threads API 連携

- API ベース：`https://graph.threads.net/v1.0/{userId}`
- 投稿フロー：
  1. メディアコンテナ作成（`/threads`）
  2. （カルーセル時）子コンテナの `status` が `FINISHED` になるまで待機
  3. 公開（`/threads_publish`、`creation_id` を渡す）
- トークン：
  - 交換：`th_exchange_token`（短期→長期）
  - 更新：`th_refresh_token`（30 日ごと自動）

---

## 8. デプロイ（CI/CD）

`.github/workflows/deploy-gas.yml`

- **トリガー**：`main` ブランチへの push（`*.gs` / `*.html` / `appsscript.json` 変更時）
- **手順**：clasp インストール → 認証（`CLASPRC_JSON` Secret）→ `clasp push --force`
- **必要 Secret**：`CLASPRC_JSON`（GitHub Secrets）

---

## 9. 制約・注意事項

- Threads API には **トピック設定・下書き保存のエンドポイントは存在しない**（下書きはアプリ側スプレッドシートで独自管理）。
- 画像は **最大 4 枚**。
- Web App のアクセス権限は `MYSELF`（デプロイ者本人のみ）、実行は `USER_DEPLOYING`。
- 予約投稿は 1 分粒度のトリガーで処理（厳密な秒単位ではない）。

---

## 10. 関連ファイル

| ファイル | 役割 |
|----------|------|
| `Code.gs` | GAS バックエンド（API・データ操作・トリガー） |
| `Index.html` | フロントエンド（UI・ロジック） |
| `appsscript.json` | GAS マニフェスト（タイムゾーン・Web App 設定） |
| `.clasp.json` | clasp 設定（scriptId） |
| `.github/workflows/deploy-gas.yml` | 自動デプロイ設定 |

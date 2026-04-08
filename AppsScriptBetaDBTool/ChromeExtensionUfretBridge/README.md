# ChromeExtensionUfretBridge

`manifest.json` を Apps Script プロジェクト配下に置くと、Apps Script 側が JSON マニフェストとして解釈しエラーになる場合があります。

このためリポジトリ上では `manifest.chrome.json.txt` として保存しています。

Chrome拡張として読み込む手順:
1. このフォルダをローカルへコピー
2. `manifest.chrome.json.txt` を `manifest.json` にリネーム
3. `chrome://extensions` で「デベロッパーモード」ON
4. 「パッケージ化されていない拡張機能を読み込む」でこのフォルダを指定

## 現在の運用方針（中間仕様）
- Apps Script 画面との direct bridge（`postMessage` / `bridge-ready`）には依存しません。
- ユーザーが拡張ポップアップを手動実行し、取得した chord + 歌詞テキストを Apps Script 画面へ手動コピペします。
- Apps Script 側の「Uフレットからコード取得」ボタンは、参照URL確認と手動コピペ案内のために使います。

## v1.0.2 変更点（履歴）
- Apps Script の iframe / about:blank / sandbox 構成でも content script が届きやすいよう、`all_frames` に加えて `match_about_blank` / `match_origin_as_fallback` を有効化。
- `content_bridge.js` に `bridge-ready` / `bridge-received` の最小限イベントを追加（タイムアウト切り分け用）。
- `bridge-ping` を受けると `bridge-ready` を返すヘルスチェックを追加。
- `requestId` を持つ応答は success / failure の最終1回のみ。

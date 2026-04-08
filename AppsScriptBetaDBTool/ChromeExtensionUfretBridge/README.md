# ChromeExtensionUfretBridge

`manifest.json` を Apps Script プロジェクト配下に置くと、Apps Script 側が JSON マニフェストとして解釈しエラーになる場合があります。

このためリポジトリ上では `manifest.chrome.json.txt` として保存しています。

Chrome拡張として読み込む手順:
1. このフォルダをローカルへコピー
2. `manifest.chrome.json.txt` を `manifest.json` にリネーム
3. `chrome://extensions` で「デベロッパーモード」ON
4. 「パッケージ化されていない拡張機能を読み込む」でこのフォルダを指定

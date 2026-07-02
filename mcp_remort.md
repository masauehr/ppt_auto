# jma-mcp-remote とは

- jma_mcp（ローカルstdio版）をベースに、HTTP/SSE通信へ切り替えたリモートデプロイ版
- Renderにデプロイし、Claude.ai Web版・デスクトップアプリ・iPhone版から利用できる
- SSEエンドポイント: `https://jma-mcp-remote.onrender.com/sse`
- プロトコル: MCP（Model Context Protocol）/ HTTP + SSEベース

## ローカル版との違い

- 通信方式: stdio（ローカル版）→ HTTP+SSE（リモート版）
- 起動方法: Claude Codeがサブプロセス起動 → Render上で常駐
- 対応クライアント: Claude Code（CLI）のみ → Web版・デスクトップ・iPhoneすべて
- コスト: 無料（ローカル実行） vs Render無料プラン（スリープあり）
- ツールの実装内容は同一（server.pyを共有、全21種）

## Renderへのデプロイ手順

- Renderダッシュボード（dashboard.render.com）で新規Web Serviceを作成
- GitHubリポジトリ `masauehr/jma-mcp-remote` を接続
- Build Command: `pip install -r requirements.txt` / Start Command: `python server.py`
- ログに「Uvicorn running on http://0.0.0.0:XXXXX」と出れば起動成功

## Claude.aiへの接続手順

- 設定 → コネクタ → カスタムコネクタを追加
- URLに `https://jma-mcp-remote.onrender.com/sse` を登録
- Web版で登録した設定はデスクトップアプリ・iPhone版にも自動共有される
- iPhone版もコネクタをONにすれば動作する

> 同じURLを再登録しようとすると「A server with this URL already exists.」と出るが、登録済みという意味なので問題ない

## 出典リンク表示設定

- Claude.aiではCLAUDE.mdが読み込まれないため、出典URLの表示にはProjectsの設定が必要
- Projects → 手順を編集、のカスタム指示欄に「ツール結果末尾の出典URLをそのまま表示すること」という指示を追加する
- この設定はWeb版・デスクトップアプリ・iPhone版すべてで共有される

## クライアント別対応状況

- Claude Code（CLI）: ローカル版を使用（`.mcp.json` で接続）
- Claude.ai Web版・デスクトップアプリ: リモート版を使用（コネクタ登録済み）
- Claude iPhone版: リモート版を使用（コネクタをONにすれば動作）
- いずれのクライアントもMCP対応・気象庁データ取得が可能

## 注意事項・今後の運用 <!-- summary -->

- Render無料プランは15分間アクセスがないとスリープし、次回アクセス時に起動まで30〜60秒かかる
- ローカル版（jma_mcp）にツールを追加した場合は、リモート版（jma_mcp_remote/server.py）にも同じ変更を反映する必要がある（差分は起動部分のみ）
- ローカル版とリモート版を併用し、用途（Claude CodeでのCLI利用か、Web・スマホでの利用か）に応じて使い分けるのが実用的

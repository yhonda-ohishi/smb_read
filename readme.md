# smb_read プロジェクト

このプロジェクトは、SMB共有フォルダのファイル一覧取得・フィルタ・POST送信などを行う Python スクリプト群です。

## smb_listener.py

### 概要
SMBサーバー上の指定フォルダのファイル一覧を取得し、JSON形式でコンソール出力、または指定URLへPOST送信します。ファイルの更新日時や作成日時でフィルタも可能です。

### 主な機能
- SMB共有フォルダ内のファイル・ディレクトリ一覧取得
- ISO 8601形式日時以降の更新/作成ファイルのみフィルタ
- ファイル情報をJSON形式で出力
- 指定URLへPOST送信（オプション）

### 使用方法
```powershell
python smb_listener.py --server <SMBサーバーIP/ホスト名> --share <共有名> --username <ユーザー名> --password <パスワード> [--domain <ドメイン>] [--post-url <URL>] [--since-time <ISO8601日時>] [--folder-path <フォルダパス>]
```

#### オプション例
- `--server` : SMBサーバーのIPアドレスまたはホスト名（必須）
- `--share` : 共有フォルダ名（必須）
- `--username` : ユーザー名（必須）
- `--password` : パスワード（必須）
- `--domain` : Active Directoryドメイン名（省略可）
- `--post-url` : JSONデータをPOSTするURL（省略可）
- `--since-time` : ISO 8601形式日時以降のファイルのみ抽出（例: '2023-01-01T10:00:00+09:00'）
- `--folder-path` : SMB共有内のフォルダパス（例: '/'、省略可）

## smb_processor_and_poster.py

### 概要
SMB共有フォルダから取得したファイル情報を加工し、指定したWeb API等へPOST送信する処理を行うスクリプトです。
電子車検証の出力情報から必要なデータを抽出し、外部APIへ送信することを目的としています。
JSONとPDFの両方の形式を特定のurlにPOST送信することが可能です。
JSONデータを先に送信します。
JSONデータとPDFデータがそろっていないとPOST送信は行われません。

### 主な機能
- SMB共有フォルダからファイル情報取得
- 必要に応じてデータ加工
- 外部API等へのPOST送信

### 使用方法
（詳細はスクリプト内のコメントや引数定義を参照してください。基本的な使い方は`smb_listener.py`と類似です）

---

## 依存パッケージ
- `smbprotocol` または `pysmb`（SMB接続用）
- `requests`（POST送信用）

インストール例:
```powershell
pip install pysmb requests
```

## 注意事項
- Windows環境での動作を想定しています。
- Python 3.11以上推奨（asyncio.run対応）
- SMBサーバーへの接続には適切な権限が必要です。

## ライセンス
MIT License

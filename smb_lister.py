# smb_lister.py

from smb.SMBConnection import SMBConnection
import os
import socket # socketモジュールをインポート
import datetime # datetimeモジュールをインポート
import json # jsonモジュールをインポート

async def list_shared_files():
    # SMB接続情報
    server_ip_or_name = 'server_ip_or_name' # ここにサーバーのIPアドレスまたはホスト名を指定
    share_name = "share_name" # 画像で確認した日本語の共有名
    username = 'user' # ユーザー名
    password = 'password' # パスワード
    domain = 'domain' # ドメイン名、必要な場合は指定
    client_name = socket.gethostname() # Windowsでクライアントのホスト名を取得

    file_details = [] # JSONデータとして格納するためのリスト

    # SMBConnectionのインスタンスを作成
    # use_ntlmv2=True はNTLMv2認証を強制します。Active Directory環境では推奨されます。
    conn = SMBConnection(username, password, client_name, server_ip_or_name, domain=domain, 
            use_ntlm_v2=True,#https://murochi.hateblo.jp/entry/2018/06/30/pysmb%25e3%2581%25ae%25e8%25a8%25ad%25e5%25ae%259a%25e3%2581%25a8smb%25e3%2583%2590%25e3%2583%25bc%25e3%2582%25b8%25e3%2583%25a7%25e3%2583%25b3%25e5%25b7%25ae%25e7%2595%25b0%25e3%2581%25ab%25e3%2581%25a4
            is_direct_tcp=True,)
    try:
        connected = conn.connect(server_ip_or_name, 445)
        if not connected:
            print(f"サーバー {server_ip_or_name} へ接続に失敗しました。")
            return

        # print(f"サーバー {server_ip_or_name} に接続しました。") # JSON出力時は不要なログ

        initial_file_list = conn.listPath(share_name, '/')

        # print(f"共有フォルダ '{share_name}' のファイル一覧:") # JSON出力時は不要なログ
        # print(f"{'種類':<5} {'ファイル名':<30} {'サイズ':<15} {'更新日時':<25} {'作成日時':<25}")
        # print("-" * 100)

        for file in initial_file_list:
            if file.filename in ['.', '..']:
                continue

            full_path = os.path.join('/', file.filename)
            detailed_file_info = None

            try:
                detailed_file_info = conn.getAttributes(share_name, full_path)
            except Exception as attr_err:
                # print(f"DEBUG: Failed to get attributes for {full_path}: {attr_err}") # デバッグログ
                detailed_file_info = None

            item_type = ""
            file_size = None # バイト単位の数値、取得できない場合はNone
            last_write_timestamp = None # datetimeオブジェクト、取得できない場合はNone
            creation_timestamp = None   # datetimeオブジェクト、取得できない場合はNone

            if file.isDirectory:
                item_type = "directory"
            else:
                item_type = "file"

            if detailed_file_info:
                # サイズの取得
                if not detailed_file_info.isDirectory and hasattr(detailed_file_info, 'file_size'):
                    file_size = detailed_file_info.file_size # 数値のまま保持

                # タイムスタンプの取得
                if hasattr(detailed_file_info, 'last_write_time'):
                    try:
                        utc_dt = datetime.datetime.utcfromtimestamp(detailed_file_info.last_write_time)
                        last_write_timestamp = utc_dt.astimezone() # タイムゾーン対応のdatetimeオブジェクト
                    except Exception:
                        pass # 変換エラーの場合はNoneのまま

                if hasattr(detailed_file_info, 'create_time'): # 'creation_time' ではなく 'create_time' を使用
                    try:
                        utc_dt = datetime.datetime.utcfromtimestamp(detailed_file_info.create_time)
                        creation_timestamp = utc_dt.astimezone() # タイムゾーン対応のdatetimeオブジェクト
                    except Exception:
                        pass # 変換エラーの場合はNoneのまま

            # 各項目を辞書として格納
            item_data = {
                "name": file.filename,
                "type": item_type,
                "path": full_path, # 共有ルートからの相対パス
                "size_bytes": file_size, # バイト単位の数値
                "last_write_time": last_write_timestamp.isoformat() if last_write_timestamp else None, # ISO 8601形式の文字列
                "creation_time": creation_timestamp.isoformat() if creation_timestamp else None,     # ISO 8601形式の文字列
            }
            file_details.append(item_data)

        # 全ての情報をJSON形式で出力
        # indent=2 で整形されたJSONが出力されます
        print(json.dumps(file_details, indent=2, ensure_ascii=False))

    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        if 'conn' in locals() and conn: # conn が定義されており、None でないことを確認
            conn.close()
        print("SMB接続を閉じました。")

# asyncio を使って非同期関数を実行
# Python 3.7+ では以下の書き方で良いです
if __name__ == '__main__':
    import asyncio
    asyncio.run(list_shared_files())
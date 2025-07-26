# smb_lister.py

import os
import sys
import socket
import asyncio
import datetime
import json
import argparse
import requests


# 日時オブジェクトをISO 8601形式の文字列に変換するヘルパー関数
class DateTimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        return json.JSONEncoder.default(self, obj)


async def list_and_post_shared_files(
    server_ip_or_name,
    share_name,
    username,
    password,
    domain,
    post_url,
    since_time_str,
    folder_path="/",
):
    from smb.SMBConnection import SMBConnection

    conn = None
    file_details = []

    # フィルター用の日時をパース
    filter_since_time = None
    if since_time_str:
        try:
            # ISO 8601 形式の文字列を datetime オブジェクトに変換
            filter_since_time = datetime.datetime.fromisoformat(since_time_str)
            # タイムゾーン情報がない場合は、ローカルタイムゾーンとして扱う
            if filter_since_time.tzinfo is None:
                filter_since_time = filter_since_time.astimezone()
            print(f"Filtering data since: {filter_since_time.isoformat()}")
        except ValueError as e:
            raise ValueError(
                f"Invalid --since-time format. Use ISO 8601 (e.g., '2023-01-01T10:00:00+09:00'): {e}"
            )

    try:
        conn = SMBConnection(
            username,
            password,
            socket.gethostname(),
            server_ip_or_name,
            domain=domain,
            use_ntlm_v2=True,
            is_direct_tcp=True,
        )

        connected = conn.connect(server_ip_or_name, 445)
        if not connected:
            raise Exception(f"サーバー {server_ip_or_name} へ接続に失敗しました。")

        initial_file_list = conn.listPath(share_name, folder_path)

        max_creation_time = None
        max_last_write_time = None

        for file in initial_file_list:
            if file.filename in [".", ".."]:
                continue

            full_path = os.path.join(folder_path, file.filename)
            detailed_file_info = None

            try:
                detailed_file_info = conn.getAttributes(share_name, full_path)
            except Exception:
                detailed_file_info = None

            item_type = ""
            file_size = None
            last_write_timestamp = None
            creation_timestamp = None

            if file.isDirectory:
                item_type = "directory"
            else:
                item_type = "file"

            if detailed_file_info:
                if not detailed_file_info.isDirectory and hasattr(
                    detailed_file_info, "file_size"
                ):
                    file_size = detailed_file_info.file_size

                if hasattr(detailed_file_info, "last_write_time"):
                    try:
                        utc_dt = datetime.datetime.fromtimestamp(
                            detailed_file_info.last_write_time, datetime.UTC
                        )
                        last_write_timestamp = utc_dt.astimezone()
                    except Exception:
                        pass

                if hasattr(detailed_file_info, "create_time"):
                    try:
                        utc_dt = datetime.datetime.fromtimestamp(
                            detailed_file_info.create_time, datetime.UTC
                        )
                        creation_timestamp = utc_dt.astimezone()
                    except Exception:
                        pass

            # フィルターロジック
            # filter_since_time が指定されている場合、last_write_time または creation_time のいずれかが条件を満たす
            should_include = True
            if filter_since_time:
                is_newer_by_write = False
                is_newer_by_create = False
                if last_write_timestamp and last_write_timestamp >= filter_since_time:
                    is_newer_by_write = True
                if creation_timestamp and creation_timestamp >= filter_since_time:
                    is_newer_by_create = True

                if not (is_newer_by_write or is_newer_by_create):
                    should_include = False  # どちらの条件も満たさない場合は含めない

            if should_include:
                item_data = {
                    "name": file.filename,
                    "type": item_type,
                    "path": full_path,
                    "size_bytes": file_size,
                    "last_write_time": last_write_timestamp,
                    "creation_time": creation_timestamp,
                }
                file_details.append(item_data)

                # 最大タイムスタンプの更新
                if last_write_timestamp and (
                    max_last_write_time is None
                    or last_write_timestamp > max_last_write_time
                ):
                    max_last_write_time = last_write_timestamp

                if creation_timestamp and (
                    max_creation_time is None or creation_timestamp > max_creation_time
                ):
                    max_creation_time = creation_timestamp

        # 最終的なJSON出力の構造
        final_output_data = {
            "data": file_details,
            "last_create_time": max_creation_time,
            "last_write_time": max_last_write_time,
        }

        # コンソールに出力
        print(
            json.dumps(
                final_output_data, indent=2, ensure_ascii=False, cls=DateTimeEncoder
            )
        )

        # POSTリクエストを送信
        if post_url:
            print(f"\nSending data to {post_url}...")
            headers = {"Content-Type": "application/json"}
            # final_output_data を直接ダンプして送信
            response = requests.post(post_url, json=final_output_data, headers=headers)
            response.raise_for_status()
            print(f"POST successful! Status Code: {response.status_code}")
            print(f"Response: {response.text}")

    except Exception as e:
        error_output = {"status": "error", "message": str(e)}
        print(json.dumps(error_output, indent=2, ensure_ascii=False))
        sys.exit(1)
    finally:
        if conn:
            conn.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="SMB共有フォルダのファイル一覧を取得し、JSON形式で指定されたURLにPOSTします。"
    )
    parser.add_argument(
        "--server",
        required=True,
        help="SMBサーバーのIPアドレスまたはホスト名 (例: 172.18.21.102)",
    )
    parser.add_argument("--share", required=True, help="共有フォルダ名 (例: 共有)")
    parser.add_argument(
        "--username",
        required=True,
        help="SMB共有にアクセスするためのユーザー名 (例: yhonda)",
    )
    parser.add_argument("--password", required=True, help="ユーザーのパスワード")
    parser.add_argument(
        "--domain",
        default="",
        help="Active Directoryドメイン名 (例: ohishi.local, 省略可能)",
    )
    parser.add_argument("--post-url", help="JSONデータをPOSTするURL (省略可能)")
    parser.add_argument(
        "--since-time",
        help="指定されたISO 8601形式の日時以降に更新/作成されたファイルのみをフィルター (例: '2023-01-01T10:00:00+09:00')",
    )
    parser.add_argument(
        "--folder-path",
        default="/",
        help="SMB共有内のフォルダパス (例: '/', 省略可能)",
    )
    args = parser.parse_args()

    if sys.version_info >= (3, 11):
        asyncio.run(
            list_and_post_shared_files(
                args.server,
                args.share,
                args.username,
                args.password,
                args.domain,
                args.post_url,
                args.since_time,
                args.folder_path,
            )
        )
    else:
        loop = asyncio.get_event_loop()
        try:
            loop.run_until_complete(
                list_and_post_shared_files(
                    args.server,
                    args.share,
                    args.username,
                    args.password,
                    args.domain,
                    args.post_url,
                    args.since_time,
                )
            )
        finally:
            loop.close()

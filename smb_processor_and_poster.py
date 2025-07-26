# smb_processor_and_poster.py

import argparse
import json
import requests
import sys
import os
import asyncio
import datetime
import socket
from collections import defaultdict
import tempfile
import shutil
import base64
import mimetypes
from dotenv import load_dotenv  # .env ファイルから環境変数を読み込むために必要

# Windowsイベントログに書き込むためのモジュール (Windows環境でのみ利用可能)
try:
    import win32evtlog
    import win32evtlogutil
    import win32con

    IS_WINDOWS = True
except ImportError:
    IS_WINDOWS = False
    print(
        "Warning: pywin32 not found. Windows Event Log logging will be skipped.",
        file=sys.stderr,
    )


load_dotenv()

# 最後の実行日時を保存するファイル名
# LAST_RUN_TIMESTAMP_FILE = "last_run_timestamp.json"

# スクリプトが置かれているディレクトリの絶対パスを取得
# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# LAST_RUN_TIMESTAMP_FILE = os.path.join(SCRIPT_DIR, "last_run_timestamp.json")

# 最後の実行日時を保存するファイル名
# ★ここを修正します★
if getattr(sys, "frozen", False):
    # PyInstallerでフリーズされた実行ファイルの場合
    # pyinstaller でビルドされた実行ファイルでは、sys.executable が実行ファイルのパスを指す
    # 実行ファイルが存在するディレクトリを取得
    APP_DIR = os.path.dirname(sys.executable)
else:
    # 通常のPythonスクリプトとして実行される場合
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

LAST_RUN_TIMESTAMP_FILE = os.path.join(APP_DIR, "last_run_timestamp.json")

# Windowsイベントログのソース名とログファイル名
EVENT_SOURCE_NAME = "SMBProcessorAndPoster"
EVENT_LOG_NAME = "Application"  # 通常はApplicationログに書き込む

# イベントIDの定義
EVENT_ID_SUCCESS = 1000
EVENT_ID_ERROR = 1001
EVENT_ID_START = 1002
EVENT_ID_WARNING = 1003


# Windowsイベントログにメッセージを記録する関数
def log_to_windows_event_log(message: str, event_id: int, event_type: int):
    if not IS_WINDOWS:
        return

    try:
        # イベントソースが登録されているか確認し、なければ登録
        # この処理は一度だけ行えば良いですが、スクリプト内で安全に呼び出すためにtry-exceptで囲みます
        try:
            # ソースが存在するかを試す (OpenEventLogが失敗しないことで確認)
            h = win32evtlog.OpenEventLog(None, EVENT_LOG_NAME)
            win32evtlog.CloseEventLog(h)
        except win32evtlog.error:
            # ソースが存在しない場合、登録を試みる
            print(f"Registering Event Log Source: {EVENT_SOURCE_NAME}", file=sys.stderr)
            win32evtlogutil.AddSourceToRegistry(
                EVENT_SOURCE_NAME,
                os.path.abspath(sys.argv[0]),  # 実行中のスクリプトのパス
                EVENT_LOG_NAME,  # ログファイル名
            )
            print(f"Event Log Source {EVENT_SOURCE_NAME} registered.", file=sys.stderr)

        # イベントログにメッセージを書き込む
        win32evtlogutil.ReportEvent(
            EVENT_SOURCE_NAME,
            event_id,
            eventType=event_type,
            strings=[message],
            data=None,  # data 引数は None のまま
        )
        # print(f"Logged to Event Log (ID: {event_id}, Type: {event_type}): {message}")
    except Exception as e:
        print(f"Error writing to Windows Event Log: {e}", file=sys.stderr)


log_to_windows_event_log(
    f"Using last run timestamp file: {LAST_RUN_TIMESTAMP_FILE} (absolute path: {os.path.abspath(LAST_RUN_TIMESTAMP_FILE)}) ",
    EVENT_ID_START,
    win32con.EVENTLOG_INFORMATION_TYPE,
)


# SMBConnection を扱うためのクラス
class SMBClient:
    def __init__(
        self,
        server_ip_or_name: str,
        share_name: str,
        username: str,
        password: str,
        domain: str,
    ):
        self.server_ip_or_name = server_ip_or_name
        self.share_name = share_name
        self.username = username
        self.password = password
        self.domain = domain
        self.conn = None  # SMBConnection インスタンスを保持

    async def connect(self):
        """SMBサーバーへの接続を確立します。"""
        from smb.SMBConnection import SMBConnection  # 遅延インポート

        try:
            # SMBConnection のインスタンスを作成
            # client_name にはスクリプトを実行しているPCのホスト名を使用
            self.conn = SMBConnection(
                self.username,
                self.password,
                socket.gethostname(),
                self.server_ip_or_name,
                self.domain,
                use_ntlm_v2=True,  # NTLMv2認証を強制
                is_direct_tcp=True,
            )  # 直接TCPポート445を使用

            # サーバーに接続
            connected = self.conn.connect(self.server_ip_or_name, 445)
            if not connected:
                raise Exception(
                    f"SMBサーバー {self.server_ip_or_name} へ接続に失敗しました。"
                )
            print(f"SMBサーバー {self.server_ip_or_name} に接続しました。")
        except Exception as e:
            log_to_windows_event_log(
                f"SMB接続エラー: {e}", EVENT_ID_ERROR, win32con.EVENTLOG_ERROR_TYPE
            )
            raise Exception(f"SMB接続エラー: {e}")

    async def list_files(self, target_folder_path: str = "/") -> list[dict]:
        """
        指定されたSMBフォルダ内のファイルとディレクトリの情報を取得します。

        Args:
            target_folder_path (str): 共有ルートからの対象フォルダのパス (例: '/PDFファイル配車', デフォルトは '/')

        Returns:
            list[dict]: 各ファイル/ディレクトリの詳細情報を含む辞書のリスト
        """
        if not self.conn:
            raise Exception(
                "SMBConnectionが確立されていません。接続を先に確立してください。"
            )

        # パスを正規化 (先頭に / がない場合は追加し、末尾に / がない場合は追加)
        if not target_folder_path.startswith("/"):
            target_folder_path = "/" + target_folder_path
        if not target_folder_path.endswith("/") and target_folder_path != "/":
            target_folder_path += "/"

        print(f"Listing files in SMB path: {self.share_name}{target_folder_path}")
        file_list_raw = self.conn.listPath(self.share_name, target_folder_path)
        detailed_file_list = []

        for file in file_list_raw:
            # 特殊なディレクトリ '.' と '..' はスキップ
            if file.filename in [".", ".."]:
                continue

            # SMBパスを構築 (例: /PDFファイル配車/document.pdf)
            full_path = os.path.join(target_folder_path, file.filename).replace(
                "\\", "/"
            )  # SMBは通常 / を使う

            detailed_file_info = None
            try:
                # 各ファイル/ディレクトリの詳細属性を取得
                detailed_file_info = self.conn.getAttributes(self.share_name, full_path)
            except Exception as attr_err:
                # 属性取得に失敗した場合、エラーメッセージを表示してNoneとして扱う
                log_to_windows_event_log(
                    f"Warning: Failed to get attributes for {full_path}: {attr_err}",
                    EVENT_ID_WARNING,
                    win32con.EVENTLOG_WARNING_TYPE,
                )
                print(
                    f"Warning: Failed to get attributes for {full_path}: {attr_err}",
                    file=sys.stderr,
                )
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
                # ファイルサイズを取得 (ディレクトリの場合は None)
                if not detailed_file_info.isDirectory and hasattr(
                    detailed_file_info, "file_size"
                ):
                    file_size = detailed_file_info.file_size

                # 最終書き込み日時 (更新日時) を取得し、タイムゾーン対応の datetime オブジェクトに変換
                if hasattr(detailed_file_info, "last_write_time"):
                    try:
                        # Unixエポックからの秒数をUTC datetimeオブジェクトに変換 (Python 3.11+ 推奨)
                        utc_dt = datetime.datetime.fromtimestamp(
                            detailed_file_info.last_write_time, datetime.UTC
                        )
                        last_write_timestamp = (
                            utc_dt.astimezone()
                        )  # ローカルタイムゾーンに変換
                    except Exception as e:
                        log_to_windows_event_log(
                            f"Warning: Failed to convert last_write_time for {full_path}: {e}",
                            EVENT_ID_WARNING,
                            win32con.EVENTLOG_WARNING_TYPE,
                        )
                        print(
                            f"Warning: Failed to convert last_write_time for {full_path}: {e}",
                            file=sys.stderr,
                        )
                        pass  # 変換エラーの場合は None のまま

                # 作成日時を取得し、タイムゾーン対応の datetime オブジェクトに変換
                if hasattr(
                    detailed_file_info, "create_time"
                ):  # pysmbは 'create_time' を使用
                    try:
                        utc_dt = datetime.datetime.fromtimestamp(
                            detailed_file_info.create_time, datetime.UTC
                        )
                        creation_timestamp = (
                            utc_dt.astimezone()
                        )  # ローカルタイムゾーンに変換
                    except Exception as e:
                        log_to_windows_event_log(
                            f"Warning: Failed to convert create_time for {full_path}: {e}",
                            EVENT_ID_WARNING,
                            win32con.EVENTLOG_WARNING_TYPE,
                        )
                        print(
                            f"Warning: Failed to convert create_time for {full_path}: {e}",
                            file=sys.stderr,
                        )
                        pass  # 変換エラーの場合は None のまま

            # 取得した情報を辞書としてリストに追加
            item_data = {
                "name": file.filename,
                "type": item_type,
                "path": full_path,  # 共有ルートからの完全なSMBパス
                "size_bytes": file_size,
                "last_write_time": last_write_timestamp,
                "creation_time": creation_timestamp,
            }
            detailed_file_list.append(item_data)

        return detailed_file_list

    def download_file(self, smb_path: str, local_dir: str) -> str | None:
        """
        SMB共有フォルダからファイルをダウンロードします。

        Args:
            smb_path (str): 共有ルートからのファイルのパス (例: /PDFファイル配車/document.pdf)
            local_dir (str): ダウンロード先のローカルディレクトリ

        Returns:
            str: ダウンロードされたファイルのローカルパス。失敗した場合は None。
        """
        if not self.conn:
            raise Exception("SMBConnectionが確立されていません。")

        local_file_name = os.path.basename(smb_path)
        local_full_path = os.path.join(local_dir, local_file_name)

        # ダウンロード先ディレクトリが存在しない場合は作成
        os.makedirs(local_dir, exist_ok=True)

        try:
            print(f"Downloading '{smb_path}' to '{local_full_path}'...")
            with open(local_full_path, "wb") as fp:
                # pysmb の retrieveFile は同期処理
                self.conn.retrieveFile(self.share_name, smb_path, fp)
            print(f"Successfully downloaded: {local_full_path}")
            return local_full_path
        except Exception as e:
            log_to_windows_event_log(
                f"Error downloading '{smb_path}': {e}",
                EVENT_ID_ERROR,
                win32con.EVENTLOG_ERROR_TYPE,
            )
            print(f"Error downloading '{smb_path}': {e}", file=sys.stderr)
            return None

    def close(self):
        """SMB接続を閉じます。"""
        if self.conn:
            self.conn.close()
            print("SMB接続を閉じました。")


# JSONエンコーダー (datetimeオブジェクトをISO 8601形式の文字列に変換)
class DateTimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        return json.JSONEncoder.default(self, obj)


# ファイル名から共通の部分を抽出するヘルパー関数
def extract_common_name(filename: str) -> str | None:
    """
    ファイル名から共通の識別子部分を抽出します。
    例: "20250725135757_北九州１００え５０４３.pdf" -> "北九州１００え５０４３"
    """
    if "_" in filename:
        parts = filename.split("_", 1)  # 最初の '_' で分割
        if len(parts) > 1:
            name_with_ext = parts[1]
            if "." in name_with_ext:
                return name_with_ext.split(".", 1)[0]  # 最初の '.' の前まで
    return None  # パターンに合わない場合


# 拡張子の優先順位を定義 (低い値が優先)
EXT_ORDER = {
    ".json": 0,
    ".pdf": 1,
}


# ファイルを拡張子でソートするためのキー関数
def sort_by_extension(item: dict) -> int:
    """
    ファイルの拡張子に基づいてソート順序を返します。
    .json が .pdf より優先されます。
    """
    ext = os.path.splitext(item.get("name", ""))[1].lower()
    return EXT_ORDER.get(ext, 999)  # 定義されていない拡張子は最後に


def save_last_run_timestamp(timestamp: datetime.datetime):
    """
    指定されたタイムスタンプをファイルに保存します。
    """
    try:
        with open(LAST_RUN_TIMESTAMP_FILE, "w", encoding="utf-8") as f:
            json.dump({"last_write_time": timestamp.isoformat()}, f, indent=2)
        print(f"Last run timestamp saved to {LAST_RUN_TIMESTAMP_FILE}")
        log_to_windows_event_log(
            f"Last run timestamp saved: {timestamp.isoformat()} {LAST_RUN_TIMESTAMP_FILE}",
            EVENT_ID_SUCCESS,
            win32con.EVENTLOG_INFORMATION_TYPE,
        )
    except Exception as e:
        log_to_windows_event_log(
            f"Error saving last run timestamp: {e}",
            EVENT_ID_ERROR,
            win32con.EVENTLOG_ERROR_TYPE,
        )
        print(f"Warning: Failed to save last run timestamp: {e}", file=sys.stderr)


def load_last_run_timestamp() -> str | None:
    """
    ファイルから前回の実行タイムスタンプを読み込みます。
    """
    if not os.path.exists(LAST_RUN_TIMESTAMP_FILE):
        return None
    try:
        with open(LAST_RUN_TIMESTAMP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            log_to_windows_event_log(
                f"Loaded last run timestamp: {data.get('last_write_time', 'None')} from {LAST_RUN_TIMESTAMP_FILE}",
                EVENT_ID_SUCCESS,
                win32con.EVENTLOG_INFORMATION_TYPE,
            )
            return data.get("last_write_time")
    except (json.JSONDecodeError, FileNotFoundError, KeyError) as e:
        log_to_windows_event_log(
            f"Error loading last run timestamp: {e}",
            EVENT_ID_ERROR,
            win32con.EVENTLOG_ERROR_TYPE,
        )
        print(
            f"Warning: Failed to load last run timestamp from {LAST_RUN_TIMESTAMP_FILE}: {e}",
            file=sys.stderr,
        )
        return None


async def process_and_post_data(
    post_url: str,
    smb_server: str,
    smb_share: str,
    smb_user: str,
    smb_pass: str,
    smb_domain: str,
    folder_path: str,
    since_time_str: str,
):
    """
    SMB共有フォルダからファイル情報を取得し、フィルター、ダウンロード、結果をPOSTします。
    """
    smb_client = None
    temp_download_dir = None  # 一時ディレクトリのパスを保持する変数

    # 環境変数からCF-Accessヘッダーを取得
    cf_id = os.environ.get("CF_Access_Client_Id")
    cf_secret = os.environ.get("CF_Access_Client_Secret")

    if not cf_id or not cf_secret:
        log_to_windows_event_log(
            "CF_Access_Client_Id or CF_Access_Client_Secret environment variables are not set.",
            EVENT_ID_WARNING,
            win32con.EVENTLOG_WARNING_TYPE,
        )
        print(
            "Warning: CF_Access_Client_Id or CF_Access_Client_Secret environment variables are not set.",
            file=sys.stderr,
        )
        if post_url:
            print(
                "POSTing without CF-Access headers. This might lead to authentication issues.",
                file=sys.stderr,
            )

    # フィルター用の日時をパース
    filter_since_time = None
    if since_time_str:
        try:
            filter_since_time = datetime.datetime.fromisoformat(since_time_str)
            # タイムゾーン情報がない場合は、ローカルタイムゾーンとして扱う
            if filter_since_time.tzinfo is None:
                filter_since_time = filter_since_time.astimezone()
            print(f"Filtering data since: {filter_since_time.isoformat()}")
        except ValueError as e:
            log_to_windows_event_log(
                f"Invalid --since-time format: {e}",
                EVENT_ID_ERROR,
                win32con.EVENTLOG_ERROR_TYPE,
            )
            raise ValueError(
                f"Invalid --since-time format. Use ISO 8601 (e.g., '2023-01-01T10:00:00+09:00'): {e}"
            )

    # 処理開始をイベントログに記録
    log_to_windows_event_log(
        f"SMBProcessorAndPoster started for share '{smb_share}' path '{folder_path}'.",
        EVENT_ID_START,
        win32con.EVENTLOG_INFORMATION_TYPE,
    )

    try:
        # 1. SMBサーバーに接続
        smb_client = SMBClient(smb_server, smb_share, smb_user, smb_pass, smb_domain)
        await smb_client.connect()

        # 2. SMB共有フォルダから直接ファイル情報を取得
        print(
            f"SMB共有フォルダ '{smb_share}' 内のパス '{folder_path}' からファイル情報を取得中..."
        )
        initial_file_list = await smb_client.list_files(folder_path)
        print(f"合計 {len(initial_file_list)} 個のアイテムが見つかりました。")

        # 3. フィルター、グループ化、ダウンロード、情報更新を行うための準備
        raw_grouped_files = defaultdict(list)  # フィルター前の全グループ
        processed_other_items = (
            []
        )  # グループ化されないアイテム (ディレクトリなど) を格納

        max_creation_time = None
        max_last_write_time = None

        for item in initial_file_list:
            # フィルターロジック
            should_include = True
            if filter_since_time:
                # ★修正点: フィルターは last_write_time のみで、厳密に大きい (>) を使用★
                # creation_time はフィルター条件に含めない
                if (
                    item["last_write_time"] is None
                    or item["last_write_time"] <= filter_since_time
                ):
                    should_include = False

            if should_include:
                # 最大タイムスタンプの更新 (フィルターされたアイテムのみ)
                if item["last_write_time"] and (
                    max_last_write_time is None
                    or item["last_write_time"] > max_last_write_time
                ):
                    max_last_write_time = item["last_write_time"]

                # creation_time の最大値も引き続き追跡 (これはフィルターには使わないが、情報として)
                if item["creation_time"] and (
                    max_creation_time is None
                    or item["creation_time"] > max_creation_time
                ):
                    max_creation_time = item["creation_time"]

                # グループ化またはその他のアイテムとして分類
                if item.get("type") == "file" and item.get("path"):
                    common_name = extract_common_name(item["name"])
                    if common_name:
                        raw_grouped_files[common_name].append(item)
                    else:
                        # パターンに合わないファイルはそのまま processed_other_items に追加 (ダウンロードはしない)
                        processed_other_items.append(item)
                else:  # ディレクトリはダウンロードしないが、情報として含める
                    processed_other_items.append(item)

        # グループごとに処理 (JSON/PDF揃いチェック、ソート、ダウンロード、情報追加)
        processed_groups_list = []  # 最終的な、JSON/PDFが揃ったグループのみを格納

        # ファイルダウンロードのための一時ディレクトリを作成
        temp_download_dir = tempfile.mkdtemp(prefix="smb_download_")
        print(f"一時ダウンロードディレクトリを作成しました: {temp_download_dir}")

        for common_name in sorted(raw_grouped_files.keys()):
            group = raw_grouped_files[common_name]

            # JSONとPDFが存在するかチェック
            has_json = False
            has_pdf = False
            for item in group:
                ext = os.path.splitext(item.get("name", ""))[1].lower()
                if ext == ".json":
                    has_json = True
                elif ext == ".pdf":
                    has_pdf = True

            if not (has_json and has_pdf):
                print(f"Skipping group '{common_name}': Missing .json or .pdf file.")
                log_to_windows_event_log(
                    f"Skipping group '{common_name}': Missing .json or .pdf file.",
                    EVENT_ID_WARNING,
                    win32con.EVENTLOG_WARNING_TYPE,
                )
                continue  # JSONまたはPDFが揃っていない場合はこのグループをスキップ

            print(f"Processing group '{common_name}': Both .json and .pdf found.")

            group.sort(key=sort_by_extension)  # 揃っている場合はソート

            current_group_info = {"common_name": common_name, "files": []}

            for item in group:
                smb_path = item["path"]

                local_path = None
                download_success = False
                file_blob_base64 = None
                file_mime_type = None

                if item.get("type") == "file":
                    local_path = smb_client.download_file(smb_path, temp_download_dir)
                    download_success = local_path is not None

                    if download_success:
                        try:
                            with open(local_path, "rb") as f:
                                file_content = f.read()
                                file_blob_base64 = base64.b64encode(
                                    file_content
                                ).decode("utf-8")
                            file_mime_type, _ = mimetypes.guess_type(local_path)
                            if file_mime_type is None:
                                file_mime_type = (
                                    "application/octet-stream"  # 不明な場合は汎用タイプ
                                )
                        except Exception as e:
                            log_to_windows_event_log(
                                f"Error reading or encoding file {local_path}: {e}",
                                EVENT_ID_ERROR,
                                win32con.EVENTLOG_ERROR_TYPE,
                            )
                            print(
                                f"Error reading or encoding file {local_path}: {e}",
                                file=sys.stderr,
                            )
                            file_blob_base64 = None
                            file_mime_type = None
                            download_success = (
                                False  # エンコード失敗もダウンロード失敗とみなす
                            )

                new_item = item.copy()
                new_item["local_path"] = local_path  # ダウンロード先のローカルパス
                new_item["download_success"] = (
                    download_success  # ダウンロード成否フラグ
                )
                new_item["blob"] = file_blob_base64  # Base64エンコードされたデータ
                new_item["mime_type"] = file_mime_type  # MIMEタイプ
                current_group_info["files"].append(new_item)

            processed_groups_list.append(
                current_group_info
            )  # 揃ったグループのみを最終リストに追加

        # POSTするペイロードのリストを構築 (JavaScriptのsendData形式に合わせる)
        post_payloads = []
        for group_info in processed_groups_list:
            for file_item in group_info["files"]:
                if (
                    file_item["download_success"] and file_item["blob"]
                ):  # ダウンロード成功したファイルのみ
                    send_data_item = {
                        "blob": file_item["blob"],
                        "filename": file_item["name"],  # multi.filename に相当
                        "type": file_item["mime_type"],  # multi.type に相当
                        # 必要であれば、ここに他のメタデータも追加可能
                        "smb_path": file_item["path"],
                        "common_name": group_info["common_name"],
                        "last_write_time": file_item["last_write_time"],
                        "creation_time": file_item["creation_time"],
                        "size_bytes": file_item["size_bytes"],
                    }
                    post_payloads.append(send_data_item)

        # コンソールに出力するJSONは、POSTするペイロードのリスト (blobは省略)
        console_payloads = []
        for payload_item in post_payloads:
            console_item = payload_item.copy()  # オリジナルをコピー
            if "blob" in console_item:
                console_item["blob"] = "..."  # blob の代わりに省略記号
            console_payloads.append(console_item)

        print(
            json.dumps(
                console_payloads, indent=2, ensure_ascii=False, cls=DateTimeEncoder
            )
        )

        # POSTリクエストを送信
        if post_url:
            print(f"\nSending data to {post_url}...")

            headers = {
                "Content-Type": "application/json",
                "CF-Access-Client-Id": cf_id,
                "CF-Access-Client-Secret": cf_secret,
            }

            for (
                payload_item
            ) in post_payloads:  # こちらは blob を含むオリジナルリストを使用
                print(f"POSTing file: {payload_item['filename']}")
                try:
                    # json=payload_item の代わりに data=json.dumps(...) を使用し、DateTimeEncoder を適用
                    response = requests.post(
                        post_url,
                        data=json.dumps(payload_item, cls=DateTimeEncoder),
                        headers=headers,
                    )
                    response.raise_for_status()  # HTTPエラーがあれば例外を発生させる
                    log_to_windows_event_log(
                        f"POST successful for {payload_item['filename']}. Status: {response.status_code}",
                        EVENT_ID_SUCCESS,
                        win32con.EVENTLOG_INFORMATION_TYPE,
                    )
                    print(
                        f"  POST successful for {payload_item['filename']}! Status Code: {response.status_code}"
                    )
                    # print(f"  Response: {response.text}") # レスポンスが長い場合はコメントアウト
                except requests.exceptions.RequestException as e:
                    log_to_windows_event_log(
                        f"Error POSTing {payload_item['filename']}: {e}",
                        EVENT_ID_ERROR,
                        win32con.EVENTLOG_ERROR_TYPE,
                    )
                    print(
                        f"  Error POSTing {payload_item['filename']}: {e}",
                        file=sys.stderr,
                    )
                except Exception as e:
                    log_to_windows_event_log(
                        f"Unexpected error during POST for {payload_item['filename']}: {e}",
                        EVENT_ID_ERROR,
                        win32con.EVENTLOG_ERROR_TYPE,
                    )
                    print(
                        f"  Unexpected error during POST for {payload_item['filename']}: {e}",
                        file=sys.stderr,
                    )

            print("All POST requests attempted.")

        # 正常終了時に last_write_time の最大値を保存
        if max_last_write_time:
            save_last_run_timestamp(max_last_write_time)
            log_to_windows_event_log(
                f"Script completed successfully. Last processed timestamp: {max_last_write_time.isoformat()}",
                EVENT_ID_SUCCESS,
                win32con.EVENTLOG_INFORMATION_TYPE,
            )
        else:
            log_to_windows_event_log(
                "Script completed successfully. No new files processed.",
                EVENT_ID_SUCCESS,
                win32con.EVENTLOG_INFORMATION_TYPE,
            )

    except Exception as e:
        error_message = f"Script execution failed: {e}"
        log_to_windows_event_log(
            error_message, EVENT_ID_ERROR, win32con.EVENTLOG_ERROR_TYPE
        )
        error_output = {"status": "error", "message": str(e)}
        print(json.dumps(error_output, indent=2, ensure_ascii=False))
        sys.exit(1)  # エラー終了
    finally:
        # SMB接続を閉じる
        if smb_client:
            smb_client.close()
        # 一時ダウンロードディレクトリのクリーンアップ
        if temp_download_dir and os.path.exists(temp_download_dir):
            try:
                print(
                    f"一時ダウンロードディレクトリを削除しています: {temp_download_dir}"
                )
                shutil.rmtree(temp_download_dir)
                print("一時ディレクトリの削除が完了しました。")
            except Exception as e:
                log_to_windows_event_log(
                    f"Warning: Failed to clean up temporary directory {temp_download_dir}: {e}",
                    EVENT_ID_WARNING,
                    win32con.EVENTLOG_WARNING_TYPE,
                )
                print(
                    f"Warning: 一時ディレクトリ {temp_download_dir} の削除に失敗しました: {e}",
                    file=sys.stderr,
                )


if __name__ == "__main__":
    # コマンドライン引数のパーサーを設定
    parser = argparse.ArgumentParser(
        description="SMB共有フォルダからファイル情報を直接取得し、フィルター、ダウンロード、結果をPOSTします。",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--post-url", help="処理済みデータをPOSTするターゲットURL (省略可能)。"
    )

    # SMB接続情報
    parser.add_argument(
        "--smb-server", required=True, help="SMBサーバーのIPアドレスまたはホスト名。"
    )
    parser.add_argument("--smb-share", required=True, help="SMB共有フォルダ名。")
    parser.add_argument("--smb-user", required=True, help="SMB接続用ユーザー名。")
    parser.add_argument("--smb-pass", required=True, help="SMB接続用パスワード。")
    parser.add_argument(
        "--smb-domain", default="", help="SMB接続用ドメイン名 (省略可能)。"
    )
    parser.add_argument(
        "--folder-path",
        default="/",
        help="SMB共有内の対象フォルダのパス (例: /PDFファイル配車, デフォルトはルート /)。",
    )

    # --since-time 引数のデフォルト値を設定
    loaded_since_time = (
        load_last_run_timestamp()
    )  # ファイルから前回のタイムスタンプを読み込む
    parser.add_argument(
        "--since-time",
        default=loaded_since_time,
        help=f"""指定されたISO 8601形式の日時以降に更新/作成されたファイルのみをフィルター。
(例: '2023-01-01T10:00:00+09:00').
デフォルトは前回実行時の最大更新日時 ({loaded_since_time if loaded_since_time else 'なし'})。
注意: フィルターは last_write_time のみが基準となり、厳密に新しい (>) ファイルのみが処理されます。
これにより、重複処理を防ぎます。""",
    )

    args = parser.parse_args()

    # post-url が指定されていない場合は警告を表示
    if not args.post_url:
        print(
            "Warning: --post-url was not provided. Data will be printed to console but not POSTed."
        )

    # asyncio.run の呼び出し (Python 3.11+ とそれ以前のバージョンに対応)
    if sys.version_info >= (3, 11):
        asyncio.run(
            process_and_post_data(
                args.post_url,
                args.smb_server,
                args.smb_share,
                args.smb_user,
                args.smb_pass,
                args.smb_domain,
                args.folder_path,
                args.since_time,
            )
        )
    else:
        loop = asyncio.get_event_loop()
        try:
            loop.run_until_complete(
                process_and_post_data(
                    args.post_url,
                    args.smb_server,
                    args.smb_share,
                    args.smb_user,
                    args.smb_pass,
                    args.smb_domain,
                    args.folder_path,
                    args.since_time,
                )
            )
        finally:
            loop.close()

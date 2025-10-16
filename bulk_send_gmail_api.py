from __future__ import annotations
import os #ファイル操作、パス操作
import csv #csvファイルを扱う
import random
import time #sleepなどを行う
import base64 #バイナリデータをBase64(テキスト形式)にエンコード/デコードする(Gmail APIは本文をBase64形式で送信する必要がある)
import mimetypes #ファイルの拡張子からMIMEタイプ(Content-type)を推測する。添付ファイルの種類を自動判定する
from typing import Dict, List, Optional

from email.mime.text import MIMEText #プレーンテキスト、HTMLテキストをMIME形式のオブジェクトにする
from email.mime.base import MIMEBase #添付ファイルのMIMEオブジェクトを作る基本クラス。ファイルのバイナリデータを格納してBase64エンコードするのに使う
from email.mime.multipart import MIMEMultipart #複数パート(text+html+添付)のメールをまとめるクラス。mixedやalternativeといったMIME構造を組み立てる
from email import encoders #添付ファイルをBase64でエンコードするための補助

from googleapiclient.discovery import build #Google各種APIを呼び出すためのServiceオブジェクトを作成する関数
from googleapiclient.errors import HttpError #Gmail APIへのリクエスト時に起きるHTTPエラーをキャッチするための例外クラス。
from google_auth_oauthlib.flow import InstalledAppFlow #OAuth2.0認証フロー(ローカル環境でブラウザを開いて認証する)を簡単に扱うクラス。初回実行時にgoogleアカウントで許可を得てトークンを発行するために使用
from google.auth.transport.requests import Request #資格情報(token)のリフレッシュ処理で使用するHTTPリクエストオブジェクト
from google.oauth2.credentials import Credentials #認証済みのユーザー資格情報(access token, refresh token, 期限など)を保持するクラス。token.jsonとの読み書きで使用

SCOPES = ['https://www.googleapis.com/auth/gmail.send'] #Gmail APIを使う場合、OAuth2.0認証時にどの権限を使うかを指定

# 全メール共通の添付（複数OK）
# assetsフォルダ内に保存
COMMON_ATTACHMENTS = [
]

def get_gmail_service() -> any: #-> any は「戻り値は何でもあり」, 実際は Gmail API service オブジェクト が返ります
    creds: Optional[Credentials] = None #creds は「資格情報（ログイン情報）」を入れる変数, Optional[Credentials] は「Credentials 型か、無ければ None」
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES) #token.jsonから資格情報を取得する, SCOPESの権限を使う
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request()) #Request() は「Googleに再発行をお願いするためのリクエストオブジェクト」
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES) #credentials.json（Google Cloud Console から取得した OAuth クライアント情報）を使って認証の準備
            # 初回はブラウザで許可 → token.json を自動保存
            creds = flow.run_local_server(port=0) #flow.run_local_server(port=0) を実行すると、ブラウザが開きます。
        with open('token.json', 'w') as f:
            f.write(creds.to_json()) #新しい資格情報を token.json に保存します。
    return build('gmail', 'v1', credentials=creds)

    """
	最後に、Google API クライアントライブラリの build を使って Gmail サービスを作ります。
	•	'gmail' → Gmail API
	•	'v1' → バージョン1
	•	credentials=creds → さっき準備した認証情報を使う
	この戻り値が「serviceオブジェクト」で、実際の送信に使います。
    例service.users().messages().send(...).execute()
    """

def make_mime_message(sender: str,
                      to_addr: str,
                      subject: str,
                      text_body: Optional[str] = None,
                      html_body: Optional[str] = None,
                      attachments: Optional[List[str]] = None):
    """
    プレーン/HTML/添付に対応するMIMEメッセージを作成
    """
    if attachments or (text_body and html_body):
        msg = MIMEMultipart('mixed')  # 添付ありや、text+htmlのときはmultipartに
        alt = MIMEMultipart('alternative')
        if text_body:
            alt.attach(MIMEText(text_body, 'plain', _charset='utf-8')) 
        if html_body:
            alt.attach(MIMEText(html_body, 'html', _charset='utf-8'))
        msg.attach(alt) #外側の mixed に内側の alternative を入れる
    else:
        # どちらか片方のみ
        # MIMEText は「本文だけのメールオブジェクト」を作るクラス。
        if html_body:
            msg = MIMEText(html_body, 'html', _charset='utf-8')
        else:
            msg = MIMEText(text_body or '', 'plain', _charset='utf-8')

    if isinstance(msg, MIMEMultipart): #msg が MIMEMultipart（複数パートのメール）なら
        msg['To'] = to_addr
        msg['From'] = sender
        msg['Subject'] = subject
    else:
        msg['To'] = to_addr
        msg['From'] = sender
        msg['Subject'] = subject

    # 添付ファイル
    if attachments:
        for path in attachments:
            path = path.strip()
            if not path:
                continue
            if not os.path.exists(path):
                print(f'[WARN] 添付が見つかりません: {path}')
                continue
            ctype, encoding = mimetypes.guess_type(path) 
            """
            もし ctype が None(判別できない)
            または encoding がある(例:gzip圧縮など)場合は、
            → 無理に推測せずに "application/octet-stream" （「ただのバイナリデータ」）として扱う。
            """
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1) #ctype = "image/jpeg" → maintype = "image", subtype = "jpeg"
            with open(path, 'rb') as fp: #画像やPDFなどの添付ファイルは必ず バイナリモードで開きます
                part = MIMEBase(maintype, subtype) #MIMEBase は 添付ファイル用のMIMEオブジェクトを作るクラス
                part.set_payload(fp.read()) #part の「中身」にファイルの生データを詰め込む
            encoders.encode_base64(part) #添付ファイルの中身（バイナリデータ）を Base64 という文字列形式に変換する
            filename = os.path.basename(path) #フルパスから最後のファイル名だけを返します
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            # ルートがmultipartでない場合に備え、multipartに包む
            if not isinstance(msg, MIMEMultipart):
                wrapper = MIMEMultipart('mixed')
                wrapper['To'] = msg['To']
                wrapper['From'] = msg['From']
                wrapper['Subject'] = msg['Subject']
                wrapper.attach(msg)
                msg = wrapper
            msg.attach(part)

    return msg

def to_gmail_raw(msg) -> Dict:
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode() #「メールの生データ（バイト列）」に変換 -> バイトデータを Base64 でエンコード -> decodeで文字列 (str) に直します
    #urlsafe_b64encode:通常のbase64で使われる+と/を-と_に置き換えら形式 -> URLやJSONでも安全に扱えるようにする
    return {'raw': raw} #Gmail APIのmessage.sendメソッドではJSONリクエストボディを送る必要がある

def send_with_retry(service, user_id: str, body: Dict, max_retries: int = 5):
    """
    429/403(レート/クォータ)などに指数バックオフでリトライ
    """
    delay = 1.5
    for attempt in range(1, max_retries + 1):
        try:
            return service.users().messages().send(userId=user_id, body=body).execute() #メールの送信
        except HttpError as e:
            status = getattr(e, 'status_code', None) or (e.resp.status if hasattr(e, 'resp') else None)
            if status in (403, 429, 500, 503):
                print(f'[RETRY {attempt}] status={status} {e}. {delay:.1f}s待機…')
                time.sleep(delay)
                delay *= 2
                continue
            raise

def read_csv_rows(csv_path: str) -> List[Dict]:
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        # newline = '' : Pythonのcsvモジュールを使うときのおすすめ設定。改行コードの違い（WindowsのCRLF、MacのLF）による不具合を防ぐ
        #UTF-8（BOM付きもOK）で読み込む
        return list(csv.DictReader(f)) #csv.DictReader は CSV ファイルを 1行ずつ「辞書型」で読み込むクラス

def main(csv_path: str,
         sender: str,
         dry_run: bool = False,
         throttle_sec: float = 0.0):
    """
    csv_path: CSVのパス
    sender  : 送信者のGmailアドレス
    dry_run : Trueなら実際は送らず内容だけ表示
    throttle_sec : 送信間隔（レート調整用）
    """
    rows = read_csv_rows(csv_path)
    if not rows:
        print('CSVが空です。')
        return

    service = get_gmail_service()

    success, fail = 0, 0
    batch_size = 50            # 何通ごとに休憩するか
    batch_pause_sec = 60       # 何秒休憩するか

    for i, row in enumerate(rows, start=1): #enumerate()番号と要素を同時に扱える
        # CSV列を辞書として利用。差し込みOK（{name} など）
        email = (row.get('Email') or '').strip()
        if not email:
            print(f'[{i}] Email列が空なのでスキップ')
            continue

        # 件名・本文
        # subject_tpl = (row.get('subject') or '（件名未設定）').strip()  ##CSVファイルにsubject rowがある場合
        subject_tpl = 'Subject' ##固定のsubject
        # body_tpl = (row.get('body') or '').strip() ##CSVファイルにbodyがある場合
        with open("send.txt", "r", encoding="utf-8") as f:
            body_tpl = f.read()
        body_html_tpl = (row.get('body_html') or '').strip()

        # 差し込み（存在しないキーがあるとKeyErrorになるのでフォールバック）
        def safe_format(tpl: str, data: Dict) -> str:
            try:
                return tpl.format(**data) #テンプレートに辞書式データの差し込んだ文章を返す
            except KeyError as ke:
                # 足りないキーはそのまま残す
                return tpl

        subject = safe_format(subject_tpl, row)
        body_text = safe_format(body_tpl, row) if body_tpl else None
        body_html = safe_format(body_html_tpl, row) if body_html_tpl else None

        # 添付リスト
        attachments_str = row.get('attachments') or ''
        # CSV側にも複数指定がある想定（; または , 区切りを許容）
        if attachments_str:
            for sep in [';', ',']:
                attachments_str = attachments_str.replace(sep, '\n')
        row_attachments = [p.strip() for p in attachments_str.splitlines() if p.strip()]
        # 共通 + 行ごとの添付（重複除去 & 空要素除去）
        attachments = list(dict.fromkeys([p for p in (COMMON_ATTACHMENTS or []) + row_attachments if p]))

        # MIME作成
        msg = make_mime_message(
            sender=sender,
            to_addr=email,
            subject=subject,
            text_body=body_text,
            html_body=body_html,
            attachments=attachments
        )
        raw = to_gmail_raw(msg)

        if dry_run:
            print(f'[{i}] (DRY-RUN) to={email}, subject="{subject}", attachments={len(attachments)}')
        else:
            try:
                send_with_retry(service, 'me', raw)
                print(f'[{i}] 送信成功: {email}')
                success += 1
            except Exception as e:
                print(f'[{i}] 送信失敗: {email} -> {e}')
                fail += 1
        if throttle_sec > 0:
            jitter = random.uniform(0.0, throttle_sec * 0.6)
            time.sleep(throttle_sec + jitter)

        # バッチ休止
        if i % batch_size == 0 and not dry_run:
            print(f'--- バッチ休止: {batch_pause_sec}秒 ---')
            time.sleep(batch_pause_sec)

    print(f'完了: 成功 {success} / 失敗 {fail} / 合計 {len(rows)}')

if __name__ == '__main__':
    # 使い方例：
    # python bulk_send_gmail_api.py
    CSV_PATH = 'test.csv'            # 送る人のCSVファイル
    SENDER   = 'sender@example.com'          # SenderのEmail
    DRY_RUN  = True       # 先に True で内容確認→問題なければ False
    THROTTLE = 0.7         # 軽いウェイト（レート保護）
    main(CSV_PATH, SENDER, DRY_RUN, THROTTLE)
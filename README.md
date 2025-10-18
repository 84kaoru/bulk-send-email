# Bulk Gmail Sender using Gmail API

**Gmail API** を利用して CSV ファイルに基づき複数の宛先へメールを一括送信する Python スクリプトです。  
プレーンテキスト／HTML本文、共通・個別添付ファイル、送信間隔の調整、エラー時の自動リトライなど、実運用を想定した機能を備えています。

---

## 1. 必要環境

- Python 3.8 以上
- Gmail アカウント
- Gmail API を有効にした Google Cloud Console プロジェクト
- OAuth クライアント ID（`credentials.json`）

---

## 2. セットアップ

### 2.1 リポジトリの準備

スクリプトを任意のフォルダに配置します。

### 2.2 仮想環境（任意）

```bash
python -m venv venv
```
有効化  
  
macOS  
```bash
source venv/bin/activate
```

Windows
```bash
venv\Scripts\activate
```
### 2.3 必要ライブラリのインストール
```bash
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```
## 3. Gmail API の有効化と OAuth 認証
1.	Google Cloud Console にアクセス
2.	「APIとサービス」→「ライブラリ」→「Gmail API」を検索して有効化
3.	「認証情報」を開き、「OAuth クライアント ID」を作成
- アプリケーションの種類：「デスクトップアプリ」
4.	credentials.json をダウンロードしてスクリプトと同じフォルダに配置
5.	初回実行時にブラウザが開き、Google アカウントで認証
•	認証後、自動的に token.json が生成され、以降は再認証不要
## 4. CSV ファイルの準備

送信対象と本文・添付情報を CSV で管理します。例：test.csv  
```
Email,name,subject,body,body_html,attachments
taro@example.com,太郎さん,ご案内,{name}、いつもありがとうございます。,,
```

## 5. 添付ファイル設定

全メール共通の添付は COMMON_ATTACHMENTS に設定します。
```python
COMMON_ATTACHMENTS = [
    'assets/common1.pdf',
    'assets/common2.jpg',
]
```
CSV の添付指定と併用可能で、重複は自動的に除去されます。
## 6. スクリプト設定と実行

### 6.1 設定

スクリプト末尾の以下の部分を編集します。
```python
if __name__ == '__main__':
    CSV_PATH = 'test.csv'            # 送る人のCSVファイル
    SENDER   = 'sender@example.com'          # SenderのEmail
    DRY_RUN  = True       # 先に True で内容確認→問題なければ False
    THROTTLE = 0.7         # 軽いウェイト（レート保護）
    main(CSV_PATH, SENDER, DRY_RUN, THROTTLE)
```
### 6.2 実行
```bash
python bulk_send_gmail_api.py
```

## 7. Dry Run モード

DRY_RUN = True に設定すると、送信先・件名・添付ファイル数などを確認するだけで、実際には送信しません。
本番前のテストに使用します。

## 8. レート制限とリトライ

Gmail API にはレート制限があります。
このスクリプトでは以下の方法で制御します。
- THROTTLE で送信間隔を設定
- 50通ごとに 60 秒間自動休止
- 403 / 429 / 500 / 503 エラー時は指数バックオフで自動リトライ

## 9. 注意事項
- Gmail API の送信上限に注意（1日・1分あたり）
- 添付ファイルの合計サイズは Gmail の制限（25MB）を超えないようにする
- CSV は UTF-8（BOM付き）推奨
- Google の利用規約・法令を遵守すること

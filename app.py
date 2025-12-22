import streamlit as st
import fitz  # PyMuPDF
import io
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# 1. APIの権限範囲（スコープ）の設定
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.title("📄 PDFをGoogleスライドに変換 (画像貼り付け)")
st.caption("PDFの各ページを高画質な画像として、新しいGoogleスライドに1枚ずつ貼り付けます。")

# 認証処理の関数
def authenticate_google():
    creds = None
    # Streamlitのセッション内で認証情報を保持
    if 'google_creds' in st.session_state:
        creds = st.session_state.google_creds

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # secrets.toml または Streamlit Cloud の Secrets から情報を取得
            client_config = {
                "installed": {
                    "client_id": st.secrets["google_oauth"]["client_id"],
                    "project_id": st.secrets["google_oauth"]["project_id"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                    "client_secret": st.secrets["google_oauth"]["client_secret"],
                    "redirect_uris": ["http://localhost"]
                }
            }
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            # ローカル実行時はサーバーを立て、クラウド時はURLを表示
            creds = flow.run_local_server(port=0)
        st.session_state.google_creds = creds
    return creds

# メイン処理
if st.button("Googleアカウントでログイン"):
    try:
        st.session_state.creds = authenticate_google()
        st.success("ログインに成功しました！")
    except Exception as e:
        st.error(f"ログインエラー: {e}")

uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

if uploaded_file and 'creds' in st.session_state:
    if st.button("🚀 スライド作成を開始"):
        creds = st.session_state.creds
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        try:
            # 1. 新規スライドの作成
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            
            # PDFの読み込み
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, page in enumerate(doc):
                status_text.text(f"処理中: {i+1} / {total_pages} ページ目")
                
                # 2. PDFページを画像に変換（高画質設定）
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # 3. 画像をGoogleドライブに一時保存
                file_metadata = {
                    'name': f'temp_slide_img_{i}.png',
                    'parents': ['root'] # ルート直下に保存
                }
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                file_id = file.get('id')
                
                # 4. Slides APIからアクセスできるように権限を一時公開
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                
                # 画像の直リンクURL
                file_url = f"https://drive.google.com/uc?id={file_id}"

                # 5. スライドの追加と画像の挿入
                requests = [
                    {'createSlide': {'objectId': f'page_{i}'}}, # スライド作成
                    {'createImage': {
                        'elementProperties': {'pageObjectId': f'page_{i}'},
                        'url': file_url
                    }}
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                
                # 6. 使い終わった一時画像ファイルを削除（ドライブを汚さないため）
                drive_service.files().delete(fileId=file_id).execute()
                
                progress_bar.progress((i + 1) / total_pages)

            # 最初の空スライド（デフォルトの1枚目）を削除（任意）
            # slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': [{'deleteObject': {'objectId': 'p'}}]}).execute()

            st.balloons()
            st.success("✅ スライドが完成しました！")
            st.markdown(f"### [作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

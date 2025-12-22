import streamlit as st
import fitz  # PyMuPDF
import io
import os
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# http通信（localhost）を許可する設定
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# 1. APIの権限範囲（スコープ）の設定
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.title("📄 PDFをGoogleスライドに変換 (画像貼り付け)")
st.caption("PDFの各ページを高画質な画像として、新しいGoogleスライドに1枚ずつ貼り付けます。")

# --- 認証処理の関数 ---
def authenticate_google():
    creds = None
    if 'google_creds' in st.session_state:
        creds = st.session_state.google_creds

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.google_creds = creds
            except:
                creds = None

        if not creds:
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
            
            # Flowを初期化
            flow = Flow.from_client_config(
                client_config, 
                scopes=SCOPES,
                redirect_uri='http://localhost'
            )
            
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            
            st.info("💡 Google認証が必要です。")
            st.markdown(f"**手順1:** [👉 ここをクリックしてGoogle認証を開く]({auth_url})")
            st.write("**手順2:** 認証後、ブラウザがエラーになります。その時の**アドレスバー（URL）の内容をすべてコピー**して貼り付けてください。")
            
            # 入力欄
            auth_response = st.text_input("**手順3:** コピーしたURLをここに貼り付けてEnter:")
            
            if auth_response:
                try:
                    # 【重要】URLから code= の後の文字列だけを抽出して、直接コードで認証します。
                    # これにより (mismatching_state) エラーを回避できます。
                    if "code=" in auth_response:
                        auth_code = auth_response.split("code=")[1].split("&")[0]
                    else:
                        auth_code = auth_response
                    
                    # authorization_response ではなく code を使うのがポイントです
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    st.session_state.google_creds = creds
                    st.success("認証に成功しました！")
                    st.rerun()
                except Exception as e:
                    st.error(f"認証に失敗しました。もう一度リンクからやり直してください。: {e}")
    return creds

# --- メイン画面 ---
creds = authenticate_google()

uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 スライド作成を開始"):
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
                
                # 2. PDFページを画像に変換
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # 3. 画像をGoogleドライブに一時保存
                file_metadata = {'name': f'temp_img_{i}.png', 'parents': ['root']}
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                file_id = file.get('id')
                
                # 4. Slides APIからアクセスできるように権限を一時公開
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}"

                # 5. スライドの追加と画像の挿入
                page_id = f"page_{i}"
                requests = [
                    {'createSlide': {'objectId': page_id}},
                    {'createImage': {
                        'elementProperties': {'pageObjectId': page_id},
                        'url': file_url
                    }}
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                
                # 6. 一時ファイルを削除
                drive_service.files().delete(fileId=file_id).execute()
                
                progress_bar.progress((i + 1) / total_pages)

            st.balloons()
            st.success("✅ スライドが完成しました！")
            st.markdown(f"### [作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

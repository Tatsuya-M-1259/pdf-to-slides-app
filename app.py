import streamlit as st
import fitz  # PyMuPDF
import io
import os
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# セキュリティチェックを緩和（localhost対策）
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# APIの権限範囲
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.title("📄 PDFをGoogleスライドに変換 (画像貼り付け)")
st.caption("PDFの各ページを高画質な画像として、新しいGoogleスライドに1枚ずつ貼り付けます。")

def authenticate_google():
    creds = None
    if 'google_creds' in st.session_state:
        creds = st.session_state.google_creds

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.google_creds = creds
                return creds
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
            
            # Flowをセッションに保持
            if 'auth_flow' not in st.session_state:
                st.session_state.auth_flow = Flow.from_client_config(
                    client_config, 
                    scopes=SCOPES,
                    redirect_uri='http://localhost'
                )
            
            flow = st.session_state.auth_flow
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            
            st.info("💡 Google認証が必要です。")
            st.markdown(f"**手順1:** [👉 ここをクリックしてGoogle認証を開く]({auth_url})")
            st.write("**手順2:** 認証後、ブラウザがエラーになります。その時の**URL欄（アドレスバー）の内容をすべてコピー**してください。")
            
            auth_response = st.text_input("**手順3:** コピーしたURLをここに貼り付けてEnter:", key="auth_input_final")
            
            if auth_response:
                try:
                    # 【ここが解決の鍵】URLから code= 以降だけを抜き出し、stateチェックをバイパスします
                    if "code=" in auth_response:
                        auth_code = auth_response.split("code=")[1].split("&")[0]
                    else:
                        auth_code = auth_response
                    
                    # fetch_token(code=...) を使うことで CSRF Warning を回避
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    st.session_state.google_creds = creds
                    st.success("認証に成功しました！🎉")
                    st.rerun()
                except Exception as e:
                    st.error(f"認証に失敗しました。もう一度リンクからやり直してください。: {e}")
                    if st.button("認証を最初からやり直す"):
                        if 'auth_flow' in st.session_state:
                            del st.session_state.auth_flow
                        st.rerun()
    return creds

# --- メイン画面 ---
creds = authenticate_google()

uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 スライド作成を開始"):
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        try:
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, page in enumerate(doc):
                status_text.text(f"処理中: {i+1} / {total_pages} ページ目")
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # ドライブに一時保存
                file_metadata = {'name': f'temp_{i}.png', 'parents': ['root']}
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                file_id = file.get('id')
                
                # 一時的に誰でも閲覧可能にしてSlides APIに渡す
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}"

                page_id = f"page_obj_{i}"
                requests = [
                    {'createSlide': {'objectId': page_id}},
                    {'createImage': {
                        'elementProperties': {'pageObjectId': page_id},
                        'url': file_url
                    }}
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                drive_service.files().delete(fileId=file_id).execute()
                progress_bar.progress((i + 1) / total_pages)

            st.balloons()
            st.success("✅ スライドが完成しました！")
            st.markdown(f"### [作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

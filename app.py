import streamlit as st
import fitz  # PyMuPDF
import io
import os
import time
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# セキュリティ設定の緩和
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# Google API スコープ
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

# Googleスライドの標準サイズ (16:9)
SLIDE_W = 720
SLIDE_H = 405

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #fcfcfc; }
    h1 { color: #444; border-bottom: 2px solid #444; }
    </style>
    """, unsafe_allow_html=True)

st.title("📄 PDF to Google Slides (安定版 ver 8.1)")
st.info("比率を維持したまま中央に配置する安定バージョンです。エラーを修正しました。")

# --- 認証処理 ---
def authenticate_google():
    creds = None
    if 'google_creds' in st.session_state:
        creds = st.session_state.google_creds

    if "code" in st.query_params and not creds:
        try:
            flow = Flow.from_client_config(
                {"web": {
                    "client_id": st.secrets["google_oauth"]["client_id"],
                    "project_id": st.secrets["google_oauth"]["project_id"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_secret": st.secrets["google_oauth"]["client_secret"],
                    "redirect_uris": [st.secrets["google_oauth"]["redirect_uri"]]
                }},
                scopes=SCOPES,
                redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
            )
            flow.fetch_token(code=st.query_params["code"])
            creds = flow.credentials
            st.session_state.google_creds = creds
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"認証エラー: {e}")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.google_creds = creds
            except: creds = None
        
        if not creds:
            flow = Flow.from_client_config(
                {"web": {
                    "client_id": st.secrets["google_oauth"]["client_id"],
                    "project_id": st.secrets["google_oauth"]["project_id"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "client_secret": st.secrets["google_oauth"]["client_secret"],
                    "redirect_uris": [st.secrets["google_oauth"]["redirect_uri"]]
                }},
                scopes=SCOPES,
                redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
            )
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            st.link_button("🔑 Googleアカウントでログイン", auth_url)
            st.stop()
    return creds

# --- メイン処理 ---
creds = authenticate_google()
uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 スライドを作成する (中央配置)"):
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        try:
            # 1. 新規プレゼンテーション作成
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            first_slide_id = presentation.get('slides')[0].get('objectId')
            
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            progress_bar = st.progress(0)

            for i, page in enumerate(doc):
                # 中央揃え・比率維持の計算
                pdf_w = page.rect.width
                pdf_h = page.rect.height
                scale = min(SLIDE_W / pdf_w, SLIDE_H / pdf_h)
                new_w = pdf_w * scale
                new_h = pdf_h * scale
                pos_x = (SLIDE_W - new_w) / 2
                pos_y = (SLIDE_H - new_h) / 2

                # 高画質画像化
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
                img_data = pix.tobytes("png")
                
                # Googleドライブへアップ
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body={'name': f'slide_{int(time.time())}_{i}.png'}, media_body=media, fields='id').execute()
                file_id = file.get('id')
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}"

                # スライド作成
                page_id = f"slide_{i}_{int(time.time())}"
                requests = [
                    {'createSlide': {'objectId': page_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
                    {
                        'createImage': {
                            'elementProperties': {
                                'pageObjectId': page_id,
                                'size': {
                                    'width': {'magnitude': new_w, 'unit': 'PT'},
                                    'height': {'magnitude': new_h, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1, 'scaleY': 1,
                                    'translateX': pos_x, 'translateY': pos_y, 'unit': 'PT'
                                }
                            },
                            'url': file_url
                        }
                    }
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                
                # 【修正箇所】drive_service を正しく使用
                drive_service.files().delete(fileId=file_id).execute()
                progress_bar.progress((i + 1) / total_pages)

            # 最初の空白スライド削除
            slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': [{'deleteObject': {'objectId': first_slide_id}}]}).execute()
            
            st.balloons()
            st.success("✅ 作成完了しました！")
            st.markdown(f"### [👉 スライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

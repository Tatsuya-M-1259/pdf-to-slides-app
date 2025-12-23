import streamlit as st
import fitz  # PyMuPDF
import io
import os
import time
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# セキュリティ緩和
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive.file']

# スライド標準サイズ
SLIDE_W = 720
SLIDE_H = 405

# 画面デザインの変更（確実に更新されたことを示すため、オレンジのヘッダーにします）
st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #fffaf0; }
    h1 { color: #ff4b4b; border-bottom: 3px solid #ff4b4b; }
    </style>
    """, unsafe_allow_html=True)

st.title("📄 PDF to Google Slides (完ぺき版 ver 7.0)")
st.warning("⚠️ PDFの『余白』を自動で切り落として、中身を最大化して貼り付けます。")

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
                scopes=SCOPES, redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
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
            try: creds.refresh(Request())
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
                scopes=SCOPES, redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
            )
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            st.link_button("🔑 Googleアカウントでログイン", auth_url)
            st.stop()
    return creds

creds = authenticate_google()
uploaded_file = st.file_uploader("PDFファイルをアップロード", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 余白をカットしてスライド作成"):
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)
        try:
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            first_slide_id = presentation.get('slides')[0].get('objectId')
            
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            progress_bar = st.progress(0)

            for i, page in enumerate(doc):
                # 【新機能】内容がある領域を自動検知してトリミング
                content_box = page.get_bbox() # 描画がある最小の枠を取得
                # 少しだけマージン（5%）を残してカット
                padding = 20
                crop_rect = fitz.Rect(content_box[0]-padding, content_box[1]-padding, content_box[2]+padding, content_box[3]+padding)
                
                # トリミングした領域を高画質で画像化
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=crop_rect)
                img_data = pix.tobytes("png")
                
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body={'name': f'trim_{int(time.time())}_{i}.png'}, media_body=media, fields='id').execute()
                file_id = file.get('id')
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}&t={time.time()}"

                page_id = f"slide_{int(time.time())}_{i}"
                requests = [
                    {'createSlide': {'objectId': page_id, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}},
                    {
                        'updatePageProperties': {
                            'objectId': page_id,
                            'pageProperties': {
                                'pageBackgroundFill': {
                                    'stretchedPictureFill': {'contentUrl': file_url}
                                }
                            },
                            'fields': 'pageBackgroundFill'
                        }
                    }
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                drive_service.files().delete(fileId=file_id).execute()
                progress_bar.progress((i + 1) / total_pages)

            slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': [{'deleteObject': {'objectId': first_slide_id}}]}).execute()
            
            st.balloons()
            st.success("✅ 余白をカットし、中身を最大化したスライドが完成しました！")
            st.markdown(f"### [👉 作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

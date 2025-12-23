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

# スライド標準サイズ (16:9)
SLIDE_W = 720
SLIDE_H = 405

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
# 確実に更新されたことを示すため、今回は「薄いミントグリーン」にします
st.markdown("""
    <style>
    .stApp { background-color: #f5fffa; }
    h1 { color: #2e8b57; border-bottom: 2px solid #2e8b57; }
    </style>
    """, unsafe_allow_html=True)

st.title("📄 PDF to Google Slides (完全解決版 ver 7.2)")
st.caption("エラーを完全に回避するロジックに変更しました。余白をカットして中身を最大化します。")

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
    if st.button("🚀 内容を最大化してスライド作成"):
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
                # --- 【最重要】バージョン依存しない内容検知ロジック ---
                blocks = page.get_text("blocks")
                if blocks:
                    # 描画がある最小・最大の座標を自前で計算
                    x0 = min(b[0] for b in blocks)
                    y0 = min(b[1] for b in blocks)
                    x1 = max(b[2] for b in blocks)
                    y1 = max(b[3] for b in blocks)
                    
                    padding = 10
                    crop_rect = fitz.Rect(
                        max(0, x0 - padding),
                        max(0, y0 - padding),
                        min(page.rect.width, x1 + padding),
                        min(page.rect.height, y1 + padding)
                    )
                else:
                    crop_rect = page.rect
                # --------------------------------------------------
                
                # トリミングした領域を高画質で画像化
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=crop_rect)
                img_data = pix.tobytes("png")
                
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body={'name': f'fs_{int(time.time())}_{i}.png'}, media_body=media, fields='id').execute()
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
            st.success("✅ 全面表示のスライドが完成しました！")
            st.markdown(f"### [👉 作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

import streamlit as st
import fitz  # PyMuPDF
import io
import os
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# セキュリティチェックを緩和（localhost対策ですがWebアプリ形式でも念のため保持）
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

# 1. APIの権限範囲の設定
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.title("📄 PDFをGoogleスライドに変換 (自動認証版)")
st.caption("ログインボタンを押すだけで認証が完了し、PDFをスライド化できます。")

# --- 画像位置を中央にリセットする関数 ---
def reset_images_position(presentation_id, creds):
    slides_service = build('slides', 'v1', credentials=creds)
    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    slides = presentation.get('slides', [])
    requests = []
    SLIDE_W, SLIDE_H = 720, 405 # 16:9の標準サイズ

    for slide in slides:
        elements = slide.get('pageElements', [])
        for element in elements:
            if 'image' in element:
                obj_id = element['objectId']
                img_w = element['size']['width']['magnitude']
                img_h = element['size']['height']['magnitude']
                requests.append({
                    'updatePageElementTransform': {
                        'objectId': obj_id,
                        'applyMode': 'ABSOLUTE',
                        'transform': {
                            'scaleX': 1, 'scaleY': 1,
                            'translateX': (SLIDE_W - img_w) / 2,
                            'translateY': (SLIDE_H - img_h) / 2,
                            'unit': 'PT'
                        }
                    }
                })
    if requests:
        slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
        return True
    return False

# --- 認証処理（自動取得版） ---
def authenticate_google():
    creds = None
    # セッション内に認証情報があるか確認
    if 'google_creds' in st.session_state:
        creds = st.session_state.google_creds

    # 1. Googleから戻ってきた際のURLパラメータ「code」を自動取得
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
            st.query_params.clear() # URLを綺麗にする
            st.rerun()
        except Exception as e:
            st.error(f"認証コードの取得に失敗しました: {e}")

    # 2. 認証が必要な場合の表示
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.google_creds = creds
            except:
                creds = None
        
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
            
            st.info("💡 PDFをスライドに変換するにはGoogleログインが必要です。")
            st.link_button("🔑 Googleアカウントでログイン", auth_url)
            st.stop() # ログインするまで以降の処理を止める
            
    return creds

# --- メイン画面 ---
creds = authenticate_google()

uploaded_file = st.file_uploader("PDFファイルをアップロードしてください", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 スライド作成を開始"):
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        try:
            # 新規スライド作成
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            first_slide_id = presentation.get('slides')[0].get('objectId') # 最初の空白ページ
            
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            progress_bar = st.progress(0)

            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body={'name': f'temp_{i}.png'}, media_body=media, fields='id').execute()
                file_id = file.get('id')
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}"

                page_id = f"slide_{i}"
                requests = [
                    {'createSlide': {'objectId': page_id}},
                    {'createImage': {
                        'elementProperties': {
                            'pageObjectId': page_id,
                            'size': {'height': {'magnitude': 350, 'unit': 'PT'}, 'width': {'magnitude': 600, 'unit': 'PT'}},
                            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 60, 'translateY': 27, 'unit': 'PT'}
                        },
                        'url': file_url
                    }}
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                drive_service.files().delete(fileId=file_id).execute()
                progress_bar.progress((i + 1) / total_pages)

            # 最初の空白ページを削除
            slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': [{'deleteObject': {'objectId': first_slide_id}}]}).execute()
            
            st.session_state.last_presentation_id = presentation_id
            st.balloons()
            st.success("✅ スライドが完成しました！")
            st.markdown(f"### [👉 作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

    # 中央リセット機能
    if 'last_presentation_id' in st.session_state:
        st.divider()
        if st.button("🖼️ 画像の位置を中央にリセットする"):
            if reset_images_position(st.session_state.last_presentation_id, creds):
                st.toast("全スライドの画像を中央に配置しました！")

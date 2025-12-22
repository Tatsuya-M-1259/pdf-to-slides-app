import streamlit as st
import fitz  # PyMuPDF
import io
import os
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request

# http通信を許可する設定
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive.file'
]

st.set_page_config(page_title="PDF to Google Slides", layout="wide")
st.title("📄 PDFをGoogleスライドに変換 (改良版)")

# --- 認証処理 ---
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
            
            if 'auth_flow' not in st.session_state:
                st.session_state.auth_flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri='http://localhost')
            
            flow = st.session_state.auth_flow
            auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
            
            st.info("💡 Google認証が必要です。")
            st.markdown(f"**手順1:** [👉 認証を開始する]({auth_url})")
            auth_response = st.text_input("**手順2:** エラー画面のURLをここに貼り付けてEnter:", key="auth_input")
            
            if auth_response:
                try:
                    if "code=" in auth_response:
                        auth_code = auth_response.split("code=")[1].split("&")[0]
                    else:
                        auth_code = auth_response
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    st.session_state.google_creds = creds
                    st.rerun()
                except Exception as e:
                    st.error(f"認証失敗: {e}")
    return creds

# --- 画像位置を中央にリセットする関数 ---
def reset_images_position(presentation_id, creds):
    slides_service = build('slides', 'v1', credentials=creds)
    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    slides = presentation.get('slides', [])
    
    requests = []
    # Googleスライドの標準サイズ (16:9) は 720pt x 405pt
    SLIDE_W = 720
    SLIDE_H = 405

    for slide in slides:
        elements = slide.get('pageElements', [])
        for element in elements:
            if 'image' in element:
                obj_id = element['objectId']
                # 中央配置のための計算
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

# --- メイン処理 ---
creds = authenticate_google()
uploaded_file = st.file_uploader("PDFをアップロード", type="pdf")

if uploaded_file and creds:
    if st.button("🚀 スライド作成を開始"):
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        try:
            # 1. 新規スライド作成
            presentation = slides_service.presentations().create(body={'title': uploaded_file.name}).execute()
            presentation_id = presentation.get('presentationId')
            # 最初の空白スライドのIDを記憶しておく
            first_slide_id = presentation.get('slides')[0].get('objectId')
            
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            total_pages = len(doc)
            progress_bar = st.progress(0)

            for i, page in enumerate(doc):
                # 2. PDFを画像化
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # 3. ドライブに保存
                media = MediaIoBaseUpload(io.BytesIO(img_data), mimetype='image/png')
                file = drive_service.files().create(body={'name': f't_{i}.png'}, media_body=media, fields='id').execute()
                file_id = file.get('id')
                drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
                file_url = f"https://drive.google.com/uc?id={file_id}"

                # 4. スライド追加と画像の中央配置
                # 画像サイズをスライドの高さ(405pt)に合わせる計算
                requests = [
                    {'createSlide': {'objectId': f'pg_{i}'}},
                    {'createImage': {
                        'elementProperties': {
                            'pageObjectId': f'pg_{i}',
                            'size': {'height': {'magnitude': 350, 'unit': 'PT'}, 'width': {'magnitude': 600, 'unit': 'PT'}},
                            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 60, 'translateY': 27, 'unit': 'PT'}
                        },
                        'url': file_url
                    }}
                ]
                slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
                drive_service.files().delete(fileId=file_id).execute()
                progress_bar.progress((i + 1) / total_pages)

            # 5. 最後に最初の空白スライドを削除
            slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': [{'deleteObject': {'objectId': first_slide_id}}]}).execute()
            
            st.session_state.last_presentation_id = presentation_id
            st.balloons()
            st.success("✅ 完成しました！")
            st.markdown(f"### [👉 作成されたスライドを開く](https://docs.google.com/presentation/d/{presentation_id})")

        except Exception as e:
            st.error(f"エラー: {e}")

    # 位置リセットボタン（作成完了後に表示）
    if 'last_presentation_id' in st.session_state:
        st.divider()
        st.subheader("🛠️ スライドの微調整")
        if st.button("🖼️ 全ての画像の位置を中央にリセットする"):
            if reset_images_position(st.session_state.last_presentation_id, creds):
                st.toast("画像の位置を中央に戻しました！")
            else:
                st.warning("リセット対象の画像が見つかりませんでした。")

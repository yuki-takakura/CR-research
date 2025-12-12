import streamlit as st
import os
import cv2
import xlsxwriter
import tempfile
import easyocr
import ssl
import datetime
from PIL import Image

# --- セキュリティ設定（Mac用） ---
ssl._create_default_https_context = ssl._create_unverified_context
# -----------------------------

from scenedetect import detect, ContentDetector

st.set_page_config(page_title="動画分析DBツール", layout="wide")
st.title("📊 動画分析データベース作成ツール（軽量・埋め込み版）")
st.write("画像を物理的にリサイズしてセルに密着させ、データの蓄積・コピーに最適化します。")

# --- AIモデル設定 ---
@st.cache_resource
def load_model():
    return easyocr.Reader(['ja', 'en'], gpu=False)

uploaded_file = st.file_uploader("分析する動画ファイルをアップロード", type=["mp4", "mov"])

if uploaded_file:
    tfile = tempfile.NamedTemporaryFile(delete=False)
    tfile.write(uploaded_file.read())
    video_path = tfile.name
    original_filename = uploaded_file.name

    if st.button("分析レポートを作成する"):
        status_box = st.empty()
        bar = st.progress(0)
        
        status_box.text("🚀 AIモデルをロード中...")
        reader = load_model()

        status_box.text("🎬 シーン検出中...")
        scene_list = detect(video_path, ContentDetector())
        
        status_box.text(f"✅ {len(scene_list)} シーン検出。Excel生成開始...")

        # Excel準備
        wb = xlsxwriter.Workbook("creative_db_lite.xlsx")
        ws = wb.add_worksheet("Database")
        
        # --- 書式設定 ---
        font_name = 'Meiryo UI'
        
        # 見出し
        fmt_header = wb.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E79', 
            'align': 'center', 'valign': 'vcenter', 'border': 1, 
            'font_name': font_name, 'font_size': 11
        })
        # データセル
        fmt_center = wb.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1, 
            'font_name': font_name, 'font_size': 10
        })
        fmt_text = wb.add_format({
            'text_wrap': True, 'valign': 'top', 'align': 'left',
            'border': 1, 'font_name': font_name, 'font_size': 10
        })
        fmt_gray = wb.add_format({
            'text_wrap': True, 'valign': 'top', 'font_color': '#555555',
            'border': 1, 'font_name': font_name, 'font_size': 9
        })
        fmt_yellow = wb.add_format({
            'text_wrap': True, 'valign': 'top', 'bg_color': '#FFFFCC', 
            'border': 1, 'font_name': font_name, 'font_size': 10
        })

        # --- ヘッダー作成（A列に項目名） ---
        START_ROW = 0
        
        # メタ情報
        today = datetime.datetime.now().strftime('%Y/%m/%d')
        ws.write(0, 0, "分析日", fmt_header)
        ws.write(0, 1, today, fmt_center)
        ws.write(1, 0, "ファイル名", fmt_header)
        ws.write(1, 1, original_filename, fmt_text)
        
        # 項目見出し（3行目からデータ開始）
        START_DATA_ROW = 3
        headers = ["キャプチャ", "秒数", "抽出テキスト", "注釈", "コメント"]
        
        # A列に見出しを配置
        ws.set_column('A:A', 20)
        for i, h in enumerate(headers):
            ws.write(START_DATA_ROW + i, 0, h, fmt_header)

        # --- 画像設定（物理リサイズ用） ---
        TARGET_HEIGHT = 160  # 目標とする画像の高さ（ピクセル）
        PADDING = 10         # セル内の余白
        
        # 行の高さを設定（画像高さ + 余白）
        # Excelの行高さはポイント単位 (1 px = 0.75 point)
        ROW_HEIGHT_PT = (TARGET_HEIGHT + PADDING * 2) * 0.75
        
        ws.set_row(START_DATA_ROW, ROW_HEIGHT_PT)     # キャプチャ行
        ws.set_row(START_DATA_ROW + 1, 25)            # 秒数行
        ws.set_row(START_DATA_ROW + 2, 100)           # テキスト行
        ws.set_row(START_DATA_ROW + 3, 50)            # 注釈行
        ws.set_row(START_DATA_ROW + 4, 60)            # コメント行

        cap = cv2.VideoCapture(video_path)
        if not os.path.exists('images'): os.makedirs('images')

        # --- ループ処理 ---
        for i, scene in enumerate(scene_list):
            status_box.text(f"📸 処理中: シーン {i+1} / {len(scene_list)}")
            col = i + 1
            
            # 時間取得
            start = scene[0].get_seconds()
            end = scene[1].get_seconds()
            mid = (start + end) / 2
            
            cap.set(cv2.CAP_PROP_POS_MSEC, mid * 1000)
            ret, frame = cap.read()
            
            if ret:
                # 1. OpenCV(BGR) -> Pillow(RGB)変換
                img_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                pil_img = Image.fromarray(img_rgb)
                
                # 2. 画像を物理的にリサイズ（軽量化）
                # アスペクト比を維持して高さをTARGET_HEIGHTに合わせる
                aspect_ratio = pil_img.width / pil_img.height
                new_width = int(TARGET_HEIGHT * aspect_ratio)
                pil_img_resized = pil_img.resize((new_width, TARGET_HEIGHT), Image.LANCZOS)
                
                # 3. リサイズした画像を保存
                img_path = f"images/scene_{i}.jpg"
                pil_img_resized.save(img_path, quality=85)
                
                # 4. 列幅を画像幅に合わせて調整
                # Excelの列幅は文字数換算 (概算: pixels / 7 + 余白)
                col_width = (new_width + PADDING * 2) / 7.0
                ws.set_column(col, col, col_width)
                
                # 5. AI文字認識（元の高画質フレームを使用すると重いので、リサイズ前を使うか検討だが、ここではリサイズ前を使う）
                main_texts = []
                note_texts = []
                try:
                    results = reader.readtext(frame, detail=1) # AIには元の高画質を渡す
                    frame_h = frame.shape[0]
                    for (bbox, text, prob) in results:
                        if prob < 0.3: continue
                        box_h = bbox[2][1] - bbox[1][1]
                        ratio = box_h / frame_h
                        if ratio > 0.035: main_texts.append(text)
                        elif ratio > 0.012: note_texts.append(text)
                except: pass

                str_main = "\n".join(main_texts) if main_texts else ""
                str_note = "\n".join(note_texts) if note_texts else ""

                # --- Excel書き込み ---
                # 画像の貼り付け（物理リサイズ済みなので scale=1 でOK）
                ws.insert_image(START_DATA_ROW, col, img_path, 
                                {'x_offset': PADDING, 'y_offset': PADDING, 
                                 'object_position': 1}) # 1 = Move and size with cells
                
                ws.write(START_DATA_ROW + 1, col, f"{start:.1f}s - {end:.1f}s", fmt_center)
                ws.write(START_DATA_ROW + 2, col, str_main, fmt_text)
                ws.write(START_DATA_ROW + 3, col, str_note, fmt_gray)
                ws.write(START_DATA_ROW + 4, col, "", fmt_yellow)

            bar.progress((i + 1) / len(scene_list))

        wb.close()
        cap.release()
        
        status_box.text("✨ 完了！軽量化＆埋め込み完了しました。")
        st.success("分析完了！")
        with open("creative_db_lite.xlsx", "rb") as f:
            st.download_button("Excelレポートをダウンロード", f, "creative_db_lite.xlsx")

import streamlit as st
import os
import cv2
import xlsxwriter
import tempfile
import easyocr
import ssl
import datetime
import shutil
import numpy as np
from PIL import Image
from difflib import SequenceMatcher

# --- セキュリティ設定 ---
ssl._create_default_https_context = ssl._create_unverified_context
# ---------------------

from scenedetect import detect, ContentDetector

st.set_page_config(page_title="動画CR解析ツール", page_icon="🎞️", layout="wide")

# --- デザインCSS ---
st.markdown("""
    <style>
    html, body, [class*="css"] {font-family: 'Helvetica Neue', 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif;}
    header, footer {visibility: hidden;}
    .block-container {padding-top: 2rem; padding-bottom: 5rem;}
    h1 {color: #1F4E79; border-bottom: 2px solid #1F4E79; padding-bottom: 10px;}
    .stFileUploader {border: 2px dashed #A0A0A0; border-radius: 12px; padding: 30px; background-color: #F8F9FA;}
    .stButton>button {
        color: white; background: linear-gradient(135deg, #1F4E79 0%, #007BFF 100%);
        border: none; border-radius: 8px; height: 3.5em; width: 100%; font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)

st.title("動画CR解析ツール")
st.markdown("##### 視覚情報(キャプチャ・テキスト)を高画質かつ精密に抽出・解析します")

# --- AIモデル設定 ---
@st.cache_resource
def load_ocr_model():
    return easyocr.Reader(['ja', 'en'], gpu=False)

# --- 類似度判定 ---
def is_text_different(text1, text2, threshold=0.8):
    if not text1 and not text2: return False
    if not text1 or not text2: return True
    
    ratio = SequenceMatcher(None, text1, text2).ratio()
    len_diff = abs(len(text1) - len(text2))
    if len_diff > 5: return True
        
    return ratio < threshold

uploaded_files = st.file_uploader("分析する動画（複数可）", type=["mp4", "mov"], accept_multiple_files=True)

if uploaded_files:
    with st.sidebar:
        st.write("### ⚙️ 解析設定")
        st.info("ナチュラル高画質モード")
        scan_interval = st.slider("カット内監視間隔（秒）", 0.5, 2.0, 1.0, 0.5)

    if st.button("🚀 解析スタート"):
        status_box = st.empty()
        total_bar = st.progress(0)
        
        status_box.info("🤖 AIエンジンをロード中...")
        ocr_reader = load_ocr_model()

        # Excel準備
        output_filename = f"CR_Analysis_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb = xlsxwriter.Workbook(output_filename)
        ws = wb.add_worksheet("Analysis")
        
        # --- 書式設定 ---
        font_name = 'Meiryo UI'
        border_style = 1
        
        fmt_header = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1F4E79', 'align': 'center', 'valign': 'vcenter', 'border': border_style, 'font_name': font_name, 'font_size': 11})
        fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': border_style, 'font_name': font_name, 'font_size': 10})
        fmt_text = wb.add_format({'text_wrap': True, 'valign': 'top', 'align': 'left', 'border': border_style, 'font_name': font_name, 'font_size': 10})
        fmt_gray = wb.add_format({'text_wrap': True, 'valign': 'top', 'font_color': '#555555', 'border': border_style, 'font_name': font_name, 'font_size': 9})
        fmt_title = wb.add_format({'bold': True, 'bg_color': '#E2E2E2', 'font_size': 12, 'font_name': font_name, 'border': border_style, 'valign': 'vcenter'})
        fmt_meta_value = wb.add_format({'font_size': 10, 'font_name': font_name, 'border': border_style, 'valign': 'vcenter'})

        if os.path.exists('images'): shutil.rmtree('images')
        os.makedirs('images')

        CURRENT_ROW = 0
        
        for file_idx, uploaded_file in enumerate(uploaded_files):
            status_box.info(f"▶️ File {file_idx + 1}/{len(uploaded_files)}: {uploaded_file.name} を解析中...")
            
            tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
            tfile.write(uploaded_file.read())
            video_path = tfile.name

            status_box.info(f"🎬 シーン構造を解析中... ({uploaded_file.name})")
            scene_list = detect(video_path, ContentDetector())
            
            # --- レイアウト定義 ---
            TITLE_ROW = CURRENT_ROW
            META_ROW = CURRENT_ROW + 1
            DATA_START_ROW = CURRENT_ROW + 2
            
            # A列見出し
            ws.set_column('A:A', 25)
            ws.write(META_ROW, 0, "分析情報", fmt_header)
            
            headers = ["キャプチャ", "時間", "抽出テキスト", "注釈", "コメント"]
            for i, h in enumerate(headers):
                ws.write(DATA_START_ROW + i, 0, h, fmt_header)
            
            # ★ サイズ調整（ここを変更）
            # 余計なフィルターを排除し、単純に解像度を高く保存する
            IMAGE_SAVE_HEIGHT = 480 # 保存サイズ（かなり大きめ）
            DISPLAY_HEIGHT = 220    # Excel上の表示サイズ（160→220へ大型化）
            
            PADDING = 10
            ws.set_row(META_ROW, 25)
            ws.set_row(DATA_START_ROW, (DISPLAY_HEIGHT + PADDING * 2) * 0.75) 
            ws.set_row(DATA_START_ROW + 1, 25)  
            ws.set_row(DATA_START_ROW + 2, 120) 
            ws.set_row(DATA_START_ROW + 3, 50)  
            ws.set_row(DATA_START_ROW + 4, 50)  
            
            cap = cv2.VideoCapture(video_path)
            col = 1 
            
            for i, scene in enumerate(scene_list):
                scene_start = scene[0].get_seconds()
                scene_end = scene[1].get_seconds()
                duration = scene_end - scene_start

                if duration <= scan_interval:
                    check_points = [(scene_start + scene_end) / 2]
                else:
                    check_points = np.arange(scene_start, scene_end, scan_interval).tolist()
                    if check_points[-1] > scene_end - 0.2:
                        check_points.pop()

                prev_text_content = ""

                for pt in check_points:
                    cap.set(cv2.CAP_PROP_POS_MSEC, pt * 1000)
                    ret, frame = cap.read()
                    if not ret: continue

                    # OCR
                    try:
                        results = ocr_reader.readtext(frame, detail=1, mag_ratio=2.0, text_threshold=0.3, low_text=0.3)
                        results.sort(key=lambda x: x[0][0][1])
                        
                        main_texts = []
                        note_texts = []
                        frame_h = frame.shape[0]
                        for (bbox, text, prob) in results:
                            if prob < 0.05: continue
                            box_h = bbox[2][1] - bbox[1][1]
                            ratio = box_h / frame_h
                            if ratio > 0.030: main_texts.append(text)
                            elif ratio > 0.005: note_texts.append(text)
                    except: 
                        main_texts = []
                        note_texts = []

                    ocr_text = "\n".join(main_texts)
                    note_text = "\n".join(note_texts)
                    
                    current_text_content = (ocr_text + note_text).replace('\n', '').replace(' ', '')
                    
                    is_first_in_scene = (pt == check_points[0])
                    if not is_first_in_scene:
                        if not is_text_different(prev_text_content, current_text_content, threshold=0.8):
                            continue 

                    # ★ 変更点：フィルターを全廃止し、自然なリサイズのみ実行
                    img_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    pil_img = Image.fromarray(img_rgb)
                    
                    # アスペクト比を維持して高画質リサイズ
                    aspect = pil_img.width / pil_img.height
                    save_width = int(IMAGE_SAVE_HEIGHT * aspect)
                    
                    # 補正なしのLANCZOSリサイズのみ（これが一番自然で綺麗です）
                    pil_img_resized = pil_img.resize((save_width, IMAGE_SAVE_HEIGHT), Image.LANCZOS)
                    
                    img_filename = f"img_f{file_idx}_s{i}_{int(pt*100)}.jpg"
                    img_path = os.path.join("images", img_filename)
                    
                    # 保存品質をMAXに
                    pil_img_resized.save(img_path, quality=95, subsampling=0)
                    
                    # Excel列幅の調整
                    display_width = (save_width * (DISPLAY_HEIGHT / IMAGE_SAVE_HEIGHT))
                    col_width = (display_width + PADDING * 2) / 7.0
                    ws.set_column(col, col, col_width)

                    # 書き込み（表示サイズを大きく設定）
                    scale_ratio = DISPLAY_HEIGHT / IMAGE_SAVE_HEIGHT
                    
                    ws.insert_image(DATA_START_ROW, col, img_path, 
                                    {'x_offset': PADDING, 'y_offset': PADDING, 
                                     'x_scale': scale_ratio, 'y_scale': scale_ratio, 
                                     'object_position': 1})
                    
                    ws.write(DATA_START_ROW + 1, col, f"{pt:.1f}s", fmt_center)
                    ws.write(DATA_START_ROW + 2, col, ocr_text, fmt_text)
                    ws.write(DATA_START_ROW + 3, col, note_text, fmt_gray)
                    ws.write(DATA_START_ROW + 4, col, "", fmt_text)

                    ws.write(META_ROW, col, "", fmt_meta_value)

                    prev_text_content = current_text_content
                    col += 1

            # --- 後処理 ---
            max_col = col - 1
            if max_col > 0:
                ws.merge_range(TITLE_ROW, 0, TITLE_ROW, max_col, f"■ 分析対象: {uploaded_file.name}", fmt_title)
                
                now = datetime.datetime.now()
                wk = ["月", "火", "水", "木", "金", "土", "日"][now.weekday()]
                date_str = f"{now.year}年{now.month}月{now.day}日（{wk}）{now.hour:02}:{now.minute:02}"
                
                if max_col >= 1: ws.write(META_ROW, 1, f"抽出日時: {date_str}", fmt_meta_value)
                if max_col >= 2: ws.write(META_ROW, 2, "担当者: [          ]", fmt_meta_value)
                if max_col >= 3: ws.write(META_ROW, 3, "案件名: [          ]", fmt_meta_value)
            
            cap.release()
            total_bar.progress((file_idx + 1) / len(uploaded_files))
            CURRENT_ROW += 8

        wb.close()
        
        status_box.success("✨ 解析完了！高画質レポートをダウンロードできます。")
        st.balloons()
        
        with open(output_filename, "rb") as f:
            st.download_button("📥 解析レポートをダウンロード", f, output_filename)

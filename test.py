from PIL import Image
from google.cloud import vision
import os
import shutil
import gspread
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import requests
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, filedialog

# Google Sheets設定
SHEET_ID = '1TKZDqycs5QPoUyIdlEtHplEZUmm-uMUJIo2DwgwkGgQ'
SHEET_NAME = '2024　11月'

# Google Cloud Vision APIのクライアントを作成
client = vision.ImageAnnotatorClient()

# 認証情報の設定
creds = Credentials.from_service_account_file(
    'C:/Users/Windows 10/Desktop/KanColle_Score/kancollescore-5438219c3c3b.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)

# Google Sheetsクライアント
gs_client = gspread.authorize(creds)
sheet = gs_client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# Google Photos API用スコープと認証
PHOTOS_SCOPES = ['https://www.googleapis.com/auth/photoslibrary.readonly']

def authenticate_google_photos():
    creds = None
    # トークンファイルが存在するか確認
    if os.path.exists('token_photos.pickle'):
        with open('token_photos.pickle', 'rb') as token:
            creds = pickle.load(token)
    # トークンが無効の場合は再認証
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', PHOTOS_SCOPES
            )
            creds = flow.run_local_server(port=0)
        # 認証情報を保存
        with open('token_photos.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

# Google Photos APIから画像を取得する関数
def fetch_photos(creds):
    headers = {
        'Authorization': f'Bearer {creds.token}'
    }
    response = requests.get(
        'https://photoslibrary.googleapis.com/v1/mediaItems',
        headers=headers
    )
    if response.status_code == 200:
        media_items = response.json().get('mediaItems', [])
        if media_items:
            print("Retrieved images from Google Photos:")
            for item in media_items:
                print(f"Image URL: {item['baseUrl']}")
            # 例として最初の画像を表示
            first_image_url = media_items[0]['baseUrl']
            print(f"Displaying first image: {first_image_url}")
        else:
            print("No images found in Google Photos.")
    else:
        print(f"Failed to retrieve images: {response.status_code}")

# ローカル画像ファイルを選択する関数
def select_image():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="画像ファイルを選択してください",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")]
    )
    return file_path

# ローカル画像の処理
def process_local_image(file_path):
    if file_path:
        print(f"Selected local image: {file_path}")
        # ここで必要に応じて画像の処理を行う
        image = Image.open(file_path)
        image.show()  # 例として画像を表示
    else:
        print("No file selected.")

if __name__ == "__main__":
    # 選択ダイアログを表示
    choice = messagebox.askquestion(
        "画像ソース選択", 
        "Google Photosから画像を選択しますか？\n[はい]を選ぶとGoogle Photosから画像を取得します。[いいえ]を選ぶとローカルから画像を選択します。"
    )
    
    if choice == 'yes':
        # Google Photos認証と画像取得
        photos_creds = authenticate_google_photos()
        fetch_photos(photos_creds)
    else:
        # ローカル画像の選択と処理
        image_path = select_image()
        if image_path:  # ファイルが選択された場合のみ処理
            process_local_image(image_path)
        else:
            print("画像が選択されませんでした。プログラムを終了します。")
            exit()

    # 画像から特定範囲を切り取って OCR を行う関数
def crop_and_ocr(image, top_left, bottom_right):
    cropped_image = image.crop((*top_left, *bottom_right))
    with open("temp_image.png", "wb") as temp_file:
        cropped_image.save(temp_file.name)
    with open(temp_file.name, "rb") as image_file:
        content = image_file.read()
        vision_image = vision.Image(content=content)
        response = client.text_detection(image=vision_image)
        texts = response.text_annotations
        return texts[0].description.splitlines() if texts else []

# 画像作成日時を取得し、0:00～3:00の場合は前日の日付として扱う
image_creation_time = datetime.fromtimestamp(os.path.getctime(image_path))
if 0 <= image_creation_time.hour < 3:
    creation_date = (image_creation_time - timedelta(days=1)).strftime("%m/%d").lstrip("0").replace("/0", "/")
else:
    creation_date = image_creation_time.strftime("%m/%d").lstrip("0").replace("/0", "/")
creation_hour = image_creation_time.hour
print(f"画像の作成日時: {creation_date}, 時間: {creation_hour}")

# シートのA列から対応する行を見つける
date_column = sheet.col_values(1)
try:
    target_row = date_column.index(creation_date) + 1
    print(f"一致する日付が見つかりました: {creation_date}（行: {target_row}）")
except ValueError:
    print("指定の日付に対応する行が見つかりません。")
    exit()

# 時間に応じてセルの範囲を決定
if 3 <= creation_hour <= 15:
    score_column = "B"
    rank_column = "E"
else:
    score_column = "H"
    rank_column = "K"

# スコアと順位のセルを定義
score_cell = f"{score_column}{target_row}"
rank_cell = f"{rank_column}{target_row}"

# 画像のサイズを取得
image = Image.open(image_path)
width, height = image.size
print(f"画像サイズ: {width}x{height}")

# 画像サイズに応じてOCR範囲を設定
if (width, height) == (1200, 720):
    filter1_top_left = (270, 230)
    filter1_bottom_right = (340, 675)
    filter2_top_left = (343, 230)
    filter2_bottom_right = (573, 675)
    filter3_top_left = (1075, 230)
    filter3_bottom_right = (1150, 675)
elif (width, height) == (2520, 1080):
    filter1_top_left = (670, 345)
    filter1_bottom_right = (870, 1002)
    filter2_top_left = (870, 345)
    filter2_bottom_right = (1220, 1002)
    filter3_top_left = (1970, 345)
    filter3_bottom_right = (2082, 1002)
else:
    print("対応していない画像サイズです。プログラムを終了します。")
    exit()

# OCRを実行
texts1 = crop_and_ocr(image, filter1_top_left, filter1_bottom_right)
texts2 = crop_and_ocr(image, filter2_top_left, filter2_bottom_right)
texts3 = crop_and_ocr(image, filter3_top_left, filter3_bottom_right)

# 配列に格納し、トランスポーズ
all_texts = [texts1, texts2, texts3]
transposed_texts = list(map(list, zip(*all_texts)))

# プレイヤー名が「KP」「kp」「Kp」「kP」のいずれかに一致する行を抽出
target_names = {"KP", "kp", "Kp", "kP"}
matching_texts = [entry for entry in transposed_texts if entry[1] in target_names]

# 結果を出力
print("Transposed Texts:")
print(transposed_texts)
print("\nPlayer Information Matching 'KP', 'kp', 'Kp', or 'kP':")
print(matching_texts)

# データチェック
if len(matching_texts) < 1 or any(len(item) == 0 for item in matching_texts[0]):
    print("エラー: matching_textsに必要なデータの項目が不足しています。")
    exit()

# Google Sheetsに書き込み
try:
    rank_value = matching_texts[0][0]
    player_name = matching_texts[0][1]
    score_value = matching_texts[0][2]

    print(f"Rank: {rank_value}, Player Name: {player_name}, Score: {score_value}")

    sheet.update(range_name=score_cell, values=[[score_value]])  # スコア
    sheet.update(range_name=rank_cell, values=[[rank_value]])    # 順位

    print("データがGoogle Sheetsに正常に書き込まれました。")

except IndexError as e:
    print("エラー: 必要なデータ項目が欠けています。")
    exit()

# オリジナルの画像を指定のディレクトリにコピー
output_dir = "C:/Users/Windows 10/Desktop/KanColle_Score/image/score"
os.makedirs(output_dir, exist_ok=True)
destination_path = os.path.join(output_dir, os.path.basename(image_path))
shutil.copy(image_path, destination_path)
print(f"Original image copied to {output_dir}")

# コピーが正常に行われたか確認し、元画像を削除
if os.path.exists(destination_path):
    os.remove(image_path)
    print(f"Original image at {image_path} has been deleted.")
else:
    print("Failed to copy the image. Original image not deleted.")


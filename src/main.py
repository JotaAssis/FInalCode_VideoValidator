import cv2
import pandas as pd
import os
from ultralytics import YOLO
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# YOLO
MODEL_PATH = "models/yolov8n.pt"
TARGET_CLASSES = {"cell phone", "radio", "smartphone"}

model = YOLO(MODEL_PATH)

# Analisar vídeo
def analyze_video(video_url):
    inicio = time.time()
    cap = cv2.VideoCapture(video_url)
    if not cap.isOpened():
        return "Erro"

    found = False
    frame_count = 0

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret or frame_count > 300:
            break

        results = model(frame)
        for r in results:
            for box in r.boxes:
                cls = model.names[int(box.cls[0])]
                if cls in TARGET_CLASSES:
                    found = True
                    break
            if found:
                break

        frame_count += 1

    cap.release()
    fim = time.time()
    print(f"Tempo de execução: {fim - inicio:.4f} segundos")
    return "Verdadeiro" if found else "Falso"

# Processar vídeos
def process_videos(input_file, output_file):
    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
        video_url = row["Evidência"]
        evidency = analyze_video(video_url)

        df.at[index, "Evidencia"] = evidency
        df.at[index, "Status"] = "Validado" if evidency in ["Verdadeiro", "Falso"] else "Erro"

    df.to_excel(output_file, index=False)
    print("Processo concluído! Resultados salvos em:", output_file)

if __name__ == "__main__":
    # Configurar a autenticação com o Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("c:/Projects/Yolov_Video_IA/src/project-segsat-yolo-d7578484bf24.json", scope)
    client = gspread.authorize(creds)

    # Abrir a planilha do Google Sheets
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1RFzS2RsmkMhsCmqiCzqdk_7IhUzYI6b4eB1f0SvKHDU/edit?gid=799806951#gid=799806951")
    worksheet = sheet.get_worksheet(0)

    try:
        # Baixar a planilha como um DataFrame do pandas
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)

        # Salvar o DataFrame em um arquivo Excel temporário
        input_excel = "data/temp_input.xlsx"
        df.to_excel(input_excel, index=False)

        output_excel = "data/resultado.xlsx"

        os.makedirs("models", exist_ok=True)
        os.makedirs("data", exist_ok=True)

        inicio = time.time()
        process_videos(input_excel, output_excel)
        fim = time.time()
        print(f"Tempo total de execução: {fim - inicio:.4f} segundos")

        # Carregar o arquivo Excel de saída e atualizar a planilha do Google Sheets
        result_df = pd.read_excel(output_excel)
        worksheet.update([result_df.columns.values.tolist()] + result_df.values.tolist())
        print("Planilha do Google Sheets atualizada com os resultados.")
    except Exception as e:
        print(f"Erro ao obter registros da planilha: {e}")
import requests
import pandas as pd
import os

from dotenv import load_dotenv

load_dotenv()

AZURE_SPEECH_KEY = os.environ.get('AZURE_SPEECH_KEY')
AZURE_REGION = os.environ.get('AZURE_REGION')  # Регион

LANGUAGE = "es-ES"  # Испанский язык
VOICE = "es-ES-ElviraNeural"  # Голос для синтеза речи
OUTPUT_DIR = ""  # Папка для сохранения файлов

PROMPTS_FILENAME = "название_файла_для_генерации.xlsx"


def get_access_token():
    token_url = f"https://{AZURE_REGION}.api.cognitive.microsoft.com/sts/v1.0/issuetoken"
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_SPEECH_KEY,
        "Content-Length": "0"
    }
    response = requests.post(token_url, headers=headers)
    if response.status_code == 200:
        return response.text
    else:
        raise Exception(f"Error getting token: {response.status_code} - {response.text}")


def synthesize_speech(prompt_name, text, output_filename):
    url = f"https://{AZURE_REGION}.tts.speech.microsoft.com/cognitiveservices/v1"
    access_token = get_access_token()
    print("Access token obtained.")
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/ssml+xml",
        "X-Microsoft-OutputFormat": "riff-24khz-16bit-mono-pcm"
    }

    ssml = f"""
    <speak version='1.0' xml:lang='{LANGUAGE}'>
        <voice name='{VOICE}'>{text}</voice>
    </speak>
    """

    response = requests.post(url, headers=headers, data=ssml)
    if response.status_code == 200:
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        with open(output_path, "wb") as audio_file:
            audio_file.write(response.content)
        print(f"File saved: {output_path}")
    else:
        print(f"Error synthesizing speech: {response.status_code} - {response.text}, prompt: {prompt_name}")


def main():
    table_path = PROMPTS_FILENAME
    df = pd.read_excel(table_path).astype(str)
    
    for index, row in df.iterrows():
        text = row["prompt_text"]
        filename = row["prompt_filename"]
        if not filename or filename == 'nan':
            continue
        if not filename.endswith(".wav"):
            filename += ".wav"

    try:
        synthesize_speech(filename, text, filename)
    except Exception as e:
        print(f"Error processing {filename}: {e}")

if __name__ == "__main__":
    main()

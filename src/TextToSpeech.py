import requests
import json

class TextToSpeech:
    def __init__(self, voicevox_url="http://127.0.0.1:50021"):
        self.voicevox_url = voicevox_url

    def tts(self, text, speed=1.1, speaker=3):
        res1 = requests.post(
            "http://127.0.0.1:50021/audio_query",
            params={"text": text, "speaker": speaker},
        )
        data = res1.json()
        data["speedScale"] = speed
        res2 = requests.post(
            "http://127.0.0.1:50021/synthesis",
            params={"speaker": speaker},
            data=json.dumps(data),
        )
        return res2.content


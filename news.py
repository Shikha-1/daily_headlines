import json
import requests
import time

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == "__main__":
    try:
        news = requests.get("https://gnews.io/api/v4/search?&q=None&lang=en&country=in&token=8496e24ceabec0c97216424af7001348").text
        news_dict = json.loads(news)
        speak("Hello everyone! I'm here to present today's headlines:")
        for news in news_dict["articles"]:
            for i in range(0, 10):
                speak(news_dict["articles"][i]["title"])
                time.sleep(0.5)
            speak("That's the end of today's headlines. Thank you!")
            break
    except Exception as e:
        speak(e)
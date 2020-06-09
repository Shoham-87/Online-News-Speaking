from win32com.client import Dispatch
import requests
import json
def speak(str):
    speak=Dispatch("SAPI.spVoice")
    speak.Speak(str)

if __name__ == '__main__':
    """Gets the object by request and then parsing is done by jason loads
    then spaling works"""
    r=requests.get('https://newsapi.org/v2/top-headlines?sources=bbc-news&apiKey=0e8810e8368c4a90ba28fefec20ccfd6')
    p=json.loads(r.text)
    speak("WELCOME MASTER NAME!TODAYS NEWS ARE.")
    tags=["First","Second","Third","Fourth","Fifth"]
    for i in range(5):
        speak(f"The {tags[i]} news is:")
        speak(p["articles"][i]["title"])
    speak("Thank You!")

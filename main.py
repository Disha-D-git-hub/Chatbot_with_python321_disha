import speech_recognition as sr
import os
import webbrowser
import win32com.client
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occured,Could you please try that again?"


if __name__ == '__main__':
    print('Pycharm')
    speaker.Speak("Hello I am Jarvis A I, your personal assistant")
    while True:
        print(f"Listening...")
        query = takeCommand()
        sites=[["youtube","https://youtube.com"],["wikipedia","https://wikipedia.com"],["google","https://google.com"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} right away!")
                webbrowser.open_new_tab(site[1])
                # Moved the break statement after the for loop.
                break
        if "time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"The time is: {strfTime}")
            
        if "music" in query:
            musicpath=r"\Users\disha\Music\NightChanges.mp3"
            os.system(f"start{musicpath}")
        if query.lower() == "stop":
            break
# import pyttsx3
# import speech_recognition as sr
# import webbrowser
# import os
# import random
# import datetime
# engine=pyttsx3.init("sapi5")
# voice=engine.getProperty("voices")
# engine.setProperty(voice,"voices[1].id")


# def speak(audio):
#     engine.say(audio)
#     engine.runAndWait()
# def takecommand():
#     r=sr.Recognizer()
#     with sr.Microphone() as source:
#         speak("listening")
#         r.pause_threshold=1
#         audio=r.listen(source)
#     try:
#         speak("reconizing")
#         quary=r.recognize_google(audio,language="eng-in")
#         print(f"user said{quary}")
#         return quary
#     except:
#         speak("not understand tell again")
# if __name__=="__main__":
#     while True:
#         quary=takecommand().lower()
#         sites=[["open google","google.com"],["open youtube","youtube.com"]]
#         for site in sites:
#             if site[0] in quary:
#                 webbrowser.open(site[1])
#         if "play music" in quary:
#             musicdir="D:\\music"
#             song=os.listdir(musicdir)
#             os.startfile(os.path.join(musicdir,random.choice(song)))
#         elif "open word" in quary:
#             os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word.lnk")
#         elif "time"  in quary:
#             time=datetime.datetime.now().strftime("%h %M ")
#             speak(time)
#         elif "open amazon" in quary:
#             os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Amazon.com.lnk")
#         elif "open notepad" in quary:
#             os.startfile("C:\\Windows\\notepad.exe")
        



#it with win32com.client
# import pyttsx3
import win32com.client
import speech_recognition as sr
import webbrowser
import os
import random
import datetime
# engine=pyttsx3.init("sapi5")
# voice=engine.getProperty("voices")
# engine.setProperty(voice,"voices[1].id")
speaker=win32com.client.Dispatch("SAPI.spvoice")


def speak(audio):
    # engine.say(audio)
    # engine.runAndWait()
    speaker.speak(audio)
def takecommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        speak("listening")
        r.pause_threshold=1
        audio=r.listen(source)
    try:
        speak("reconizing")
        quary=r.recognize_google(audio,language="eng-in")
        print(f"user said{quary}")
    except:
        speak("not understand tell again")
        return "non"
    return quary
if __name__=="__main__":
    while True:
        quary=takecommand().lower()
        sites=[["open google","google.com"],["open youtube","youtube.com"]]
        for site in sites:
            if site[0] in quary:
                webbrowser.open(site[1])
        if "play music" in quary:
            musicdir="D:\\music"
            song=os.listdir(musicdir)
            os.startfile(os.path.join(musicdir,random.choice(song)))
        elif "open word" in quary:
            os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word.lnk")
        elif "time"  in quary:
            time=datetime.datetime.now().strftime("%h %M ")
            speak(time)
        elif "open amazon" in quary:
            os.startfile("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Amazon.com.lnk")
        elif "open notepad" in quary:
            os.startfile("C:\\Windows\\notepad.exe")
        elif "quit" in quary:
            speak("thanks you")
            exit()
        












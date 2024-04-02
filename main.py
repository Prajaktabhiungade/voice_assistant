import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia 
import webbrowser
import os

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[1].id)
engine.setProperty('voice',voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning")
    elif hour>=12 and hour<18:
        speak("Good afternoon")
    else:
        speak("Good Evening")
    speak("I am jarvis mam.please tell me how may i help you")
def takeCommand():
    #it takes microphone input from the user and returns string output

    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("listening.....")
        r.pause_threshold=1
        audio=r.listen(source)


    try:
        print("Recognizing...")
        query=r.recognize_google(audio,language='en-in')
        print(f"User Said: {query}\n")

    except Exception as e:
       # print(e)
        print("say that again please.....")
        return "None"
    return query
    
if _name=="main_":
   wishMe()
#takeCommand()
while True:
    #if 1:
    query = takeCommand().lower()

    if 'wikipedia' in query:
        speak('Searching Wikipedia....')
        query = query.replace("wikipedia","")
        results = wikipedia.summary(query,sentences=2)
        speak("According to Wikipedia")
        print(results)
        speak(results)
    elif 'open youtube' in query:
        webbrowser.open("youtube.com")

    elif 'open google' in query:
        webbrowser.open("google.com")

    elif 'open stackoverflow' in query:
        webbrowser.open("stackoverflow.com")

    elif 'play music' in query:
         music_dir='D:\\Non critical\\songs\\Favorite Songs2'
         songs=os.listdir(music_dir)
         print(songs)
         os.startfile(os.path.join(music_dir, songs[0]))

    elif 'the time' in query:
        #codePath="C:\\Users\\Haris\\AppData\\Local\\Programs\\Microsoft VS code\\Code.exe"
        strTime=datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"sir,the time is{strTime}")
        #os.startfile(codePath)

    elif 'email to harry' in query:
        try:
            speak("what should i say?")
            content=takeCommand()
            to="harryyourEmail@gmail.com"
            sendEmail(to,content)
            speak("Email Has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry my friend harry bhai.i am not able to send this email")
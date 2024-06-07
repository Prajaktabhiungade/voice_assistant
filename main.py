import pyttsx3
import subprocess
import winshell
from   ecapture import ecapture as ec
import win32com.client as wincl
from urllib.request import urlopen
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import sys
from tkinter import *

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("I am Jarvis ma'am. Please tell me how may I help you")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User Said: {query}\n")

    except Exception as e:
        print("Say that again please...")
        return "None"
    return query

def exit_voice_assistant():
    print("Exiting Voice Assistant...")
    sys.exit()

def execute_command():
    query = takeCommand().lower()
    if 'wikipedia' in query:
        try:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        except Exception as e:
            print(e)
            speak("Sorry there is something wrong")
    elif 'open youtube' in query:
        try:
            webbrowser.open("youtube.com")
        except Exception as e:
            print(e)
            speak("Sorry there is something wrong")
    

    elif 'open google' in query:
            try:
                webbrowser.open("google.com")
                
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")

    elif 'open stack overflow' in query:
            try:
                webbrowser.open("stackoverflow.com")
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")


    elif 'play music' in query:
            try:
                music_dir = 'C:\\Users\\rites\\Desktop\\songs'
                songs = os.listdir(music_dir)
                print(songs)
                os.startfile(os.path.join(music_dir, songs[0]))
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")
                 
    elif 'open file' in query:
            try:
                music_dir = 'C:\\Users\\rites\\Desktop\\java'
                songs = os.listdir(music_dir)
                print(songs)
                os.startfile(os.path.join(music_dir, songs[0]))
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")
                 
    elif 'how are you' in query:
            try:
                speak("I am fine, Thank you")
                speak("How are you, Sir")
            except Exception as e:
                print(e)
                speak("Sorry, but i didn't understand ,can you ask again")
 
    elif 'who you are' in query or "what is your name" in query :
            try:
                speak("my name is jarvis sir")
                speak("for better friendship i also need to know your name, Sir")
            except Exception as e:
                print(e)
                speak("Sorry, but i didn't understand ,can you ask again")
    elif 'my name is' in query or "i am" in query or "myself" in query :
            try:
                speak("oh nice name, Thank you")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't get it can you please tell again")
    elif 'fine' in query or "good" in query:
            try:
                speak("It's good to know that your fine")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't get it can you please tell again")

    elif 'the time' in query:
            try:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"sir, the time is {strTime}")
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")

    elif 'shutdown system' in query:
            try:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")
    elif "restart" in query:
            try:
                subprocess.call(["shutdown", "/r"])
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")
    elif "hibernate" in query or "sleep" in query:
            try:
                speak("Hibernating")
                subprocess.call("shutdown / h")
            except Exception as e:
                print(e)
                speak("Sorry there is something wrong")
 
    elif 'empty recycle bin' in query:
            try:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")
            except Exception as e:
                print(e)
                speak("Sorry, I unable to empty the  recycle bin")
    elif "don't listen" in query or "stop listening" in query:
            try:
                speak("for how much time you want to stop jarvis from listening commands")
                a = int(takeCommand())
                time.sleep(a)
                print(a)
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't get it please say again")
    elif "open map" in query:
            try:
                query = query.replace("open map", "")
                location = query
                speak("User asked to open map")
                speak(location)
                webbrowser.open("https://www.google.com/maps/search/?api=1&parameters" + location + "")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't open the file")
    elif "write a note" in query:
            try:
                speak("What should i write, sir")
                note = takeCommand()
                file = open('jarvis.txt', 'w')
                speak("Sir, Should i include date and time")
                snfm = takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)
            except Exception as e:
                print(e)
                speak("Sorry, there is problem while writing , can you please try again")

    elif "show note" in query:
            try:
                speak("Showing Notes")
                file = open("jarvis.txt", "r") 
                print(file.read())
                speak(file.read(6))
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't show note")
    elif "click my photo" in query:
            try:
                    pyautogui.press("super")
                    pyautogui.typewrite("camera")
                    pyautogui.press("enter")
                    pyautogui.sleep(2)
                    speak("SMILE")
                    pyautogui.press("enter")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't open camera")
    elif "screenshot" in query:
            try:
                import pyautogui #pip install pyautogui
                im = pyautogui.screenshot()
                im.save("ss.jpg")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't take screenshot")
    elif "exit" in query:
            speak("Thanks for giving me your time")
            exit()

'''class VoiceAssistantGUI:
    def _init_(self, root):
        self.root = root
        self.root.title("Voice Assistant")
        self.root.geometry("400x400")
        self.label = Label(root, text="Voice Assistant", font=("Helvetica", 16))
        self.label.pack(pady=20)
        self.text = Text(root, height=10, width=40)
        self.text.pack(pady=10)
        self.start_button = Button(root, text="Start", command=self.start_assistant)
        self.start_button.pack(pady=10)
        self.stop_button = Button(root, text="Stop", command=exit_voice_assistant)
        self.stop_button.pack(pady=10)

    def start_assistant(self):
        self.text.insert(END, "Assistant started...\n")
        wishMe()
        while True:
            execute_command()'''

if __name__ == "__main__":
     wishMe()
     #takeCommand()
     execute_command()
     exit_voice_assistant()
    
from win32com.client import Dispatch
import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import time
import subprocess
import ecapture as ec
import wolframalpha
import json
import requests

number = 1
print('Loading Walmart Siri lol')

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice', 'voices[0].id')




speak = Dispatch("SAPI.SpVoice").Speak

def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
         speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good evening!")



def takeCommand():
 while number == 1:   
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio=r.listen(source)

        try:
            statement=r.recognize_google(audio,language='en-in')
            print(f"user said:{statement}\n")

        except Exception as e:
            print ("Walmart Siri didn't hear you lol")
            return "none"
        return statement
        time.sleep(.01)
    
    
        


speak("Loading Walmart Siri")
wishMe()



if __name__=='__main__':

    while True:
        statement = takeCommand().lower()
        if statement==0:
            continue

        if "goodbye" in statement or "okay bye" in statement or "turn off" in statement:
            speak ('see you later!')
            break


        if 'wikipedia' in statement:
            speak('Searching the wiki')
            statement =statement.replace("wikipedia", "")
            results = wikipedia.summary(statement, sentences=3)
            speak("according to the wiki")
            speak(results)
            
        

        elif 'open youtube' in statement:
            webbrowser.open_new_tab("https://www.youtube.com/watch?v=GFq6wH5JR2A")
            speak("youtube is now open")
            time.sleep(3)
            
        elif 'open google' in statement:
            webbrowser.open_new_tab("https://www.google.com")
            speak("Google is open")
            time.sleep(3)

        elif 'open gmail' in statement:
            webbrowser.open_new_tab("https://www.gmail.com")
            speak("gmail is now open")
            time.sleep(3)
            
        elif 'open hack the box' in statement:
            webbrowser.open_new_tab("https://www.hackthebox.com")
            speak("hack the box is now open")
            time.sleep(3)
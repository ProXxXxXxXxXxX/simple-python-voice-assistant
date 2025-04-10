##NOTE: I dont really care what you do with this code. If you want to take it and implement it into your own projects, go ahead.
## all I ask is if you use it in another public project just credit this original post somewhere in your project. 

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
## voice assistant command to check what time of day it is and say "good morning" or "good evening respectivly."
##  TODO: make a function to tell the time when user asks "what time is it"
def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
         speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")
    else:
        speak("Good evening!")

## when program is started, start listening

def takeCommand():
 while number == 1:   
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio=r.listen(source)

        try:
            statement=r.recognize_google(audio,language='en-in')
            print(f"user said:{statement}\n")
## if voice assistant didn't hear or heard something it doesnt have a command for, print "didn't hear you"
## TODO: make the voice assistant respond audibly "I didn't hear you" but dont make it annoying
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
## turns off if you say turn off
        if "goodbye" in statement or "okay bye" in statement or "turn off" in statement:
            speak ('see you later!')
            break

## when it hears "wikipidia" searches the wiki and responds with 3 sentences 
        if 'wikipedia' in statement:
            speak('Searching the wiki')
            statement = statement.replace("wikipedia", "")
            sentences = 3
            try:
                results = wikipedia.summary(statement, sentences=sentences)
                speak("according to the wiki")
                speak(results)
## asks if user would like the assistant to continue talking
                while True:
                    speak('continue?')
                    response = takeCommand()
                    if 'yes' in response or 'sure' in response or 'ok' in response:
                        sentences += 3  # increase the number of sentences to fetch
                        try:
                            results = wikipedia.summary(statement, sentences=sentences)
                            speak(results)
                        except wikipedia.exceptions.DisambiguationError: #handling for errors
                            speak("Multiple results found. Please be more specific.")
                            break
                        except wikipedia.exceptions.PageError: #handling for errors
                            speak("No results found. Please try again.")
                            break
                    else:
                        speak('okay')
                        break
            except wikipedia.exceptions.PageError: #handling for errors
                speak('No results found. Please try again.')
                
            except wikipedia.exceptions.DisambiguationError: #handling for errors
                speak ('multiple results found. Please be more specific.')
            
#below opens pages when asked 


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
            
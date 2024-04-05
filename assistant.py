import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import pyaudio


engin = pyttsx3.init("sapi5")
voices = engin.getProperty("voices")
engin.setProperty("voice",voices[0].id)

def speak(audio):
    engin.say(audio)
    engin.runAndWait()
    
def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("good morning sir!")
        
    elif hour >=12 and hour < 18:
        speak("good afternoon sir!")
        
    else:
        speak("good evening sir!")
        
    assname = ("jarvis")
    speak("I am your assistant")
    speak(assname)
    
def username():
    speak("What should i call you sir")
    uname = takecommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))
    
    speak("how can i help you?")
    
def takecommand():
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        print("listening....")
        r.pause_threshold = 1
        audio = r.listen(source)
        
    try:
        print("recognizing....")
        query = r.recognize_google(audio,language = "en-in")
        print(f"User said: {query}\n")

    except Exception as e :
        print(e)
        print("unable to recognize your voice.")
        return "none"
    
    return query

def sendemail(to, content):
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.ehlo()
    server.starttls()
    
    server.login("*******","*********")
    server.sendmail("******",to,content)
    server.close()
    
    
if __name__ =="__main__":
    clear = lambda: os.system("cls")
    clear()
    wishme()
    username()
    
    while True:
        query = takecommand().lower()
        
        if "wikipedia" in query:
            speak("searching wikipedia","")
            result = wikipedia.summary(query, sentences = 3)
            speak ("according to wikipedia")
            print(result)
            speak(result)
            
        elif "open youtube" in query:
            speak("here you go to youtube \n")
            webbrowser.Chrome.open("youtube.com")
            
        elif "open google" in query:
            speak("here you go to google \n.")
            webbrowser.Chrome.open("google.com")
            
        elif "stackoverflow" in query:
            speak("here you go to stackoverflow. happy coding \n")
            webbrowser.Chrome.open("stackoverflow.com")
        
        elif "playmusic" in query:
            speak("here you are. chill.\n")
            music_dir = r"D:\13-MUSIC\RAP"
            song = os.listdir(music_dir)
            print(song)
            random = os.startfile(os.path.join(music_dir, song[1]))
            
        elif "the time" in query:
            strttime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"Sir, the time is {strttime}")


        elif "how are you " in query:
            speak("i'm fine ,thank you")
            speak("what about you sir?")
            
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
            
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
            
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takecommand()
            speak("Thanks for naming me"+ assname)
            
        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)
            
        elif 'exit' in query:
            speak("ok, call me when you need help.")
            exit()
            
        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by you sir.")
            
        elif 'joke' in query:
            speak(pyjokes.get_joke())    
        
        elif "calculate" in query: 
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            print("The answer is " + answer) 
            speak("The answer is " + answer)
        
        elif 'search' in query or 'play' in query:
            query = query.replace("search", "") 
            query = query.replace("play", "")          
            webbrowser.open(query)
            
        elif "why you came to world" in query:
            speak("Thanks to alireza. further It's a secret")
            
        elif 'power point presentation' in query:
            speak("opening Power Point presentation")
            power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power)
 
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
            
        elif "who are you" in query:
            speak("I am your virtual assistant created by alireza")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project by you.but about the reason : you will see. time will show you. ")
            
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                
        elif "emty recycle bin" in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = True, sound = True)
            speak("Recycle Bin Recycled")
            
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takecommand())
            time.sleep(a)
            print(a)
        
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("you asked to Locate")
            speak(location)
            webbrowser.Chrome.open("https://www.google.nl / maps / place/" + location + "")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
            
        elif "jarvis" in query:
            wishme()
            speak("Jarvis in your service Mister")
            speak(assname)

        elif "i love you" in query:
            speak("love? it will kill you never fell in love with someone sir.")
import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import random
import win32com.client as wincl
import json
import time
import subprocess
import ctypes #pip install ctypes
from email.mime import text
from pythonwin.pywin.framework import app



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am jarvis Sir. Please tell me how may I help you")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

 
        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to google\n")
            webbrowser.open("google.com")

        elif 'open facebook' in query:
             speak("Here you go to Facebook\n")
             webbrowser.open("facebook.com")

        elif 'open amazon' in query:
             speak("Here you go to Amazon\n")
             webbrowser.open("amazon.in")

        elif 'open flipkart' in query:
             speak("Here you go to Flipkart\n")
             webbrowser.open("flipkart.com")

        elif 'open gmail' in query:
            speak("Here you go to Gmail\n")
            webbrowser.open("gmail.com")



            #basic question

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "mast" in query:
             speak("It's good to know that your fine")

        elif "who made you" in query or "who created you" in query:
             speak("I have been created by saqlain .")

        elif "who i am" in query or "main kaun hun" in query:
              speak("If you talk then definately your human.")
 
        elif 'what is love' in query or " what's love" in query:
             speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
             speak("I am your Personal assistant created by saqlain ")

        elif 'reason' in query:
             speak("I was created as a Minor project by Master saqlain")

        elif "why you came to world" in query:
             speak("Thanks to saqlain. further It's a secret")

        elif "what is my gaming name" in query:
             speak("your gaming name is DeathStroke")     
        
        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")   
        
    
        elif 'open news' in query:
            speak("Here you go to news\n")
            webbrowser.open('news.google.co.in')

        elif 'open wikipedia' in query:
            speak("Here you go to wikipedia\n")
            webbrowser.open('www.wikipedia.org')

        elif 'open weather' in query:
            speak("Here you go to weather\n")
            webbrowser.open('imdmumbai.gov.in')

        elif "who is hasnain" in query:
            speak("hasnain is saqlain smaller brother")

        elif "tell me about more about hasnain" in query:
            speak("hasnain is studing in 10 standard, he want to be youtuber but fact is he dont doing any thing to be a youtuber he only want mobile phone. ussi ki rattt laga k rakha h. In my suggestion please dont give mobile hasnain has two freint one kaif kaliya and another rafi samosawala ")

        elif "who is shaheen" in query:
            speak("shaheen full name is sayed shaheen fatima ak sheeba bano and she is mother of saqlain and she also have two sautela aulad hasnain shahreen")

        elif "tell me more about shaheen" in query:
            speak("shaheen is housewife , she like playing ludo , block. she use facebook and instagram whatsapp daily she is totally addicted. but she is a working hard women she never hurt anyone and smilling always. she is goodlistner she listen all bakwas of family. but she only like saqlain se kaam karana ")

        elif "who is shahreen" in query:
             speak("shahreen full name is sayed shahreen fatima urff papa ki pari she is elder sister of saqlain")

        elif "tell me more about shahreen" in  query:
            speak("shahreen is a working woman and she think that she is most smartest person in the world but she is always sounding like frustated and tired he rarely smile in day. she is the boss of my family . raaz ki baat yeah h isse koi behas nhi karna chata qki bahot sunati h.. but dil ki bahot achi h ghr ki zimadari k wajai aisi h isi chakkar mai usko daadii ati h. hahahahahahahhahahaha ")

        elif " who is ayaz" in query:
           speak("ayaz full name is sayed ayaz husain he is a father of saqlain,hasnain,shahreen")

        elif "tell me more about ayaz" in query:
            speak("ayaz is a very kind good person, he like to watch chote miyan in free time he do all household work with full shiddat , he rarely got angry and he has strong beleave in allah, but woh apnai app ko thoda kam samjhte h qki woh kam padhai likhe h or unko nicotec ki laat padh gyi h mgr woh nhi mante h yeah baad sirf tension h yeah bahana deke khate h")


        elif 'play music' in query:
            music_dir = 'D:\\Non Critical\\songs\\Favorite Songs2'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))
            

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak("Sir, the time is {strTime}")

        elif 'open youtube music' in query:
              speak("Here you go to Youtube Music\n")
              webbrowser.open("music.youtube.com")

        elif 'open code' in query:
            codePath = "C:\\Users\\saqlain\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'email to saqlain' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "sayedsaqlain16@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry my friend harry bhai. I am not able to send this email")   

                exit()

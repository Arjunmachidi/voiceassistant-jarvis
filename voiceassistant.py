import subprocess
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
import wolframalpha

# Initialize the pyttsx3 engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    """Function to speak the given audio string."""
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    """Function to wish the user based on the time of day."""
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning Sir!")

    elif 12 <= hour < 18:
        speak("Good Afternoon Sir!")

    else:
        speak("Good Evening Sir!")

    assname = "Jarvis"
    speak("I am your Assistant")
    speak(assname)

def username():
    """Function to ask and greet the user by their name."""
    speak("What should I call you, sir?")
    uname = takeCommand()
    speak(f"Welcome, Mister {uname}")
    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print(f"Welcome Mr. {uname}".center(columns))
    print("#####################".center(columns))
    speak("How can I help you, Sir?")

def takeCommand():
    """Function to capture voice commands and return the recognized text."""
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
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query.lower()

def sendEmail(to, content):
    """Function to send an email using SMTP."""
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('email', 'password')
    server.sendmail('email', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')  # Function to clear the console screen
    clear()
    wishMe()
    username()

    while True:
        query = takeCommand()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube")
            webbrowser.open("https://www.youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google")
            webbrowser.open("https://www.google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Overflow. Happy coding!")
            webbrowser.open("https://stackoverflow.com")

        elif 'play music' in query or 'play song' in query:
            speak("Here you go with music")
            music_dir = "C:\\Users\\ARJUN\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random_song = os.path.join(music_dir, random.choice(songs))
            os.startfile(random_song)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")


        elif 'email to arjun' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("Whom should I send it to?")
                to = input()  # You can use takeCommand() to get the recipient from voice
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that you are fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir?")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "what is your name" in query:
            speak("My friends call me")
            speak(assname)
            print(f"My friends call me {assname}")

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Arjun.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = "API"
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

        elif "who i am" in query:
            speak("If you talk then definitely you are human.")

        elif "why you came to world" in query:
            speak("Thanks to Arjun. Further, it's a secret")

        elif 'power point presentation' in query:
            speak("Opening Power Point presentation")
            power = r"C:\ProgramData\Microsoft\Windows\start Menu\Programs\PowerPoint.lnk"
            os.startfile(power)

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Arjun")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister Arjun")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, "Location of wallpaper", 0)
            speak("Background changed successfully")


        elif 'news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1

                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("For how much time do you want me to stop listening to commands?")
            pause_duration = int(takeCommand())
            time.sleep(pause_duration)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak(f"User asked to Locate {location}")
            webbrowser.open(f"https://www.google.nl/maps/place/{location}")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the applications are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should I write, sir?")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, should I include date and time?")
            include_datetime = takeCommand()
            if 'yes' in include_datetime or 'sure' in include_datetime:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime + " - ")
            file.write(note)
            file.close()

        elif "show note" in query:
            try:
                speak("Showing Notes")
                file = open("jarvis.txt", "r")
                print(file.read())
                speak(file.read())
                file.close()
            except Exception as e:
                print(e)
                speak("I am unable to show notes at the moment")

        elif "update assistant" in query:
            speak("After downloading the file, please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)
            with open("Voice.py", "wb") as Pypdf:
                total_length = int(r.headers.get('content-length'))
                for ch in progress.bar(r.iter_content(chunk_size=2391975), expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)

        elif "jarvis" in query:
            wishMe()
            speak("Jarvis at your service, Mister")
            speak(assname)

        elif "weather" in query:
            api_key = "API key"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("Which city's weather do you want to know?")
            city_name = takeCommand()
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url)
            x = response.json()
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidity = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                speak(f"Temperature: {current_temperature} Kelvin")
                speak(f"Atmospheric pressure: {current_pressure} hPa")
                speak(f"Humidity: {current_humidity} %")
                speak(f"Description: {weather_description}")
            else:
                speak("City not found")

        elif "send message" in query:
            speak("You need to create an account on Twilio to use this service")
            account_sid = 'Account Sid key'
            auth_token = 'Auth token'
            client = Client(account_sid, auth_token)
            speak("What message would you like to send?")
            message_content = takeCommand()
            speak("From which number should I send this message?")
            sender_number = input()  # You can use takeCommand() to get the sender's number from voice
            speak("To which number should I send this message?")
            receiver_number = input()  # You can use takeCommand() to get the receiver's number from voice

            message = client.messages.create(
                body=message_content,
                from_=sender_number,
                to=receiver_number
            )

            print(message.sid)

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you, Mister?")
            speak(assname)

        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "how are you" in query:
            speak("I'm fine, glad you asked me that")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query or "who is" in query:
            client = wolframalpha.Client("API")  # Use the same API key that you have generated earlier
            res = client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")
                speak("No results")



       



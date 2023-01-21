#import all the required for the project
import ctypes
import json
import shutil
import smtplib
from urllib.request import urlopen
import requests
import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import pyjokes
import subprocess
import wolframalpha
import tkinter
import webbrowser
import os
import winshell
import time
import win32com.client as wincl


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

assname = "Prabhat Ranjan"
uname = ""
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishme():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good Morning Sir !")

    elif 12 <= hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    speak("I am your Assistant")
    speak(assname)


def username():
    speak("What should I call you sir?")
    uname = takecommand()
    speak("Welcome Sir")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    speak("How can I Help you" + uname)


def takecommand():
    listener = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        listener.adjust_for_ambient_noise(source, duration=1)
        audio = listener.listen(source)

    try:
        print("Recognizing...")
        query = listener.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice. Please repeat.")
        return "None"
    return query


def sendemail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('#userName', '#password')
    server.sendmail('#email address', to, content)
    server.close()

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wishme()
    username()
while True:
    query = takecommand().lower()
    if 'wikipedia' in query:
        speak('Searching Wikipedia...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=3)
        speak("According to Wikipedia")
        print(results)
        speak(results)

    elif 'open youtube' in query:
        speak("Here you go to Youtube\n")
        webbrowser.open("youtube.com")

    elif 'open google' in query:
        speak("Here you go to Google\n")
        webbrowser.open("google.com")

    elif 'open stack overflow' in query:
        speak("Here you go to Stack Over flow. Happy coding")
        webbrowser.open("stackoverflow.com")

    if 'youtube' in query:
        query = query.replace("youtube", "")
        speak('Playing ' + query)
        pywhatkit.playonyt(query)

    elif 'play music' in query or "play song" in query:
        speak("Here you go with music")
        music_dir = "C:\Stuff\Songs\English\Taylor" \
                    " Swift - reputation (2017) (Mp3 320kbps)" \
                    " [Clean Version]\Taylor Swift - reputation (2017)"
        songs = os.listdir(music_dir)
        #print(songs)
        random = os.startfile(os.path.join(music_dir, songs[len(songs)-1]))
    elif 'the time' in query:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f"Sir, the time is {strTime}")

    elif 'email to Prabhat' in query:
        try:
            speak("What should I say?")
            content = takecommand()
            to = "#email address"
            sendemail(to, content)
            speak("Email has been sent !")
        except Exception as e:
            print(e)
            speak("I am not able to send this email")

    elif 'send a mail' in query:
        try:
            speak("What should I say?")
            content = takecommand()
            speak("whom should I send")
            to = '#email address'
            sendemail(to, content)
            speak("Email has been sent !")
        except Exception as e:
            print(e)
            speak("I am not able to send this email")

    elif 'how are you' in query:
        speak("I am fine, Thank you, glad you asked me that")
        speak("How are you, Sir")

    elif 'fine' in query:
        speak("It's good to know that your fine")

    elif "change my name" in query:
        speak("What would you like me to call you, Sir ")
        uname = takecommand()
        speak("From now onwards, I will call you" +uname)

    elif "change your name to" in query:
        speak("What would you like to call me, Sir ")
        assname = takecommand()
        speak("Thanks for naming me")

    elif "what's your name" in query or "what is your name" in query:
        speak("My friends call me")
        speak(assname)
        print("My friends call me", assname)

    elif 'stop the program' in query:
        speak("Thanks for giving me your time. Have a nice day")
        exit()

    elif "who made you" in query or "who created you" in query:
        speak("I have been created by Prabhat Ranjan")

    elif 'joke' in query:
        speak(pyjokes.get_joke())

    elif "result" in query:

        app_id = "#token id"
        client = wolframalpha.Client(app_id)
        indx = query.lower().split().index('result')
        query = query.split()[indx + 1:]
        res = client.query(' '.join(query))
        answer = next(res.results).text
        speak("The answer is " + answer)

    elif 'search' in query:
        query = query.replace("search", "")
        webbrowser.open(query)

    elif "will you be my girlfriend" in query or "will you be my valentine" in query:
        speak("I'm not sure about, may be you should give me some time")

    elif "i love you" in query:
        speak("It's hard to understand")

    elif "who am i" in query:
        speak("If you can talk then definitely you are a human.")

    elif "why you came to this world" in query:
        speak("Thanks to Prabhat. further It's a secret. Ask him if you want")

    elif 'resume' in query:
        speak("opening Power Point presentation")
        power = r"C:\Stuff\Documents\Resume\Resume.pdf"
        os.startfile(power)

    elif 'is love' in query:
        speak("It is 7th sense that destroy all other senses")

    elif "who are you" in query:
        speak("I am your virtual assistant created by Prabhat")

    elif "good Morning" in query or "good evening" in query or "good night" in query:
        speak("A very" + query)
        speak("How are you Mister")
        speak(uname)

    elif 'change background' in query:
        directory = r"C:\Users\Prabhat\OneDrive\Pictures\Wallpaper"
        imagePath = directory + "\Hayley.jpg"
        ctypes.windll.user32.SystemParametersInfoW(20, 0, imagePath, 0)
        speak("Background changed succesfully")

    elif 'india news' in query:

        try:
            jsonObj = urlopen('''https://newsapi.org/v2/top-headlines?country=in&apiKey=#keyID''')
            data = json.load(jsonObj)
            i = 1

            speak('here are some top news from the times of india')
            print('''=============== INDIA NEWS  ============''' + '\n')

            for item in data['articles']:
                print(str(i) + '. ' + item['title'] + '\n')
                print(item['description'] + '\n')
                speak(str(i) + '. ' + item['title'] + '\n')
                i += 1
        except Exception as e:
            print(str(e))

    elif 'world news' in query:
        try:
            query_params = {
                "source": "bbc-news",
                "sortBy": "top",
                "apiKey": "#api key"
                            }
            main_url = " https://newsapi.org/v1/articles"
            res = requests.get(main_url, params=query_params)
            open_bbc_page = res.json()
            article = open_bbc_page["articles"]
            results = []

            for ar in article:
                results.append(ar["title"])

            for i in range(len(results)):
                print(i + 1, results[i])

            speak(results)
        except Exception as e:

            print(str(e))

    elif 'lock window' in query:
        speak("locking the device")
        ctypes.windll.user32.LockWorkStation()

    elif "restart" in query:
        subprocess.call(["shutdown", "/r"])

    elif "hibernate" in query or "sleep" in query:
        speak("Hibernating")
        subprocess.call("shutdown / h")

    elif "log off" in query or "sign out" in query:
        speak("Make sure all the application are closed before sign-out")
        time.sleep(5)
        subprocess.call(["shutdown", "/l"])

    elif 'shutdown system' in query:
        speak("Hold On a Sec ! Your system is on its way to shut down")
        subprocess.call('shutdown / p /f')

    elif 'empty recycle bin' in query:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        speak("Recycle Bin Recycled")

    elif "don't listen" in query or "please stop listening" in query:
        speak("for how much time you want to stop" +(assname)+ "from listening commands")
        a = int(takecommand())
        time.sleep(a)
        print(a)

    elif "where is" in query:
        content = takecommand()
        query = query.replace("where is", "")
        location = query
        speak("User asked to Locate")
        speak(location)
        webbrowser.open("https://www.google.nl/maps/place/" + location + "")

    # elif "camera" in query or "take a photo" in query:
    #     ec.capture(0, "Jarvis Camera ", "img.jpg")

    elif "write a note" in query:
        speak("What should i write, sir")
        note = takecommand()
        file = open('bot.txt', 'w')
        speak("Sir, Should i include date and time")
        snfm = takecommand()
        if 'yes' in snfm or 'sure' in snfm:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            file.write(strTime)
            file.write(" :- ")
            file.write(note)
        else:
            file.write(note)

    elif "show note" in query:
        speak("Showing Notes")
        file = open("jarvis.txt", "r")
        print(file.read())
        speak(file.read(6))

    elif "weather" in query:

        City_API_endpoint = "http://api.openweathermap.org/data/2.5/weather?q="
        join_key = "&appid=" + "#api key"
        units = "&units=metric"
        speak("What is the city name")
        City = takecommand()
        current_city_weather = City_API_endpoint + City + join_key + units
        print(current_city_weather)

        json_data = requests.get(current_city_weather).json()

        if json_data["cod"] != "404":
            y = json_data['main']
            current_temperature = json_data['main']['temp']
            current_pressure = json_data['main']['pressure']
            current_humidity = json_data['main']['humidity']
            feel_like = json_data['main']['feels_like']
            z = json_data['weather']
            weather_description = z[0]['description']

            speak(" Temperature (in celsius) = " + str(
                current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                current_pressure) + "\n humidity (in percentage) = " + str(current_humidity) + "\n description = " + str(
                weather_description) +"\n feels like = " +str(feel_like))

        else:
            speak(" City Not Found ")

    elif "whatsapp #name" in query:
        speak("what is the message you want to send")
        message = takecommand()
        sample = "This is Alexa, this is for checking whatsapp functionality \n I am trying to integrate whatsapp to the Alexa"
        finalmsg ="This is my message: " + message + "\n" +sample
        # strTime = datetime.datetime.now().strftime("%H)
        hour = time.strftime("%H")
        minute = time.strftime("%M")
        print(hour, minute)
        pywhatkit.sendwhatmsg(f"#mobile number", finalmsg, int(hour), int(int(minute)+2))




    # elif "send message " in query:
    #     # You need to create an account on Twilio to use this service
    #     account_sid = 'Account Sid key'
    #     auth_token = 'Auth token'
    #     client = Client(account_sid, auth_token)






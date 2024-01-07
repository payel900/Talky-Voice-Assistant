import subprocess 
import wolframalpha
import pyttsx3
import random
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import time
import ecapture as ec
import winshell
import pyjokes
import feedparser
import smtplib
import datetime 
import json
import requests
from email.message import EmailMessage
from twilio.rest import Client
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import tkinter 
from tkinter import *
import shutil


voiceEngine = pyttsx3.init('sapi5')
voices = voiceEngine.getProperty('voices')
voiceEngine.setProperty('voice', voices[1].id)

def speak(text):
	voiceEngine.say(text)
	voiceEngine.runAndWait()

def wish():
    print("Wishing.")
    time = int(datetime.datetime.now().hour)
    global uname,asname
    if time>= 0 and time<12:
        speak("Good Morning !")

    elif time<18:
        speak("Good Afternoon !")

    else:
        speak("Good Evening !")

    speak(" I am Talky virtual Assistant. Tell me how may I help you.")
    
    
def getName():
    global uname
    speak("Can I please know your name?")
    uname = takeCommand()
    if uname != "None":
        print("Name:",uname)
        speak("I am glad to know you!")
        columns = shutil.get_terminal_size().columns
        speak("How can i Help you, ")
        speak(uname)
    else:
        speak("Please tell your name?")

def takeCommand():
    global showCommand
    showCommand.set("Listening....")
    
    # cmdLabel.config(textvariable=points)

    rec = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        rec.pause_threshold = 1

        audio = rec.listen(source)

        try:
            print("Recognizing...")
            query = rec.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")
        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand audio")
            return "None"
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service; {0}".format(e))
            return "None"

        except Exception as e:
            print(e)
            print("Say that again please...")
            return "None"
        return query

def getWeather(city_name):
    cityName = takeCommand() 
    
    #getting input of name of the place from user
    
    api = "http://api.openweathermap.org/data/2.5/weather?q=" + cityName + "&appid=eea37893e6d01d234eca31616e48c631"
    w_data = requests.get(api).json()
    weather = w_data['weather'][0]['main']
    temp = int(w_data['main']['temp'] - 273.15)
    temp_min = int(w_data['main']['temp_min'] - 273.15)
    temp_max = int(w_data['main']['temp_max'] - 273.15)
    pressure = w_data['main']['pressure']
    humidity = w_data['main']['humidity']
    visibility = w_data['visibility']
    wind = w_data['wind']['speed']
    sunrise = time.strftime("%H:%M:%S", time.gmtime(w_data['sys']['sunrise'] + 19800))
    sunset = time.strftime("%H:%M:%S", time.gmtime(w_data['sys']['sunset'] + 19800))
    all_data1 = f"Condition: {weather} \nTemperature: {str(temp)}°C\n"
    all_data2 = f"Minimum Temperature: {str(temp_min)}°C \nMaximum Temperature: {str(temp_max)}°C \n" \
                f"Pressure: {str(pressure)} millibar \nHumidity: {str(humidity)}% \n\n" \
                f"Visibility: {str(visibility)} metres \nWind: {str(wind)} km/hr \nSunrise: {sunrise}  " \
                f"\nSunset: {sunset}"
    speak(f"Gathering the weather information of {cityName}...")
    print(f"Gathering the weather information of {cityName}...")
    print(all_data1)
    speak(all_data1)
    print(all_data2)
    speak(all_data2)

def getNews():
    try:
        response = requests.get('https://www.bbc.com/news')
  
        b4soup = BeautifulSoup(response.text, 'html.parser')
        headLines = b4soup.find('body').find_all('h3')
        unwantedLines = ['BBC World News TV', 'BBC World Service Radio',
                    'News daily newsletter', 'Mobile app', 'Get in touch']

        for x in list(dict.fromkeys(headLines)):
            if x.text.strip() not in unwantedLines:
                print(x.text.strip())
    except Exception as e:
        print(str(e))
    
def callVoiceAssistant():

    uname=''
    asname=''
    os.system('cls')
    wish()
    getName()
    print(uname)

    while True:

        command = takeCommand().lower()
        print(command)
        home_user_dir = os.path.expanduser("~")

        if "start" in command:
            wish()
            
        elif 'how are you' in command:
            speak("I am fine, Thank you")
            speak("How are you, ")
            speak(uname)

        elif "good morning" in command or "good afternoon" in command or "good evening" in command:
            speak("A very" +command)
            speak("Thank you for wishing me! Hope you are doing well!")

        elif 'fine' in command or "good" in command:
            speak("It's good to know that your fine")
       
        elif "who are you" in command:
            speak("I am Talky virtual Assistant developed on python , "
                        "and Subhas Nath Sir is Our project guide,"
                        "and my team members are Vishwajit Kumar,Pooja Arora, Arpita Pal,Payel Patra,Saurodip Majee and this is our minor project.")

        elif "change my name to" in command:
            speak("What would you like me to call you, Sir or Madam ")
            uname = takeCommand()
            speak('Hello again,')
            speak(uname)
        
        elif "change name" in command:
            speak("What would you like to call me, Sir or Madam ")
            assname = takeCommand()
            speak("Thank you for naming me!")

        elif "what's your name" in command:
            speak("People call me")
            speak(assname)
        
        elif 'time' in command:
            strTime = datetime.datetime.now()
            curTime=str(strTime.hour)+"hours"+str(strTime.minute)+"minutes"+str(strTime.second)+"seconds"
            speak(uname)
            speak(f" the time is {curTime}")
            print(curTime)

        elif 'wikipedia' in command:
            speak('Searching Wikipedia')
            command = command.replace("wikipedia", "")
            results = wikipedia.summary(command, sentences = 3)
            speak("These are the results from Wikipedia")
            print(results)
            speak(results)
    
        elif 'open youtube' in command:
            speak("Here you go, the Youtube is opening\n")
            webbrowser.open("youtube.com")

        elif 'open google' in command:
            speak("Opening Google\n")
            webbrowser.open("google.com")

        elif 'play music' in command or "play song" in command:
            speak("Enjoy the music!")
            n=random.randint(0,50)
            print(n)
            music_dir = 'D:\music'
            song= os.listdir(music_dir)
            os.startfile(os.path.join(music_dir, song[n]))

        elif 'joke' in command:
            speak(pyjokes.get_joke())
            
        elif 'mail' in command:
            def send_mail(receiver, subject, message):
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.ehlo()
                    server.starttls()
                    server.login('vishwajitkumar102@gmail.com','oeaxgzaekkwxvzcd')
                    email = EmailMessage()
                    email['From'] = 'vishwajitkumar102@gmail.com'
                    email['To'] = receiver
                    email['Subject'] = subject
                    email.set_content(message)
                    server.send_message(email)
                    
            email_list = {
                        'Umesh': 'ums945891@gmail.com',
                        'Rishabh': 'rishabhran123@gmail.com',
                        'name': 'something@something.com',
                        'assitant': 'something@something.com',
                        'Abhishek':'anandabhishek045@gmail.com',
                        'Vishwajeet':'vishwajitkumar102@gmail.com',
                        'Bishnu':'bishnusahcool@gmail.com',
                        'Rohit':'rohitbxr212@gmail.com'    
                    }
            
            def get_mail_info():
                try:
                    speak('To whom you want to send email')
                    name = takeCommand()
                    receiver = email_list[name]
                    print(receiver)
                    speak('What is the subject of your email?')
                    subject = takeCommand()
                    speak('Tell me the text in your email')
                    message = takeCommand()
                    send_mail(receiver, subject, message)
                    speak('Hey lazy person. Your email is sent Successfully.')
                    speak('Do you want to send more email?')
                    send_more = takeCommand()
                    if 'yes' in send_more:
                        get_mail_info()
                except Exception as e:
                    print(e)
                    speak("I am sorry, not able to send this email")
        
            get_mail_info()
            

        elif 'exit' in command:
            speak("Thanks for giving me your time")
            exit()

        elif "will you be my gf" in command or "will you be my bf" in command:
            speak("I'm not sure about that, may be you should give me some time")

        elif "i love you" in command:
            speak("Thank you! But, It's a pleasure to hear it from you.")

        elif "weather" in command:
            speak(" Please tell your city name ")
            print("City name : ")
            cityName = takeCommand()
            getWeather(cityName)

        elif "what is" in command or "who is" in command:
            
            client = wolframalpha.Client("API_ID")
            res = client.query(command)

            try:
                print (next(res.results).text)
                speak (next(res.results).text)
            except StopIteration:
                print ("No results")

        elif 'search' in command:
            command = command.replace("search", "")
            webbrowser.open(command)

        elif 'news' in command:
            getNews()
        
        elif "don't listen" in command or "stop listening" in command:
            speak("Stopped listening to commands for 5 mins")
            time.sleep(5)
            print("Stopped listening to commands for 5 mins")

        elif "camera" in command or "take a photo" in command:
            ec.capture(0, "Talky Camera ", "img.jpg")
        
        elif 'shutdown system' in command:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

        elif "restart" in command:
            subprocess.call(["shutdown", "/r"])

        elif "sleep" in command:
            speak("Setting in sleep mode")
            subprocess.call("shutdown / h")
            
        #code for to write a note     
            
        elif "write a note" in command:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('Talky.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
                
        #code to open whatsApp(automation)
               
        elif 'open whatsapp' in command:
            os.startfile(home_user_dir + "\\AppData\\Local\\WhatsApp\\WhatsApp.exe") 
                   
        else:
            speak("Sorry, I am not able to understand you, please try again!")


#Creating the main window 

wn = tkinter.Tk() 
wn.title("Talky Voice Assistant")
wn.geometry('1920x1080')
wn.config(bg='LightBlue1')
  
Label(wn, text='Welcome and  meet the Voice Assistant by Talky ', bg='LightBlue1',
      fg='black', font=('Courier', 15)).place(x=50, y=10)

#Button to convert PDF to Audio form

Button(wn, text="Lets Begin...", bg='gray',font=('Courier', 15), fg='white',highlightcolor='black',
       command=callVoiceAssistant).place(x=270, y=100)

showCommand=StringVar()
cmdLabel=Label(wn, textvariable=showCommand, bg='LightBlue1',
      fg='black', font=('Courier', 15))
cmdLabel.place(x=250, y=150)

#Runs the window till it is closed

wn.mainloop()
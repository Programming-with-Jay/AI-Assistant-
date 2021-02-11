import ctypes
import datetime
import os
import random
import time
import urllib.request
import webbrowser
import pronouncing
import pyjokes
import pyttsx3
import pywhatkit as kit
import requests
import speech_recognition as sr
import wikipedia
import winshell
from PyDictionary import PyDictionary
from gekko import GEKKO
from nltk.corpus import wordnet
from translate import Translator
from win32com.client import Dispatch
from pygame import mixer
import time
import holidays
import time
import keyboard
from pyautogui import *
import pyautogui
import os
import pandas as pd
from datetime import datetime


print ( "Initializing Alexis" )
engine = pyttsx3.init ( 'sapi5' )
voices = engine.getProperty ( 'voices' )
engine.setProperty ( 'voice' , voices [ 0 ].id )
comment = ('amazing' , 'awesome' , 'like it' , 'keep it up' , 'great' , 'good')


# speak will speak the str

def speak(text):
    
    engine.say ( text )
    engine.runAndWait ()


def takeCommand():
    r = sr.Recognizer ()
    with sr.Microphone () as source:
        print ( "listening..." )
        audio = r.listen ( source )

    try:
        print ( "Recognizing" )
        query = r.recognize_google ( audio , language='en=IN' )
        print ( f'user said : {query}\n' )
    except Exception:
        query = None
    return query


def greet():
    hour = int (datetime.now().hour )
    if 0 <= hour < 12:
        speak ( 'Good morning sir' )

    elif 12 <= hour < 18:
        speak ( 'Good afternoon sir' )
    else:
        speak ( "Good evening sir " )

def wiki(query):
    speak ( "Ok sir" )
    speak ( "Searching for Wikipedia" )
    query = query.replace ( "wikipedia" , "" )
    result = wikipedia.summary ( query , sentences=2 )
    print ( result )
    speak ( result )
greet ()

speak ( '"I am Alexis   How   may I help You?"' )
run = True
while run:
    query = takeCommand ()
    try:
        if 'wikipedia' in query.lower():
            wiki(query)
        if 'classroom' in query.lower ():
            speak ( "Ok sir" )
            webbrowser.open ( "https://classroom.google.com/c/MTA5NTkwNDg0NzQ2" )

        elif ' YouTube' and ' youtube' in query.lower ():
            speak ( "Searching on youtube" )
            search = query
            speak ( "Here's what I got for you sir" + str ( search ) )
            kit.playonyt ( search )
        elif 'open YouTube ' in query.lower ():
            speak ( "Yes sir" )
            y = 'https://youtube.com'
            webbrowser.get ().open ( y )
        elif 'meaning' in query.lower ():
            dictionary = PyDictionary ()
            word = query
            s = word.split ()
            b = s [ 5 ]
            print ( dictionary.meaning ( b ) )
            speak ( dictionary.meaning ( b ) )
        elif 'rhyming' in query.lower ():
            pro = query
            w = pro.split ()
            v = w [ 3 ]
            a = pronouncing.rhymes ( v )
            speak ( a )
            print ( a )
        elif 'translate' in query.lower ():
            speak ( 'What is the word?' )
            d = takeCommand ()
            speak ( "To which language should I convert to?" )
            lan = takeCommand ()
            translator = Translator ( to_lang=lan )
            t = translator.translate ( d )
            print ( t )
            speak ( t )
        elif 'linear equation' in query.lower ():
            speak ( "what is the equation?" )
            equation = input ( "Enter the equation:  " )
            m = GEKKO ()
            x = m.Var ()
            y = m.Var ()
            m.Equation ( equation )
            m.solve ( disp=False )
            print ( x.value , y.value )
        elif 'holidays' in query.lower():
            india=holidays.India(years=2020).items()
            if datetime.datetime.now() in india:
                speak('Sir it is holiday')
            else:
                speak('Sir we do not have a holiday ')

        elif 'change voice' in query.lower ():
            engine.setProperty ( 'voice' , voices [ 1 ].id )
            speak ( 'Hello sir' )
        elif 'meeting' in query.lower ():

            x=978
            y=506
            def sign_in(meeting_id,paswd):
                os.startfile('C:/Users/Jay/AppData/Roaming/Zoom/bin/Zoom')
                time.sleep(5)
                pyautogui.click(907,414)
                time.sleep(5)
                keyboard.write(meeting_id)
                keyboard.press('enter')
                time.sleep(5)
                keyboard.write(paswd)
                keyboard.press('enter')


            df = pd.read_csv('timings.csv')
            sign=True
            while sign==True:
                # checking of the current time exists in our csv file
                now = datetime.now().strftime("%H:%M")
                if now in str(df['timings']):
                    row = df.loc[df['timings'] == now]
                    m_id = str(row.iloc[0,1])
                    m_pswd = str(row.iloc[0,2])

                    sign_in(m_id, m_pswd)
                    print('signed in')
                    sign=False


        elif 'flipkart' in query.lower ():
            so = query
            da = so.split ()
            abc = da [ 2 ]
            uv = 'https://www.flipkart.com/search?q='
            suv = uv + abc + '&as=on&as-show=on&otracker=AS_Query_TrendingAutoSuggest_5_0_na_na_na&otracker1=AS_Query_TrendingAutoSuggest_5_0_na_na_na&as-pos=5&as-type=TRENDING&suggestionId=' + abc + '&requestId=1569eb47-6bd0-43a0-a215-0060ab9079dd'
            webbrowser.open ( suv )
        elif 'amazon' in query.lower ():
            gf = query
            df = gf.split ()
            ff = df [ 2 ]
            vur = 'https://www.amazon.in/s?k='
            vus = vur + ff + '&ref=nb_sb_noss_2'
            webbrowser.open ( vus )

        elif 'shut down' in query.lower ():
            speak ( 'bye sir' )
            os.system ( "shutdown /s /t 1" )
        elif 'restart' in query.lower ():
            speak ( 'bye sir' )
            os.system ( "shutdown /r /t 1" )
        elif 'lock' in query.lower ():
            ctypes.windll.user32.LockWorkStation ()
        elif 'news' in query.lower ():
            def News():
                main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=4dbc17e007ab436fb66416009dfb59a8"
                page = requests.get ( main_url ).json ()
                art = page [ "articles" ]
                res = [ ]
                for ar in art:
                    res.append ( ar [ "title" ] )
                for i in range ( len ( res ) ):
                    print ( i + 1 , res [ i ] )

                speak = Dispatch ( "SAPI.Spvoice" )
                speak.Speak ( res )


            if __name__ == '__main__':
                News ()
        elif 'empty recycle bin' in query:
            winshell.recycle_bin ().empty ( confirm=False , show_progress=False , sound=True )
            speak ( "Recycle Bin Recycled" )
        elif "go to sleep" in query.lower ():
            speak ( "ok sir" )
            time.sleep ( 10 )
            sn=input('press a to awake')
            if sn=='a':
                a=takeCommand()

        elif 'what is the weather' in query.lower ():
            api_adress = 'https://api.openweathermap.org/data/2.5/weather?appid=0c42f7f6b53b244c78a418f4f181282a&q='
            city = query
            sp = city.split ()
            a = sp [ 5 ]
            url = api_adress + a.capitalize ()
            resp = requests.get ( url )
            x = resp.json ()
            speak ( 'Sir what do you want to know : humidity weather wind speed or temperature' )
            want = takeCommand ()
            if want == 'weather':
                y = x [ 'weather' ]
                c_w = y [ 0 ] [ 'description' ]
                speak ( f'sir the weather {c_w}' )
            elif want == 'temperature':
                z = x [ 'main' ]
                c_t = z [ 'temp' ]
                c = c_t - 273.15
                speak ( f'sir the temperature {c}' )
            elif want == 'humidity':
                n = x [ 'main' ]
                c_h = n [ 'humidity' ]
                speak ( f'sir the humidity is {c_h}' )
            elif want == 'wind speed':
                m = x [ 'wind' ]
                c_s = m [ 'speed' ]
                speak ( f'sir the speed is {c_s}' )

        elif ' gmail' in query.lower ():
            speak ( "Ok sir" )
            webbrowser.open ( "https://gmail.com/" )
        elif ' open whatsapp' in query.lower ():
            search = query.lower ()
            speak ( "Ok sir" )
            webbrowser.open ( 'https://web.whatsapp.com/' )
        elif 'play music' in query.lower ():
            speak ( "Ok sir" )
            songs_dir = "C:/Users/Jay/Music/Playlists"
            songs = os.listdir ( songs_dir )
            os.startfile ( os.path.join ( songs_dir , songs [ 4 ] ) )
        elif 'stop music' in query.lower ():
            speak ( "Ok sir" )
            quit ( os.startfile )
        elif 'hi' in query.lower ():
            a = ('Ola' , 'Hello' , 'Hi' , 'Yo')
            b = random.choice ( a )
            speak ( f'{b} Sir' )
        elif 'the time' in query.lower ():
            time = datetime.datetime.now ().strftime ( '%H:%M:%S' )
            print ( time )
            speak ( f'Sir the time is {time}' )
        elif 'joke' in query.lower ():
            speak ( pyjokes.get_joke () )

        elif 'prime videos' in query.lower ():
            speak ( "This is what I got for you sir" )
            webbrowser.open ( 'https://www.primevideo.com/region/eu' )
        elif 'exit' in query.lower ():
            speak ( "Bye sir" )
            quit ()
        elif 'alarm' in query.lower ():
            def sound():
                mixer.init ()
                mixer.music.load ( 'J:/short-alarm-clock-sound.mp3' )


            def alarm():
                speak ( 'sir please enter th hour' )
                hor = int ( input ( 'Enter The hour- ' ) )
                speak ( 'sir please enter the minute' )
                minn = int ( input ( 'Enter the min- ' ) )
                speak ( 'sir please enter the second' )
                sec = int ( input ( 'Enter the sec- ' ) )
                n = 5
                print ( 'alarm set for' , str ( hor ) , str ( minn ) , str ( sec ) )
                speak ( f'alarm set for {str ( hor )} hour  {str ( minn )} minute   {str ( sec )} second' )
                while True:
                    if time.localtime ().tm_hour == hor and time.localtime ().tm_min == minn and time.localtime ().tm_sec:
                        print ( 'wake up' )
                        speak ( 'wake up sir' )
                        speak(f'it is {str ( hor )} hour  {str ( minn )} minute   {str ( sec )} second')
                        api_adress = 'https://api.openweathermap.org/data/2.5/weather?appid=0c42f7f6b53b244c78a418f4f181282a&q='
                        city = query
                        sp = city.split ()
                        a = sp [ 5 ]
                        url = api_adress + a.capitalize ()
                        resp = requests.get ( url )
                        x = resp.json ()
                        z = x [ 'main' ]
                        c_t = z [ 'temp' ]
                        c = c_t - 273
                        speak ( f'sir the temperature  outside is {c} degrees celsius' )
                        y = x [ 'weather' ]
                        c_w = y [ 0 ] [ 'description' ]
                        speak ( f'sir it is a nice {c_w} day' )
                        time.sleep(5)
                        for i in range(10):
                            speak('sir you should wake up')
                            time.sleep(5)



                        break
                sound ()
                while n > 0:
                    mixer.music.play ()
                    time.sleep ( 2 )
                    n = n - 1
                sn = takeCommand ()
                if 'sleep' in sn:
                    speak ( 'as you wish sir' )
                    return
                else:
                    return


            alarm ()

        elif "google" in query.lower ():
            speak ( "Searching for Google" )
            search = query
            result = kit.search ( search )
            print ( result )
            speak ( result )

    except Exception as e:
        pass

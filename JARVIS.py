from win32com.client import Dispatch
import time
import random
import speech_recognition as sr
import webbrowser as web
import wikipedia
import os
import datetime
import pyjokes
engine = Dispatch('Sapi.Spvoice')
def speak(stirng:str):
    engine.speak(stirng)

def printNspeak(string:str):
    print(string)
    speak(string)

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour > 6 and hour < 12:
        speak('Good Morning Sir')
    elif hour > 12 and hour < 18:
        speak('Good Afternoon Sir')
    else:
        speak('Good Evening Sir')
    speak('I Am JARVIS Please Tell Me How Can I Help You ?')

def voice2string()->str:
    r = sr.Recognizer()
    r.energy_threshold = 6000
    r.pause_threshold = 2
    with sr.Microphone() as source:
        print('Listening...')
        audio = r.listen(source)
    try:
        print('Recognizing...')
        query = r.recognize_google(audio,language='en-in')
        print(f'User Said : {query}\n')
    except:
        printNspeak('Sir Please Say That Again')
        return 'None'
    return query

if __name__ == "__main__":
    wishMe()
    quitL = ['quit','exit','close']
    while True:
        query = voice2string().lower()
        if 'search' in query:
            printNspeak('Searching In Google...')
            query = query.replace('search','')
            web.open(f'https://google.com/search?q={query}')
        elif 'wikipedia' in query:
            query = query.replace('wikipedia','')
            try:
                printNspeak('Searching In Wikipedia...')
                wiki = wikipedia.summary(query,sentences=2)
                printNspeak(wiki)
            except:
                printNspeak('Page Not Found')
        elif 'browse' in query:
            query = query.replace('browse ','')
            printNspeak(f'Opening {query}...')
            web.open(f'www.{query}')
        elif 'open youtube' in query:
            web.open('www.youtube.com')
        elif 'open stackoverflow' in query:
            web.open('www.stackoverflow.com')
        elif 'open instagram' in query:
            web.open('www.instagram.com')
        elif 'open facebook' in query:
            web.open('www.facebook.com')
        elif 'open github' in query:
            web.open('www.github.com')
        elif 'open amazon' in query:
            web.open('www.amazon.in')
        elif 'joke' in query or 'jokes' in query:
            printNspeak(pyjokes.get_joke(category='all'))
        elif 'play song' in query or 'play music' in query:
            music_dir = r'C:\Users\MUDRA\Music'
            dir_contents = os.listdir(music_dir)
            ext = ['.mp3','.wav']
            songs = [file for file in dir_contents if os.path.splitext(file)[1].lower() in ext]
            random.shuffle(songs)
            printNspeak('Playing A Song...')
            os.startfile(os.path.join(music_dir,songs[0]))
        elif 'open code' in query:
            codePath = r'C:\Users\MUDRA\AppData\Local\Programs\Microsoft VS Code\Code.exe'
            printNspeak('Opening Visual Studio Code...')
            os.startfile(codePath)
        elif 'open chrome' in query:
            chromePath = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            printNspeak('Opening Google Chrome...')
            os.startfile(codePath)
        elif "don't listen" in query or "stop listening" in query:
            printNspeak('How Long Time I Stop Listening ???')
            try:
                sleepTime = int(input('Sir Please Give Time In Seconds\n'))
                printNspeak('OK Sir I Will Not Listen...')
                time.sleep(sleepTime)
                printNspeak('OK Sir I Now I Can Listen...')
            except:
                printNspeak('Sorry Sir You Entered A Invalid Input !!!')
        elif 'clear screen' in query:
            os.system('cls')
            speak('Sir I Cleared Screen')
        elif 'dice' in query:
            printNspeak(f'Sir Here Is Number {random.randint(1,6)}')
        elif 'open service' in query or 'open services' in query:
            printNspeak('Opening Services.msc ...')
            servicePath = r'C:\Windows\System32\services.msc'
            os.startfile(servicePath)
        elif 'command' in query:
            printNspeak('Sir,What Is Your Command ...')
            cmd = input(f'{os.getcwd()}>')
            os.system(cmd)
        elif query in quitL: 
            quit()
        elif 'none' == query:
            pass
        else:
            printNspeak('Sir Sorry To Say That I Cannot Do That You Said')
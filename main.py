import speech_recognition as sr
import webbrowser
import datetime
import os
import win32com.client

from pywikihow import search_wikihow

import requests
from bs4 import BeautifulSoup

import psutil

import Myalarm

speaker = win32com.client.Dispatch("SAPI.SpVoice")

'''
def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for: {prompt}\n\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    # print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)
'''


def takeCommand():
    print('Listening...')
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            # Language can be changed
            command = r.recognize_google(audio, language="en-in")
            print(f"User said: {command}")
            return command
        except Exception as e:
            print("Sorry! Some error occurred.")
            # speaker.Speak("Sorry! Some error occurred.")
            return "Sorry! Some error occurred."


if __name__ == '__main__':
    print("Hello I am Jarvis AI, How may I assist you?")
    speaker.Speak("Hello I am Jarvis AI, How may I assist you?")
    while True:
        query = takeCommand()

        # Opening various sites online -> Other websites can be added
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"],
                 ["google", "https://google.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} Sir...")
                webbrowser.open(site[1])

        # Opening music -> Make list for music files
        if "open music".lower() in query.lower():
            musicPath = "C:/Users/Ayush/Downloads/racing_into_the_night.mp3"
            os.system(f"start {musicPath}")

        # Telling the current time
        elif "time".lower() in query.lower():
            strTime = datetime.datetime.now().strftime("%H:%M")
            print(f"Sir the time is {strTime}")
            speaker.Speak(f"Sir the time is {strTime}")

        # telling the temperature
        elif "temperature".lower() in query.lower():
            search = "temperature in delhi"
            url = f"https://www.google.com/search?q={search}"
            r = requests.get(url)
            data = BeautifulSoup(r.text, "html.parser")
            temp = data.find("div", "BNeawe").text
            print(f"Current {search} is {temp}")
            speaker.Speak(f"Current {search} is {temp}")

        # Activating how to do mode
        elif "how to do mode".lower() in query.lower():
            speaker.Speak("How to do mode is activated")
            while True:
                speaker.Speak("Please tell me what do you want to know?")
                how = takeCommand()
                try:
                    if "close" in how:
                        speaker.Speak("How to do mode is closed")
                        break
                    else:
                        max_results = 1
                        how_to = search_wikihow(how, max_results)
                        assert len(how_to) == 1
                        how_to[0].print()
                        speaker.Speak(how_to[0].summary)
                except Exception as e:
                    speaker.Speak("Sorry sir, I cannot find this")

        # Finding out the current battery percentage
        elif "battery" in query.lower():
            battery = psutil.sensors_battery()
            percentage = battery.percent
            print(f"Sir, our system has {percentage} percent power left")
            speaker.Speak(f"Sir, our system has {percentage} percent power left")
            if percentage >= 60:
                print("We have enough power to continue our work")
                speaker.Speak("We have enough power to continue our work")
            elif percentage >= 30:
                print("Please plug in the system to charging")
                speaker.Speak("Please plug in the system to charging")
            else:
                print("The system may shut down soon")
                speaker.Speak("The system may shut down soon")

        # Make a list for apps
        elif "open notepad".lower() in query.lower():
            npath = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(npath)

        # Setting an alarm
        elif "alarm" in query.lower():
            print("Sir please tell the time to set the alarm for. For example 'set alarm to 5:30 am'")
            speaker.Speak("Sir please tell the time to set the alarm for. For example 'set alarm to 5:30 am'")
            tt = takeCommand()
            tt = tt.replace("set alarm to ", "")
            tt = tt.replace(".", "")
            tt = tt.upper()
            Myalarm.alarm(tt)

        # Closing the program
        elif "sleep".lower() in query.lower():
            speaker.Speak("Bye sir, Have a good day!")
            exit()

        # speaker.Speak(query)

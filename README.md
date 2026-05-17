# jarvish-voice-assistence
Overview

JARVIS is a Python-based Virtual Assistant that performs voice-controlled tasks such as opening applications, searching the web, telling the time, answering queries, and automating basic desktop activities. The project is inspired by the AI assistant shown in Iron Man movies.

This project was developed to improve practical knowledge of Python, speech recognition, automation, and AI-based interaction systems.

Features
Voice recognition using microphone input
Text-to-speech response system
Open websites and applications using voice commands
Google and YouTube search
Tell current time and date
Play music
Perform system-related tasks
User-friendly interaction
Technologies Used
Python
SpeechRecognition
pyttsx3
pywhatkit
datetime
webbrowser
os module
Installation
1. Clone the Repository
git clone https://github.com/your-username/jarvis-virtual-assistant.git
2. Navigate to Project Folder
cd jarvis-virtual-assistant
3. Install Required Libraries
pip install -r requirements.txt
Required Libraries
speechrecognition
pyttsx3
pywhatkit
pyaudio
wikipedia
Run the Project
python main.py
Example Commands
"Open YouTube"
"Search Python tutorial"
"What is the time?"
"Play music"
"Open Google"
Future Improvements
Add ChatGPT API integration
Add GUI interface
Add face recognition
Add smart home automation
Add multilingual support
Learning Outcomes

Through this project, we learned:

Python automation
Speech processing
Voice assistant architecture
API integration basics
Real-time command handling
Contributors
Suthar Pradip
Jigar Chavda
Mali Mehul
Vataliya Dhruvit
Vanol Dhruvitra
License

This project is developed for educational purposes only.

Contact

Email: pradipsuthar9429@gmail.com

GitHub: Add your GitHub profile link here

JARVIS Virtual Assistant
Overview

JARVIS is a Python-based AI Virtual Assistant designed to perform voice-controlled and automation tasks. Initially, the project used the SpeechRecognition library, but due to accuracy and microphone detection issues, we upgraded the system by integrating a more reliable speech processing library for better voice recognition performance and faster response time.

This project helped us understand real-world implementation of voice assistants, automation, and AI interaction systems.

Features
Voice command recognition
Text-to-speech response system
Open websites and desktop applications
Search information on Google and YouTube
Tell current date and time
Play music and videos
Smart command execution
Improved voice recognition accuracy
Faster response compared to the previous version
Technologies Used
Python
pyttsx3
pywhatkit
webbrowser
datetime
os module
Advanced Speech Recognition Library
Automation modules
Why We Changed the Speech Recognition System

Initially, we used the SpeechRecognition library, but it had several limitations:

Poor voice detection in noisy environments
Slow response time
Microphone compatibility issues
Frequent command recognition errors

To improve performance and user experience, we integrated a newer and more efficient voice recognition library that provides:

Better accuracy
Faster speech processing
Improved microphone handling
More stable real-time recognition
Project Structure
Jarvis-Virtual-Assistant/
│
├── main.py
├── requirements.txt
├── README.md
├── modules/
│   ├── assistant.py
│   ├── commands.py
│   ├── speech_engine.py
│   └── automation.py
│
└── assets/
    └── screenshots/
Installation
Clone the Repository
git clone https://github.com/your-username/jarvis-virtual-assistant.git
Move to Project Folder
cd jarvis-virtual-assistant
Install Dependencies
pip install -r requirements.txt
Required Libraries
pyttsx3
pywhatkit
pyaudio
vosk
sounddevice
wikipedia
Run the Project
python main.py
Example Voice Commands
“Open YouTube”
“Search Python tutorial”
“Play music”
“What is the time?”
“Open Google”
“Tell me today’s date”
Future Enhancements
ChatGPT API integration
GUI dashboard
Face recognition login
Home automation support
Mobile application version
Multilingual voice support
Learning Outcomes

This project helped us learn:

Python automation
Voice processing systems
AI assistant workflow
Real-time speech recognition
Error handling and optimization
Team collaboration in software development
Contributors
Suthar Pradip
Jigar Chavda
Mali Mehul
Vataliya Dhruvit
Vanol Dhruvitra
License

This project is developed for educational and learning purposes only.

Contact

Email: pradipsuthar9429@gmail.com

GitHub: Add your GitHub repository link here

import speech_recognition as sr
import webbrowser
import win32com.client
import time
import requests
import wikipedia
import musiclibrary

# --------------------
# Speech Engine
# --------------------

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print("Jarvis:", text)
    speaker.Speak(text)


# --------------------
# Setup
# --------------------

recognizer = sr.Recognizer()

newsapi = "02c2d21b767141f2bcc65edf0de3093f"

wikipedia.set_lang("en")


# --------------------
# Listen Function
# --------------------

def listen():

    with sr.Microphone() as source:

        print("Listening...")

        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        recognizer.pause_threshold = 1
        recognizer.energy_threshold = 300

        try:
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=6)
        except:
            return None

    try:
        text = recognizer.recognize_google(audio)
        print("Heard:", text)
        return text.lower()

    except sr.UnknownValueError:
        print("Could not understand")
        return None

    except sr.RequestError:
        speak("Speech recognition service unavailable")
        return None


# --------------------
# Command Processor
# --------------------

def process_command(command):

    if command is None:
        return

    # OPEN WEBSITES
    if "open google" in command:
        speak("Opening Google")
        webbrowser.open("https://google.com")

    elif "open youtube" in command:
        speak("Opening YouTube")
        webbrowser.open("https://youtube.com")

    elif "open linkedin" in command:
        speak("Opening LinkedIn")
        webbrowser.open("https://linkedin.com")

    elif "open whatsapp" in command:
        speak("Opening WhatsApp")
        webbrowser.open("https://web.whatsapp.com")


    # PLAY MUSIC
    elif command.startswith("play"):

        song = command.replace("play", "").strip().lower()

        if song in musiclibrary.music:
            speak("Playing " + song)
            webbrowser.open(musiclibrary.music[song])
        else:
            speak("Song not found in library")


    # NEWS
    elif "news" in command:

        speak("Here are the latest headlines")

        try:

            r = requests.get(
                f"https://newsapi.org/v2/top-headlines?country=us&apiKey={newsapi}"
            )

            data = r.json()

            for article in data["articles"][:5]:
                speak(article["title"])

        except:
            speak("Unable to fetch news")


    # WIKIPEDIA
    elif "who is" in command or "what is" in command:

        query = command.split("is")[-1].strip()

        try:

            result = wikipedia.summary(query, sentences=1)

            speak(result)

        except wikipedia.DisambiguationError as e:

            result = wikipedia.summary(e.options[0], sentences=1)

            speak(result)

        except:
            speak("I could not find information about that")


    # WEATHER
    elif "weather" in command:

        speak("Opening weather forecast")

        webbrowser.open("https://www.google.com/search?q=weather")


    # EXIT
    elif "stop" in command or "exit" in command or "shutdown" in command:

        speak("Shutting down")

        exit()


    # FALLBACK SEARCH
    else:

        speak("Searching Google")

        webbrowser.open(f"https://www.google.com/search?q={command}")


# --------------------
# Main Loop
# --------------------

def main():

    speak("Initializing Jarvis")

    while True:

        wake = listen()

        if wake and ("jarvis" in wake or "service" in wake):

            speak("Yes, how can I help?")

            for _ in range(8):

                command = listen()

                if command is None:
                    continue

                process_command(command)

        time.sleep(0.5)


# --------------------

if __name__ == "__main__":

    try:
        main()

    except KeyboardInterrupt:

        speak("Jarvis shutting down")

Your code is actually decent for a beginner-to-intermediate virtual assistant project. The structure is clean, commands are separated properly, and you already implemented:

Voice recognition
TTS
News API
Wikipedia
Browser automation
Music system

That is already stronger than most basic college projects.

But there are some important problems in your current code that you should fix before uploading to GitHub professionally.

Problems in Current Code
1. API Key Exposed

This is a mistake:

newsapi = "02c2d21b767141f2bcc65edf0de3093f"

Never upload API keys publicly.

Use:

import os

newsapi = os.getenv("NEWS_API_KEY")

Then create .env file.

2. Too Many Bare except:

This is bad practice:

except:
    return None

Because it hides real errors.

Better:

except Exception as e:
    print(e)
    return None
3. SpeechRecognition Delay Problem

You were correct earlier.

recognize_google() is:

internet dependent
slower
unstable sometimes

That is why many developers move to:

vosk
whisper
faster-whisper

This is actually a smart improvement, not a small change.

Your README Should Mention This Properly

Add this section:

Improved Speech Recognition System

Initially, the assistant used the SpeechRecognition library with Google Speech API for voice processing. However, it created issues such as:

Slow response time
Internet dependency
Recognition failures in noisy environments
Microphone instability

To improve performance and reliability, the system was upgraded with a more advanced speech processing approach, resulting in:

Faster recognition
Better accuracy
Improved real-time interaction
More stable command processing
Strong Points in Your Project

Your project already includes:

Modular command handling
Voice activation ("Jarvis")
External APIs
Browser automation
Dynamic searching
Music integration

This is enough for a strong beginner GitHub project.

What You Should Improve Next
Priority Improvements
1. Convert into Modules

Right now everything is inside one file.

Professional structure:

Jarvis/
│
├── main.py
├── commands.py
├── speech.py
├── config.py
├── utils.py
└── musiclibrary.py

This matters a lot for resume quality.

2. Add GUI

Use:

Tkinter (easy)
CustomTkinter (best for modern UI)

Even a simple animated interface increases project quality massively.

3. Add Offline AI

Best upgrade:

vosk
whisper

This makes your project more unique.

4. Add Hotword Detection

Instead of continuously recording.

Use:

Porcupine
Vosk wake word
openWakeWord
5. Add ChatGPT Integration

This can make your project resume-worthy for AI/ML internships.

One More Important Thing

Your current code is automation-focused, not true AI.

Do not write:

"AI-powered advanced intelligent assistant"

That would be exaggerated.

Correct description:

"Python-based voice-controlled virtual assistant with speech recognition and automation capabilities."

That sounds professional and honest.

you give me of this readme file
JARVIS Virtual Assistant
Overview

JARVIS is a Python-based voice-controlled virtual assistant developed to perform automation and daily productivity tasks using speech commands. The assistant can open websites, search information, play music, provide news updates, answer questions using Wikipedia, and perform browser-based automation.

The project was designed to improve practical knowledge of Python automation, speech recognition, APIs, and real-time command processing.

Features
Voice-controlled interaction
Text-to-speech response system
Open websites and web applications
Google search automation
YouTube automation
Music playback system
Latest news headlines using News API
Wikipedia search support
Weather search functionality
Wake word activation ("Jarvis")
Real-time speech command processing
Technologies Used
Python
SpeechRecognition
pyttsx3 / SAPI Voice Engine
Wikipedia API
News API
webbrowser module
requests module
win32com.client
Improved Speech Recognition System

Initially, the project used the SpeechRecognition library with Google Speech API for speech processing. However, several issues were identified during testing:

Slow response time
Internet dependency
Recognition failures in noisy environments
Microphone detection instability

To improve the overall performance and user experience, the speech processing system was optimized with better microphone handling, ambient noise adjustment, and improved command recognition logic.

Future versions of the project aim to integrate advanced offline speech recognition systems such as:

Vosk
Whisper AI
Faster-Whisper

This will improve:

Recognition accuracy
Real-time processing speed
Offline functionality
Stability of voice interaction
Project Structure
Jarvis-Virtual-Assistant/
│
├── main.py
├── musiclibrary.py
├── requirements.txt
├── README.md
│
└── assets/
    └── screenshots/
Installation
1. Clone the Repository
git clone https://github.com/your-username/jarvis-virtual-assistant.git
2. Navigate to Project Folder
cd jarvis-virtual-assistant
3. Install Dependencies
pip install -r requirements.txt
Required Libraries
SpeechRecognition
pywin32
requests
wikipedia
pyaudio
Run the Project
python main.py
Example Voice Commands
"Open Google"
"Open YouTube"
"Play Believer"
"What is Artificial Intelligence"
"Who is Elon Musk"
"Tell me the news"
"Open LinkedIn"
"Show weather"
Future Improvements
GUI-based interface
Offline speech recognition
ChatGPT integration
Smart desktop automation
Face recognition authentication
AI conversation memory
Mobile application support
Multilingual support
Learning Outcomes

This project helped us understand:

Python automation
Speech recognition systems
API integration
Voice assistant architecture
Browser automation
Real-time command handling
Error handling and optimization
Contributors
Suthar Pradip
Jigar Chavda
Mali Mehul
Vataliya Dhruvit
Vanol Dhruvitra
License

This project is developed for educational and learning purposes only.

Contact

Email: pradipsuthar9429@gmail.com

GitHub: Add your GitHub repository link here

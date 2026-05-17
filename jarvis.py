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
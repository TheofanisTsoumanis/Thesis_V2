
import speech_recognition as sr
import time


def Select_Language(): 
    rec = sr.Recognizer()
    rec.pause_threshold = 0.7
    rec.energy_threshold = 300
    try:
        with sr.Microphone() as mic:

            rec.adjust_for_ambient_noise(mic, duration = 1)
            print("Ενεργό μικρόφωνο / Active mic:")
            audio = rec.listen(mic, timeout = 4, phrase_time_limit = 10)
            
            text_gr = str(rec.recognize_google(audio, language = "el-GR"))
            text_en = str(rec.recognize_google(audio, language = "en-US"))
            text_gr = text_gr.capitalize()
            text_en = text_en.capitalize()
            return  text_gr, text_en
            
    except Exception as e:
        print(f"{type(e).__name__}: {e}\n")
        return None, None

def Speech(lang):
    rec = sr.Recognizer()
    rec.pause_threshold = 1.5

    try:
        with sr.Microphone() as mic:

            rec.adjust_for_ambient_noise(mic, duration = 0.5)
            audio = rec.listen(mic, timeout = 5, phrase_time_limit = 13)

            text = str(rec.recognize_google(audio, language = lang))

            text = text.capitalize()
            print(text)

            return text
        
    except Exception as e:
        print(f"{type(e).__name__}: {e}\n")
        return None



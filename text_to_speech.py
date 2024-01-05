import os

import pygame
from gtts import gTTS
from playsound import playsound

def text_to_speech(text, language='en', filename='output.mp3'):
    try:
        tts = gTTS(text=text, lang=language, slow=False)
        tts.save(filename)
        playsound(filename)
        os.remove(filename)
    except Exception as e:
        print(e)

import whisper
import pyautogui
import pyttsx3
import webbrowser
import sys
import time
import threading
import numpy as np
import wave
import pyaudio
import os
import re

# Setup
pyautogui.FAILSAFE = True
engine = pyttsx3.init()
engine.setProperty('rate', 150)

print("Loading Whisper model...")
model = whisper.load_model("small")

# Websites
WEBSITES = {
    "youtube": "https://www.youtube.com",
    "twitter": "https://twitter.com",
    "x": "https://x.com",
    "instagram": "https://www.instagram.com",
    "messenger": "https://www.messenger.com",
    "gmail": "https://mail.google.com",
    "google": "https://www.google.com",
}

# Movement state
moving = False
move_thread = None
stop_event = threading.Event()
current_speed = 12

# Audio settings
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 16000
p = pyaudio.PyAudio()

# Spoken numbers
number_words = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    "thirty": 30, "forty": 40, "fifty": 50
}

def word_to_number(word):
    return number_words.get(word.lower())

def speak(text):
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def record_audio():
    stream = p.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK)
    print("🔴 Listening (peak >2000 required)...")
    frames = []
    
    voice_threshold = 2000
    silence_threshold = 1500
    silence_count = 0
    max_silence_frames = 30
    min_speech_frames = 40
    peak_volume = 0

    for _ in range(int(RATE / CHUNK * 20)):
        data = stream.read(CHUNK, exception_on_overflow=False)
        frames.append(data)
        
        audio_np = np.frombuffer(data, dtype=np.int16)
        volume = np.max(np.abs(audio_np))
        peak_volume = max(peak_volume, volume)
        
        print(f"\rCurrent: {volume:<5} | Peak: {peak_volume:<5}", end="")
        
        if volume < silence_threshold:
            silence_count += 1
        else:
            silence_count = 0
        
        if silence_count > max_silence_frames and len(frames) > min_speech_frames:
            print("\n✅ Speech ended")
            break

    stream.stop_stream()
    stream.close()

    if peak_volume < voice_threshold:
        print(f"\n❌ Ignored background (peak {peak_volume})")
        return None

    print(f"\n🎤 Voice detected (peak {peak_volume})")
    temp_file = "temp.wav"
    wf = wave.open(temp_file, 'wb')
    wf.setnchannels(CHANNELS)
    wf.setsampwidth(p.get_sample_size(FORMAT))
    wf.setframerate(RATE)
    wf.writeframes(b''.join(frames))
    wf.close()
    return temp_file

def listen():
    audio_file = record_audio()
    if audio_file is None:
        return ""
    
    result = model.transcribe(audio_file, language="en", fp16=False)
    text = result["text"].strip().lower()
    
    if len(text.split()) < 2 or text in ["thank you", "thanks", "music"]:
        print(f"🚫 Filtered junk: '{text}'")
        os.remove(audio_file)
        return ""
    
    print(f"🎯 Command: '{text}'")
    os.remove(audio_file)
    return text

# Normal mouse movement
def continuous_move(dx, dy):
    global moving
    stop_event.clear()
    moving = True
    while not stop_event.is_set():
        pyautogui.moveRel(dx, dy, duration=0.1)
        time.sleep(0.05)
    moving = False

def start_normal_move(direction):
    global moving, move_thread, current_speed
    current_speed = 12
    if moving:
        stop_all()
    
    speed = current_speed
    dx, dy = 0, 0
    if direction == "left": dx = -speed
    elif direction == "right": dx = speed
    elif direction == "up": dy = -speed
    elif direction == "down": dy = speed
    else: return
    
    speak(f"Moving {direction} continuously")
    moving = True
    stop_event.clear()
    move_thread = threading.Thread(target=continuous_move, args=(dx, dy))
    move_thread.daemon = True
    move_thread.start()

def stop_all():
    global moving
    if moving:
        stop_event.set()
        moving = False
        speak("Movement stopped")

def execute_command(cmd):
    global current_speed
    if not cmd:
        return
    
    # === HIGHLIGHT HAS HIGHEST PRIORITY ===
    if "highlight" in cmd:
        if "right" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'right')
            speak("Highlighted word right")
            return
        if "left" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'left')
            speak("Highlighted word left")
            return
        if "down" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'down')
            speak("Highlighted line down")
            return
        if "up" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'up')
            speak("Highlighted line up")
            return
    
    # Select All
    if "select all" in cmd:
        pyautogui.hotkey('ctrl', 'a')
        speak("Selected all")
        return
    
    # Normal mouse move (only if no highlight)
    if "left" in cmd:
        start_normal_move("left")
        return
    if "right" in cmd:
        start_normal_move("right")
        return
    if "up" in cmd:
        start_normal_move("up")
        return
    if "down" in cmd:
        start_normal_move("down")
        return
    
    # Stop
    if "stop" in cmd:
        stop_all()
        return
    
    # Space
    if "space" in cmd:
        pyautogui.press("space")
        speak("Space added")
        return
    
    # Punctuation
    if "question mark" in cmd:
        pyautogui.typewrite("?")
        speak("Question mark")
        return
    if "exclamation mark" in cmd or "exclamation point" in cmd:
        pyautogui.typewrite("!")
        speak("Exclamation mark")
        return
    
    # Delete left (Backspace)
    left_delete_variants = ["delete", "remove", "believe", "delight", "elite", "delayed", "the lead", "delhi"]
    left_delete_match = re.search(r"(" + "|".join(left_delete_variants) + r")\s+((?:\w+\s)?\w+)\s*times?", cmd)
    if left_delete_match:
        number_part = left_delete_match.group(2).strip()
        count = word_to_number(number_part) or (int(number_part) if number_part.isdigit() else 1)
        count = max(1, min(50, count))
        pyautogui.press(["backspace"] * count)
        speak(f"Deleted left {count} times")
        return
    
    if any(v in cmd for v in left_delete_variants):
        pyautogui.press("backspace")
        speak("Deleted left")
        return
    
    # Click
    if "click" in cmd:
        if "double" in cmd:
            pyautogui.doubleClick()
            speak("Double clicked")
        elif "right" in cmd:
            pyautogui.rightClick()
            speak("Right clicked")
        else:
            pyautogui.click()
            speak("Clicked")
        return
    
    # Enter
    if "enter" in cmd or "send" in cmd:
        pyautogui.press("enter")
        speak("Enter pressed")
        return
    
    # Scroll
    if "scroll down" in cmd:
        pyautogui.scroll(-600)
        speak("Scrolled down")
    elif "scroll up" in cmd:
        pyautogui.scroll(600)
        speak("Scrolled up")
    
    # Open
    if "open" in cmd:
        for site in WEBSITES:
            if site in cmd:
                webbrowser.open(WEBSITES[site])
                speak(f"Opening {site}")
                return
    
    # Type
    if "type" in cmd or "message" in cmd:
        parts = cmd.split(" ", 1)
        text = parts[1] if len(parts) > 1 else ""
        pyautogui.typewrite(text)
        speak("Typed")
        return
    
    speak("Try: highlight right | right | stop | select all | click | enter")

# Start
speak("Priority fixed! 'highlight right' now works correctly — no more normal move.")
print("\n=== PRIORITY FIXED ===\n")
print("Say 'highlight right' → highlights word | 'right' → normal move")

while True:
    cmd = listen()
    execute_command(cmd)
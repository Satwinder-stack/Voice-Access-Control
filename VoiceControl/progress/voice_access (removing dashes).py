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

# Speed & Scroll settings
NORMAL_SPEED = 12
SLOW_SPEED = 3
SCROLL_AMOUNT = 600

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
    word = word.lower()
    if word.isdigit():
        return int(word)
    return number_words.get(word)

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

# Timed continuous movement
def timed_move(dx, dy, duration):
    global moving
    stop_event.clear()
    moving = True
    end_time = time.time() + duration
    
    while time.time() < end_time and not stop_event.is_set():
        pyautogui.moveRel(dx, dy, duration=0.1)
        time.sleep(0.05)
    
    moving = False

# Continuous (infinite) movement
def continuous_move(dx, dy):
    global moving
    stop_event.clear()
    moving = True
    while not stop_event.is_set():
        pyautogui.moveRel(dx, dy, duration=0.1)
        time.sleep(0.05)
    moving = False

def start_move(direction, speed=NORMAL_SPEED, duration=None):
    global moving, move_thread
    
    if moving:
        stop_all()
    
    dx, dy = 0, 0
    if direction == "left":
        dx = -speed
    elif direction == "right":
        dx = speed
    elif direction == "up":
        dy = -speed
    elif direction == "down":
        dy = speed
    else:
        return
    
    is_slow = (speed == SLOW_SPEED)
    speed_name = "slowly" if is_slow else "continuously"
    time_info = f" for {duration} seconds" if duration else ""
    
    speak(f"Moving {direction} {speed_name}{time_info}")
    
    moving = True
    if duration:
        move_thread = threading.Thread(target=timed_move, args=(dx, dy, duration))
    else:
        move_thread = threading.Thread(target=continuous_move, args=(dx, dy))
    
    move_thread.daemon = True
    move_thread.start()

def stop_all():
    global moving
    if moving:
        stop_event.set()
        time.sleep(0.1)
        moving = False
        speak("Movement stopped")

def execute_command(cmd):
    if not cmd:
        return
    
    # === HIGHLIGHT / SELECT ===
    highlight_variants = ["highlight", "high light", "hi light", "hilight", "highlite", "select", "mark"]
    if any(v in cmd for v in highlight_variants):
        if "right" in cmd:
            pyautogui.moveRel(60, 0, duration=0.2)
            pyautogui.doubleClick()
            pyautogui.moveRel(-40, 0, duration=0.1)
            speak("Selected next word")
            return
        elif "left" in cmd:
            pyautogui.moveRel(-60, 0, duration=0.2)
            pyautogui.doubleClick()
            pyautogui.moveRel(40, 0, duration=0.1)
            speak("Selected previous word")
            return
        elif "line" in cmd or "paragraph" in cmd or "down" in cmd:
            pyautogui.tripleClick()
            speak("Selected line")
            return
        else:
            pyautogui.doubleClick()
            speak("Selected word")
            return
    
    # Select All
    if "select all" in cmd:
        pyautogui.hotkey('ctrl', 'a')
        speak("Selected all")
        return
    
    # === TAB CONTROL ===
    if "tab" in cmd:
        if "next" in cmd or "right" in cmd:
            pyautogui.hotkey('ctrl', 'tab')
            speak("Next tab")
            return
        if "previous" in cmd or "back" in cmd or "left" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'tab')
            speak("Previous tab")
            return
        if "new" in cmd or "open" in cmd:
            pyautogui.hotkey('ctrl', 't')
            speak("New tab opened")
            return
        if "close" in cmd:
            pyautogui.hotkey('ctrl', 'w')
            speak("Closed tab")
            return
    
    # === VOLUME CONTROL ===
    if "volume" in cmd:
        count = 1
        match = re.search(r"(higher|lower)\s*(?:by\s*)?(\w+)?", cmd)
        if match:
            direction = match.group(1)
            number_part = match.group(2)
            if number_part:
                count = word_to_number(number_part)
                if count is None:
                    count = 1
            count = max(1, min(50, count))
            
            if direction == "higher":
                pyautogui.press(['volumeup'] * count)
                speak(f"Volume higher by {count}")
                return
            elif direction == "lower":
                pyautogui.press(['volumedown'] * count)
                speak(f"Volume lower by {count}")
                return
        
        if "higher" in cmd:
            pyautogui.press('volumeup')
            speak("Volume higher")
            return
        if "lower" in cmd:
            pyautogui.press('volumedown')
            speak("Volume lower")
            return
        if "mute" in cmd:
            pyautogui.press('volumemute')
            speak("Muted")
            return
    
    # === SCROLLING ===
    if "scroll" in cmd:
        count = 1
        match = re.search(r"(up|down)\s*(?:by\s*)?(\w+)?", cmd)
        if match:
            direction = match.group(1)
            number_part = match.group(2)
            if number_part:
                count = word_to_number(number_part)
                if count is None:
                    count = 1
            count = max(1, min(30, count))
        
        if "down" in cmd:
            pyautogui.scroll(-SCROLL_AMOUNT * count)
            speak(f"Scrolled down{f' {count} times' if count > 1 else ''}")
            return
        elif "up" in cmd:
            pyautogui.scroll(SCROLL_AMOUNT * count)
            speak(f"Scrolled up{f' {count} times' if count > 1 else ''}")
            return
    
    # === TIMED OR CONTINUOUS MOVEMENT ===
    if "move" in cmd:
        direction = None
        for d in ["left", "right", "up", "down"]:
            if d in cmd:
                direction = d
                break
        if not direction:
            return
        
        speed = SLOW_SPEED if "slowly" in cmd else NORMAL_SPEED
        
        duration = None
        match = re.search(r"(\w+)\s*seconds?", cmd)
        if match:
            num_str = match.group(1)
            seconds = word_to_number(num_str)
            if seconds:
                duration = seconds
        
        start_move(direction, speed, duration)
        return
    
    # === SIMPLE DIRECTION (continuous normal speed) ===
    if cmd in ["left", "right", "up", "down"]:
        start_move(cmd, NORMAL_SPEED)
        return
    
    # Stop
    if "stop" in cmd:
        stop_all()
        return
    
    # === OTHER COMMANDS ===
    if "space" in cmd:
        pyautogui.press("space")
        speak("Space added")
        return
    
    if "question mark" in cmd:
        pyautogui.typewrite("?")
        speak("Question mark")
        return
    if "exclamation mark" in cmd or "exclamation point" in cmd:
        pyautogui.typewrite("!")
        speak("Exclamation mark")
        return
    
    # Delete
    left_delete_variants = ["delete", "remove", "believe", "delight", "elite", "delayed", "the lead", "delhi"]
    if any(v in cmd for v in left_delete_variants):
        count = 1
        match = re.search(r"(" + "|".join(left_delete_variants) + r")\s+((?:\w+\s)?\w+)\s*times?", cmd)
        if match:
            num_part = match.group(2).strip()
            count = word_to_number(num_part) or (int(num_part) if num_part.isdigit() else 1)
            count = max(1, min(50, count))
        pyautogui.press(["backspace"] * count)
        speak(f"Deleted left {count} times")
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
    
    # Open website
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
    
    speak("Try: next tab | previous tab | new tab | close tab | move right 5 seconds | scroll down ten")

# Start
speak("Tab switching added! Say 'next tab', 'previous tab', or 'new tab'.")
print("\n=== TAB SWITCHING ADDED ===\n")
print("Commands:")
print("• 'next tab' → Ctrl+Tab")
print("• 'previous tab' or 'back tab' → Ctrl+Shift+Tab")
print("• 'new tab' → Ctrl+T")
print("• 'close tab' → Ctrl+W")

while True:
    cmd = listen()
    execute_command(cmd)
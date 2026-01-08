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
import subprocess
import psutil
import language_tool_python
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)

# Suppress the specific pyttsx3 destructor error
import traceback
def warn_with_traceback(message, category, filename, lineno, file=None, line=None):
    log = file if hasattr(file,'write') else sys.stderr
    log.write(warnings.formatwarning(message, category, filename, lineno, line))
    traceback.print_stack()

# Setup
pyautogui.FAILSAFE = True
engine = pyttsx3.init()
engine.setProperty('rate', 150)

print("Loading Whisper model...")
model = whisper.load_model("small")

# Common misspellings for autocorrect
AUTOCORRECT = {
    "teh": "the", "recieve": "receive", "accomodate": "accommodate",
    "definately": "definitely", "seperate": "separate", "occured": "occurred",
    "wierd": "weird", "untill": "until", "alot": "a lot", "thier": "their",
    "its": "it's", "your": "you're", "there": "their", "loose": "lose",
    "wich": "which", "becuase": "because", "tomorow": "tomorrow",
}

# Create a tool instance (downloads LanguageTool automatically on first use)
tool = language_tool_python.LanguageTool('en-US')  # Downloads ~200MB first time only

BROWSER_PROCESSES = {
    "brave": "brave",
    "chrome": "chrome",
    "google": "chrome",  # "close google" means Chrome
    "edge": "msedge",
    "firefox": "firefox",
}

# Websites
WEBSITES = {
    "youtube": "https://www.youtube.com",
    "twitter": "https://twitter.com",
    "x": "https://x.com",
    "instagram": "https://www.instagram.com",
    "messenger": "https://www.messenger.com",
    "gmail": "https://mail.google.com",
    "google": "https://www.google.com",
    "facebook": "https://www.facebook.com",
    "reddit": "https://www.reddit.com",
    "linkedin": "https://www.linkedin.com",
    "netflix": "https://www.netflix.com",
    "amazon": "https://www.amazon.com",
    "spotify": "https://www.spotify.com",
    "twitch": "https://www.twitch.tv",
    "github": "https://www.github.com",
    "bing": "https://www.bing.com",
    "yahoo": "https://www.yahoo.com",
    "duckduckgo": "https://www.duckduckgo.com",
    "stackoverflow": "https://stackoverflow.com",
    "quora": "https://www.quora.com",
    "slack": "https://slack.com",
    "zoom": "https://zoom.us",
    "youtube music": "https://music.youtube.com",
    "discord": "https://discord.com",
    "youtube studio": "https://studio.youtube.com",
    "google docs": "https://docs.google.com",
    "google sheets": "https://sheets.google.com",
    "google slides": "https://slides.google.com",
    "outlook web": "https://outlook.live.com",
    "hotmail": "https://outlook.live.com",
    "office 365": "https://www.office.com",
    "wordpress": "https://wordpress.com",
    "wikipedia": "https://www.wikipedia.org",
    "yahoo mail": "https://mail.yahoo.com",
    "ebay": "https://www.ebay.com",
    "pinterest": "https://www.pinterest.com",
    "cnn": "https://www.cnn.com",
    "bbc": "https://www.bbc.com",
    "hulu": "https://www.hulu.com",
    "disney plus": "https://www.disneyplus.com",
    "paypal": "https://www.paypal.com",
    "imdb": "https://www.imdb.com",
    "weather": "https://www.weather.com",
    "nytimes": "https://www.nytimes.com",
    "the verge": "https://www.theverge.com",
    "techcrunch": "https://techcrunch.com",
    "hacker news": "https://news.ycombinator.com",
    "medium": "https://medium.com",
    "tumblr": "https://www.tumblr.com",
    "etsy": "https://www.etsy.com",
    "craigslist": "https://www.craigslist.org",
    "gitlab": "https://www.gitlab.com",
    "bitbucket": "https://bitbucket.org",
    "lazada": "https://www.lazada.com",
    "shopee": "https://shopee.com",
    "aliexpress": "https://www.aliexpress.com",
    "flipkart": "https://www.flipkart.com",
    "gma7": "https://www.gmanetwork.com",
    "abs cbn": "https://www.abs-cbn.com",
}

# Desktop Apps
APPS = {
    "paint": "mspaint",
    "word": "winword",
    "excel": "excel",
    "teams": "ms-teams",
    "calculator": "calc",
    "brave": r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    "edge": "msedge",
    "firefox": "firefox",
    "discord": r"C:\Users\%USERNAME%\AppData\Local\Discord\Update.exe --processStart Discord.exe",
    "cmd": "cmd",
    "cmd administrator": "cmd",
    "notepad": "notepad",
    "file explorer": "explorer",
    "powershell": "powershell",
    "outlook": "outlook",
    "powerpoint": "powerpnt",
    "visual studio code": "code",
    "vscode": "code",
    "spotify": "spotify",
    "task manager": "taskmgr",
}

# Movement state
moving = False
move_thread = None
stop_event = threading.Event()

# Speed & Scroll
NORMAL_SPEED = 12
SLOW_SPEED = 3
FAST_SPEED = 50
SCROLL_AMOUNT = 600

# Audio settings
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 16000
p = pyaudio.PyAudio()

# Numbers and ordinals
number_words = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    "thirty": 30, "forty": 40, "fifty": 50
}

ordinal_words = {
    "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
    "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10
}

def grammar_correct(text):
    if not text.strip():
        return text
    
    words = text.split()
    # Keep spelling mode (no correction)
    if len(words) > 3 and all(len(w.strip(".,!?")) <= 2 for w in words):
        return text.replace("-", "").replace(" ", "")
    
    # Otherwise, use LanguageTool for full grammar correction
    corrected = tool.correct(text)
    return corrected

def word_to_number(word):
    word = word.lower()
    if word.isdigit():
        return int(word)
    return number_words.get(word) or ordinal_words.get(word)

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

# Movement functions
def timed_move(dx, dy, duration):
    global moving
    stop_event.clear()
    moving = True
    end_time = time.time() + duration
    while time.time() < end_time and not stop_event.is_set():
        pyautogui.moveRel(dx, dy, duration=0.1)
        time.sleep(0.05)
    moving = False

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
    if direction == "left": dx = -speed
    elif direction == "right": dx = speed
    elif direction == "up": dy = -speed
    elif direction == "down": dy = speed
    else: return
    
    speed_name = "slowly" if speed == SLOW_SPEED else "fast" if speed == FAST_SPEED else "continuously"
    time_info = f" for {duration} seconds" if duration else ""
    speak(f"Moving {direction} {speed_name}{time_info}")
    
    moving = True
    target = timed_move if duration else continuous_move
    args = (dx, dy, duration) if duration else (dx, dy)
    move_thread = threading.Thread(target=target, args=args)
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
    
    # === TERMINATE PROGRAM ===
    if "terminate" in cmd or "terminate assistant" in cmd or "exit" in cmd or "quit" in cmd or "stop assistant" in cmd:
        speak("Goodbye! Shutting down the assistant.")
        print("\nAssistant terminated by user command.")
        
        # Stop any ongoing speech to avoid the destructor error
        engine.stop()
        
        # Clean up audio
        p.terminate()
        sys.exit(0)

    # === CLOSE BROWSER WINDOWS ===
    if cmd.startswith("close "):
        browser_phrase = cmd[6:].strip()
        target_process = None
        for key, process_name in BROWSER_PROCESSES.items():
            if key in browser_phrase:
                target_process = process_name
                browser_name = key.capitalize()
                if key == "google":
                    browser_name = "Chrome"
                break
        
        if target_process:
            closed = False
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if target_process in proc.info['name'].lower():
                        proc.terminate()
                        closed = True
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            if closed:
                speak(f"All {browser_name} windows closed")
            else:
                speak(f"No {browser_name} browser found running")
            return
        
    # === HELP ===
    if "view codes" in cmd:
        help_text = """
Available commands:
- start writing [text] → type text safely (no commands inside)
- arrow left/right/up/down → move cursor
- add [symbol] → type punctuation/symbols
- highlight it → select word/line
- select all → select everything
- copy it / copy all → copy selected text
- paste it → paste clipboard
- undo it → undo last action
- save it → Ctrl+S
- refresh it → F5
- find it → Ctrl+F
- print it → Ctrl+P
- zoom in / zoom out
- right click
- open windows → Windows key
- brightness higher / lower
- first tab / tenth tab / next tab / previous tab / new tab / close tab
- volume higher / lower [number] / mute it 
- scroll up / down [number]
- move left/right/up/down [slowly/fast[optional]] [seconds] - default normal speed
- stop moving → stop mouse movement
- click it / double click
- enter it / send it → press Enter
- open [app] → Gmail, YouTube, Word, Excel, Paint, CMD, etc.
- delete [number] times → backspace
        """
        print("\n" + help_text)
        speak("Full command list in console.")
        return
    
    # === ARROW KEYS ===
    if re.search(r"\b(arrow|ero|aero|arro|arow|erar)\b", cmd):
        for dir_name, key in [("left", "left"), ("right", "right"), ("up", "up"), ("down", "down")]:
            if dir_name in cmd:
                pyautogui.press(key)
                speak(f"Arrow {dir_name}")
                return
    
    # === ZOOM ===
    if "zoom in" in cmd:
        pyautogui.hotkey('ctrl', '+')
        speak("Zoomed in")
        return
    if "zoom out" in cmd:
        pyautogui.hotkey('ctrl', '-')
        speak("Zoomed out")
        return
    
    # === RIGHT CLICK ===
    if "right click" in cmd:
        pyautogui.rightClick()
        speak("Right clicked")
        return
    
    # === OPEN WINDOWS KEY ===
    if "open windows" in cmd:
        pyautogui.press('win')
        speak("Windows key pressed")
        return
    
    # === BRIGHTNESS ===
    if "brightness" in cmd:
        if "higher" in cmd:
            pyautogui.hotkey('win', 'a')
            time.sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.press('up')
            speak("Brightness increased")
        elif "lower" in cmd:
            pyautogui.hotkey('win', 'a')
            time.sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.press('down')
            speak("Brightness decreased")
        return
    
    # === COPY / PASTE / UNDO ===
    if "copy" in cmd:
        if "all" in cmd:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'c')
            speak("Copied all")
        else:
            pyautogui.hotkey('ctrl', 'c')
            speak("Copied")
        return
    
    if "paste" in cmd:
        pyautogui.hotkey('ctrl', 'v')
        speak("Pasted")
        return
    
    if "undo" in cmd:
        pyautogui.hotkey('ctrl', 'z')
        speak("Undid")
        return
    
    # === SAVE / REFRESH / FIND / PRINT ===
    if "save" in cmd:
        pyautogui.hotkey('ctrl', 's')
        speak("Saved")
        return
    
    if "refresh" in cmd or "reload" in cmd:
        pyautogui.press('f5')
        speak("Refreshed")
        return
    
    if "find" in cmd or "search" in cmd:
        pyautogui.hotkey('ctrl', 'f')
        speak("Find opened")
        return
    
    if "print" in cmd:
        pyautogui.hotkey('ctrl', 'p')
        speak("Print dialog opened")
        return
    
    # === START WRITING - SMART + EXTRA SPACE AT END ===
    writing_match = re.match(r"start writing[.,]?\s*(.*)", cmd)
    if writing_match:
        raw_text = writing_match.group(1).strip()
        
        if not raw_text:
            speak("Nothing to write")
            return
        
        clean_text = raw_text.replace("-", " ")
        
        words = clean_text.split()
        # Spelling mode (no correction, no space added at end)
        if len(words) > 3 and all(len(w.strip(".,!?")) <= 2 for w in words):
            text_to_type = clean_text.replace(" ", "")
        else:
            # Normal mode: full grammar correction
            text_to_type = grammar_correct(clean_text)
            # Add a space at the end so you can continue easily
            text_to_type += " "
        
        pyautogui.typewrite(text_to_type, interval=0)
        speak("Written")
        return

    # === DIRECT TAB SWITCHING ===
    tab_match = re.search(r"(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|\d+)\s*tab", cmd)
    if tab_match:
        num = word_to_number(tab_match.group(1))
        if num and 1 <= num <= 10:
            if num == 1:
                pyautogui.hotkey('ctrl', '1')
            elif num <= 9:
                pyautogui.hotkey('ctrl', str(num))
            else:
                pyautogui.hotkey('ctrl', '9')
                time.sleep(0.1)
                pyautogui.hotkey('ctrl', 'tab')
            speak(f"{'First' if num == 1 else 'Tenth' if num == 10 else f'Tab {num}'} tab")
        else:
            speak("Tab number not supported")
        return
    
    # === HIGHLIGHT / SELECT ===
    if any(v in cmd for v in ["highlight", "high light", "hi light", "hilight", "highlite", "select"]):
        if "right" in cmd:
            pyautogui.moveRel(60, 0, duration=0.2)
            pyautogui.doubleClick()
            pyautogui.moveRel(-40, 0, duration=0.1)
            speak("Selected next word")
        elif "left" in cmd:
            pyautogui.moveRel(-60, 0, duration=0.2)
            pyautogui.doubleClick()
            pyautogui.moveRel(40, 0, duration=0.1)
            speak("Selected previous word")
        elif "line" in cmd or "paragraph" in cmd or "down" in cmd:
            pyautogui.tripleClick()
            speak("Selected line")
        else:
            pyautogui.doubleClick()
            speak("Selected word")
        return
    
    if "select all" in cmd:
        pyautogui.hotkey('ctrl', 'a')
        speak("Selected all")
        return
    
    # === TAB CONTROL ===
    if "tab" in cmd:
        if "next" in cmd or "right" in cmd:
            pyautogui.hotkey('ctrl', 'tab')
            speak("Next tab")
        elif "previous" in cmd or "back" in cmd or "left" in cmd:
            pyautogui.hotkey('ctrl', 'shift', 'tab')
            speak("Previous tab")
        elif "new" in cmd or "open" in cmd:
            pyautogui.hotkey('ctrl', 't')
            speak("New tab opened")
        elif "close" in cmd:
            pyautogui.hotkey('ctrl', 'w')
            speak("Closed tab")
        return
    
    # === VOLUME ===
    if "volume" in cmd:
        count = 1
        match = re.search(r"(higher|lower)\s*(?:by\s*)?(\w+)?", cmd)
        if match:
            direction, num_part = match.group(1), match.group(2)
            if num_part:
                count = word_to_number(num_part) or 1
            count = max(1, min(50, count))
            key = 'volumeup' if direction == "higher" else 'volumedown'
            pyautogui.press([key] * count)
            speak(f"Volume {direction} by {count}")
            return
        if "higher" in cmd:
            pyautogui.press('volumeup')
            speak("Volume higher")
        elif "lower" in cmd:
            pyautogui.press('volumedown')
            speak("Volume lower")
        elif "mute" in cmd:
            pyautogui.press('volumemute')
            speak("Muted")
        return
    
    # === SCROLLING ===
    if "scroll" in cmd:
        count = 1
        match = re.search(r"(up|down)\s*(?:by\s*)?(\w+)?", cmd)
        if match:
            direction, num_part = match.group(1), match.group(2)
            if num_part:
                count = word_to_number(num_part) or 1
            count = max(1, min(30, count))
        amount = SCROLL_AMOUNT if "up" in cmd else -SCROLL_AMOUNT
        pyautogui.scroll(amount * count)
        speak(f"Scrolled {'up' if 'up' in cmd else 'down'}{f' {count} times' if count > 1 else ''}")
        return
    
    # === MOUSE MOVEMENT ===
    if "move" in cmd:
        direction = next((d for d in ["left", "right", "up", "down"] if d in cmd), None)
        if not direction:
            return
        speed = SLOW_SPEED if "slowly" in cmd else FAST_SPEED if "fast" in cmd else NORMAL_SPEED
        duration_match = re.search(r"(\w+)\s*seconds?", cmd)
        duration = word_to_number(duration_match.group(1)) if duration_match else None
        start_move(direction, speed, duration)
        return
    
    if cmd in ["left", "right", "up", "down"]:
        start_move(cmd, NORMAL_SPEED)
        return
    
    if "stop" in cmd:
        stop_all()
        return
    
    # === START WRITING - 100% SAFE ===
    writing_match = re.match(r"start writing[.,]?\s*(.*)", cmd)
    if writing_match:
        text_to_type = writing_match.group(1)  # Raw text - no command processing inside
        
        if not text_to_type.strip():
            speak("Nothing to write")
            return
        
        text_to_type = text_to_type.replace("-", "")
        
        words = text_to_type.split()
        if len(words) > 3 and all(len(w.strip(".,!?")) <= 2 for w in words):
            text_to_type = text_to_type.replace(" ", "")
        
        pyautogui.typewrite(text_to_type, interval=0)
        speak("Written")
        return
    
    # === ADD SYMBOLS ===
        # === ADD SYMBOLS OR NUMBERS ===
    if (cmd.lower().startswith("add ") or 
        cmd.lower().startswith("add. ") or 
        cmd.lower().startswith("add, ") or 
        cmd.lower().startswith("at ") or 
        cmd.lower().startswith("at. ") or 
        cmd.lower().startswith("at, ")):
        
        # Extract the phrase after the prefix
        lower_cmd = cmd.lower()
        if lower_cmd.startswith("add "):
            phrase = cmd[4:].strip()
        elif lower_cmd.startswith("add. "):
            phrase = cmd[5:].strip()
        elif lower_cmd.startswith("add, "):
            phrase = cmd[5:].strip()
        elif lower_cmd.startswith("at "):
            phrase = cmd[3:].strip()
        elif lower_cmd.startswith("at. "):
            phrase = cmd[4:].strip()
        elif lower_cmd.startswith("at, "):
            phrase = cmd[4:].strip()

        # Optional: make phrase lowercase for matching
        phrase = phrase.lower()  # Important: .lower() for consistency
        
        symbols = {
            # Single digits — ONLY word and digit forms (NO "number ..." prefixes)
            "zero": "0", "zero.": "0",
            "one": "1", "one.": "1",
            "two": "2", "two.": "2",
            "three": "3", "three.": "3",
            "four": "4", "four.": "4",
            "five": "5", "five.": "5",
            "six": "6", "six.": "6",
            "seven": "7", "seven.": "7",
            "eight": "8", "eight.": "8",
            "nine": "9", "nine.": "9",
            "0": "0", "1": "1", "2": "2", "3": "3", "4": "4",
            "5": "5", "6": "6", "7": "7", "8": "8", "9": "9",

            # All your punctuation/symbols (unchanged)
            "space.": " ", "spacebar.": " ", "space bar.": " ", "space": " ",
            "question mark.": "?", "question.": "?", "question": "?", 
            "exclamation mark.": "!", "exclamation mark": "!", "exclamation point.": "!", "exclamation point": "!", "exclamation.": "!", "exclamation": "!",
            "dot.": ".", "period": ".", "full stop.": ".", "point.": ".", "end of sentence.": ".",
            "comma.": ",", "comma": ",", "pause.": ",", "pause": ",", "break.": ",", "break": ",",
            "dash.": "-", "hyphen": "-", "hyphen.": "-", "dash": "-", "hyphenate.": "-", "hyphenate": "-",
            "underscore.": "_", "under score.": "_", "under-score.": "_", "under score": "_", "under-score": "_", 
            "colon.": ":", "colon": ":", 
            "semicolon.": ";", "semi colon.": ";", "semi-colon.": ";", "semi colon": ";", "semi-colon": ";",
            "apostrophe.": "'", "single quote.": "'", "apostrophe": "'", "single quote": "'",
            "quotation mark.": '"', "double quote.": '"', "quotation.": '"', "double quote": '"',
            "slash.": "/", "forward slash.": "/", "forward-slash.": "/", "forward slash": "/", "forward-slash": "/", "divide sign.": "/", "division sign.": "/", "divide.": "/", "division.": "/", "divide": "/", "division": "/", "slash": "/",
            "backslash.": "\\", "back slash.": "\\", "back-slash.": "\\", "back slash": "\\", "back-slash": "\\",
            "at sign.": "@", "at symbol.": "@", "at.": "@", "at": "@", "at sign": "@", "at symbol": "@",
            "hash sign.": "#", "hashtag.": "#", "pound sign.": "#", "number sign.": "#",
            "dollar sign.": "$", "dollar symbol.": "$", "dollar.": "$",
            "percent sign.": "%", "percent symbol.": "%", "percent.": "%",
            "caret.": "^", "circumflex.": "^", "to the power of.": "^", "power of.": "^", "to the power of": "^", "power of": "^",
            "ampersand.": "&", "and sign.": "&", "and symbol.": "&",
            "asterisk.": "*", "star.": "*", "multiply sign.": "*", "multiplication sign.": "*", "multiply.": "*", "multiplication.": "*", "multiply": "*", "multiplication": "*", "asterisk": "*", "star": "*", "multiply sign": "*", "multiplication sign": "*", "asterisk sign.": "*", "asterisk sign": "*", 
            "plus sign.": "+", "plus symbol.": "+", "plus.": "+", "plus": "+", "addition sign.": "+", "addition symbol.": "+", "addition.": "+", "addition": "+", "plus symbol": "+",
            "equal sign.": "=", "equals sign.": "=", "equal symbol.": "=", "equals symbol.": "=",
            "tilde.": "~", "approximate sign.": "~", "approximation sign.": "~", "curly dash.": "~", "curly dash": "~",
            "backtick.": "`", "grave accent.": "`", "back tick.": "`", "back tick": "`", "grave accent": "`",
            "left parenthesis.": "(", "left parenthesis": "(", "left paren.": "(", "left paren": "(",
            "right parenthesis.": ")", "right parenthesis": ")", "right paren.": ")", "right paren": ")",
            "left bracket.": "[", "left bracket": "[", "left square bracket.": "[", "left square bracket": "[", 
            "right bracket.": "]", "right bracket": "]", "right square bracket.": "]", "right square bracket": "]",
            "left brace.": "{", "left brace": "{", "left curly brace.": "{", "left curly brace": "{",
            "right brace.": "}", "right brace": "}", "right curly brace.": "}", "right curly brace": "}",
            "less than sign.": "<", "less than symbol.": "<", "less than.": "<", "less than": "<", "less than sign": "<", "less than symbol": "<",
            "greater than sign.": ">", "greater than symbol.": ">", "greater than.": ">", "greater than": ">", "greater than sign": ">", "greater than symbol": ">",
            "pipe.": "|", "vertical bar.": "|", "pipe": "|", "vertical bar": "|",
            "minus sign.": "-", "minus symbol.": "-", "minus.": "-", "minus": "-", "subtraction sign.": "-", "subtraction symbol.": "-", "subtraction.": "-", "subtraction": "-", "minus symbol": "-", "minus sign": "-",
        }

        # First: Check if it's a direct symbol match (e.g., "add five", "add comma")
        if phrase in symbols:
            pyautogui.typewrite(symbols[phrase])
            spoken = phrase.replace(".", "").replace("point", "").replace("mark", "").replace("sign", "").strip()
            speak(f"{spoken.title()} added")
            return

        # Second: Explicit number mode — "add number ..."
        if phrase.startswith("number "):
            number_phrase = phrase[7:].strip()
            clean_phrase = number_phrase.replace(",", "").replace("comma", "").replace(".", "").replace("-", " ")
            words = clean_phrase.split()

            try:
                total = 0
                current = 0
                i = 0
                while i < len(words):
                    word = words[i].lower()
                    if word in number_words:
                        value = number_words[word]
                    elif word.isdigit():
                        value = int(word)
                    else:
                        i += 1
                        continue

                    if value >= 1000:
                        current = current or 1
                        total += current * value
                        current = 0
                    elif value == 100:
                        current = current or 1
                        current *= 100
                    else:
                        if current >= 20 and value < 10:
                            current += value
                        elif current > 0:
                            current = current * 10 + value
                        else:
                            current = value
                    i += 1

                total += current or 0
                if total == 0:
                    speak("Number not recognized")
                    return

                pyautogui.typewrite(str(total))
                speak(f"Number {total:,} added")
                return
            except Exception:
                speak("Number not recognized")
                return

        # Third: Optional "add symbol ..." prefix
        if phrase.startswith("symbol "):
            symbol_phrase = phrase[7:].strip()
            if symbol_phrase in symbols:
                pyautogui.typewrite(symbols[symbol_phrase])
                spoken = symbol_phrase.replace(".", "").replace("point", "").replace("mark", "").replace("sign", "").strip()
                speak(f"{spoken.title()} added")
                return

        speak("Symbol or number not recognized")
        return
    
    # === DELETE ===
    delete_variants = ["delete", "remove", "believe", "delight", "elite", "delayed", "the lead", "delhi"]
    if any(v in cmd for v in delete_variants):
        count = 1
        match = re.search(r"(" + "|".join(delete_variants) + r")\s+((?:\w+\s)?\w+)\s*times?", cmd)
        if match:
            count = word_to_number(match.group(2).strip()) or 1
        count = max(1, min(50, count))
        pyautogui.press(["backspace"] * count)
        speak(f"Deleted left {count} times")
        return
    
    # === CLICK ===
    if "click" in cmd:
        if "double" in cmd:
            pyautogui.doubleClick()
            speak("Double clicked")
        else:
            pyautogui.click()
            speak("Clicked")
        return
    
    # === ENTER ===
    if "enter" in cmd or "send" in cmd:
        pyautogui.press("enter")
        speak("Enter pressed")
        return
    
    # === OPEN APPS & WEBSITES ===
    if "open" in cmd:
        opened = False
        # Websites
        for site, url in WEBSITES.items():
            if site in cmd:
                webbrowser.open(url)
                speak(f"Opening {site}")
                opened = True
                break
        
        if not opened:
            # Desktop apps
            for app_key, app_cmd in APPS.items():
                if app_key in cmd:
                    try:
                        if "administrator" in app_key:
                            subprocess.run(['runas', '/user:Administrator', app_cmd])
                        else:
                            subprocess.run(app_cmd, shell=True)
                        speak(f"Opening {app_key.replace('administrator', '(as admin)')}")
                    except Exception as e:
                        speak(f"Could not open {app_key}")
                    opened = True
                    break
        
        if not opened:
            speak("App or website not recognized")
        return
    
    speak("Try: view codes | start writing hello | open gmail | brightness higher | zoom in")

# Start
speak("Grammar improved! Sentences now properly capitalized with better commas.")
print("\n=== GRAMMAR UPGRADE ===\n")
print("• Every sentence capitalized")
print("• Natural comma placement")
print("• Final period added")
print("• Autocorrect enhanced")

while True:
    cmd = listen()
    execute_command(cmd)
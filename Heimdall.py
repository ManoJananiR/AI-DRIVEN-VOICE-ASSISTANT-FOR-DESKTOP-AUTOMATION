import os 
import webbrowser
import speech_recognition as sr
import pywhatkit
import wikipedia
import pyautogui
import datetime
import requests
import re
import urllib.parse
import time
import pywinauto
import pygetwindow as gw
from pywinauto import Application,Desktop
from pywinauto.findwindows import find_element
import platform
import uiautomation as auto
import psutil
import traceback
from pywinauto.findwindows import ElementNotFoundError
import math
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from urllib.parse import quote
import uiautomation as automation
import pytesseract
import win32con
import win32gui
import pyttsx3
import win32ui
from google_trans_new import google_translator
from pytesseract import image_to_string
import pytesseract
from bs4 import BeautifulSoup
from googletrans import Translator
import tkinter as tk
from tkinter import scrolledtext,Label, Entry, Frame,Button,Canvas,Scrollbar,font,PhotoImage
import threading
from PIL import Image, ImageTk,ImageFont,ImageGrab
import itertools
import subprocess
import platform
import cv2
import phonenumbers
import numpy as np
import pyaudio
import json

recording = False  # Global flag to track recording status
output_file = None  # Global variable to store filename

# Set the path for Tesseract-OCR (update this path if installed elsewhere)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# GUI Setup
root = tk.Tk()
root.title("Heimdall - Voice Assistant")
logo = tk.PhotoImage(file=r"E:\Desktop assistant\images\Capture.png")
root.iconphoto(True, logo)
root.geometry("720x650+300+20")
root.resizable(False, False)

# Light and Dark Mode Colors
light_bg = "white"
light_text = "#000000"
dark_bg = "#1e1e1e"
dark_text = "#ffffff"
user_color_light = "#C6E0B4"
assistant_color_light = "#DDEBF7"
user_color_dark = "#4CAF50"
assistant_color_dark = "#0c8294"
root.configure(bg=light_bg)
custom_font = ("Delius", 12,"bold")

dark_mode = False  # Theme flag

def toggle_theme():
    """Switch between light and dark mode"""
    global dark_mode
    dark_mode = not dark_mode
    new_bg = dark_bg if dark_mode else light_bg
    root.configure(bg=new_bg)
    chat_frame.configure(bg=new_bg)
    canvas.configure(bg=new_bg)
    scrollable_frame.configure(bg=new_bg)
    gif_label.configure(bg=new_bg)
    toggle_btn.configure(bg="black" if dark_mode else "#dddddd", fg=dark_text if dark_mode else light_text, text="Light Mode ☀" if dark_mode else "Dark Mode 🌙")

def create_chat_bubble(parent, text, bg_color, fg_color, align_right=False):
    """Creates a simple rectangular message bubble with correct padding."""
    text_label = tk.Label(
        parent,
        text=text,
        font=("Delius", 12, "bold"),
        fg=fg_color,
        bg=bg_color,
        padx=5, pady=6,  # Ensure text padding is consistent
        wraplength=350,  # Ensures text wraps properly within a max width
        # relief="flat",  # No border
        bd=0
    )

    # Correct alignment: right for user, left for assistant
    text_label.pack(anchor="e" if align_right else "w", padx=5, pady=5)

    return text_label

def add_chat_message(text, sender):
    """Display chat messages with correct padding based on its own text."""
    bg_color = user_color_dark if sender == "user" and dark_mode else user_color_light if sender == "user" else assistant_color_dark if dark_mode else assistant_color_light
    fg_color = "#000" if sender == "assistant" else "#333"

    is_user = sender == "user"
    bubble = create_chat_bubble(scrollable_frame, text, bg_color, fg_color, align_right=is_user)
    
    # Ensures proper spacing between user & assistant messages
    bubble.pack(fill="x", padx=(350, 10) if is_user else (10, 350), pady=5)

    # Scroll to the latest message
    canvas.update_idletasks()
    canvas.yview_moveto(1.0)

# Load GIFs for light and dark mode
light_gif_path = "E:\\Desktop assistant\\images\\light.gif"
dark_gif_path = "E:\\Desktop assistant\\images\\dark.gif"

gif_image_light = Image.open(light_gif_path)
gif_image_dark = Image.open(dark_gif_path)
frames_light, frames_dark = [], []

for frame in range(gif_image_light.n_frames):
    gif_image_light.seek(frame)
    gif_image_dark.seek(frame)
    frames_light.append(ImageTk.PhotoImage(gif_image_light.copy()))
    frames_dark.append(ImageTk.PhotoImage(gif_image_dark.copy()))

gif_index = 0
def update_gif():
    """Animate GIF based on theme"""
    global gif_index
    gif_label.config(image=frames_dark[gif_index] if dark_mode else frames_light[gif_index])
    gif_index = (gif_index + 1) % len(frames_light)
    root.after(100, update_gif)

gif_label = Label(root, bg=light_bg)
gif_label.pack(pady=10)
root.after(0, update_gif)

# Toggle Button

toggle_btn = Button(root, text="Dark Mode 🌙", command=toggle_theme, font=custom_font , padx=10, pady=5, bg="#ddd", fg=light_text, relief="flat")
toggle_btn.pack(pady=5)

# Scrollable Chat Area
chat_frame = Frame(root, bg=light_bg)
chat_frame.pack(fill="both", expand=True, padx=10, pady=5)

canvas = Canvas(chat_frame, bg=light_bg, highlightthickness=0)
scrollbar = Scrollbar(chat_frame, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollable_frame = Frame(canvas, bg=light_bg)
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Initialize TTS Engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

def talk(text):
    """Talk the provided text and display it in chat"""
    add_chat_message(text, "assistant")
    root.update()
    engine.say(text)
    engine.runAndWait()

def take_command():
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 50  # Increased sensitivity
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.5  # 🔹 Must be >= non_talking_duration
    recognizer.non_talking_duration = 0.2  # 🔹 Keep this lower

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)

        print("Listening...")
        while True:
            try:
                audio = recognizer.listen(source, timeout=None, phrase_time_limit=60)
                print("Processing...")

                command = recognizer.recognize_google(audio, language='en-US')
                command = command.lower()

                add_chat_message(command, "user")
                return command  
            
            except sr.UnknownValueError:
                print("Didn't catch that. Listening...")
                continue  
            
            except sr.RequestError:
                talk("There was an issue with the speech recognition service. Check your internet connection.")
                return "error"  
    
exchange_rates = {
    'USD': 82.0,
    'INR': 1.0,
}

# Amazon domains based on location
amazon_domains = {
    'india': 'amazon.in',
    'mumbai': 'amazon.in',
    'chennai': 'amazon.in',
    'america': 'amazon.com',
    'usa': 'amazon.com',
    'united states': 'amazon.com',
    # Add more locations as needed
}

def get_amazon_domain(location):
    for loc in amazon_domains:
        if loc in location:
            return amazon_domains[loc]
    return 'amazon.in'  # Default to India if location not found

def get_price_from_amazon(product_name, location):
    api_key = "bdf3f0ad02a5690191bcae56cf0f4134"  # ScraperAPI key
    base_url = "https://api.scraperapi.com"
    amazon_domain = get_amazon_domain(location)
    target_url = f"https://{amazon_domain}/s?k={product_name.replace(' ', '+')}"
    scraper_url = f"{base_url}?api_key={api_key}&url={target_url}"

    response = requests.get(scraper_url)
    soup = BeautifulSoup(response.content, "html.parser")

    price = None
    try:
        price = soup.find("span", {"class": "a-offscreen"}).text
    except AttributeError:
        talk("Sorry, I couldn't find the price for that item.")

    if price:
        price = re.sub(r'[^\d.]', '', price)  # Remove non-numeric characters
        # Convert price to rupees based on location
        currency = 'USD' if 'amazon.com' in target_url else 'INR'
        price_in_rupees = float(price) * exchange_rates[currency]

        return price_in_rupees
    return None

def check_wifi():
    """Check which Wi-Fi the laptop is connected to and announce it."""
    wifi_name = None

    if platform.system() == "Windows":
        try:
            result = subprocess.check_output(["netsh", "wlan", "show", "interfaces"], encoding="utf-8")
            for line in result.split("\n"):
                if "SSID" in line and "BSSID" not in line:
                    wifi_name = line.split(":")[1].strip()
                    break
        except Exception as e:
            talk(f"Error checking Wi-Fi: {e}")
            return

    if wifi_name:
        talk(f"Your laptop is connected to {wifi_name}.")
    else:
        talk("Your laptop is not connected to any Wi-Fi.")

def set_volume_by_percentage(percentage):
    """
    Sets the system volume to the specified percentage (0 to 100).
    Returns the actual volume level set in percentage.
    """
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    
    # Ensure percentage is within range 0-100
    scalar = max(0.0, min(percentage / 100.0, 1.0))
    volume.SetMasterVolumeLevelScalar(scalar, None)
    
    # Get actual volume set
    current_volume = round(volume.GetMasterVolumeLevelScalar() * 100)
    talk(f"Volume set to {current_volume} percent.")
    return current_volume

def process_volume_command():
    """Listens for a volume command and sets the volume accordingly."""
    command = take_command()

    if command:
        words = command.split()
        for word in words:
            if word.isdigit():  # Check if the word is a number
                volume_level = int(word)
                set_volume_by_percentage(volume_level)
                return

    talk("I did not catch a valid number. Please try again.")

def get_split_phone_number():
    while True:
        talk("Please say the first 5 digits of the phone number.")
        first_part = take_command()
        first_digits = ''.join(filter(str.isdigit, first_part))

        talk("Now say the last 5 digits of the phone number.")
        second_part = take_command()
        second_digits = ''.join(filter(str.isdigit, second_part))

        combined = first_digits + second_digits

        if len(combined) == 10:
            full_number = f"+91{combined}"
            talk(f"You said {full_number}. Is that correct? Say yes or no.")
            confirm = take_command().lower()

            if "yes" in confirm:
                return full_number
            else:
                talk("Okay, let's try again.")
        else:
            talk("That doesn't seem like a valid 10-digit number. Let's try again.")

def send_whatsapp_message():
    phone_number = get_split_phone_number()

    talk("What message should I send?")
    message = take_command()

    now = datetime.datetime.now() + datetime.timedelta(minutes=1)
    hour = now.hour
    minute = now.minute

    talk(f"Sending your message to {phone_number}")
    pywhatkit.sendwhatmsg(phone_number, message, hour, minute, wait_time=10, tab_close=True)

    talk("Your message has been scheduled.")

# Read active window content
def read_active_window_content():
    try:
        hwnd = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(hwnd)
        talk(f"The active window is titled: {window_title}")

        rect = win32gui.GetWindowRect(hwnd)
        x, y, x1, y1 = rect

        # Capture the window area
        window_image = ImageGrab.grab(bbox=(x, y, x1, y1))

        # Use OCR to extract text
        extracted_text = pytesseract.image_to_string(window_image)

        if extracted_text.strip():
            talk("The content of the active window is as follows:")
            talk(extracted_text[:1000])  # Read only the first 1000 characters
        else:
            talk("The active window has no readable content.")
    except Exception as e:
        talk("An error occurred while reading the active window's content.")
        print(f"Error: {e}")

# Function to get device information
def get_device_information():
    try:
        # Get system details
        system_info = {
            "System": platform.system(),
            "Node Name": platform.node(),
            "Release": platform.release(),
            "Version": platform.version(),
            "Machine": platform.machine(),
            "Processor": platform.processor(),
            "RAM": f"{round(psutil.virtual_memory().total / (1024 ** 3), 2)} GB"
        }

        # Read the system details
        talk("Here is your device information:")
        for key, value in system_info.items():
            talk(f"{key}: {value}")
    except Exception as e:
        talk("An error occurred while retrieving device information.")
        print(f"Error: {e}")

recording = False  # Global flag to control recording
output_file = None  # File name for saving the recording

def record_screen(file_name):
    """Records the screen and saves it to a file."""
    global recording

    recording = True  # Start recording
    screen_size = pyautogui.size()  # Get screen resolution

    # Use MJPG codec for better compatibility
    fourcc = cv2.VideoWriter_fourcc(*"MJPG")
    out = cv2.VideoWriter(file_name, fourcc, 20.0, screen_size)

    if not out.isOpened():
        talk("Error initializing video writer. Recording failed.")
        return

    talk(f"Recording started. File will be saved as {file_name}")

    try:
        while recording:
            img = pyautogui.screenshot()
            frame = np.array(img)
            frame = cv2.cvtColor(frame, cv2.COLOR_RGB2BGR)
            out.write(frame)  # Write frame to video
            time.sleep(1 / 20)  # 20 FPS

    except Exception as e:
        print(f"⚠️ Error during recording: {e}")
    
    out.release()  # Release video writer
    talk(f"Recording saved as {file_name}")

# Function to take a screenshot and save with user-provided name
def take_screenshot():
    try:
        # Ask the user for the file name to save the screenshot
        talk("What name would you like to save the screenshot as?")
        file_name = take_command().strip()
        
        if file_name:
            screenshot = pyautogui.screenshot()
            screenshot.save(f"{file_name}.png")  # Save screenshot with the user-provided name
            talk(f"Screenshot taken and saved as {file_name}.png.")
            print(f"Screenshot saved as {file_name}.png")
        else:
            talk("No file name provided. Please try again.")

    except Exception as e:
        talk("An error occurred while taking the screenshot.")
        print(f"Error taking screenshot: {e}")

# Function to translate text and talk the result
def translate_text():
    try:
        talk("Which word do you want to translate?")
        word = take_command()

        if not word:
            return

        talk("To which language should I translate?")
        language = take_command()

        if not language:
            return

        talk(f"Translating '{word}' to {language}...")

        # Language codes dictionary
        lang_codes = {
            "bengali": "bn", "chinese": "zh-cn", "english": "en", "french": "fr", "german": "de",
            "greek": "el", "hindi": "hi", "italian": "it", "japanese": "ja", "korean": "ko",
            "portuguese": "pt", "russian": "ru", "spanish": "es", "tamil": "ta", "turkish": "tr",
            "hebrew": "iw"
        }

        lang_code = lang_codes.get(language)

        if not lang_code:
            talk(f"Sorry, I don't support the language '{language}'. Please try again.")
            return translate_text()

        # Perform translation using googletrans
        translator = Translator()
        translated_word = translator.translate(word, src='auto', dest=lang_code).text

        # talk and print the result
        talk(f"The translation for '{word}' in {language} is '{translated_word}'.")

    except Exception as e:
        talk("An error occurred while translating.")
        print(f"Error: {e}")
# Function to fetch weather information
def fetch_weather():
    talk("Please tell me the city name.")
    city = take_command()
    if city:
        api_key = "05dd4d28db225f9c4f1b2f7adeb2a774"  # Replace with your actual OpenWeather API key
        url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
        try:
            response = requests.get(url).json()
            if "main" in response:
                temperature = response["main"]["temp"]
                description = response["weather"][0]["description"]
                talk(f"The weather in {city} is {description} with a temperature of {temperature} degrees Celsius.")
            else:
                talk("I couldn't fetch the weather details. Please try again.")
        except Exception as e:
            talk("There was an error fetching the weather details.")
            print(f"Error fetching weather: {e}")
# Function to search Wikipedia
def search_wikipedia(command):
    topic = command.replace("wikipedia", "").strip()
    if topic:
        try:
            info = wikipedia.summary(topic, sentences=2)
            talk(info)
        except wikipedia.exceptions.DisambiguationError:
            talk("The topic is ambiguous. Please be more specific.")
        except wikipedia.exceptions.PageError:
            talk("I couldn't find any page on Wikipedia for that topic.")
            
def take_spelled_filename():
    talk("Please spell the file name, letter by letter.")
    filename = ""
    while True:
        letter = take_command()
        if "stop" in letter:  # Stop input when user says "stop"
            break
        elif len(letter) == 1:  # Ensure the input is a single letter
            filename += letter
            talk(f"Added letter {letter}.")
        else:
            talk("Please say a single letter.")
    return filename

def open_notepad():
    talk("Opening Notepad and starting to write. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("notepad.exe")  # Open Notepad
    time.sleep(1)  # Allow some time for Notepad to open
    content = ""
    start_time = time.time()
    while True:
        # Listen for content until the user says 'stop writing'
        command = take_command()
        if "stop writing" in command:
            break
        elif command:  # If the command is not empty, add to content
            content += command + "\n"
            pyautogui.write(command)  # Simulate typing in Notepad
        # Break after 20 seconds of no input or wait for the next command
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()  # Reset the timer
        time.sleep(3)  # Add a break of 3 seconds after each listening
    # Ask for file name to save the content
    talk("What name would you like to save the file as? Please say the file name.")
    file_name = take_command().strip()
    # If the recognition of the filename fails or is empty, provide alternative options
    if not file_name:
        talk("I couldn't recognize the filename. Would you like to spell it out instead? Say 'yes' to spell it or 'no' to enter it manually.")
        response = take_command().strip().lower()
        if "yes" in response:
            file_name = take_spelled_filename()
        elif "no" in response:
            talk("Please type the file name manually.")
            file_name = input("Enter file name manually: ").strip()
        else:
            talk("Invalid response. Please try again later.")
            return
    if file_name:
        # Fixed directory path where the file will be saved
        save_directory = "C:\\Users\\User\\Documents"
        save_path = os.path.join(save_directory, f"{file_name}.txt")
        try:
            # Save the content to the file
            with open(save_path, "w") as f:
                f.write(content)
            talk(f"Content saved as {file_name}.txt in your Documents folder.")
            print(f"Content saved as {file_name}.txt in {save_directory}")
        except Exception as e:
            talk(f"An error occurred while saving the file: {e}")
    else:
        talk("No file name provided. The content was not saved.")
def close_notepad():
    try:
        # Terminate all Notepad instances
        os.system("taskkill /im notepad.exe /f")
        talk("Notepad has been closed.")
        print("Notepad closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Notepad.")
        print(f"Error closing Notepad: {e}")
def open_word():
    talk("Opening Microsoft Word. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("winword.exe")  # Open Word
    time.sleep(5)  # Allow some time for Word to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            content += command + "\n"
            pyautogui.write(command)
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the document as?")
    file_name = take_command().strip()
    if file_name:
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Document saved as {file_name}.")
    else:
        talk("No file name provided. The document was not saved.")
def close_word():
    try:
        os.system("taskkill /im winword.exe /f")
        talk("Microsoft Word has been closed.")
        print("Word closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Word.")
        print(f"Error closing Word: {e}")
def open_excel():
    talk("Opening Microsoft Excel. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("excel.exe")  # Open Excel
    time.sleep(5)  # Allow some time for Excel to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            pyautogui.write(command)
            pyautogui.press("tab")
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the spreadsheet as?")
    file_name = take_command().strip()
    if file_name:
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Spreadsheet saved as {file_name}.")
    else:
        talk("No file name provided. The spreadsheet was not saved.")
def close_excel():
    try:
        os.system("taskkill /im excel.exe /f")
        talk("Microsoft Excel has been closed.")
        print("Excel closed successfully.")
    except Exception as e:
        talk("An error occurred while closing Excel.")
        print(f"Error closing Excel: {e}")
def open_powerpoint():
    talk("Opening Microsoft PowerPoint. talk your content, and say 'stop writing' when you want to stop.")
    os.startfile("powerpnt.exe")  # Open PowerPoint
    time.sleep(5)  # Allow some time for PowerPoint to open
    content = ""
    start_time = time.time()
    while True:
        command = take_command()
        if "stop writing" in command:
            break
        elif command:
            pyautogui.write(command)
            pyautogui.press("enter")
        if time.time() - start_time > 20:
            talk("Listening for your next input...")
            start_time = time.time()
        time.sleep(3)
    talk("What name would you like to save the presentation as?")
    file_name = take_command().strip()
    if file_name:
        pyautogui.hotkey("ctrl", "s")
        time.sleep(2)
        pyautogui.write(file_name)
        pyautogui.press("enter")
        talk(f"Presentation saved as {file_name}.")
    else:
        talk("No file name provided. The presentation was not saved.")
def close_powerpoint():
    try:
        os.system("taskkill /im powerpnt.exe /f")
        talk("Microsoft PowerPoint has been closed.")
        print("PowerPoint closed successfully.")
    except Exception as e:
        talk("An error occurred while closing PowerPoint.")
        print(f"Error closing PowerPoint: {e}")
def get_active_window():
    """Get the title of the currently active window."""
    try:
        active_window = gw.getActiveWindow()
        return active_window.title if active_window else "Unknown Window"
    except Exception as e:
        print("Error getting active window:", e)
        return "Unknown Window"
def switch_window():
    """Switch to the next open window using Alt + Tab."""
    talk("Switching window...")
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    pyautogui.keyUp("alt")

def switch_tab():
    """Switch between tabs based on the active window."""
    active_window = get_active_window().lower()

    if "chrome" in active_window:
        talk("Switching tab in Chrome")
        pyautogui.hotkey("ctrl", "tab")  # Next tab in Chrome

    elif "visual studio code" in active_window or "vs code" in active_window:
        talk("Switching tab in VS Code")
        pyautogui.hotkey("ctrl", "pageup")  # Next tab in VS Code

    else:
        talk("This application does not support tab switching.")

def fetch_news():
    try:
        talk("What topic are you interested in?")
        topic = take_command().strip()
        if not topic:
            talk("You didn't provide a topic. Please try again.")
            return

        api_key = "74ba339c1205483cbf2d4205c51cd690"  # Replace with your actual API key
        url = f"https://newsapi.org/v2/everything?q={topic}&language=en&pageSize=5&apiKey={api_key}"
        print(f"Request URL: {url}")

        response = requests.get(url).json()
        print(f"Response: {response}")

        if response.get("status") == "error":
            talk(f"Error fetching news: {response.get('message')}")
            return

        if response.get("articles"):
            talk(f"Here are the top news articles on {topic}:")
            for i, article in enumerate(response["articles"][:5], 1):
                title = article.get("title", "No title available")
                description = article.get("description", "No description available")
                print(f"News {i}: {title} - {description}")
                talk(f"News {i}: {title}. {description}")
        else:
            talk("No news articles found. Please try a different topic.")
    except Exception as e:
        talk("There was an error fetching the news.")
        print(f"Error fetching news: {e}")
def check_battery_status():
    """Check battery percentage and charging status."""
    battery = psutil.sensors_battery()
    if battery is None:
        talk("Sorry, I couldn't fetch battery details.")
        return

    percent = battery.percent
    charging = battery.power_plugged
    status = "charging" if charging else "not charging"
    
    talk(f"The battery is at {percent} percent and it is currently {status}.")

def refresh_windows():
    try:
        pyautogui.press('f5')  # Simulate pressing F5
        talk("The desktop has been refreshed.")
        print("Desktop refreshed.")
    except Exception as e:
        talk("An error occurred while refreshing the desktop.")
        print(f"Error refreshing desktop: {e}")
# Function to perform actions based on commands
def run_assistant():
     global recording, output_file  # Needed to modify recording status
     while True:
        command = take_command()
        if not command:  # Check if command is None or empty
             continue
            # root.after(100, run_assistant)
            # return  # Return to listening state if no command is detected
        if "refresh windows" in command or "refresh desktop" in command:
            refresh_windows()
        elif "charge percentage" in command or "battery status" in command:
            check_battery_status()
        elif "time" in command:
            current_time = datetime.datetime.now().strftime('%I:%M %p')
            talk(f"The current time is {current_time}.")
        elif "search" in command:
            query = command.replace("search", "").strip()
            webbrowser.open(f"https://www.google.com/search?q={query}")
            talk(f"Searching for {query}.")
        elif "who are you" in command:
            talk("I am Heimdall, your personal assistant. How can I help you?")
        elif "how are you" in command:
            talk("I am fine and hope you are too, boss. How can I help you?")
        elif "tell me a joke" in command:
            talk("Why was the math book sad? Because it had too many problems.")
        elif "play" in command:
            song = command.replace("play", "").strip()
            talk(f"Playing {song} on YouTube.")
            pywhatkit.playonyt(song)
        elif "Wi-Fi" in command or "wi-fi" in command:
            check_wifi()
        elif "price of" in command:
            parts = command.split(" in ")
            if len(parts) == 2:
                item_name = parts[0].replace("price of", "").strip()
                location = parts[1].strip()
                price = get_price_from_amazon(item_name, location)
                if price:
                    talk(f"The price of {item_name} in {location} is {price} rupees.")
                else:
                    talk("Sorry, I couldn't find the price for that item.")
            else:
                talk("Please specify the location for the price.")
        elif "get weather" in command:
            fetch_weather()
        elif "device information" in command:
            get_device_information()
        elif "set volume to" in command:
            process_volume_command()
        elif "record screen" in command and not recording:
            talk("Say the file name to save the recording.")
            file_name = take_command().replace(" ", "_") + ".avi"
            output_file = file_name if file_name.strip() else "screen_recording.avi"
        # Start screen recording in a new thread
            threading.Thread(target=record_screen, args=(output_file,), daemon=True).start()
        elif "stop" in command and recording:
            recording = False  # Stop recording
            talk("Recording stopped.")
        elif "open google" in command or "open Google" in command or "Open Google" in command:
            webbrowser.open("https://www.google.com")
            talk("Opening Google.")
        elif "open w3schools" in command or "open W3Schools" in command or "Open W3Schools" in command:
            webbrowser.open("https://www.w3schools.com/")
            talk("Opening W3Schools.")
        elif "wikipedia" in command:
            search_wikipedia(command)
        elif "translate" in command:
            translate_text()
        elif "take a screenshot" in command:
            take_screenshot()
        elif "where is" in command:
            location = command.replace("where is", "").strip()
            webbrowser.open(f"https://www.google.com/maps/place/{location}")
            talk(f"Here is the location of {location}.")
        elif "send message" in command or "send whatsapp message" in command:
            send_whatsapp_message()
        elif "read window content" in command:
            read_active_window_content()
        elif "move window" in command or "move" in command:
            switch_window()
        elif "move tab" in command:
            switch_tab()
        # elif "open vscode" in command or "open Vscode" in command
        elif "open notepad" in command:
            open_notepad()
        elif "close notepad" in command:
            close_notepad()
        elif "open word" in command:
            open_word()
        elif "close word" in command:
            close_word()
        elif "open excel" in command:
            open_excel()
        elif "close excel" in command:
            close_excel()
        elif "open powerpoint" in command:
            open_powerpoint()
        elif "close powerpoint" in command:
            close_powerpoint()
        elif "news" in command:
            fetch_news()
        elif "goodbye" in command or "exit" in command:
            talk("Good bye Boss, See you later!")
            root.after(1000, root.destroy)  # Close GUI after message
            break
        else:
            talk("I didn't catch that. Please try again.")
       #root.after(100, run_assistant)  # Schedule next listening

def start_assistant():
    """Start assistant in a separate thread"""
    talk("Hello Boss, how can I assist you now?")
    threading.Thread(target=run_assistant, daemon=True).start()

# Start Assistant After GUI Loads
root.after(1000, start_assistant)

# Start Main Event Loop
root.mainloop()
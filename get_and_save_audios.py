from tkinter import ttk, messagebox
import tkinter as tk
import pyttsx3
from gtts import gTTS
import traceback
from util_functions import threading, requests

#SOMETHING DOES NOT WORK

'''
Play Audio
'''
def playAudio(text_field_Description):
    try:
        engine = pyttsx3.init()
        text = text_field_Description.get("1.0", tk.END)
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
       messagebox.showerror("Play Audio Error", str(e))

'''
Save Audio
'''
def saveAudio(text_field_Description, text_field_ID):
    try:
        engine = pyttsx3.init()
        text = text_field_Description.get("1.0", tk.END)
        #id = text_field_ID.get()
        engine.save_to_file(text, f"{text_field_ID.get()}.mp3")
        engine.runAndWait()
        messagebox.showinfo("Save Audio", "Audio saved as id.mp3")
    except Exception as e:
        messagebox.showerror("Save Audio Error", str(e))

def SaveAllAudios(text_field):
    update_dress_list = []
    # get dress numbers from text field
    get_text_field = text_field.get("1.0", "end-1c").split(',')

    # add to list
    for number in get_text_field:
        if (number.strip().isnumeric()):
            update_dress_list.append(int(number.strip()))
    for num in update_dress_list:
        print(num)
        fetchTextAndSaveAudio(num)

def fetchTextAndSaveAudio(id):
    headers = {
        'Accept': '*/*',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
    }
    try:
        # Replace with the actual API endpoint you're supposed to hit
        url = f"https://abcd2.projectabcd.com/api/getinfo.php?id={id}"
        response = requests.get(url, headers=headers)
        if response.ok:
            data = response.json()
            description = data['data']['description']

            # Convert the description to audio
            tts = gTTS(description, lang='en')
            # Save the audio file named as id.mp3
            audio_file_path = f"{id}.mp3"
            tts.save(audio_file_path)
        else:
            print(f"Failed to fetch data for ID {id}: {response.status_code}, {response.reason}")
    except Exception as e:
        print(f"Error fetching data for ID {id}: {e}")

'''
Get Text
'''
def fetchText(root, text_field_Description, id):
    headers = {
        'Accept': '*/*',  # Use wildcard or specific type based on API requirement
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'
    }
    try:
        url = f"https://abcd2.projectabcd.com/api/getinfo.php?id={id}"
        response = requests.get(url, headers=headers)
        # print("Request Headers:", response.request.headers)
        if response.ok:
            # print("API Response:", response.text)
            data = response.json()
            # description = data.get("description", "No description found.")
            description = data['data']['description']
            text_field_Description.delete(1.0, tk.END)  # This clears all text from the widget
            root.after(0, lambda: text_field_Description.insert(tk.END, description))
        else:
            print("Failed to fetch data:", response.status_code, response.reason)
            messagebox.showerror("Error", f"Failed to fetch data from the server: {response.reason}, Status Code: {response.status_code}")
    except Exception as e:
        traceback.print_exc()
        print("Error:", str(e))
        messagebox.showerror("Error", str(e))

'''
Spins up new thread to run getText function
'''
def getTextThread(root, text_field_Description, text_field_ID):
    dress_id = text_field_ID.get()
    if dress_id:
        thread = threading.Thread(target=fetchText(root, text_field_Description, id))
        thread.start()
    else:
        messagebox.showerror("Error", "Please enter a dress ID.")

'''
Spins up new thread to run playAudio function
'''
def playAudioThread(root, button, text_field_Description):
    button.config(state='disabled')
    get_audio_thread = threading.Thread(target=playAudio(text_field_Description))
    get_audio_thread.start()

'''
Spins up new thread to run saveAudio function
'''
def saveAudioThread(root, text_field_Description, button, text_field_ID):
    button.config(state='disabled')
    get_audio_thread = threading.Thread(target=saveAudio(text_field_Description, text_field_ID))
    get_audio_thread.start()
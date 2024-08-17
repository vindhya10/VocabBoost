import pyttsx3
from plyer import notification
import time
import pandas as pd
import inflect
from tkinter import *

# Initialize inflect engine
p = inflect.engine()

# Path to the Excel file
file_location = r"C:\Users\USER\PycharmProjects\Vocabulary_tutor\list.xlsx"
try:
    df = pd.read_excel(file_location, engine='openpyxl')
    print(df.head())  # Debug print to check if the data is loaded correctly
except Exception as e:
    print(f"Error opening file: {e}")
    exit()

# Notify function
def notify(title, message):
    notification.notify(title=title, message=message, app_icon=r"C:\Users\USER\PycharmProjects\Vocabulary_tutor\icon3.ico", timeout=15)

# Text-to-Speech initialization
engine = pyttsx3.init()
voices = engine.getProperty('voices')
# Use a default voice index
engine.setProperty('voice', voices[1].id)  # You can change the index if needed
engine.setProperty("rate", 135)
engine.setProperty("volume", 1.0)

# Function to display vocabulary
def values(name, timer):
    i = 1
    for _, row in df.iterrows():
        try:
            word = row[0]
            meaning = row[1]
            msg = f"_WORD: {word}\n\nMEANING_: {meaning}"
            text = f"Hello {name}, here is your {p.ordinal(i)} word: {word} - {meaning}"
            print(f"Debug Text: {text}")  # Debug print statement
            print(f"Debug Msg: {msg}")  # Debug print statement
            engine.say(text)
            notify("VOCABULARY", msg)
            engine.runAndWait()
            i += 1
            time.sleep(timer)
        except Exception as e:
            print(f"Error: {e}")
            break

# Tkinter GUI setup
root = Tk()
root.title('Vocabulary Tutor')
root.geometry("700x700")
root.configure(background='gray27')

def submit():
    try:
        name = my_box1.get()
        timer = int(my_box2.get())
        if not name or timer <= 0:
            raise ValueError("Please provide a valid name and timer.")
        values(name, timer)
    except ValueError as e:
        print(f"Error: {e}")

read = """Welcome to Vocabulary Tutor,
this app allows you to set the timer.
You will be notified with an English
word and its meaning after every
specified time. Please fill the below
requirements to continue..."""

my_label0 = Label(root, text=read, font=("Courier", 18, "bold italic"), fg="light cyan", bg="gray27")
my_label0.pack(pady=20)

my_label1 = Label(root, text="Enter your name", font=("Helvetica", 18), fg="gold", bg="gray27")
my_label1.pack(pady=20)

my_box1 = Entry(root)
my_box1.pack(pady=20)

my_label = Label(root, text="Set the timer (in sec)", font=("Times bold", 18), fg="gold", bg="gray27")
my_label.pack(pady=20)

my_box2 = Entry(root)
my_box2.pack(pady=20)

my_button = Button(root, text="Submit", fg='gray1', bg='PaleVioletRed1', activebackground='lawn green', command=submit)
my_button.pack(pady=50)

root.mainloop()

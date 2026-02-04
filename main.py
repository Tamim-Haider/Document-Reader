import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import docx
import pyttsx3
import os
import threading  # For handling speech in the background

# Setup Text-to-Speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)  # Speed (lower is slower)

# Functions to read files
def read_text_file(path):
    with open(path, 'r', encoding='utf-8') as file:
        return file.read()

def read_pdf_file(path):
    doc = fitz.open(path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def read_word_file(path):
    doc = docx.Document(path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

def read_document(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension == '.txt':
        return read_text_file(file_path)
    elif extension == '.pdf':
        return read_pdf_file(file_path)
    elif extension == '.docx':
        return read_word_file(file_path)
    else:
        return "Unsupported file type."

def speak_text(text):
    engine.say(text)
    engine.runAndWait()

def speak_text_threaded(text):
    """Function to run text-to-speech in a separate thread."""
    # Stop any previous speech before starting a new one
    engine.stop()
    # Start a new speech thread
    thread = threading.Thread(target=speak_text, args=(text,))
    thread.start()

# Function to stop speaking
def stop_speaking(event=None):
    engine.stop()  # Stop the engine
    print("Speech stopped!")

# When the button is clicked
def open_and_read_file():
    file_path = filedialog.askopenfilename(
        title="Select a document",
        filetypes=(("PDF Files", "*.pdf"), ("Text Files", "*.txt"), ("Word Files", "*.docx"), ("All files", "*.*"))
    )
    if file_path:
        try:
            print(f"Selected file: {file_path}")  # Debug
            content = read_document(file_path)
            if content.strip() == "":
                messagebox.showwarning("Warning", "The document is empty or could not be read properly.")
            else:
                messagebox.showinfo("Document Content", content[:1000] + ("..." if len(content) > 1000 else ""))
                # Start speaking the content
                speak_text_threaded(content[:1000])  # Speak the first part of the document
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the document.\n\nError: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("Desktop Assistant - Document Reader")
root.geometry("400x250")  # Increased height a bit

# "Read Document" Button
read_button = tk.Button(root, text="📄 Read and Speak Document", command=open_and_read_file, font=('Arial', 14))
read_button.pack(pady=20)

# "Stop Speaking" Button
stop_button = tk.Button(root, text="🛑 Stop Speaking", command=stop_speaking, font=('Arial', 14), bg='red', fg='white')
stop_button.pack(pady=20)

# Bind the Escape key to stop speaking
root.bind('<Escape>', stop_speaking)

root.mainloop()

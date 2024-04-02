import tkinter as tk
from PIL import Image, ImageTk
import pyttsx3

class VoiceAssistantGUI:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.root = tk.Tk()
        self.root.title("JarvisUI")

        # Load and display image
        image = Image.open(r"C:\Users\HP\Desktop\jarvis_GUI.jpg")  
        photo = ImageTk.PhotoImage(image)
        image_label = tk.Label(self.root, image=photo)
        image_label.image = photo
        image_label.pack()
       
        self.root.mainloop()

if __name__ == "__main__":
    app = VoiceAssistantGUI()






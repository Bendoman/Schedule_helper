
from tkinter import *
from tkinter import ttk
import customtkinter  

app = customtkinter.CTk()
app.geometry("720x480")

for i in range(0, 25):
    app.Label(text="Hello World!").grid(column=0, row=i)

app.mainloop()

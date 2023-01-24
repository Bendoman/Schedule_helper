from tkinter import *
from tkinter import ttk

# if you are still working under a Python 2 version, 
# comment out the previous line and uncomment the following line
# import Tkinter as tk

root = Tk()
root.title('Hello world')
root.geometry('720x500')

# Create a main frame
main_frame = Frame(root)
main_frame.pack(fill=BOTH, expand=1)

# Create a canvase
my_canvas = Canvas(main_frame)
my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

# Add a scrollbar to the Canvas
my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
my_scrollbar.pack(side=RIGHT, fill=Y)

# configure the canvas
my_canvas.configure(yscrollcommand=my_scrollbar.set)
my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox('all')))

# create another frame inside the canvas
second_frame = Frame(my_canvas)

my_canvas.create_window((0,0), window=second_frame, anchor="nw")




for i in range(100):
    Label(second_frame, text=f'label {i}').grid(row=i, column=0, pady=10, padx=10)
    Button(second_frame, text=f'Button {i}').grid(row=i, column=1, pady=10, padx=10)

root.mainloop()
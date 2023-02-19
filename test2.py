import tkinter
from tkinter import ttk

def event_handler(event):
    pass

app = tkinter.Tk()
menu = ttk.Menubutton()

for event_type in tkinter.EventType.__members__.keys():
    event_seq= "<" + event_type + ">"
    try:
        menu.bind_all(event_seq, event_handler)
        print(event_type)
    except tkinter.TclError:
        #print("bind error:", event_type)
        pass

app.mainloop()
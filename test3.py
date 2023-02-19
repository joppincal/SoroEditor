import tkinter as tk

def menucallback(event):
    check = root.call(event.widget, "index","active")
    print('menucallback')
    if check != "none":
        menu_index.set(check)
def foo():
    display.configure(text=menu_index.get())
    index = menu_index.get()
    if index == 0:
        display.configure(bg='blue')
    if index == 1:
        display.configure(bg='red')
    display.wait_variable(menu_index)
    foo()
    
    

root = tk.Tk()
display = tk.Label(root, text='default')
display.pack(fill='x')

menu_index = tk.IntVar()

menubar =tk.Menu(root)
menu1 = tk.Menu(menubar, tearoff=0)
menu1.add_command(label="Button1")
menu1.add_command(label="Button2")
menubar.add_cascade(label='Menu1', menu=menu1)

menu1.bind('<<MenuSelect>>', menucallback)
tk.Tk.config(root, menu=menubar)

foo()
root.mainloop()
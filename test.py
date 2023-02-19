from ttkbootstrap import PanedWindow
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText



app = ttk.Window(themename='darkly')

pw1 = PanedWindow(app, height=300, width=500, style='info', orient=HORIZONTAL)
st1 = ScrolledText(pw1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
st2 = ScrolledText(pw1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
st3 = ScrolledText(pw1, padding=5, autohide=True, wrap=CHAR, font=('', 10))
pw1.add(st1, weight=1)
pw1.add(st2, weight=1)
pw1.add(st3, weight=1)

# pw2 = PanedWindow(app, height=300, width=500, orient=HORIZONTAL)
pw1.pack(fill=BOTH, expand=True)
# pw2.pack(fill=BOTH, expand=True)

app.mainloop()

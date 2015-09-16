from tkinter import *
from tkinter.ttk import *

class Assessment(Tk):
    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()
        self.title('VCPI Assessment')

    def initialize(self):
        self.grid()

if __name__ == "__main__":
    app = Assessment(None)
    app.mainloop()
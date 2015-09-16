from tkinter import *
from tkinter.font import *
from tkinter.ttk import *

class Assessment(Tk):
    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.title('VCPI Assessment')
        self.grid()
        self.font = Font(family='New Time Roman', size=9)

        self.entry_one = Entry(self, font=self.font)
        self.entry_one.grid(column=0, row=0)

        self.entry_two = Entry(self, font=self.font)
        self.entry_two.grid(column=0, row=1)

        self.entry_three = Entry(self, font=self.font)
        self.entry_three.grid(column=0, row=2)

        self.entry_four = Entry(self, font=self.font)
        self.entry_four.grid(column=0, row=3)

        self.button_one = tkinter.Button(self, text='Query', font=self.font)
        self.button_one.grid(column=0, row=4, sticky='W')

        self.button_two = tkinter.Button(self, text='Save',font=self.font)
        self.button_two.grid(column=0, row=5, sticky='W')

if __name__ == "__main__":
    app = Assessment(None)
    app.mainloop()
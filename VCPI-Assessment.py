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
        self.entry_one.grid(column=1, row=0, sticky='E')

        self.entry_two = Entry(self, font=self.font)
        self.entry_two.grid(column=1, row=1, sticky='E')

        self.entry_three = Entry(self, font=self.font)
        self.entry_three.grid(column=1, row=2, sticky='E')

        self.entry_four = Entry(self, font=self.font)
        self.entry_four.grid(column=1, row=3, sticky='E')

        self.button_one = tkinter.Button(self, text='Query', font=self.font)
        self.button_one.grid(column=0, row=4, sticky='W')

        self.button_two = tkinter.Button(self, text='Save',font=self.font)
        self.button_two.grid(column=0, row=5, sticky='W')

        self.label_one = Label(self, text='Facility ID', font=self.font)
        self.label_one.grid(column=0, row=0, sticky='W')

        self.label_two = Label(self, text='PC ID', font=self.font)
        self.label_two.grid(column=0, row=1, sticky='W')

        self.label_three = Label(self, text='Location', font=self.font)
        self.label_three.grid(column=0, row=2, sticky='W')

        self.label_four = Label(self, text='Comment', font=self.font)
        self.label_four.grid(column=0, row=3, sticky='W')

if __name__ == "__main__":
    app = Assessment(None)
    app.set_font_size(5)
    app.mainloop()
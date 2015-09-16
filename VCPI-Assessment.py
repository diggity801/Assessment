from tkinter import *
from tkinter.font import *
from tkinter.ttk import *
import wmi
import datetime

class Assessment(Tk):
    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()

    def initialize(self):
        self.title('VCPI Assessment')
        self.wmi = wmi.WMI()
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

    def get_os_name(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.caption

    def get_os_version(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.version

    def get_manufacturer(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.manufacturer

    def get_name(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.name

    def get_model(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.model

    def get_install_date(self):
        for os in self.wmi.Win32_OperatingSystem():
            return str(datetime.datetime.strptime(str(os.installdate)[0:8], '%Y%m%d')).split()[0]

    def get_architecture(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.osarchitecture

    def get_domain(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.domain

    def get_processor(self):
        for processor in self.wmi.Win32_Processor():
            return processor.name

    def get_memory(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.totalvisiblememorysize

    def get_serial(self):
        for bios in self.wmi.Win32_Bios():
            return bios.serialnumber

    def get_network_address(self):
        return self.wmi.Win32_NetworkAdapterConfiguration()[1].ipaddress[0]

if __name__ == "__main__":
    app = Assessment(None)
    print(app.get_os_name())
    print(app.get_os_version())
    print(app.get_name())
    print(app.get_model())
    print(app.get_architecture())
    print(app.get_domain())
    print(app.get_network_address())
    print(app.get_processor())
    print(app.get_install_date())
    app.mainloop()
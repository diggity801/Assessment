from tkinter import *
from tkinter.font import *
from tkinter.ttk import *
from win32com.client import GetObject
import wmi
import datetime

class Assessment(Tk):
    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.wmi = wmi.WMI()
        self.initialize()

    def initialize(self):
        self.geometry('250x150')
        self.title('Assessment')
        self.font = Font(family='Tahoma', size=9)

        self.grid()
        self.grid_columnconfigure(0,weight=1)
        self.grid_columnconfigure(1,weight=1)
        self.grid_columnconfigure(2,weight=1)
        self.grid_rowconfigure(0,weight=5)
        self.grid_rowconfigure(4,weight=5)
        self.grid_rowconfigure(5, weight=2)
        self.grid_rowconfigure(6, weight=2)
        self.resizable(False, False)

        self.entry_one = Entry(self, font=self.font)
        self.entry_one.grid(column=2, row=0, sticky='WS')

        self.entry_two = Entry(self, font=self.font)
        self.entry_two.grid(column=2, row=1, sticky='WS')

        self.entry_three = Entry(self, font=self.font)
        self.entry_three.grid(column=2, row=2, sticky='WS')

        self.entry_four = Entry(self, font=self.font)
        self.entry_four.grid(column=2, row=3, sticky='WS')

        self.button_one = tkinter.Button(self, text='Query', font=self.font, width=6, height=1, )
        self.button_one.grid(column=1, row=4, sticky='S')

        self.button_two = tkinter.Button(self, text='Save',font=self.font, width=6, height=1)
        self.button_two.grid(column=1, row=5, sticky='N')

        self.label_one = Label(self, text='Facility ID', font=self.font)
        self.label_one.grid(column=1, row=0, sticky='WS')

        self.label_two = Label(self, text='PC ID', font=self.font)
        self.label_two.grid(column=1, row=1, sticky='WS')

        self.label_three = Label(self, text='Location', font=self.font)
        self.label_three.grid(column=1, row=2, sticky='WS')

        self.label_four = Label(self, text='Comment', font=self.font)
        self.label_four.grid(column=1, row=3, sticky='WS')

        self.int_variable_one = IntVar(self, 0)
        self.int_variable_two = IntVar(self, 0)

        self.check_one = Checkbutton(self, variable=self.int_variable_one)
        self.check_one.grid(column=2, row=4, sticky='EWS',)

        self.check_one = Checkbutton(self, variable=self.int_variable_two)
        self.check_one.grid(column=2, row=5, sticky='NEWS')

    def get_os_name(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.caption

    def get_os_version(self):
        for os in self.wmi.Win32_OperatingSystem():
            return os.version

    def get_manufacturer(self):
        for computer in self.wmi.Win32_ComputerSystem():
            if computer.manufacturer != 'To Be Filled By O.E.M.':
                return computer.manufacturer
            else:
                return None

    def get_name(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.name

    def get_model(self):
        for computer in self.wmi.Win32_ComputerSystem():
            if computer.model != 'To Be Filled By O.E.M.':
                return computer.model
            else:
                return None

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
            return "{0:.2f}".format(float(os.totalvisiblememorysize) / 1048576)

    def get_serial(self):
        for bios in self.wmi.Win32_Bios():
            if bios.serialnumber != 'To Be Filled By O.E.M.':
                return bios.serialnumber
            else:
                return None

    def get_network_address(self):
        return self.wmi.Win32_NetworkAdapterConfiguration()[1].ipaddress[0]

    def get_antivirus(self):
        self.obj_wmi = GetObject(r'winmgmts:\\.\root\SecurityCenter2').InstancesOf('AntiVirusProduct')
        for antivirus in self.obj_wmi:
            return antivirus.displayname

    def get_last_user(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.username

if __name__ == "__main__":
    app = Assessment(None)
    print(app.get_name())
    print(app.get_manufacturer())
    print(app.get_os_name())
    print(app.get_model())
    print(app.get_os_version())
    print(app.get_install_date())
    print(app.get_architecture())
    print(app.get_domain())
    print(app.get_processor())
    print(app.get_memory())
    print(app.get_serial())
    print(app.get_network_address())
    print(app.get_antivirus())
    print(app.get_last_user())
    app.mainloop()
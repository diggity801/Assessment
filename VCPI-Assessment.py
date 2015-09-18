from tkinter import *
from tkinter.font import *
from win32com.client import GetObject
import wmi
import datetime
from openpyxl import *
import win32api
import os
import subprocess
from distutils.version import StrictVersion

class Assessment(Tk):
    def __init__(self, parent):
        Tk.__init__(self, parent)
        self.parent = parent
        self.wmi = wmi.WMI()
        self.initialize()

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
                return str(None)

    def get_name(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.name

    def get_model(self):
        for computer in self.wmi.Win32_ComputerSystem():
            if computer.model != 'To Be Filled By O.E.M.':
                return computer.model
            else:
                return str(None)

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
                return str(None)

    def get_network_address(self):
        return self.wmi.Win32_NetworkAdapterConfiguration()[1].ipaddress[0]

    def get_antivirus(self):
        self.obj_wmi = GetObject(r'winmgmts:\\.\root\SecurityCenter2').InstancesOf('AntiVirusProduct')
        for antivirus in self.obj_wmi:
            return antivirus.displayname

    def get_last_user(self):
        for computer in self.wmi.Win32_ComputerSystem():
            return computer.username

    def write_excel(self, event):
        if self.button_two['state'] != 'disable' and self.button_two['state'] != 'disabled':
            if os.path.isfile('test.xlsx'):
                self.wb = load_workbook('test.xlsx')
                self.ws = self.wb.active
                self.ws.append([self.assessment_id, self.location, self.name, self.manufacturer, self.os_name, self.processor, self.memory, self.model, self.serial, self.comment, self.network])
                self.wb.save('Assessment Details.xlsx')
            else:
                self.wb = Workbook()
                self.ws = self.wb.active
                self.ws.append([self.assessment_id, self.location, self.name, self.manufacturer, self.os_name, self.processor, self.memory, self.model, self.serial, self.comment, self.network])
                self.wb.save('Assessment Details.xlsx')

    def query_system(self):
        self.button_two['state'] = 'active'
        self.manufacturer = self.get_manufacturer()
        self.name = self.get_name()
        self.os_name = self.get_os_name()
        self.model = self.get_model()
        self.version = self.get_os_version()
        self.install_date = self.get_install_date()
        self.architecture = self.get_architecture()
        self.domain = self.get_domain()
        self.processor = self.get_processor()
        self.memory = self.get_memory()
        self.serial = self.get_serial()
        self.network = self.get_network_address()
        self.antivirus = self.get_antivirus()
        self.last_user = self.get_last_user()
        self.facility = str(self.entry_one.get())
        self.assessment_id = str(self.entry_two.get())
        self.location = str(self.entry_three.get())
        self.comment = str(self.entry_four.get())

    def multi_step(self, event):
        self.query_system()
        self.install_citrix()

    def install_citrix(self):
        try:
            self.info = win32api.GetFileVersionInfo(r'C:\Program Files (x86)\Citrix\ICA Client\Receiver\Receiver.exe', '\\')
            self.ms = self.info['FileVersionMS']
            self.ls = self.info['FileVersionLS']
            self.__version__ = "{}.{}.{}".format(win32api.HIWORD(self.ms), win32api.LOWORD(self.ms), win32api.HIWORD(self.ls))
        except:
            self.__version__ = None

        if StrictVersion(self.__version__) < StrictVersion('3.3.0'):
            try:
                subprocess.call(r'ReceiverCleanupUtility.exe /silent')
            except:
                pass
            try:
                subprocess.call(r'CitrixReceiverEnterprise.exe /silent /noreboot ENABLE_SSON=No ALLOWSAVEPWD=A ENABLE_DYNAMIC_CLIENT_NAME=Yes SERVER_LOCATION=http://pnagent.vcpi.com/Citrix/PNAgent/config.xml')
            except:
                pass

    def initialize(self):
        self.geometry('250x150')
        self.title('Assessment')
        self.font = Font(family='Tahoma', size=9)

        self.grid()
        self.grid_columnconfigure(0,weight=0)
        self.grid_columnconfigure(1,weight=2)
        self.grid_columnconfigure(2,weight=3)
        self.grid_rowconfigure(0,weight=5)
        self.grid_rowconfigure(4,weight=10)
        self.grid_rowconfigure(5, weight=4)
        self.grid_rowconfigure(6, weight=3)
        self.resizable(False, False)

        self.entry_one = Entry(self, font=self.font)
        self.entry_one.grid(column=2, row=0, sticky='WS')

        self.entry_two = Entry(self, font=self.font)
        self.entry_two.grid(column=2, row=1, sticky='WS')

        self.entry_three = Entry(self, font=self.font)
        self.entry_three.grid(column=2, row=2, sticky='WS')

        self.entry_four = Entry(self, font=self.font)
        self.entry_four.grid(column=2, row=3, sticky='WS')

        self.button_one = tkinter.Button(self, text='Query', font=self.font, width=4, height=1, padx=20)
        self.button_one.grid(column=1, row=4, sticky='SE')
        self.button_one.bind('<Button-1>', self.multi_step)

        self.button_two = tkinter.Button(self, text='Save',font=self.font, width=4, height=1, padx=20, state='disabled')
        self.button_two.grid(column=1, row=5, sticky='NE')
        self.button_two.bind('<Button-1>', self.write_excel)

        self.label_one = Label(self, text='Facility ID', font=self.font)
        self.label_one.grid(column=1, row=0, sticky='WS', padx=10)

        self.label_two = Label(self, text='PC ID', font=self.font)
        self.label_two.grid(column=1, row=1, sticky='WS', padx=10)

        self.label_three = Label(self, text='Location', font=self.font)
        self.label_three.grid(column=1, row=2, sticky='WS', padx=10)

        self.label_four = Label(self, text='Comment', font=self.font)
        self.label_four.grid(column=1, row=3, sticky='WS', padx=10)

        self.int_variable_one = IntVar(self, 0)
        self.int_variable_two = IntVar(self, 0)

        self.label_five = Label(self, text='OK', font=self.font, anchor='s', pady=4)
        self.label_five.grid(column=2, row=4, sticky='SW')

        self.label_six = Label(self, text='OK', font=self.font, anchor='sw')
        self.label_six.grid(column=2, row=5, sticky='W')

        self.frame_one = Frame(self, relief='sunken', borderwidth=0)
        self.frame_one.grid(column=2, row=4, rowspan=2, columnspan=3, sticky='E')

        self.label_seven = Label(self.frame_one, text='SQL', font=self.font)
        self.label_seven.grid(column=2, row=4, sticky='SE', padx=50)

        self.label_eight = Label(self.frame_one, text='Excel', font=self.font)
        self.label_eight.grid(column=2, row=5, sticky='NE', padx=50, pady=4)

        self.check_one = Checkbutton(self.frame_one, variable=self.int_variable_one, borderwidth=0)
        self.check_one.grid(column=2, row=4, sticky='SE', padx=20)
        self.check_one.config(state='disable')

        self.check_two = Checkbutton(self.frame_one, variable=self.int_variable_two, borderwidth=0)
        self.check_two.grid(column=2, row=5, sticky='E', padx=20)
        self.check_one.config(state='disable')

if __name__ == "__main__":
    app = Assessment(None)
    app.mainloop()
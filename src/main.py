from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import asksaveasfile
from tkinter.messagebox import showinfo
from collections import OrderedDict
from subprocess import getoutput
from csv import DictWriter
from pyqrcode import pyqrcode
from table import Table

class NetshWlan:
    def __init__(self):
        """Contruct a 2d array to save all networks' information"""
        self.listName = self.__getListName()
        self.listInfo = []
        for element in self.listName:
            record = self.__getNetworkInfo(element)
            self.listInfo.append(record)
    
    def __getListName(self):
        """Get a list of SSID"""
        chunkLine = getoutput("netsh wlan show profiles").split('\n')   # split command content into multiple lines

        i = 0
        while i < len(chunkLine):
            if chunkLine[i].find('All User Profile') == -1: # remove unnecessary line
                chunkLine.remove(chunkLine[i])
                i -= 1
            else:
                chunkLine[i] = chunkLine[i].replace('    All User Profile     : ', '')  # delete header of SSID
            i += 1
        return chunkLine

    def __getNetworkInfo(self, ssid):
        """Get all information about current network"""
        rawInfo = getoutput("netsh wlan show profiles name = \"" + ssid + "\" key = clear")
        chunkInfo = rawInfo.split('\n')

        i = 0
        while i < len(chunkInfo):
            if chunkInfo[i].find("Authentication") == -1 and \
            chunkInfo[i].find("Key") == -1:
                chunkInfo.remove(chunkInfo[i])
                i -= 1
            i += 1

        chunkInfo[0] = chunkInfo[0].replace("    Authentication         : ", "")
        if chunkInfo[1].find("Authentication") != -1:
            chunkInfo.remove(chunkInfo[1])
            chunkInfo[1] = chunkInfo[1].replace("    Key Content            : ", "")
        else:
            chunkInfo[1] = "None"
        return [ssid, chunkInfo[0], chunkInfo[1]]

class IpConfig:
    def __init__(self):
        """Get all network interface cards and their infomation"""
        self.nics = self.__createNics()

    def __createNics(self):
        """Get raw network interface cards' infomation"""
        chunkLine = getoutput("ipconfig /all").split("\n")

        nics = {}
        name = ""
        for i in range(1, len(chunkLine)):
            if chunkLine[i] == "":  # if null line then skip
                continue
            elif chunkLine[i - 1] == "" and chunkLine[i + 1] == "": # if a nics' name then push to dict as a key
                name = chunkLine[i]
                if name[len(name) - 1] == ":":
                    name = name.replace(":", "")
                nics[name] = ""
            else:   # if a detail just simply append to dict[current key]
                if nics[name] != "":    # first line, do not append "\n"
                    nics[name] += "\n" 
                nics[name] += chunkLine[i].replace("   ", "")
        return nics

class WirelessProfiles:
    def __init__(self, parent):
        """Contruct a window about 802.11 profiles"""
        self._window = parent

        self._lblTitle = Label(self._window, text = "802.11 PROFILES", font = 10)
        self._lblTitle.grid(row=0, column=0, columnspan=2, padx=(10, 0), pady=(10, 0))

        self._tblInfo = Table(self._window)
        self._tblInfo.construct(["SSID", "Kind", "Key"], [150] * 3)
        self._tblInfo.grid(row=1, column=0, rowspan=5, padx=(25, 0), pady=(10, 0))
        self._tblInfo.bind("<ButtonRelease-1>", self._on_click)
        
        self._fTabRight = Frame(self._window)
        self._fTabRight.grid(row=1, column=1, sticky=NW, padx=(10, 10), pady=(10, 0))

        self._lfDetail = LabelFrame(self._fTabRight, text="Detail")
        self._lfDetail.grid(row=0, column=0, sticky=NW)

        self._lblSsid = Label(self._lfDetail, text="SSID")
        self._lblSsid.grid(row=0, column=0, sticky=NW, padx=(10, 0), pady=(10, 0))
        self._txtSsid = Entry(self._lfDetail, state="readonly")
        self._txtSsid.grid(row=0, column=1, sticky=NE, padx=(10, 10), pady=(10, 0))

        self._lblKind = Label(self._lfDetail, text="Kind")
        self._lblKind.grid(row=1, column=0, sticky=NW, padx=(10, 0), pady=(10, 0))
        self._txtKind = Entry(self._lfDetail, state="readonly")
        self._txtKind.grid(row=1, column=1, sticky=NE, padx=(10, 10), pady=(10, 0))

        self._lblKey = Label(self._lfDetail, text="Key")
        self._lblKey.grid(row=2, column=0, sticky=NW, padx=(10, 0), pady=(10, 0))
        self._txtKey = Entry(self._lfDetail, state="readonly")
        self._txtKey.grid(row=2, column=1, sticky=NE, padx=(10, 10), pady=(10, 0))

        self._lblQr = Label(self._lfDetail, text="QR")
        self._lblQr.grid(row=3, column=0, sticky=NW, padx=(10, 0), pady=(10, 0))
        self._cvQr = Canvas(self._lfDetail, width=140, height=140)
        self._cvQr.grid(row=4, column=0, columnspan=2, sticky=N, padx=(10, 0), pady=(10, 10))

        self._checked = IntVar()
        self._chkbtnShow = Checkbutton(self._fTabRight, text="Show password", variable=self._checked, command=self.displayKey)
        self._chkbtnShow.grid(row=1, column=0, sticky=NW, pady=(7, 0))
        self._btnLoad = Button(self._fTabRight, text="Load from system", width=30, command=self.loadData)
        self._btnLoad.grid(row=2, column=0, sticky=NW, padx=(0, 10), pady=(7, 0))
        self._btnSave = Button(self._fTabRight, text="Export", width=30, command=self.exportData)
        self._btnSave.grid(row=3, column=0, sticky=NW, padx=(0, 10), pady=(7, 0))
        self._btnClear = Button(self._fTabRight, text="Clear all", width=30, command=self.clearData)
        self._btnClear.grid(row=4, column=0, sticky=NW, padx=(0, 10), pady=(7, 0))

        self._lblReport = Label(self._window, text="Total: 0")
        self._lblReport.grid(row=6, column=0, columnspan=2, sticky=NW, padx=(25, 0), pady=(10, 10))

    def _on_click(self, event):
        """Binding between table and entries"""
        currentRecord = self._tblInfo.item(self._tblInfo.focus(), "values")
        
        if currentRecord != "":     # solve exception: click on header line
            self._txtSsid.configure(state=NORMAL)
            self._txtSsid.delete(0, END)
            self._txtSsid.insert(END, currentRecord[1])
            self._txtSsid.configure(state="readonly")

            self._txtKind.configure(state=NORMAL)
            self._txtKind.delete(0, END)
            self._txtKind.insert(END, currentRecord[2])
            self._txtKind.configure(state="readonly")

            self._txtKey.configure(state=NORMAL)
            self._txtKey.delete(0, END)
            self._txtKey.insert(END, currentRecord[3])
            self._txtKey.configure(state="readonly")
            # convert current key (at focusing in table) to qr code
            # do not use string in txtKey because it may be wrong when is in hidden password mode "*""
            content = self._listNetwork.listInfo[self._tblInfo.item(self._tblInfo.focus())["values"][0] - 1][2]
            codeXbm = pyqrcode.create(content).xbm(scale=4)
            self._codeBmp = BitmapImage(data=codeXbm)
            self._codeBmp.config(foreground="black", background="white")
            self._cvQr.create_image(0, 0, anchor=NW, image=self._codeBmp)

    def displayKey(self):
        """Set wifi key whether is hidden or displayed"""
        if self._checked.get():     # case all keys are displayed
            if len(self._tblInfo.get_children()) != 0:  # when data are loaded before
                i = 0
                for row in self._tblInfo.get_children():
                    self._tblInfo.set(row, column="4", value=self._listNetwork.listInfo[i][2])
                    i += 1
                if self._tblInfo.item(self._tblInfo.focus())["values"] != "":   # solve exception: there is no focusing in table
                    self._txtKey.configure(state=NORMAL)
                    self._txtKey.delete(0, END)
                    self._txtKey.insert(END, self._tblInfo.item(self._tblInfo.focus())["values"][3])
                    self._txtKey.configure(state="readonly")
        else:   # case all keys are hidden
            if len(self._tblInfo.get_children()) != 0:
                i = 0
                for row in self._tblInfo.get_children():
                    self._tblInfo.set(row, column="4", value="*"*len(self._listNetwork.listInfo[i][2]))
                    i += 1
                if self._tblInfo.item(self._tblInfo.focus())["values"] != "":   # solve exception: there is no focusing in table
                    self._txtKey.configure(state=NORMAL)
                    self._txtKey.delete(0, END)
                    self._txtKey.insert(END, self._tblInfo.item(self._tblInfo.focus())["values"][3])
                    self._txtKey.configure(state="readonly")

    def loadData(self):
        """Load data to table"""
        self.clearData()    # clear first for not being appended later

        self._listNetwork = NetshWlan()
        self._tblInfo.fill(self._listNetwork.listInfo)

        if not self._checked.get(): # if do not show key (not checked), turn all keys into "*" character
            i = 0
            for row in self._tblInfo.get_children():
                self._tblInfo.set(row, column="4", value="*"*len(self._listNetwork.listInfo[i][2]))
                i += 1
        showinfo("Infomation", "Load all profiles from system successfully")

        report = "Total: " + str(len(self._listNetwork.listInfo))
        if len(self._listNetwork.listInfo) != 0:    # there is at least one row
            frequency = {}  # dictionary to store frequency of each kind
            for kind in self._listNetwork.listInfo: # traverse kind of network profiles to count their occurences
                if kind[1] in frequency:
                    frequency[kind[1]] += 1
                else:
                    frequency[kind[1]] = 1
            frequency = OrderedDict(sorted(frequency.items()))  # cast dict into ordered and (for) sort it

            report += "      "
            for key, value in frequency.items():    # create report string
                report += key + ": " + str(value) + "      "
        self._lblReport.configure(text=report)

    def exportData(self):
        """Export wireless network information to csv file"""
        fileTypes = [("CSV UTF-8 (Comma delimited)", "*.csv"),]
        file = asksaveasfile(mode='w', filetypes=fileTypes, defaultextension=fileTypes)
        
        if file != None:    # solve exception: when press Export button and exit immediately (no Save)
            with open(file.name, mode="w", newline="") as csv_file: # newline="" for not making new line after a record
                fieldnames = ("ID", "SSID", "Kind", "Key")
                writer = DictWriter(csv_file, fieldnames=fieldnames)
                writer.writeheader()
                for i in range(len(self._listNetwork.listInfo)):
                    writer.writerow({fieldnames[0]: i + 1,
                                    fieldnames[1]: self._listNetwork.listInfo[i][0],
                                    fieldnames[2]: self._listNetwork.listInfo[i][1],
                                    fieldnames[3]: self._listNetwork.listInfo[i][2]})
    def clearData(self):
        """Clear all data in profile table"""
        for item in self._tblInfo.get_children():
            self._tblInfo.delete(item)

        self._txtSsid.configure(state=NORMAL)
        self._txtSsid.delete(0, END)
        self._txtSsid.configure(state="readonly")
        self._txtKind.configure(state=NORMAL)
        self._txtKind.delete(0, END)
        self._txtKind.configure(state="readonly")
        self._txtKey.configure(state=NORMAL)
        self._txtKey.delete(0, END)
        self._txtKey.configure(state="readonly")
        self._cvQr.delete("all")

        self._lblReport.configure(text="Total: 0")       

class NetworkInterfaceCards:
    def __init__(self, parent):
        """Contruct a list of network interface cards"""
        self._window = parent

        self._lblTitle = Label(self._window, text="NETWORK INTERFACE CARDS", font=10)
        self._lblTitle.grid(row=0, column=0, columnspan=2, padx=(10, 0), pady=(10, 0))

        self._lblName = Label(self._window, text="List of NICs")
        self._lblName.grid(row=1, column=0, sticky=NW, padx=(25, 0), pady=(10, 0))
        self._lbName = Listbox(self._window, width=40)
        self._lbName.bind("<<ListboxSelect>>", self._on_select)
        self._lbName.grid(row=2, column=0, sticky=N, padx=(25, 0), pady=(5, 0))
        
        self._fButton = Frame(self._window)
        self._fButton.grid(row=3, column=0, sticky=NW, padx=(25, 0))
        self._lblHost = Label(self._fButton, text="Hostname: None")
        self._lblHost.grid(row=0, column=0, columnspan=2, sticky=NW, pady=(10, 0))
        self._lblTotal = Label(self._fButton, text="Total: 0")
        self._lblTotal.grid(row=1, column=0, columnspan=2, sticky=NW, pady=(10, 0))
        self._lblCurrent = Label(self._fButton, text="Current: None")
        self._lblCurrent.grid(row=2, column=0, columnspan=2, sticky=NW, pady=(10, 0))

        self._btnLoad = Button(self._fButton, text="Load NICs", width=18, command=self.loadNics)
        self._btnLoad.grid(row=3, column=0, sticky=NW, padx=(0, 5), pady=(10, 0))
        self._btnClear = Button(self._fButton, text="Clear all", width=18, command=lambda:[self.clearNics(), self.clearDetails()])
        self._btnClear.grid(row=3, column=1, sticky=NE, padx=(5, 0), pady=(10, 0))
        
        self._lblInfo = Label(self._window, text="Information")
        self._lblInfo.grid(row=1, column=1, sticky=NW, padx=(15, 0), pady=(10, 0))
        self._tblInfo = Table(window=self._window, scrollbar=False, height=20)
        self._tblInfo.construct(["Properties", "Values"], [175, 250], idcolumn=False, anchor=NW)
        self._tblInfo.grid(row=2, column=1, rowspan=10, sticky=N, padx=(15, 0), pady=(5, 0))

    def _on_select(self, event):
        """View details of each network interface card"""
        try:
            # get value (info) of current key (nics' name)
            content = self._networkInterfaceCards.nics[event.widget.get(event.widget.curselection()[0])].split("\n")
            datasource = []
            for item in content:
                currentField = item.split(" : ")
                currentField[0] = currentField[0].replace(".", "")
                datasource.append(currentField)
            self.clearDetails() # clear table before filling
            self._tblInfo.fill(datasource)
        except Exception as e:
            pass

    def loadNics(self):
        """Load list of network interface cards into listbox"""
        self.clearNics()    # clear list before filling, because function here is INSERT
        self._networkInterfaceCards = IpConfig()    # get a dictionary of nics

        for key, value in self._networkInterfaceCards.nics.items():
            if key != "Windows IP Configuration":   # leave out first field, because it is not a nic
                self._lbName.insert(END, key)
            if "Lease Obtained" in value:   # save for use later, as current nic which PC is using
                current = key
        self._lblHost.configure(text="Hostname: " + getoutput("hostname"))
        self._lblTotal.configure(text="Total: " + str(len(self._networkInterfaceCards.nics) - 1))
        self._lblCurrent.configure(text="Current use: " + current)
    
    def clearNics(self):
        """Clear the list of network inteface cards"""
        self._lbName.delete(0, END)
        
    def clearDetails(self):
        """Clear all details of the current NICs"""
        for item in self._tblInfo.get_children():
            self._tblInfo.delete(item)

class About:
    def __init__(self, parent):
        self._dialog = Toplevel(parent)
        self._dialog.geometry("300x200")
        self._dialog.title("About")
        self._dialog.resizable(False, False)
        self._dialog.grab_set()

        self._textDescription = Text(self._dialog)
        self._textDescription.pack(fill=BOTH, padx=10, pady=(10, 45))
        aboutString = "Developed by Pham Hai Duong.\n" + \
                      "This project is as a simple analyzer from cmd.\n" + \
                      "Two main features: show 802.11 profiles key and list of NICs.\n" + \
                      "Improving Python as well as Tkinter library is a core purpose."
        self._textDescription.insert(END, aboutString)
        self._textDescription.configure(state=DISABLED)
        self._dialog.mainloop()

class MainApp:
    def __init__(self):
        """Create main app"""
        self._window = Tk()
        self._window.geometry("800x600")
        self._window.title("Your network")
        self._window.resizable(False, False)

        self._menubar = Menu(self._window)  # create menubar

        self._fileMenu = Menu(self._menubar, tearoff=0)
        self._menubar.add_cascade(label="File", menu=self._fileMenu)
        self._fileMenu.add_command(label="Exit", command=self._window.destroy)

        self._infoMenu = Menu(self._menubar, tearoff=0)
        self._menubar.add_cascade(label="Information", menu=self._infoMenu)
        self._infoMenu.add_command(label="About", command=self.getAbout)

        self._window.configure(menu=self._menubar)  # add menubar to window

        self._tabControl = Notebook(self._window)
        self._tabControl.grid(row=0, column=0, sticky=N, padx=(30, 0), pady=(20, 0))
        self._tabWifi = Frame(self._tabControl)
        self._tabWifi.grid(row=0, column=0, sticky=N, padx=(10, 0), pady=(10, 0))
        self._tabNics = Frame(self._tabControl)
        self._tabNics.grid(row=0, column=0, sticky=N, padx=(10, 0), pady=(10, 0))
        self._tabControl.add(self._tabWifi, text="Wifi")
        self._tabControl.add(self._tabNics, text="NICs")

        self._wirelessProfiles = WirelessProfiles(self._tabWifi)
        self._networkInterfaceCards = NetworkInterfaceCards(self._tabNics)

        self._window.mainloop()
    
    def getAbout(self):
        authorInformation = About(self._window)
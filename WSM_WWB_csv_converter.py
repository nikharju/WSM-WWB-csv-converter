##########################################################################################
#  
#       MIT License
#
#       Copyright (c) 2023 Niklas Harju
#       Written by Niklas Harju
#
#       Permission is hereby granted, free of charge, to any person obtaining a copy
#       of this software and associated documentation files (the "Software"), to deal
#       in the Software without restriction, including without limitation the rights
#       to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#       copies of the Software, and to permit persons to whom the Software is
#       furnished to do so, subject to the following conditions:
#
#       The above copyright notice and this permission notice shall be included in all
#       copies or substantial portions of the Software.
#
#       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#       IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#       AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#       LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#       OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#       SOFTWARE.
#  
##########################################################################################
#
#        ABOUT THIS PROGRAM
#        This program imports radio frequency scan files created in and exported from 
#        Sennheisers software Wireless Systems Manager (WSM) as .csv files and reformats 
#        them to a csv format that works with Shures software Wireless Workbench (WWB).
#        Simply click 'Open file...' to import the WSM file and click 'Convert to WWB...' 
#        and choose a location and name to save the new file with the new formatting.
#        After that you can import the newly created scan file into WWB.
#
#        DISCLAIMER!
#        This program is in no way affiliated with neither Sennheiser nor Shure 
#        corporations, nor is it supported by them. It is a tool created by me 
#        personally to help with the task of using Sennheiser hardware and software 
#        for scanning radio frequency spectrums and use that data for frequency planning 
#        in WWB.
#        
##########################################################################################
#
#        HISTORY
#        Version 1.0.0 2023-05-02    Initial version, Python 3.11.0
#        Version 1.0.1 2023-05-03    Some layout fixes and added app information, Python 3.11.0
#        Version 1.0.2 2023-05-08    Added cross platform support for Windows and MacOS,
#                                    Python 3.11.0
#
##########################################################################################

import csv
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, font
import tkinter.ttk as ttk

class App:
    def __init__(self, master):
### APP SETTINGS ###
        self.master = master
        self.master.title("WSM to WWB csv converter")
        self.program_version = "1.0.2"
        self.running_os = sys.platform
        self.os_in_darkmode = self.is_os_in_darkmode()
        if self.running_os == 'win32':
            # Windows app size
            self.HEIGHT = 200
            self.WIDTH  = 600
        else:
            # MacOS app size
            self.HEIGHT = 220
            self.WIDTH  = 700
        self.master.geometry(f'{self.WIDTH}x{self.HEIGHT}+{int(self.master.winfo_screenwidth()/2-self.WIDTH/2)}+{int(self.master.winfo_screenheight()/2-self.HEIGHT/2)}')
        self.master.resizable(False, False)
        self.table = list

### STYLE SETTINGS ###
        self.dark = "#3A3A3A"
        self.grey = "#d0d0d0"
        self.light_grey = "#f1f1f1"
        self.style = ttk.Style()
        self.style.element_create("plain.field", "from", "clam")
        self.style.layout("entry_fields.TEntry",
            [('Entry.plain.field', {'children': [(
                'Entry.background', {'children': [(
                    'Entry.padding', {'children': [(
                        'Entry.textarea', {'sticky': 'nswe'})],
                'sticky': 'nswe'})], 'sticky': 'nswe'})],
                'border':'2', 'sticky': 'nswe'})])
        self.style.configure('entry_fields.TEntry')
        if self.running_os == 'win32':
            # Windows style settings
            # Missing dark mode settings for Windows
            self.style.configure('TLabelframe', background = self.light_grey)
            self.style.configure('TLabelframe.Label', background = self.light_grey)
            self.style.configure('TLabel', background = self.light_grey)
            self.style.configure('TFrame', background = self.light_grey)
            self.style.map('entry_fields.TEntry',
                fieldbackground=[('readonly', self.light_grey)])
        else:
            # MacOS style settings
            if self.os_in_darkmode:
                self.style.configure(".", background = self.dark, foreground = self.light_grey)
                self.style.configure('TLabelframe', background = self.dark)
                self.style.configure('TLabelframe.Label', background = self.dark)
                self.style.configure('TLabel', background = self.dark)
                self.style.configure('TFrame', background = self.dark)
                self.style.map('entry_fields.TEntry',
                    fieldbackground=[('readonly', self.dark)])
                self.style.map('TButton',
                    foreground = [('disabled', self.dark)])
            else:
                self.style.configure(".", background = self.light_grey, foreground = "black")
                self.style.configure('TLabelframe', background = self.light_grey)
                self.style.configure('TLabelframe.Label', background = self.light_grey)
                self.style.configure('TLabel', background = self.light_grey)
                self.style.configure('TFrame', background = self.light_grey)
                self.style.map('entry_fields.TEntry',
                    fieldbackground=[('readonly', self.light_grey)])
                self.style.map('TButton',
                    foreground = [('disabled', self.grey)])
        
### MENU BAR ###
        # File menu
        self.menubar = tk.Menu(self.master)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.filemenu.add_command(label="Open file", accelerator = "Ctrl+R", command = self.open_file)
        self.filemenu.add_command(label="Export to WWB", accelerator = "Ctrl+E", command = self.convert_file)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", accelerator = "Ctrl+Q", command = self.exit_app)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        # Help menu
        self.helpmenu = tk.Menu(self.menubar, tearoff = 0)
        self.helpmenu.add_command(label="License", command = self.show_license_info)
        self.helpmenu.add_command(label="About", command = self.show_about)
        self.menubar.add_cascade(label="Help", menu = self.helpmenu)

        self.master.config(menu = self.menubar)

### GUI LAYOUT ###
        self.paned_window = tk.PanedWindow(orient = "vertical")
        self.canvas = tk.Canvas(self.paned_window, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

### GUI CONTENT ###
        # Frame
        self.label_frame = ttk.LabelFrame(self.canvas, text='Reformat WSM csv scan files to WWB')
        self.label_frame.pack(side = "top", pady=20, ipadx=20)

        #Content lives inside frame
        self.open_file_btn = ttk.Button(self.label_frame, text = "Open file...", command = self.open_file)
        self.open_file_btn.pack(pady = 10)
        self.entry_frame = ttk.Frame(self.label_frame)
        self.entry_frame.pack()
        self.show_file_path_label = ttk.Label(self.entry_frame, text="File path: ")
        self.show_file_path_label.pack(side = "left")
        self.show_file_path = ttk.Entry(self.entry_frame, width = 60, state = "readonly", style = 'entry_fields.TEntry')
        self.show_file_path.pack(side = "left")
        self.convert_btn = ttk.Button(self.label_frame, text = "Convert file...", state = "disabled", command = self.convert_file)
        self.convert_btn.pack(pady = 10)

        # Footer
        self.footer = ttk.Frame(self.master, borderwidth = 4, relief = "ridge")
        self.footer.pack(side = "bottom", fill = "both", expand = False)
        self.status = ttk.Label(self.footer, text = "")
        self.status.pack(side = "left", anchor = tk.SW)
        
        self.paned_window.add(self.canvas)
        self.paned_window.pack(fill = "both", expand = True)

### KEY COMMANDS ###
        self.master.bind('<Control-q>', lambda event: self.exit_app())
        self.master.bind('<Control-r>', lambda event: self.open_file())
        self.master.bind('<Control-e>', lambda event: self.convert_file())
### INTERRUPT [X]-BUTTON TO ASK IF USER WANTS TO CLOSE ###
        self.master.protocol('WM_DELETE_WINDOW', lambda: self.exit_app())

### FUNCTIONS ###
    def is_os_in_darkmode(self):
        if self.running_os == 'win32':
            # Check if Windows is running in dark mode
            try:
                import winreg
            except ImportError:
                return False
            registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            reg_keypath = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
            try:
                reg_key = winreg.OpenKey(registry, reg_keypath)
            except FileNotFoundError:
                return False
            for i in range(1024):
                try:
                    value_name, value, _ = winreg.EnumValue(reg_key, i)
                    if value_name == 'AppsUseLightTheme':
                        return value == 0
                except OSError:
                    break
            return False
        else:
            # Checks if MacOS is running in dark mode
            cmd = 'defaults read -g AppleInterfaceStyle'
            p = subprocess.Popen(cmd, stdout = subprocess.PIPE,
                                stderr = subprocess.PIPE, shell = True)
            return bool(p.communicate()[0])
        
    def update_entry_text(self,entry,text,state):
        # Updates text (file path) in Entry field
        entry.configure(state="normal")
        entry.delete(0,"end")
        entry.insert(0,text)
        entry.configure(state=state)

    def update_status(self, status_msg):
        # Updates status text in footer
        self.status.configure(text = f"{status_msg}")

    def open_file(self):
        # Opens CSV file and checks compatibility
        file = filedialog.askopenfilename(defaultextension = ".csv", filetypes = [("CSV file", "*.csv"), ("All files", "*.*")])
        filenameExtension = file[-4:]
        # If user cancels the open window, there is no file
        if file:
            # Check if file is CSV file
            if filenameExtension == '.csv':
                self.update_status("File opened")
                self.update_entry_text(self.show_file_path, file, "readonly")
                reader = csv.reader(open(file), delimiter=";")
                self.table = list(reader)
                # Checks if file is from WSM by looking at known headers. If cells are different it's probably not from WSM
                if self.table[6] == ['Frequency', 'RF level (%)', 'RF level', 'Memory (%)', 'Memory', 'Squelch (%)', 'Squelch']:
                    # Reformat csv content
                    self.convert_btn.config(state = "normal")
                    del self.table[2728:2741]
                    del self.table[0:7]
                    for i in self.table:
                        del i[2:7]
                        i.insert(1, i[0][:3] + "." + i[0][3:] + ",")
                        del i[0]
                        i.insert(2, str(int(float(i[1])/100*120-120)))
                        del i[1]
                else:
                    self.update_status("Error: This doesn't look like a WSM file...")
                    self.convert_btn.config(state = "disabled")
            else:
                self.update_status('Error: Wrong file type. Only .csv files exported from WSM allowed.')
                self.convert_btn.config(state = "disabled")
        else:
            self.update_status("Operation aborted")
            self.convert_btn.config(state = "disabled")

    def convert_file(self):
        # Saves the reformatted data to a new file. Default is .csv, but user can specify file type
        file = filedialog.asksaveasfilename(defaultextension = ".csv", filetypes = [("CSV file", "*.csv"), ("All files", "*.*")])
        # If user cancels the save window, there is no file
        if file:
            # Write rows to new file
            with open(file, 'w', newline='', encoding='UTF8') as file:
                writer = csv.writer(file, delimiter=' ')
                for row in self.table:
                    writer.writerow(row)
                self.update_status(f"File saved as {file.name}, encoding: {file.encoding}")
        else:
            self.update_status("Operation aborted")

### MESSAGE BOXES ###
    def show_license_info(self):
        # Show license info in a message box
        messagebox.showinfo('License information', f'''MIT License\n\nCopyright (c) 2023 nikharju\n
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.''')

    def show_about(self):
        # Show info about app in message box
        messagebox.showinfo("About program", f"Written by Niklas Harju\nniklas.harju@yahoo.com\n\nVersion: {self.program_version}\n\nDISCLAIMER!\nThis program is in no way affiliated with neither Sennheiser nor Shure corporations, nor is it supported by them. It is a tool created by me personally to help with the task of using Sennheiser hardware and software for scanning radio frequency spectrums and use that data for frequency planning in WWB.\n\nWhat is this?\n\nThis program imports radio frequency scan files created in and exported from Sennheisers software Wireless Systems Manager (WSM) as .csv files and reformats them to a csv format that works with Shures software Wireless Workbench (WWB).\n\nSimply click 'Open file...' to import the WSM file and click 'Convert to WWB...' and choose a location and name to save the new file with the new formatting.\n\nAfter that you can import the newly created scan file into WWB.")

    def exit_app(self):
        # Exit confirmation message box
        response = messagebox.askyesno('Quit?','Are you sure you want to exit the application?')
        if response:
            self.master.destroy()

root = tk.Tk()
app = App(root)
root.mainloop()

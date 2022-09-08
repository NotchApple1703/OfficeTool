import tkinter
import webbrowser
import os
from tkinter import *
from tkinter.ttk import *

def callback(url):
    webbrowser.open_new(url)

def about():
    toplv = Toplevel()
    toplv.title('About OfficeTool')
    toplv.geometry('300x200+100+100')
    toplv.iconbitmap('icon.ico')
    toplv.resizable(0, 0)
    photo = tkinter.PhotoImage(file = 'icon.png')
    imglabel = Label(toplv, image = photo, padding = 1)
    imglabel.pack()
    Label(toplv, text = '- OfficeTool (by Notch Apple) -', font = ('Arial', 12, 'bold')).pack(side = 'top', pady = 10)
    linkfb = Label(toplv, text = 'Facebook: notchapple1703', font = ('Arial', 10, 'underline'), foreground = 'blue', cursor = 'hand2')
    linkfb.pack(pady = 10)
    linkfb.bind("<Button-1>", lambda e: callback('https://www.facebook.com/notchapple1703'))
    linkgithub = Label(toplv, text = 'Github: notchapple1703', font = ('Arial', 10, 'underline'), foreground = 'blue', cursor = 'hand2')
    linkgithub.pack()
    linkgithub.bind("<Button-1>", lambda e: callback('https://www.github.com/notchapple1703'))
    toplv.mainloop()

def install():
    if products.get() == 'Office 365':
        productid = 'O365ProPlusRetail'
    elif products.get() == 'Office 2021':
        productid = 'ProPlus2021Retail'
    elif products.get() == 'Office 2019':
        productid = 'ProPlus2019Retail'
    elif products.get() == 'Office 2016':
        productid = 'ProPlusRetail'

    exlist = ['Access','Bing','Excel','Groove','Lync','OneDrive','OneNote','Outlook','Powerpoint','Publisher','Teams','Word']

    if access.get() == 1:
        exlist.remove('Access')
    if bing.get() == 1:
        exlist.remove('Bing')
    if excel.get() == 1:
        exlist.remove('Excel')
    if groove.get() == 1:
        exlist.remove('Groove')
    if lync.get() == 1:
        exlist.remove('Lync')
    if onedrive.get() == 1:
        exlist.remove('OneDrive')
    if onenote.get() == 1:
        exlist.remove('OneNote')
    if outlook.get() == 1:
        exlist.remove('Outlook')
    if powerpoint.get() == 1:
        exlist.remove('Powerpoint')
    if publisher.get() == 1:
        exlist.remove('Publisher')
    if teams.get() == 1:
        exlist.remove('Teams')
    if word.get() == 1:
        exlist.remove('Word')

    cfg = open('configuration.xml','w')
    cfg.close()
    cfg = open('configuration.xml','a')
    cfg.write('<Configuration>\n')
    cfg.write(f'  <Add OfficeClientEdition="{architect.get()}" Channel="Current">\n')
    cfg.write(f'    <Product ID="{productid}">\n')
    cfg.write('      <Language ID="MatchOS" Fallback="en-us" />\n')
    for exclude in exlist:
        cfg.write(f'      <ExcludeApp ID="{exclude}" />\n')
    cfg.write('    </Product>\n')
    cfg.write('  </Add>\n')
    cfg.write('  <Updates Enabled="true" />\n')
    cfg.write('</Configuration>')

    os.popen('install.bat')

root = Tk()
root.title('OfficeTool (by Notch Apple)')
root.iconbitmap('icon.ico')
root.geometry('350x180+50+50')
root.resizable(0, 0)

menubar = Menu(root)
help = Menu(menubar, tearoff = 0)
menubar.add_cascade(label = 'Help', menu = help)
help.add_command(label = 'About', command = about)

root.config(menu = menubar)

Label(root, text = 'Choose Product:').grid(row = 0, column = 0, sticky = W, padx = 10, pady = 10)

products = StringVar()
product = Combobox(root, textvariable = products, width = 12)

product['value'] = ('Office 365',
                    'Office 2021',
                    'Office 2019',
                    'Office 2016',)

product.grid(row = 0, column = 1)
product.current(0)

access = IntVar()
bing = IntVar()
excel = IntVar()
groove = IntVar()
lync = IntVar()
onedrive = IntVar()
onenote = IntVar()
outlook = IntVar()
powerpoint = IntVar()
publisher = IntVar()
teams = IntVar()
word = IntVar()

accessbtn = Checkbutton(root, text = 'Access', variable = access).grid(row = 1, column = 0, sticky = W, padx = 20)
bingbtn = Checkbutton(root, text = 'Bing', variable = bing).grid(row = 2, column = 0, sticky = W, padx = 20)
excelbtn = Checkbutton(root, text = 'Excel', variable = excel).grid(row = 3, column = 0, sticky = W, padx = 20)
groovetbn = Checkbutton(root, text = 'Groove', variable = groove).grid(row = 4, column = 0, sticky = W, padx = 20)
lyncbtn = Checkbutton(root, text = 'Lync', variable = lync).grid(row = 1, column = 1, sticky = W, padx = 20)
onedrivebtn = Checkbutton(root, text = 'One Drive', variable = onedrive).grid(row = 2, column = 1, sticky = W, padx = 20)
onenotebtn = Checkbutton(root, text = 'One Note', variable = onenote).grid(row = 3, column = 1, sticky = W, padx = 20)
outlookbtn = Checkbutton(root, text = 'Outlook', variable = outlook).grid(row = 4, column = 1, sticky = W, padx = 20)
powerpointbtn = Checkbutton(root, text = 'Powerpoint', variable = powerpoint).grid(row = 1, column = 2, sticky = W, padx = 20)
publisherbtn = Checkbutton(root, text = 'Publisher', variable = publisher).grid(row = 2, column = 2, sticky = W, padx = 20)
teamsbtn = Checkbutton(root, text = 'Teams', variable = teams).grid(row = 3, column = 2, sticky = W, padx = 20)
wordbtn = Checkbutton(root, text = 'Word', variable = word).grid(row = 4, column = 2, sticky = W, padx = 20)

architect = StringVar()
Label(root, text = 'Architecture:').grid(row = 5, column = 0, sticky = W, padx = 10, pady = 10)
_64bit = Radiobutton(root, text = 'x64', variable = architect, value = 64).grid(row = 5, column = 1, sticky = W, padx = 20)
_32bit = Radiobutton(root, text = 'x32', variable = architect, value = 32).grid(row = 5, column = 2, sticky = W, padx = 20)

Button(root, text = 'Install', command = install).grid(row = 0, column = 2, padx = 5, pady = 5)

root.mainloop()
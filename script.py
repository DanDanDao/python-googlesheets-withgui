import pygsheets
from tkintertoy import Window

gui = Window()
gui.setTitle('Pick two ports for your NP class!')
gui.addEntry('stu_no', 'Type in your student number')
gui.addEntry('port_1', 'Pick port 1')
gui.addEntry('port_2', 'Pick port 2')
gui.addLabel('message', '')
gui.addButton('commands')
gui.plot('stu_no', row=0)
gui.plot('port_1', row=1)
gui.plot('port_2', row=2)
gui.plot('message', row=3)
gui.plot('commands', row=4, pady=10)

#spreadsheet information
SPREADSHEET_ID = "1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU" # <Your spreadsheet ID>
RANGE_NAME = "data" # <Your worksheet name>

#setup google spreadsheet
gc = pygsheets.authorize(service_file='creds.json')
sh = gc.open_by_key(SPREADSHEET_ID)

#import spreadsheet to dataframe
wks = sh.sheet1
df = wks.get_as_df(has_header=True, index_colum=1, empty_value="")

def checkStuNo(student_no, port_1, port_2):
    # working
    if (student_no not in df.index):
        gui.set('message', "Your student number is not available")
    else:
        addPorts(student_no, port_1, port_2)


def addPorts(student_no, port_1, port_2):
    # working
    if (int(port_1) == int(port_2)):
        gui.set('message', "Port values must be different")
    if (int(port_1) < 61000 or int(port_1) > 69000 or int(port_2) < 61000 or int(port_2)):
        gui.set('message', "Port values must be in between 61000 and 61999")
    if (int(port_1) in df.values):
        gui.set('message', "Port 1 has been taken")
    elif (int(port_2) in df.values):
        gui.set('message', "Port 2 has been taken")
    else:
        df.loc[student_no, "port_1"] = port_1
        df.loc[student_no, "port_2"] = port_2
        gui.set('message', "Pick ports successfully")

while True:
    gui.waitforUser()
    if gui.content:
        checkStuNo(gui.get('stu_no'), gui.get('port_1'), gui.get('port_2'))
    else:
        break


#Write to google spreadsheet
wks.set_dataframe(df,(1,1), copy_index=True)
header = wks.cell('A1')
header.value = 'student_no'
header.update()

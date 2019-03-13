import pygsheets
from tkintertoy import Window
import smtplib
import email

#Connect to server
server = smtplib.SMTP('smtp.outlook.com:587')
server.starttls()

#Set up GUI
gui = Window()
gui.setTitle('Pick two ports for your NP class!')
gui.addEntry('stu_no', 'Type in your student number',width=30)
gui.addEntry('pass', 'Type in your student account\'s password',width=30, show="*")
gui.addEntry('port_1', 'Pick port 1',width=30)
gui.addEntry('port_2', 'Pick port 2',width=30)
gui.addLabel('message', '',width=30)
gui.addButton('commands')
gui.plot('stu_no', row=0, pady=10)
gui.plot('pass', row=1, pady=10)
gui.plot('port_1', row=2, pady=10)
gui.plot('port_2', row=3, pady=10)
gui.plot('message', row=4, pady=10)
gui.plot('commands', row=5, pady=10)

#spreadsheet information
SPREADSHEET_ID = "1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU" # <Your spreadsheet ID>
RANGE_NAME = "data" # <Your worksheet name>

#setup google spreadsheet
gc = pygsheets.authorize(service_file='creds.json')
sh = gc.open_by_key(SPREADSHEET_ID)

#import spreadsheet to dataframe
wks = sh.sheet1
df = wks.get_as_df(has_header=True, index_colum=1, empty_value="")


#Check if student number is valid
def checkStuNo(student_no, port_1, port_2, password):
    # working
    if (student_no not in df.index):
        gui.set('message', "Your student number is not correct")
    else:
        try:
            server.login(student_no + '@student.rmit.edu.au', password)
        except smtplib.SMTPAuthenticationError:
            gui.set('message', "Your password is incorrect")
        else:
            addPorts(student_no, port_1, port_2)

#Add ports
def addPorts(student_no, port_1, port_2):
    # working
    if (int(port_1) == int(port_2)):
        gui.set('message', "Two port values must be different")
    elif (int(port_1) < 61000 or int(port_1) > 61999 or int(port_2) < 61000 or int(port_2) > 61999):
        gui.set('message', "Port value must be in between 61000 and 61999")
    elif (int(port_1) in df.values):
        gui.set('message', "Port 1 has been taken")
    elif (int(port_2) in df.values):
        gui.set('message', "Port 2 has been taken")
    else:
        df.loc[student_no, "port_1"] = port_1
        df.loc[student_no, "port_2"] = port_2
        gui.set('message', "Pick ports successfully")
        # Message for email
        msg = '\r\n' + 'Student: ' + student_no + ' picks: ' + port_1 + ' and ' + port_2
        server.sendmail(student_no + '@student.rmit.edu.au', 'guangdanny@gmail.com', msg)
        server.quit()
while True:
    gui.waitforUser()
    if gui.content:
        if (gui.get('stu_no') and gui.get('port_1') and gui.get('port_2') and gui.get('pass')):
            checkStuNo(gui.get('stu_no'), gui.get('port_1'), gui.get('port_2'), gui.get('pass'))
        else:
            gui.set('message', "Please fill every boxes")
    else:
        break


#Write to google spreadsheet
wks.set_dataframe(df,(1,1), copy_index=True)
header = wks.cell('A1')
header.value = 'student_no'
header.update()


import pandas as pd
import pygsheets


#spreadsheet information
SPREADSHEET_ID = "1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU" # <Your spreadsheet ID>
RANGE_NAME = "data" # <Your worksheet name>


gc = pygsheets.authorize(service_file='creds.json')
sh = gc.open_by_key(SPREADSHEET_ID)

wks = sh.sheet1
df = wks.get_as_df(has_header=True, index_colum=1, empty_value="")

add = "61002"
student_id = "s3438953"

#working
if (student_id in df.index):
    df.loc["s343887", "port_1"] = "in dataframe"
else:
    df.loc["s343887", "port_1"] = "not in dataframe"


#not working
if (add in df.values):
    df.loc["s343889", "port_2"] = "duplicated"
else:
    df.loc["s343889", "port_2"] = "available"


#Write to google spreadsheet
wks.set_dataframe(df,(1,1), copy_index=True)

header = wks.cell('A1')
header.value = 'student_no'
header.update()
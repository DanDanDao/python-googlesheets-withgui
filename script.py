import pandas as pd
from io import BytesIO
import requests

# r = requests.get('https://docs.google.com/spreadsheet/ccc?key=1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU&output=csv')
# data = r.content


df = pd.read_csv('https://docs.google.com/spreadsheets/d/' +
                   '1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU' +
                   '/export?gid=0&format=csv', header = 0).set_index("student_no", inplace = True)


df.head()

# print(df.loc["s343887","port_1"])
#
df.loc["s343887","port_1"] = "test success"
#
# print(df.loc["s343887","port_1"])

df.to_csv('https://docs.google.com/spreadsheets/d/' +
                 '1mn93Hl0zT7Cb2YaPryj523IQEzelVa1HWLrR3ZxrJRU' +
                 '/edit#gid=614041454&format=csv')
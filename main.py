from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
import os
import time
from pywinauto.timings import WaitUntil

# set paths in one spot to simplify life.
script_folder_path = os.path.dirname(os.path.realpath(__file__))
files_to_convert_folder = os.path.abspath(os.path.join(script_folder_path,'files_to_convert'))

# Get list of hse files names into an array/list.
hse_list = []
for file in os.listdir(files_to_convert_folder):
    if file.endswith(".HSE"):
        #strip extention and add to
        hse_list.append(os.path.splitext(file)[0])
#Output list of HSE prefixes to console.
print(*hse_list,  sep = "\n")

# Will create an iterator later.. for now just using first HSE file.
hse_name = hse_list[0]
app = Application().start("C:\HOT2000 v11.3\HOT2000.exe")
app.HOT2000.menu_select("File->Open")

#hardcoding input hse folder path for now.
app.Open.Edit.type_keys("C:\\Users\\plopez\\PycharmProjects\\h2k_upgrader\\files_to_convert\\{}.HSE".format(hse_name))
app.Open.Open.click()

# spams enter to get by errors and messages. Hopefully 8 is enough?
send_keys('{ENTER 8}')
# Sleep while H2k does its thing.
time.sleep(5.0)
# Hot 2000 dialog name has changed.
new_name = "HOT2000 - [{} - General]".format(hse_name)
app[new_name].menu_select("File->Save As")
#hardcoding ouput h2k folder path for now.
app.SaveAs.Edit.type_keys("C:\\Users\\plopez\\PycharmProjects\\h2k_upgrader\\converted_files\\{}.h2k".format(hse_name))
app.SaveAs.Save.click()




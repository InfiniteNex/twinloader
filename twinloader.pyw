import win32com.client # Outlook COM
import getpass
import clipboard
import keyboard
import winsound
import win32gui
from pyWinActivate import win_activate

title = "RPG Assignment2.1.xlsm  -  Shared"

user = getpass.getuser() + "@gfk.com"

# def window_enum_handler(hwnd, resultList):
#     if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
#         resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

# def get_app_list(handles=[]):
#     mlst=[]
#     win32gui.EnumWindows(window_enum_handler, handles)
#     for handle in handles:
#         mlst.append(handle)
#     return mlst

def pool_rpg_mails():
        new_dist_codes = []

        print("Pooling")
        #access items in the TWIN folder
        Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") #Opens Microsoft Outlook
        folder = Outlook.Folders[user] #User mails
        subfolder = folder.Folders['Inbox']#Inbox
        subfolder2 = subfolder.Folders['TWIN'] #TWIN folder
        messages = subfolder2.Items #items object

        #loop trough the folder and get any unread message's code and mark it as read and categorize
        last_message = messages.GetLast() #last recieved message
        

        # get the last message unique identifiers, to use as a timestamp
        body = last_message.Body
        split = str(body).splitlines()
        timestamp = [str(last_message), split[5][0:10], split[5][11:19]]

        # loop through the folder
        # try:
        for i in messages:
            if last_message.Categories != "RPG":
                #get dist code from message name
                x = str(last_message).split("(")
                y = x[1].split(")")

                # check if value taken is not TWIN isntead of dist code
                if y[0].isalpha():
                    y = x[2].split(")")

                # continue script with extracted code
                if y[0] not in new_dist_codes:
                    new_dist_codes.append(y[0])
                #mark as read and categorize
                last_message.Unread = False
                last_message.Categories = "RPG"
                last_message.Save()
            try:
                #Go to next unread messages if any
                last_message = messages.GetPrevious()
            except:
                pass
        # except:
        #     print("No more unread messages in this folder.")

        if not new_dist_codes:
            winsound.Beep(120, 300)
            winsound.Beep(100, 300)
            winsound.Beep(80, 300)
            print("No unread messages in this folder.")

        else:
            #clean the codes from unwanted symbols
            codes = str(new_dist_codes)
            codes = codes.replace(",", "\n")       
            codes = codes.replace('\'', '')
            codes = codes.replace('[', '')
            codes = codes.replace(']', '')


            try:
                win_activate(window_title=title, partial_match=True)
                # handle = win32gui.FindWindow(0, "RPG Assignment2.1.xlsm  -  Shared - Excel")  #//paassing 0 as I dont know classname 
                # win32gui.ShowWindow(handle, True)
                # win32gui.SetForegroundWindow(handle)  #//put the window in foreground

            except:
                pass
                # handle = win32gui.FindWindow(0, "RPG Assignment2.1.xlsm  -  Shared - Saved")
                # win32gui.ShowWindow(handle, True)
                # win32gui.SetForegroundWindow(handle)  #//put the window in foreground
                
            # win32gui.ShowWindow(handle, True)
            # win32gui.SetForegroundWindow(handle)  #//put the window in foreground

            clipboard.copy(codes)
            keyboard.press_and_release("f2")

            #successful code execution
            winsound.Beep(120, 300)
            winsound.Beep(140, 300)
            winsound.Beep(160, 300)


pool_rpg_mails()

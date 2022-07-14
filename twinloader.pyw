import win32com.client # Outlook COM
import getpass
import clipboard
import keyboard
import winsound
import win32gui
import pyWinActivate as pw


title = "RPG Assignment2.1.xlsm"
user = getpass.getuser() + "@gfk.com"
window_found = False

def pool_rpg_mails():
        new_dist_codes = []

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

        if not new_dist_codes:
            pass

        else:
            #clean the codes from unwanted symbols
            codes = str(new_dist_codes)
            codes = codes.replace(",", "\n")       
            codes = codes.replace('\'', '')
            codes = codes.replace('[', '')
            codes = codes.replace(']', '')


            try:
                pw.win_activate(window_title=title, partial_match=True)
            except:
                pass

            clipboard.copy(codes)
            keyboard.press_and_release("f2")


def check_for_open_excel_window():
    window_found = pw.check_win_exist(window_title=title)
    if window_found:
        pool_rpg_mails()


check_for_open_excel_window()



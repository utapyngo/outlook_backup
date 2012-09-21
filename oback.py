'''User stories:

As a cautious Outlook user
I want to backup my Contacts, Calendar and Mail
So I don't lose them.
'''
# Requirements:
# The program should ask for the mailbox name to backup
# The program should automatically create Backup PST file
# The program should automatically choose name of the Backup PST file
# The program should create the Backup PST file in current directory
# The program should backup all Contacts folders
# The program should backup all Calendar folders
# The program should backup all Mail folders

import time
import sys
import win32com.client
from pywintypes import com_error


def main(args):
    '''If first argument is not empty it should be
    the name of the mailbox you want to backup.
    If it is empty, you will be promted to select
    a folder you want to backup.
    '''
    source_folder_name = None
    if len(args) > 0:
        source_folder_name = args[0]

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # Find source folder by name
    source_folder = None
    if source_folder_name:
        for i in range(1, 1 + len(ns.Folders)):
            folder = ns.Folders[i]
            if folder.Name == source_folder_name:
                source_folder = folder

    # Select Folder
    if not source_folder:
        source_folder = ns.PickFolder()

    if not source_folder:
        return
        
    # Create a Backup PST file
    file_name = 'Backup {0}.pst'.format(time.strftime('%Y-%m-%d %H%M%S'))
    ns.AddStore(file_name)
    try:
        backup_folder = ns.Session.Folders.GetLast()
        
        # Do backup
        subfolders = source_folder.Folders
        for i in range(1, 1 + len(subfolders)):
            subfolder = subfolders[i]
            print subfolder.Name
            try:
                subfolder.CopyTo(backup_folder)
            except com_error as e:
                for arg in e.args:
                    print arg
                continue
    finally:        
        ns.RemoveStore(backup_folder)

if __name__ == '__main__':
    main(sys.argv[1:])

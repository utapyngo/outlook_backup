# encoding: utf-8

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

from __future__ import print_function

import time
import sys
import os

import win32com.client
from pywintypes import com_error
win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 4)

def main(args):
    '''If first argument is not empty it should be
    the name of the mailbox you want to backup.
    If it is empty, you will be prompted to select
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
    source_folder_name = source_folder.Name
    file_name = 'Backup of {0} at {1}.pst'.format(source_folder_name, time.strftime('%Y-%m-%d %H-%M-%S'))
    file_name = os.path.abspath(file_name)
    print('Creating file {0}'.format(file_name))
    try:
        ns.AddStore(file_name)
    except com_error as e:
        print('Error: ' + str(e.excepinfo[2]))
        raise
    try:
        backup_folder = ns.Session.Folders.GetLast()
        
        # Do backup
        subfolders = source_folder.Folders
        for i in range(1, 1 + len(subfolders)):
            subfolder = subfolders[i]
            print(subfolder.Name)
            try:
                subfolder.CopyTo(backup_folder)
            except com_error as e:
                print('Error: ' + unicode(e.excepinfo[2]))
                continue
    finally:        
        ns.RemoveStore(backup_folder)

if __name__ == '__main__':
    main(sys.argv[1:])

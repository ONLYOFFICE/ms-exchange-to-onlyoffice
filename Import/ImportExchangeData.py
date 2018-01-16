#!/usr/bin/python
import pkg_resources, sys
from pkg_resources import DistributionNotFound, VersionConflict

def exception_handler(exception_type, exception, traceback):
    print "%s: %s" % (exception_type.__name__, exception)

sys.excepthook = exception_handler

dependencies = [
  'requests',
  'twisted',
  'ArgumentParser',
  'progressbar'
]

pkg_resources.require(dependencies)


import os, subprocess, json, string
from argparse import ArgumentParser
from getpass import getpass

parser = ArgumentParser(description="""Import data from MS Exchange into ONLYOFFICE""")
parser.add_argument('-s', dest='portal_host_scheme', required=False, default='http', help='portal url scheme')
parser.add_argument('-d', dest='portal_host', required=True, help='portal domain/ip')
parser.add_argument('-p', dest='portal_port', required=False, default='', help="portal port")
parser.add_argument('-u', dest='portal_admin', required=True, help='portal admin user account')
parser.add_argument('-pw', dest='portal_password', help='portal admin password will be prompted for if not provided')
parser.add_argument('-f', dest='data_folder', required=True, help="the folder with exported data to import")
 
args = parser.parse_args()

if not args.portal_password:
    args.portal_password = getpass()

def main():
    SCRIPTS_DIR = "./lib/"
    SCRIPT_PST2MBOX = os.path.join(SCRIPTS_DIR, "pst2mbox.sh")
    SCRIPT_CREATE_USERS = os.path.join(SCRIPTS_DIR, "create_users.py")
    SCRIPT_INSTALL_PASSFINDER = os.path.join(SCRIPTS_DIR, "install-passfinder.sh")
    SCRIPT_GET_MAILBOXES = os.path.join(SCRIPTS_DIR, "get-mailboxes.sh")
    SCRIPT_MBOX_IMPORT = os.path.join(SCRIPTS_DIR, "mbox2imap.py")
    MBOX_IMPORT_MAPPING = os.path.join(SCRIPTS_DIR, 'mapping.json')
    DATA_FOLDER = args.data_folder
    USERS_FILE = os.path.join(DATA_FOLDER, 'users.csv')
    PST_FOLDER = os.path.join(DATA_FOLDER, 'PST')

    MBOX_FOLDER = "./mbox"
    MAILBOXES_FILE = "./mailboxes.json"

    PORTAL_SCHEME = args.portal_host_scheme
    PORTAL_HOST = args.portal_host
    PORTAL_PORT = args.portal_port
    PORTAL_ADMIN_LOGIN = args.portal_admin
    PORTAL_ADMIN_PASSWORD = args.portal_password

    print('\n___ VARIABLES ___\n')
    # print("PORTAL_SCHEME: " + PORTAL_SCHEME)
    print("PORTAL_HOST: " + PORTAL_HOST)
    if PORTAL_PORT:
        print("PORTAL_PORT: " + PORTAL_PORT)
    print("PORTAL_ADMIN_LOGIN: " + PORTAL_ADMIN_LOGIN)
    # print("PORTAL_ADMIN_PASSWORD: " + PORTAL_ADMIN_PASSWORD)
    print("EXPORT_DATA_PATH: " + DATA_FOLDER)
    print("USERS_FILE_PATH: " + USERS_FILE)
    print("PST_FOLDER: " + PST_FOLDER)

    print("\n___ CHECK REQUIREMENTS ___\n")

    IS_OK = True

    if not os.path.isdir(DATA_FOLDER):
        print("ERROR: Directory '%s' does not exist\n" % DATA_FOLDER)    

    if not os.path.isfile(USERS_FILE):
        print("ERROR: File '%s' does not exist\n" % USERS_FILE)

    if not os.path.isdir(PST_FOLDER):
        print("ERROR: Directory '%s' does not exist\n" % PST_FOLDER)
    else:
        if not os.listdir(PST_FOLDER):
            print("ERROR: Directory '%s' is empty\n" % PST_FOLDER)

    if not os.path.isdir(SCRIPTS_DIR):
        print("ERROR: Directory '%s' does not exist\n" % SCRIPTS_DIR)

    if not os.path.isfile(SCRIPT_PST2MBOX):
        print("ERROR: File '%s' does not exist\n" % SCRIPT_PST2MBOX)    

    if not os.path.isfile(SCRIPT_CREATE_USERS):
        print("ERROR: File '%s' does not exist\n" % SCRIPT_CREATE_USERS)    

    if not os.path.isfile(SCRIPT_INSTALL_PASSFINDER):
        print("ERROR: File '%s' does not exist\n" % SCRIPT_INSTALL_PASSFINDER)    

    if not os.path.isfile(SCRIPT_GET_MAILBOXES):
        print("ERROR: File '%s' does not exist\n" % SCRIPT_GET_MAILBOXES)    

    if not os.path.isfile(SCRIPT_MBOX_IMPORT):
        print("ERROR: File '%s' does not exist\n" % SCRIPT_MBOX_IMPORT)    

    if not IS_OK:
        sys.exit(2)

    print("OK")

    print("\n___ START ___\n")

    print("\n(1/5) Create users with mailboxes\n")
    subprocess.check_call([SCRIPT_CREATE_USERS, "-s", PORTAL_SCHEME, "-d", PORTAL_HOST, "-p", PORTAL_PORT, '-u', PORTAL_ADMIN_LOGIN, "-pw", PORTAL_ADMIN_PASSWORD, "-f", USERS_FILE])

    print("\n(2/5) Convert PST files to mbox format\n")
    subprocess.check_call([SCRIPT_PST2MBOX, PST_FOLDER])

    if not os.path.isdir(MBOX_FOLDER):
        print("ERROR: Directory '%s' does not exist\n" % MBOX_FOLDER)
        sys.exit(2)
    else:
        if not os.listdir(MBOX_FOLDER):
            print("ERROR: Directory '%s' is empty\n" % MBOX_FOLDER)
            sys.exit(2)
   
    print("\n(3/5) Install ASC.Mail.PasswordFinder inside docker Onlyoffice Community Server\n")
    subprocess.check_call([SCRIPT_INSTALL_PASSFINDER, SCRIPTS_DIR])
    
    print("\n(4/5) Get mailboxes.json with settings for connecting to mailboxes\n")
    subprocess.check_call([SCRIPT_GET_MAILBOXES, "."])
    
    if not os.path.isfile(MAILBOXES_FILE):
        print("ERROR: File '%s' does not exist\n" % MAILBOXES_FILE)
        sys.exit(2)

    print("\n(5/5) Import mbox files to mailboxes by IMAP\n")
    
    jsonfile = open(MAILBOXES_FILE, "r")

    mailboxes = json.load(jsonfile)
    
    jsonfile.close()

    if (not mailboxes) or (len(mailboxes) == 0):
        print("mailboxes.json is empty")
        sys.exit(2)

    print("Found {0} mailboxes to import\n".format(len(mailboxes))) 

    def findFolder(email):
        localPart = email.split("@")[0]
        #print("try find folder with name '{0}' in path '{1}'".format(localPart, MBOX_FOLDER))
        for name in os.listdir(MBOX_FOLDER):
                    #print(name)
                    full = os.path.join(MBOX_FOLDER, name)
                    if os.path.isdir(full) and name.lower() == localPart.lower():
                        #print("Found path '{0}' to import by IMAP".format(full))
                        return full

    for i, mb in enumerate(mailboxes):
        email = mb['email']
        for j, s in enumerate(mb["settings"]):
            if s['type'] == "IMAP":
               path = findFolder(email)
               if path:
                   print("Importing mbox files from {0} to {1}".format(path, email))
                   host = s['host']
                   login = s['login']
                   password = s['password']
                   folder = path
                   print(host, login, password, folder)       
                   subprocess.check_call([SCRIPT_MBOX_IMPORT, "-s", host, "-u", login, "-p", password, '-m', MBOX_IMPORT_MAPPING, "folder", folder])
                
    print("___ END ___\n")

if __name__ == "__main__":
   main()

#!/usr/bin/python
import sys, getopt, datetime, socket
import requests, json, random, string, csv

from argparse import ArgumentParser
from getpass import getpass

def generate_password():
    chars = string.ascii_uppercase + string.ascii_lowercase + string.digits + '!@#$%^&*()'
    size = random.randint(8, 12)
    return ''.join(random.choice(chars) for x in range(size))

def findDomain(domain, domains):
    for index, item in enumerate(domains):
        if item["name"] == domain:
            return item
    return None

def createOnlyofficeMailbox(email, email_name, domains, api_profile_url, api_create_mailbox_url, headers):
    try:
        print("Check if email '{0}' need to be created on Onlyoffice MailServer\n".format(email))

        splited = email.split("@")

        local_part = splited[0]
        
        domain = splited[1];

        domainJson = findDomain(domain, domains);

        if(domainJson != None):
            print("Email domain '{0}' found on Onlyoffice MailServer\n".format(domain))

            response = requests.get(api_profile_url + "?email=" + email, headers=headers).json()
            
            #print(response)
            
            profile = response["response"]

            user_id = profile["id"]

            domain_id = domainJson["id"]

            data = {"name": email_name, "local_part": local_part, "domain_id": domain_id, "user_id": user_id}

            print("Try create mailbox {0} on Onlyoffice MailServer: {1}".format(email, data))

            print('')

            response = requests.post(api_create_mailbox_url, data=data, headers=headers).json()

            error = response.get("error")

            if(error == None):
                print("Mailbox '{0}'' is created successfully".format(email))
            else:
                print("Mailbox '{0}' is not created: Error: {1}".format(email, error.get("message")))

            #print(response)

        else:
            print("Email domain '{0}' not exists on Onlyoffice MailServer".format(domain))
    except:
        print("Mailbox '{0}' is not created: Error: {1}".format(email, "Unknown error has happend"))

def exception_handler(exception_type, exception, traceback):
    print "%s: %s" % (exception_type.__name__, exception)

sys.excepthook = exception_handler

parser = ArgumentParser(description="""Import data from MS Exchange into ONLYOFFICE""")
parser.add_argument('-s', dest='portal_host_scheme', required=False, default='http', help='portal url scheme')
parser.add_argument('-d', dest='portal_host', required=True, help='portal domain/ip')
parser.add_argument('-p', dest='portal_port', required=False, default='', help="portal port")
parser.add_argument('-u', dest='portal_admin', required=True, help='portal admin user account')
parser.add_argument('-pw', dest='portal_password', help='portal admin password will be prompted for if not provided')
parser.add_argument('-f', dest='users_file', required=True, help="the folder with exported data to import")

args = parser.parse_args()

#Get Command Line Arguments
def main():
    PORTAL_SCHEME = args.portal_host_scheme
    PORTAL_HOST = args.portal_host
    PORTAL_PORT = args.portal_port
    PORTAL_ADMIN_LOGIN = args.portal_admin
    PORTAL_ADMIN_PASSWORD = args.portal_password
    USERS_FILE = args.users_file

    if not PORTAL_SCHEME:
        PORTAL_SCHEME = "http"

    if PORTAL_HOST == '':
        print("Empty portal domain. Use: -? | -h | --help")
        sys.exit(2)

    try:
        socket.gethostbyname(PORTAL_HOST)
    except socket.error:
        print("ERROR: Failed to resolve portal domain '{0}'".format(PORTAL_HOST))
        sys.exit(2)

    if PORTAL_ADMIN_LOGIN == '':
        print("ERROR: Empty portal login. Use: -? | -h | --help")
        sys.exit(2)

    if PORTAL_ADMIN_PASSWORD == '':
        print("ERROR: Empty portal password. Use: -? | -h | --help")
        sys.exit(2)

    if USERS_FILE == '':
        print("ERROR: Empty path to user.csv. Use: -? | -h | --help")
        sys.exit(2)

    print('Base parameters:')
    print("PORTAL_SCHEME: " + PORTAL_SCHEME)
    print("PORTAL_HOST: " + PORTAL_HOST)

    if PORTAL_PORT:
        print("PORTAL_PORT: " + PORTAL_PORT)

    print("PORTAL_ADMIN_LOGIN: " + PORTAL_ADMIN_LOGIN)
    # print("PORTAL_ADMIN_PASSWORD: " + PORTAL_ADMIN_PASSWORD)
    print("USERS_FILE: " + USERS_FILE)

    print('')

    if PORTAL_PORT:
        API_BASE="{0}://{1}{2}".format(PORTAL_SCHEME, PORTAL_HOST, ':{0}'.format(PORTAL_PORT))
    else:
        API_BASE="{0}://{1}".format(PORTAL_SCHEME, PORTAL_HOST)

    API_AUTH_URL="{0}/api/2.0/authentication.json".format(API_BASE)
    API_GET_PROFILE_URL="{0}/api/2.0/people/email".format(API_BASE)
    API_GET_DOMAINS_URL="{0}/api/2.0/mailserver/domains/get.json".format(API_BASE)
    API_CREATE_MAILBOX_URL="{0}/api/2.0/mailserver/mailboxes/add".format(API_BASE)
    API_CREATE_USER_URL="{0}/api/2.0/people/active.json".format(API_BASE)

    print('Base urls:')
    print("API_AUTH_URL: " + API_AUTH_URL)
    print("API_GET_PROFILE_URL: " + API_GET_PROFILE_URL)
    print("API_GET_DOMAINS_URL: " + API_GET_DOMAINS_URL)
    print("API_CREATE_MAILBOX_URL: " + API_CREATE_MAILBOX_URL)
    print("API_CREATE_USER_URL: " + API_CREATE_USER_URL)

    print('')

    data={"userName":PORTAL_ADMIN_LOGIN,"password":PORTAL_ADMIN_PASSWORD}

    response = requests.post(API_AUTH_URL, data=data).json()

    if (not response) or ("response" not in response) or ("token" not in response["response"]):
        print("ERROR: Auth failed")
        print(response)
        sys.exit(2)

    authToken = response["response"]["token"]

    print("Found auth token " + authToken)

    print('')

    headers = {'Authorization': authToken}

    response = requests.get(API_GET_DOMAINS_URL, headers=headers).json()

    if (not response) or ("response" not in response):
        print("ERROR: Get domains")
        print(response)
        sys.exit(2)

    domains = response["response"]

    for index, domain in enumerate(domains):
        print("Found '{0}' mail domain".format(domain["name"]))

    print('')

    csvfile = open(USERS_FILE, 'r')

    fieldnames = ("DisplayName","EmailAddress","EmailName")
    
    reader = csv.DictReader(filter(lambda row: row[0]!='#', csvfile), fieldnames, delimiter=',', quotechar='"')
    
    next(reader, None)  # skip the headers
    
    user_str = json.dumps([ row for row in reader ])

    #print user_str

    now = datetime.datetime.now()

    outfilesuccess="results-ok.txt"
    f_ok = open(outfilesuccess, 'a')
    f_ok.write("#############{0}###################\n".format(now))

    outfileerror="results-err.txt"
    f_err = open(outfileerror, 'a')
    f_err.write("#############{0}###################\n".format(now))

    i=0
    users = json.loads(user_str)
    count=len(users)
    for user in users:
        email=user["EmailAddress"]
        try:
            splited=user["DisplayName"].split(" ")

            name=splited[0]

            if len(splited) > 1:
                lastname = splited[1] 
            else: 
                lastname = splited[0]

            password=generate_password()

            payload={"email":email, "firstname":name, "lastname":lastname, "password":password }

            i+=1

            print("#{0}/{1}: Try create user {2}".format(i, count, payload))
            print('')

            response = requests.post(API_CREATE_USER_URL, data=payload, headers=headers).json()

            if 'error' in response:
                f_err.write("{0}\t{1}\t{2}\t{3}\tERROR:{4}\n".format(name, lastname, email, password, response["error"]["message"]))
                print("User {0} not created: Error: {1}\n".format(email, response["error"]["message"]))
            else:
                f_ok.write("{0}\t{1}\t{2}\t{3}\tOK\n".format(name, lastname, email, password))
                print("User {0} created successfully".format(email))
                print('')
                createOnlyofficeMailbox(email, user["EmailName"], domains, API_GET_PROFILE_URL, API_CREATE_MAILBOX_URL, headers)
        except:
            f_err.write("{0}\t{1}\t{2}\t{3}\tERROR:{4}\n".format(name, lastname, email, password, "Unknown error has happend"))
            print("User {0} not created: Error: {1}".format(email, "Unknown error has happend"))

        print('')

    print('\nAll done')

if __name__ == "__main__":
   main()

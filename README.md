Migrating MS Exchange data to ONLYOFFICE
========================================

Introduction
------------

This article will show how you can transfer the data from MS Exchange to
ONLYOFFICE. Currently the following data types are supported for
transfer:

-   users,
-   mailboxes,
-   mail messages.

In the next releases we are going to add the support for the following
data:

-   calendars,
-   contacts,
-   tasks.

Export data from MS Exchange
----------------------------

##### Running the necessary service

First of all you need to make sure that the **Microsoft Exchange Mailbox
Replication** (or **MSExchangeMailboxReplication**) service is launched.
Run the following command in cmd console:

    Get-Service -name MSExchangeMailboxReplication

If the service you need is running, the command result will be like
this:

    Status   Name               DisplayName
    ------   ----               -----------
    Running  MSExchangeMailb... Microsoft Exchange Mailbox Replication

Or you can go to the Windows **Control Panel** - **Administrative
Tools** - **Services**, find the **Microsoft Exchange Mailbox
Replication** and run it.

##### Assigning the rights

Assign the administrative rights to the user who is going to do the
export of the mail boxes from MS Exchange:

    New-ManagementRoleAssignment –Role "Mailbox Import Export" –User <user name>

Where `<user name>` is the user name who is going to export the data. If
you are going to do that yourself, then assign this role to your
account.

    New-ManagementRoleAssignment –Role "Mailbox Import Export" –User John

After that restart the Exchange Web Services (EWS) console with the
administrator rights. This is done clicking the EWS icon in the start
menu with the right mouse button and selecting the **Run as
administrator** option.

##### Running the script

Now download the script which will do everything you need for the
correct export of the data from MS Exchange. The script is available
[here](https://bit.ly/2zLi1d1). Once downloaded, run it in the EWS
console:

    .\ExportExchangeData.ps1 -dir "C:\Temp"

Where `.\ExportExchangeData.ps1` is the path to the script, and
`-dir "C:\Temp"` is the path to the folder which will be used to export
the files.

Please note that during the script work the folder set in the `-dir`
parameter will be shared for **everyone**. This is necessary for the
correct work of the `New-MailboxExportRequest` command. After the script
finishes its work, the sharing will be removed from the folder.

The folder with the exported files then need to be transferred to the
computer with **ONLYOFFICE** installed.

If you need to make sure that the created `PST` files have the correct
data in them, you can use the free
[pst-viewer](https://www.nucleustechnologies.com/pst-viewer.html) tool
to do that.

What the ExportExchangeData.ps1 script does:

1.  First of all, the script creates the folder set with the `-dir`
    parameter above. In case the parameter is missing, the default
    `C:\Temp` folder will be used.
2.  Once created, the folder is shared for everyone, so that the
    `New-MailboxExportRequest` command could work correctly.
3.  Then the script checks if the file for the user list has been
    previously created, deletes it if found and the user list is
    exported to it using the command:

        Get-Mailbox | Select DisplayName, PrimarySmtpAddress, Alias | Export-Csv $usersListPath

4.  After that the folder for the mail boxes is created (the old one
    will be removed, if present), and the user mail box export starts.
    It will take some time depending on the number of the users, their
    boxes and the data present in their mail messages. The boxes and
    messages will be saved into the
    [PST](https://en.wikipedia.org/wiki/Personal_Storage_Table) format
    to the `\PST\` folder (it will look like `\\ServerName\Temp\PST\` to
    the other computers over the network) of the directory previously
    set with the `-dir` command.
5.  Finally the share is removed from the folder and the script finishes
    its work.

During the script ExportExchangeData.ps1 work some errors might occur. Below are the most
frequent of them and the ways to solve the issues.

-   ###### Error \#1

    The name must be unique per mailbox. There isn't a default name
    available for a new request owned by mailbox
    'mailbox.local/folder/user'. Please clean up existing requests by
    using the Remove cmdlet or specify a unique name. + CategoryInfo :
    InvalidArgument:
    (mailbox.local/folder/user:MailboxOrMailUserIdParameter)[New-MailboxExportRequest],
    NoAvailableDefaultNamePermanentException + FullyQualifiedErrorId :
    [Server=ServerName,RequestId=4a67e451-556c-4ba0-9ab9-9d2ce8a120ff,TimeStamp=11/7/2017
    9:46:58 AM]
    [FailureCategory=Cmdlet-NoAvailableDefaultNamePermanentException]
    7B0B3FF9,
    Microsoft.Exchange.Management.Migration.MailboxReplication.MailboxExportRequest.NewMailboxExportRequest
    + PSComputerName : servername.mailbox.local

    ###### Solution

    Run the following command:

        Get-MailboxExportRequest  | Remove-MailboxExportRequest

-   ###### Error \#2

    New-MailboxExportRequest : The term 'New-MailboxExportRequest' is
    not recognized as the name of a cmdlet, function, script file, or
    operable program. Check the spelling of the name, or if a path was
    included, verify that the path is correct and try again. At line:1
    char:39 + ... oreach (\$Mailbox in (Get-Mailbox)) {
    New-MailboxExportRequest -Mailbo ... +
    \~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~\~ + CategoryInfo :
    ObjectNotFound: (New-MailboxExportRequest:String) [],
    CommandNotFoundException + FullyQualifiedErrorId :
    CommandNotFoundException

    ###### Solution

    Restart the EWS console and make sure you run it with the
    administrator rights.

-   ###### Error \#3

    \\\\ServerName\\Temp\\PST\\user.pst The server or share name
    specified in the path may be invalid, or the file could be locked. +
    CategoryInfo : NotSpecified: (:) [New-MailboxExportRequest],
    RemotePermanentException + FullyQualifiedErrorId :
    [Server=ServerName,RequestId=40efdbf2-53bf-47e8-92b8-f8fa8a8f15db,TimeStamp=11/7/2017
    9:55:48 AM] [FailureCategory=Cmdlet-RemotePermanentException]
    94719DAF,
    Microsoft.Exchange.Management.Migration.MailboxReplication.MailboxExportRequest.NewMailboxExportRequest
    + PSComputerName : servername.mailbox.local

    ###### Solution

    The `\\ServerName\Temp` folder (the one used for the mailbox export
    and set with the script `-dir` parameter) must be shared for reading
    and writing access. If you shared it previously check whether it has
    the writing access only and change the access correctly.

Import the MS Exchange data to ONLYOFFICE
-----------------------------------------

##### Installing Enterprise Edition and configure Mail Server

Install the Docker version of **Enterprise Edition**. This can be done
installing **Enterprise Edition** [using the
script](/server/docker/enterprise/enterprise-script-installation.aspx)
and selecting the Docker installation variant. After that set up the
mail server as written
[here](/gettingstarted/mail.aspx#MailServer_block).

When connecting the domain, you will need to set the same domain as has
been used for email messaging with MS Exchange (the domain from the user
email addresses). If you need to change the domain name, you need to
additionally edit the `users.csv` file (which was received in the step
above), replacing all the entries for the old MS Exchange domain with
the new one. This is done with the following command:

    sed -i 's/exchange-domain.com/new-domain.com/g' users.csv

Where `exchange-domain.com` is the old domain name used with MS
Exchange, and `new-domain.com` is the new one you are now going to use.

##### Downloading script and installing/updating dependencies

Now you need to download and unpack the script which will do the data
import process. This can be done using the command:

    wget -O "ImportExchangeData.tar" "https://bit.ly/2jdOn8t" && tar -xvf ImportExchangeData.tar && cd ./Import

The command will download and unpack the file, creating the following
folder structure:

    Import
       |-lib
       |---create_users.py
       |---mbox2imap.py
       |---mapping.json
       |---pst2mbox.sh
       |---get-mailboxes.sh
       |---install-passfinder.sh
       |---ASC.Mail.PasswordFinder.tar
       |-ImportExchangeData.py
       |-requirements.txt

You will need [Python
v2.7](https://en.wikipedia.org/wiki/Python_(programming_language))
installed. It is often installed by default with various Linux
distributions, but in case it is missing, you will need to install it
yourself. This is how it is done for Debian-based distributions:

    # apt install python
    # python -V
    Python 2.7.12

Install the [pip (package
manager)](https://en.wikipedia.org/wiki/Pip_(package_manager)), also
necessary for the correct script work:

    # apt install python-pip
    # pip -V
    pip 9.0.1 from /usr/local/lib/python2.7/dist-packages (python 2.7)

And install the other required packages:

    pip install -r requirements.txt

##### Running the script

Now you can run the script specifying the necessary parameters:

    ./ImportExchangeData.py -d "<portal domain>" -u "<portal administrator email>" -pw "<portal administrator password>" -f <path to the folder with the exported data>

If your portal is connected using HTTPS, you will need additionally use
the `-s "https"` parameter when running the code.

Replace the parameters in brackets with your own portal data and run the
script:

    ./ImportExchangeData.py -d myportal.com -u "my.email.address@gmail.com" -pw "123456" -f /root/Temp/

Wait for the script to finish working. It might take some time depending
on the number of users and the amount of their data.

##### Script work results

When the script does everything it is intended for, the results will be
the following:

-   the new portal users with the email adresses from the `users.csv`
    file will be created;
-   the mailboxes at the ONLYOFFICE Mail Server will be created, which
    will have the mail messages from MS Exchange and will be connected
    in the ONLYOFFICE **Mail** module for the users listed in the
    `users.csv` file;
-   `results-ok.txt` file will be saved to the `Import` folder, it will
    contain the list of all user accounts from the `users.csv` file and
    their passwords, which were successfully created;
-   `result-err.txt` file will be saved to the `Import` folder, it will
    contain the list of all user accounts from the `users.csv` file,
    which had problems when they were imported and created;
-   `mailboxes.json` file in the `JSON` format will be saved to the
    `Import` folder, it will contain the list of settings necessary to
    connect to the newly created mailboxes from the third-party mail
    clients.

If you create some mailboxes after the import, you can also get the
settings necessary to connect them to the third-party mail client
programs. Go to the `ImportExchangeData.py` folder and run the command:

    bash ./lib/get-mailboxes.sh -j

The `mailboxes.json` file will be overwritten with the new mailboxes
data.

Show what the ImportExchangeData.py script does Hide

1.  When run, the script sets the main working files and folders: both
    for the files and folders necessary for the output files and the
    folder with the input data (which has been exported from MS
    Exchange).
2.  At the next step the script creates the users and then their
    mailboxes.
3.  Then the script converts the exported files into the `mbox` format.
4.  When done, the script installs the `ASC.Mail.PasswordFinder` program
    which will assist in creating the `mailboxes.json` file containing
    the list of settings necessary to connect to the newly created
    mailboxes from the third-party mail clients.
5.  After that the `mailboxes.json` itself is created.
6.  And finally the `mbox` format files are imported to the mailboxes
    via IMAP.

[Download](https://www.onlyoffice.com/download-enterprise.aspx?from=helpcenter)
Host on your own server Available for Docker, \
Windows, Linux and virtual machines

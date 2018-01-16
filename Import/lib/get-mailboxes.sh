#!/bin/bash

OUTPATH=$1

if [[ -z $OUTPATH ]]; then
    OUTPATH = "."
fi

DOCKERCONTAINER="onlyoffice-community-server"
BASEPATH="/var/www/onlyoffice/Services/MailPasswordFinder"

docker exec -it ${DOCKERCONTAINER} mono ${BASEPATH}/ASC.Mail.PasswordFinder.exe -je
docker cp ${DOCKERCONTAINER}:/${BASEPATH}/mailboxes.json ${OUTPATH}/mailboxes.json

echo "File mailboxes.json has been copied to local directory"

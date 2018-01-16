#!/bin/bash

BASE_PATH=$1

if [[ -z $BASE_PATH ]]; then
    BASE_PATH = "."
fi

COMMUNITY="onlyoffice-community-server"
SERVICES_PATH="/var/www/onlyoffice/Services"
AGG_PATH="${SERVICES_PATH}/MailAggregator"
APP_PATH="${SERVICES_PATH}/MailPasswordFinder"
APP_NAME="ASC.Mail.PasswordFinder"
APP="${APP_NAME}.exe"
APP_TAR="${APP_NAME}.tar"

docker exec -it "${COMMUNITY}" rm -rf "${APP_PATH}"
docker exec -it "${COMMUNITY}" rm -f "${SERVICES_PATH}/${APP_TAR}"

echo "${BASE_PATH}/${APP_TAR}"

docker cp "${BASE_PATH}/${APP_TAR}" "${COMMUNITY}:${SERVICES_PATH}"
docker exec -it "${COMMUNITY}" tar -xvf "${SERVICES_PATH}/${APP_TAR}" -C "${SERVICES_PATH}/"
docker exec -it "${COMMUNITY}" chmod +x "${APP_PATH}/${APP}"
docker exec -it "${COMMUNITY}" cp "${AGG_PATH}/ASC.Mail.Aggregator.CollectionService.exe.config" "${APP_PATH}/ASC.Mail.PasswordFinder.exe.config"
docker exec -it "${COMMUNITY}" cp "${AGG_PATH}/web.consumers.config" "${APP_PATH}/"
docker exec -it "${COMMUNITY}" cp "${AGG_PATH}/web.storage.config" "${APP_PATH}/"
docker exec -it "${COMMUNITY}" rm -f "${APP_PATH}/Mono.Security.dll"

#!/bin/bash
SOURCE="$1"
function command_exists () {
    type "$1" &> /dev/null;
}
function install(){
    if command_exists apt-get; then
	    sudo apt-get update
	    sudo apt-get -y install pst-utils readpst
        #sudo apt-get -y install readpst
        #sudo apt-get -y install convmv 
        #sudo apt-get -y install libencode-imaputf7-perl
    elif command_exists yum ; then
        sudo yum update
        sudo yum -y install libpst 
        #sudo yum -y install convmv 
        #sudo yum -y install perl-Encode-IMAPUTF7F
    fi
}

install
if ! command_exists readpst ; then
 echo "ERROR: cannot find redapst command"
 return -1
fi
if [[ ! -d $SOURCE ]] ; then
  echo "ERROR: cannot find the source folder ${SOURCE}. The process has been stopped"
  exit -1
fi

#Create the repaired  convmv version
#sudo sed -e "s/^use utf8;$/use utf8; use Encode::IMAPUTF7;/g" </usr/bin/convmv > convmv | chmod +x convmv

#Convert PST file to mbox
if [[ -z $SOURCE ]]; then
    FILES="*.pst"
else
    FILES=$(printf "%s/*.pst" $SOURCE)
fi
if [[ ! -d "./mbox" ]] ; then
    sudo mkdir mbox
fi
for f in $FILES
do
echo "MESSAGE: the ${f} is being processed... " 
sudo readpst -r -w -q -o ./mbox $f
#sudo readpst -r -w -q -o mbox $f -t e
done

#Correct the names of directories (folders)
#if [[ -f "./convmv" ]]; then
# ./convmv --notest -f utf8 -t IMAP-UTF-7 -r mbox/ # | find mbox/ -mindepth 1 -type d -name '*.*' -exec rename "s/\./_/g" {} \+
#else
# echo "ERROR: cannot find the file 'convmv' "
# exit -1
#fi

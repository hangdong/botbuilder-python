#!/bin/bash
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

# This Script provisions and deploys a Python bot with retries.

# Make errors stop the script (set -e) 
set -e
# Uncomment to debug:  print commands (set -x)
#set -x

# Define Environment Variables
WEBAPP_NAME="pyfuntest"
WEBAPP_URL=
BOT_NAME="pyfuntest"
BOT_ID="pyfuntest"
AZURE_RESOURCE_GROUP=
BOT_APPID=
BOT_PASSWORD=

usage()
{
    echo "${0##*/} [options]"
    echo ""
    echo "Runs Python DirectLine bot test."
    echo "- Deletes and recreates given Azure Resource Group (cleanup)"
    echo "- Provision and deploy python bot"
    echo "- Provision DirectLine support for bot"
    echo "- Run python directline client against deployed bot using DirectLine"
    echo "  as a test."
    echo ""
    echo "Note: Assumes you are logged into Azure."
    echo ""
    echo "options"
    echo " -a, --appid           Bot App ID"
    echo " -p, --password        Bot App Password"
    echo " -g, --resource-group  Azure Resource Group name"
    exit 1;
}

print_help_and_exit()
{
    echo "Run '${0##*/} --help' for more information."
    exit 1
}

process_args()
{
    if [ "${PWD##*/}" != 'functionaltestbot' ]; then
        echo "ERROR: Must run from '/functional-tests/functionaltestbot' directory."
        echo "Your current directory: ${PWD##*/}"
        echo ""
        echo "For example:"
        echo "$ ./scripts/deploy_webapp.sh --appid X --password Y -g Z"
        exit 1
    fi

    save_next_arg=0
    for arg in "$@"
    do
        if [[ ${save_next_arg} -eq 1 ]]; then
            BOT_APPID="$arg"
            save_next_arg=0
        elif [[ ${save_next_arg} -eq 2 ]]; then
            BOT_PASSWORD="$arg"
            save_next_arg=0
        elif [[ ${save_next_arg} -eq 3 ]]; then
            AZURE_RESOURCE_GROUP="$arg"
            save_next_arg=0
        else
            case "$arg" in
                "-h" | "--help" ) usage;;
                "-a" | "--appid" ) save_next_arg=1;;
                "-p" | "--password" ) save_next_arg=2;;
                "-g" | "--resource-group" ) save_next_arg=3;;
                * ) usage;;
            esac
        fi
    done
    if [[ -z ${BOT_APPID} ]]; then
        echo "Bot appid parameter invalid"
        print_help_and_exit
    fi    
    if [[ -z ${BOT_PASSWORD} ]]; then
        echo "Bot password parameter invalid"
        print_help_and_exit
    fi    
    if [[ -z ${AZURE_RESOURCE_GROUP} ]]; then
        echo "Azure Resource Group parameter invalid"
        print_help_and_exit
    fi    
}

###############################################################################
# Main Script Execution
###############################################################################
process_args "$@"

# Recreate Resource Group

# It's ok to fail (set +e) - script continues on error result code.
echo "Deleting Resource Group ${AZURE_RESOURCE_GROUP}.."
set +e
az group delete --name ${AZURE_RESOURCE_GROUP} -y

echo "Creating Resource Group ${AZURE_RESOURCE_GROUP}.."
n=0
until [ $n -ge 3 ]
do
   az group create --location westus --name ${AZURE_RESOURCE_GROUP} && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Could not create group ${AZURE_RESOURCE_GROUP}"
    exit 3
fi

# Push Web App
echo "Push web app.."
n=0
until [ $n -ge 3 ]
do
   az webapp up --sku P1V2 -n ${WEBAPP_NAME} -l westus --resource-group ${AZURE_RESOURCE_GROUP} && break  
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Could not create webapp ${WEBAPP_NAME}"
    exit 4
fi

# Verify it's running
echo "Verify web app is up.."
n=0
until [ $n -ge 3 ]
do
   az webapp show --name pyfuntest --resource-group pyfuntest | python3 -c "import sys, json, io; input_stream = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8'); exit(0 if json.load(input_stream)['state'] == 'Running' else 1)" && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Webapp did not get to running state ${WEBAPP_NAME}"
    exit 4
fi

echo "Create bot.."
n=0
until [ $n -ge 3 ]
do
   az bot create --appid ${BOT_APPID} --name ${BOT_NAME} --password ${BOT_PASSWORD} --resource-group ${AZURE_RESOURCE_GROUP} --sku F0 --kind registration --location westus --endpoint "https://${WEBAPP_NAME}.azurewebsites.net/api/messages" && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Could not create BOT ${BOT_NAME}"
    exit 5
fi


# Create bot settings
echo "Create bot settings.."
n=0
until [ $n -ge 3 ]
do
   az webapp config appsettings set -g ${AZURE_RESOURCE_GROUP} -n ${AZURE_RESOURCE_GROUP} --settings MicrosoftAppId=${BOT_APPID} MicrosoftAppPassword=${BOT_PASSWORD} botId=${BOT_ID} && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Could not create BOT configuration"
    exit 6
fi

# Create DirectLine
echo "Create DirectLine.."
cd tests
n=0
until [ $n -ge 3 ]
do
   az bot directline create --name ${BOT_NAME} --resource-group ${AZURE_RESOURCE_GROUP} > "DirectLineConfig.json" && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Could not create Directline configuration"
    exit 7
fi


# Run Tests
sleep 30
echo "Running tests.."
pip install requests
n=0
until [ $n -ge 3 ]
do
   python test_py_bot.py && break
   n=$[$n+1]
   sleep 25
done
if [[ $n -ge 3 ]]; then
    echo "Tests failed!"
    exit 8
fi


